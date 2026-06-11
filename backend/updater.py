# -*- coding: utf-8 -*-
"""
版本发布与自动更新模块
为打包 exe 提供检查远程版本号、下载增量包与校验更新的能力，
减少财务终端手工替换 exe 的维护成本。

功能：
  1. 远程版本检查（对比本地与远程 manifest）
  2. 增量包下载（支持断点续传、进度回调）
  3. SHA-256 完整性校验
  4. 原子更新流程（备份 → 解压 → 校验 → 回滚机制）
  5. 自动重启

远程 manifest.json 格式示例：
{
  "latest_version": "1.2.0",
  "min_supported_version": "1.0.0",
  "release_notes": "修复余额连续性校验问题",
  "packages": [
    {
      "version": "1.2.0",
      "url": "https://updates.example.com/bankcheck/1.2.0/patch.zip",
      "sha256": "a1b2c3d4e5f6...",
      "size": 5242880,
      "type": "incremental",
      "from_version": "1.1.0"
    }
  ]
}
"""

import hashlib
import json
import logging
import os
import shutil
import subprocess
import sys
import tempfile
import threading
import zipfile
from dataclasses import dataclass, field
from datetime import datetime
from typing import Callable, Dict, List, Optional, Tuple
from urllib.request import urlopen, Request, urlretrieve
from urllib.error import URLError, HTTPError

try:
    from i18n import t, set_language, get_language, get_available_languages, init_i18n
    HAS_I18N = True
except ImportError:
    HAS_I18N = False

    def t(key, **kwargs):
        return key

    def set_language(lang):
        return False

    def get_language():
        return 'zh_CN'

    def get_available_languages():
        return {'zh_CN': '简体中文'}

    def init_i18n(lang=None):
        return None


APP_VERSION = "1.0.0"

MANIFEST_FILENAME = "manifest.json"

BACKUP_DIR_NAME = "update_backup"

UPDATE_DIR_NAME = "update_temp"

logger = logging.getLogger('bankcheck.updater')


def setup_logging(log_dir: Optional[str] = None):
    if log_dir is None:
        log_dir = os.path.dirname(os.path.abspath(__file__))
    log_file = os.path.join(log_dir, 'updater.log')

    _logger = logging.getLogger('bankcheck.updater')
    _logger.setLevel(logging.INFO)

    if not _logger.handlers:
        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setLevel(logging.INFO)

        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)

        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S',
        )
        file_handler.setFormatter(formatter)
        console_handler.setFormatter(formatter)

        _logger.addHandler(file_handler)
        _logger.addHandler(console_handler)

    return _logger


@dataclass
class PackageInfo:
    version: str
    url: str
    sha256: str
    size: int
    type: str = "incremental"
    from_version: Optional[str] = None

    def to_dict(self) -> Dict:
        d = {
            'version': self.version,
            'url': self.url,
            'sha256': self.sha256,
            'size': self.size,
            'type': self.type,
        }
        if self.from_version is not None:
            d['from_version'] = self.from_version
        return d

    @classmethod
    def from_dict(cls, data: Dict) -> 'PackageInfo':
        return cls(
            version=data['version'],
            url=data['url'],
            sha256=data['sha256'],
            size=data['size'],
            type=data.get('type', 'incremental'),
            from_version=data.get('from_version'),
        )


@dataclass
class Manifest:
    latest_version: str
    min_supported_version: str
    release_notes: str = ""
    packages: List[PackageInfo] = field(default_factory=list)
    fetched_at: Optional[str] = None

    def to_dict(self) -> Dict:
        return {
            'latest_version': self.latest_version,
            'min_supported_version': self.min_supported_version,
            'release_notes': self.release_notes,
            'packages': [p.to_dict() for p in self.packages],
            'fetched_at': self.fetched_at,
        }

    @classmethod
    def from_dict(cls, data: Dict) -> 'Manifest':
        packages = [PackageInfo.from_dict(p) for p in data.get('packages', [])]
        return cls(
            latest_version=data['latest_version'],
            min_supported_version=data['min_supported_version'],
            release_notes=data.get('release_notes', ''),
            packages=packages,
            fetched_at=data.get('fetched_at'),
        )


@dataclass
class UpdateResult:
    success: bool
    message: str
    old_version: str = ""
    new_version: str = ""
    backup_path: str = ""
    rollback_available: bool = False


def parse_version(version_str: str) -> Tuple[int, ...]:
    parts = []
    for p in version_str.strip().split('.'):
        try:
            parts.append(int(p))
        except ValueError:
            parts.append(0)
    return tuple(parts)


def compare_versions(v1: str, v2: str) -> int:
    p1 = parse_version(v1)
    p2 = parse_version(v2)
    max_len = max(len(p1), len(p2))
    p1 = p1 + (0,) * (max_len - len(p1))
    p2 = p2 + (0,) * (max_len - len(p2))
    if p1 < p2:
        return -1
    elif p1 > p2:
        return 1
    return 0


def get_app_version() -> str:
    return APP_VERSION


def get_app_dir() -> str:
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def fetch_manifest(manifest_url: str, timeout: float = 30.0) -> Manifest:
    logger.info("Fetching manifest from: %s", manifest_url)
    try:
        req = Request(manifest_url, headers={'User-Agent': f'BankCheck/{APP_VERSION}'})
        with urlopen(req, timeout=timeout) as resp:
            data = json.loads(resp.read().decode('utf-8'))
        manifest = Manifest.from_dict(data)
        manifest.fetched_at = datetime.now().isoformat()
        logger.info("Manifest fetched: latest=%s, min=%s",
                     manifest.latest_version, manifest.min_supported_version)
        return manifest
    except (URLError, HTTPError) as e:
        logger.error("Failed to fetch manifest: %s", e)
        raise
    except (json.JSONDecodeError, KeyError) as e:
        logger.error("Invalid manifest format: %s", e)
        raise


def check_for_update(manifest_url: str, current_version: Optional[str] = None,
                     timeout: float = 30.0) -> Tuple[bool, Optional[Manifest]]:
    if current_version is None:
        current_version = get_app_version()

    try:
        manifest = fetch_manifest(manifest_url, timeout=timeout)
    except Exception as e:
        logger.error("Update check failed: %s", e)
        return False, None

    cmp = compare_versions(current_version, manifest.latest_version)
    if cmp < 0:
        force_needed = compare_versions(current_version, manifest.min_supported_version) < 0
        logger.info("Update available: %s -> %s (force=%s)",
                     current_version, manifest.latest_version, force_needed)
        return True, manifest

    logger.info("Already up to date: %s", current_version)
    return False, manifest


def find_update_package(manifest: Manifest, current_version: str) -> Optional[PackageInfo]:
    exact_match = None
    full_fallback = None
    version_fallback = None
    for pkg in manifest.packages:
        if compare_versions(pkg.version, current_version) <= 0:
            continue
        if pkg.from_version is not None and compare_versions(pkg.from_version, current_version) == 0:
            if exact_match is None or compare_versions(pkg.version, exact_match.version) > 0:
                exact_match = pkg
        elif pkg.from_version is None:
            if full_fallback is None or compare_versions(pkg.version, full_fallback.version) > 0:
                full_fallback = pkg
        else:
            if version_fallback is None or compare_versions(pkg.version, version_fallback.version) > 0:
                version_fallback = pkg
    if exact_match is not None:
        return exact_match
    if full_fallback is not None and compare_versions(current_version, manifest.latest_version) < 0:
        return full_fallback
    if version_fallback is not None and compare_versions(current_version, manifest.latest_version) < 0:
        return version_fallback
    return None


def compute_sha256(file_path: str, chunk_size: int = 8192) -> str:
    h = hashlib.sha256()
    with open(file_path, 'rb') as f:
        while True:
            chunk = f.read(chunk_size)
            if not chunk:
                break
            h.update(chunk)
    return h.hexdigest()


def verify_file_integrity(file_path: str, expected_sha256: str) -> bool:
    if not os.path.isfile(file_path):
        logger.error("File not found for integrity check: %s", file_path)
        return False
    actual = compute_sha256(file_path)
    match = actual.lower() == expected_sha256.lower()
    if not match:
        logger.error("Integrity check failed: expected=%s actual=%s", expected_sha256, actual)
    else:
        logger.info("Integrity check passed: %s", file_path)
    return match


def download_package(package: PackageInfo, dest_dir: str,
                     progress_callback: Optional[Callable[[int, int], None]] = None,
                     chunk_size: int = 8192, resume: bool = True) -> str:
    os.makedirs(dest_dir, exist_ok=True)
    filename = os.path.basename(package.url.split('?')[0])
    if not filename.endswith('.zip'):
        filename = f"update_{package.version}.zip"
    dest_path = os.path.join(dest_dir, filename)

    temp_path = dest_path + '.downloading'
    downloaded = 0

    if resume and os.path.exists(temp_path):
        existing_size = os.path.getsize(temp_path)
        if existing_size < package.size:
            downloaded = existing_size
            logger.info("Resuming download from %d bytes", downloaded)
        else:
            downloaded = 0

    try:
        headers = {'User-Agent': f'BankCheck/{APP_VERSION}'}
        if downloaded > 0:
            headers['Range'] = f'bytes={downloaded}-'

        req = Request(package.url, headers=headers)
        with urlopen(req) as resp:
            mode = 'ab' if downloaded > 0 else 'wb'
            with open(temp_path, mode) as f:
                while True:
                    chunk = resp.read(chunk_size)
                    if not chunk:
                        break
                    f.write(chunk)
                    downloaded += len(chunk)
                    if progress_callback:
                        try:
                            progress_callback(downloaded, package.size)
                        except Exception:
                            pass

        shutil.move(temp_path, dest_path)
        logger.info("Download complete: %s (%d bytes)", dest_path, downloaded)
        return dest_path

    except Exception:
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except OSError:
                pass
        raise


def extract_package(zip_path: str, dest_dir: str) -> str:
    os.makedirs(dest_dir, exist_ok=True)
    logger.info("Extracting %s to %s", zip_path, dest_dir)
    try:
        with zipfile.ZipFile(zip_path, 'r') as zf:
            bad = zf.testzip()
            if bad is not None:
                raise zipfile.BadZipFile(f"Corrupt entry: {bad}")
            zf.extractall(dest_dir)
    except zipfile.BadZipFile as e:
        logger.error("Bad zip file: %s", e)
        raise
    logger.info("Extraction complete: %s", dest_dir)
    return dest_dir


def backup_current_install(app_dir: str, backup_base: Optional[str] = None) -> str:
    if backup_base is None:
        backup_base = os.path.join(app_dir, BACKUP_DIR_NAME)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    backup_dir = os.path.join(backup_base, f"backup_{timestamp}")
    os.makedirs(backup_dir, exist_ok=True)

    for item in os.listdir(app_dir):
        if item in (BACKUP_DIR_NAME, UPDATE_DIR_NAME, '__pycache__', '.git'):
            continue
        src = os.path.join(app_dir, item)
        dst = os.path.join(backup_dir, item)
        try:
            if os.path.isfile(src):
                shutil.copy2(src, dst)
            elif os.path.isdir(src):
                shutil.copytree(src, dst, dirs_exist_ok=True)
        except (OSError, shutil.Error) as e:
            logger.warning("Backup skip %s: %s", item, e)

    logger.info("Backup created: %s", backup_dir)
    return backup_dir


def apply_update(extracted_dir: str, app_dir: str, version: str) -> bool:
    logger.info("Applying update from %s to app dir %s", extracted_dir, app_dir)
    try:
        for item in os.listdir(extracted_dir):
            src = os.path.join(extracted_dir, item)
            dst = os.path.join(app_dir, item)
            if os.path.isfile(src):
                shutil.copy2(src, dst)
            elif os.path.isdir(src):
                if os.path.isdir(dst):
                    shutil.rmtree(dst)
                shutil.copytree(src, dst)
        logger.info("Update applied: version %s", version)
        return True
    except Exception as e:
        logger.error("Apply update failed: %s", e)
        return False


def rollback(backup_dir: str, app_dir: str) -> bool:
    logger.info("Rolling back from %s", backup_dir)
    try:
        for item in os.listdir(backup_dir):
            src = os.path.join(backup_dir, item)
            dst = os.path.join(app_dir, item)
            if os.path.isfile(src):
                shutil.copy2(src, dst)
            elif os.path.isdir(src):
                if os.path.isdir(dst):
                    shutil.rmtree(dst)
                shutil.copytree(src, dst)
        logger.info("Rollback complete")
        return True
    except Exception as e:
        logger.error("Rollback failed: %s", e)
        return False


def cleanup_old_backups(backup_base: str, keep: int = 3):
    if not os.path.isdir(backup_base):
        return
    dirs = sorted(
        [d for d in os.listdir(backup_base) if d.startswith('backup_')],
        reverse=True,
    )
    for old in dirs[keep:]:
        old_path = os.path.join(backup_base, old)
        try:
            shutil.rmtree(old_path)
            logger.info("Removed old backup: %s", old_path)
        except OSError as e:
            logger.warning("Could not remove backup %s: %s", old_path, e)


def restart_app():
    logger.info("Restarting application...")
    if getattr(sys, 'frozen', False):
        exe = sys.executable
        subprocess.Popen([exe])
    else:
        subprocess.Popen([sys.executable] + sys.argv)
    sys.exit(0)


class Updater:
    def __init__(self, manifest_url: str, app_dir: Optional[str] = None,
                 current_version: Optional[str] = None):
        self.manifest_url = manifest_url
        self.app_dir = app_dir or get_app_dir()
        self.current_version = current_version or get_app_version()
        self._cancel_event = threading.Event()
        self._manifest: Optional[Manifest] = None

    def check_for_update(self, timeout: float = 30.0) -> Tuple[bool, Optional[Manifest]]:
        available, manifest = check_for_update(
            self.manifest_url, self.current_version, timeout=timeout
        )
        self._manifest = manifest
        return available, manifest

    def find_package(self) -> Optional[PackageInfo]:
        if self._manifest is None:
            return None
        return find_update_package(self._manifest, self.current_version)

    def download_and_verify(
        self,
        package: PackageInfo,
        progress_callback: Optional[Callable[[int, int], None]] = None,
        resume: bool = True,
    ) -> Tuple[bool, str]:
        update_dir = os.path.join(self.app_dir, UPDATE_DIR_NAME)
        download_dir = os.path.join(update_dir, 'download')
        os.makedirs(download_dir, exist_ok=True)

        try:
            zip_path = download_package(
                package, download_dir,
                progress_callback=progress_callback,
                resume=resume,
            )
        except Exception as e:
            logger.error("Download failed: %s", e)
            return False, str(e)

        if self._cancel_event.is_set():
            logger.info("Update cancelled after download")
            return False, "Cancelled"

        if not verify_file_integrity(zip_path, package.sha256):
            try:
                os.remove(zip_path)
            except OSError:
                pass
            return False, "Integrity check failed"

        extract_dir = os.path.join(update_dir, 'extracted')
        if os.path.isdir(extract_dir):
            shutil.rmtree(extract_dir)

        try:
            extract_package(zip_path, extract_dir)
        except Exception as e:
            logger.error("Extraction failed: %s", e)
            return False, str(e)

        return True, extract_dir

    def perform_update(
        self,
        package: PackageInfo,
        extracted_dir: str,
        auto_restart: bool = False,
        keep_backups: int = 3,
    ) -> UpdateResult:
        backup_base = os.path.join(self.app_dir, BACKUP_DIR_NAME)
        try:
            backup_dir = backup_current_install(self.app_dir, backup_base)
        except Exception as e:
            logger.error("Backup failed: %s", e)
            return UpdateResult(
                success=False,
                message=f"Backup failed: {e}",
                old_version=self.current_version,
                new_version=package.version,
                rollback_available=False,
            )

        success = apply_update(extracted_dir, self.app_dir, package.version)

        if success:
            cleanup_old_backups(backup_base, keep=keep_backups)
            result = UpdateResult(
                success=True,
                message=f"Updated {self.current_version} -> {package.version}",
                old_version=self.current_version,
                new_version=package.version,
                backup_path=backup_dir,
                rollback_available=True,
            )
            logger.info("Update successful: %s -> %s", self.current_version, package.version)
            if auto_restart:
                restart_app()
            return result

        logger.warning("Update apply failed, rolling back...")
        rollback_ok = rollback(backup_dir, self.app_dir)
        return UpdateResult(
            success=False,
            message="Update apply failed, rollback " + ("succeeded" if rollback_ok else "FAILED"),
            old_version=self.current_version,
            new_version=package.version,
            backup_path=backup_dir,
            rollback_available=rollback_ok,
        )

    def run_full_update(
        self,
        progress_callback: Optional[Callable[[int, int], None]] = None,
        auto_restart: bool = False,
    ) -> UpdateResult:
        available, manifest = self.check_for_update()
        if not available or manifest is None:
            return UpdateResult(
                success=False,
                message="No update available",
                old_version=self.current_version,
            )

        package = self.find_package()
        if package is None:
            return UpdateResult(
                success=False,
                message="No compatible package found",
                old_version=self.current_version,
            )

        ok, extract_or_msg = self.download_and_verify(
            package, progress_callback=progress_callback,
        )
        if not ok:
            return UpdateResult(
                success=False,
                message=f"Download/verify failed: {extract_or_msg}",
                old_version=self.current_version,
                new_version=package.version,
            )

        return self.perform_update(package, extract_or_msg, auto_restart=auto_restart)

    def cancel(self):
        self._cancel_event.set()
        logger.info("Update cancellation requested")

    def get_rollback_info(self) -> Optional[Dict]:
        backup_base = os.path.join(self.app_dir, BACKUP_DIR_NAME)
        if not os.path.isdir(backup_base):
            return None
        dirs = sorted(
            [d for d in os.listdir(backup_base) if d.startswith('backup_')],
            reverse=True,
        )
        if not dirs:
            return None
        latest = dirs[0]
        backup_dir = os.path.join(backup_base, latest)
        return {
            'backup_dir': backup_dir,
            'timestamp': latest.replace('backup_', ''),
            'available': True,
        }

    def perform_rollback(self) -> UpdateResult:
        info = self.get_rollback_info()
        if info is None or not info.get('available'):
            return UpdateResult(
                success=False,
                message="No rollback backup available",
                old_version=self.current_version,
            )
        ok = rollback(info['backup_dir'], self.app_dir)
        if ok:
            return UpdateResult(
                success=True,
                message="Rollback succeeded",
                rollback_available=True,
                backup_path=info['backup_dir'],
            )
        return UpdateResult(
            success=False,
            message="Rollback failed",
            old_version=self.current_version,
        )
