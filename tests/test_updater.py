# -*- coding: utf-8 -*-
"""
版本发布与自动更新模块 - 单元测试
"""

import hashlib
import json
import os
import shutil
import sys
import tempfile
import threading
import zipfile
from io import BytesIO
from unittest.mock import MagicMock, Mock, patch, call

import pytest

sys_path_backend = os.path.join(os.path.dirname(__file__), '..', 'backend')
if sys_path_backend not in os.sys.path:
    os.sys.path.insert(0, sys_path_backend)

import updater


SAMPLE_MANIFEST = {
    "latest_version": "1.2.0",
    "min_supported_version": "1.0.0",
    "release_notes": "修复余额连续性校验问题",
    "packages": [
        {
            "version": "1.1.0",
            "url": "https://updates.example.com/bankcheck/1.1.0/patch.zip",
            "sha256": "aaa111bbb222",
            "size": 1024,
            "type": "incremental",
            "from_version": "1.0.0",
        },
        {
            "version": "1.2.0",
            "url": "https://updates.example.com/bankcheck/1.2.0/patch.zip",
            "sha256": "ccc333ddd444",
            "size": 2048,
            "type": "incremental",
            "from_version": "1.1.0",
        },
        {
            "version": "1.2.0",
            "url": "https://updates.example.com/bankcheck/1.2.0/full.zip",
            "sha256": "eee555fff666",
            "size": 10240,
            "type": "full",
        },
    ],
}


def _make_zip_bytes(filenames_and_content):
    buf = BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        for name, content in filenames_and_content:
            zf.writestr(name, content)
    return buf.getvalue()


def _make_zip_file(path, filenames_and_content):
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for name, content in filenames_and_content:
            zf.writestr(name, content)


def _sha256_of_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


@pytest.fixture
def tmp_dir():
    d = tempfile.mkdtemp(prefix='updater_test_')
    yield d
    shutil.rmtree(d, ignore_errors=True)


@pytest.fixture
def sample_manifest():
    return updater.Manifest.from_dict(SAMPLE_MANIFEST)


class TestParseVersion:
    def test_basic(self):
        assert updater.parse_version("1.2.3") == (1, 2, 3)

    def test_single(self):
        assert updater.parse_version("5") == (5,)

    def test_two_parts(self):
        assert updater.parse_version("1.0") == (1, 0)

    def test_with_spaces(self):
        assert updater.parse_version(" 1.2.3 ") == (1, 2, 3)

    def test_non_numeric_part(self):
        assert updater.parse_version("1.a.3") == (1, 0, 3)

    def test_empty(self):
        assert updater.parse_version("") == (0,)


class TestCompareVersions:
    def test_less(self):
        assert updater.compare_versions("1.0.0", "1.1.0") == -1

    def test_greater(self):
        assert updater.compare_versions("2.0.0", "1.9.9") == 1

    def test_equal(self):
        assert updater.compare_versions("1.0.0", "1.0.0") == 0

    def test_different_length(self):
        assert updater.compare_versions("1.0", "1.0.0") == 0

    def test_major_difference(self):
        assert updater.compare_versions("1.5.0", "2.0.0") == -1


class TestPackageInfo:
    def test_from_dict(self):
        data = {
            "version": "1.2.0",
            "url": "https://example.com/patch.zip",
            "sha256": "abc123",
            "size": 1024,
            "type": "incremental",
            "from_version": "1.1.0",
        }
        pkg = updater.PackageInfo.from_dict(data)
        assert pkg.version == "1.2.0"
        assert pkg.url == "https://example.com/patch.zip"
        assert pkg.sha256 == "abc123"
        assert pkg.size == 1024
        assert pkg.type == "incremental"
        assert pkg.from_version == "1.1.0"

    def test_from_dict_defaults(self):
        data = {
            "version": "1.0.0",
            "url": "https://example.com/full.zip",
            "sha256": "def456",
            "size": 2048,
        }
        pkg = updater.PackageInfo.from_dict(data)
        assert pkg.type == "incremental"
        assert pkg.from_version is None

    def test_to_dict(self):
        pkg = updater.PackageInfo(
            version="1.2.0",
            url="https://example.com/patch.zip",
            sha256="abc123",
            size=1024,
            type="incremental",
            from_version="1.1.0",
        )
        d = pkg.to_dict()
        assert d['version'] == "1.2.0"
        assert d['from_version'] == "1.1.0"

    def test_to_dict_no_from_version(self):
        pkg = updater.PackageInfo(
            version="1.0.0",
            url="https://example.com/full.zip",
            sha256="def456",
            size=2048,
            type="full",
        )
        d = pkg.to_dict()
        assert 'from_version' not in d

    def test_roundtrip(self):
        data = {
            "version": "2.0.0",
            "url": "https://example.com/v2.zip",
            "sha256": "xyz789",
            "size": 4096,
            "type": "full",
        }
        pkg = updater.PackageInfo.from_dict(data)
        d = pkg.to_dict()
        pkg2 = updater.PackageInfo.from_dict(d)
        assert pkg2.version == pkg.version
        assert pkg2.url == pkg.url
        assert pkg2.sha256 == pkg.sha256


class TestManifest:
    def test_from_dict(self, sample_manifest):
        assert sample_manifest.latest_version == "1.2.0"
        assert sample_manifest.min_supported_version == "1.0.0"
        assert len(sample_manifest.packages) == 3
        assert sample_manifest.release_notes == "修复余额连续性校验问题"

    def test_to_dict(self, sample_manifest):
        d = sample_manifest.to_dict()
        assert d['latest_version'] == "1.2.0"
        assert len(d['packages']) == 3
        assert 'fetched_at' in d

    def test_roundtrip(self):
        data = {
            "latest_version": "3.0.0",
            "min_supported_version": "2.0.0",
            "release_notes": "Major release",
            "packages": [
                {
                    "version": "3.0.0",
                    "url": "https://example.com/v3.zip",
                    "sha256": "aaa",
                    "size": 8192,
                    "type": "full",
                }
            ],
        }
        m = updater.Manifest.from_dict(data)
        d = m.to_dict()
        m2 = updater.Manifest.from_dict(d)
        assert m2.latest_version == "3.0.0"
        assert len(m2.packages) == 1


class TestComputeSha256:
    def test_correct_hash(self, tmp_dir):
        content = b"hello world"
        fpath = os.path.join(tmp_dir, "test.txt")
        with open(fpath, 'wb') as f:
            f.write(content)
        expected = hashlib.sha256(content).hexdigest()
        assert updater.compute_sha256(fpath) == expected

    def test_empty_file(self, tmp_dir):
        fpath = os.path.join(tmp_dir, "empty.txt")
        with open(fpath, 'wb') as f:
            pass
        expected = hashlib.sha256(b"").hexdigest()
        assert updater.compute_sha256(fpath) == expected


class TestVerifyFileIntegrity:
    def test_valid(self, tmp_dir):
        content = b"test content for integrity"
        fpath = os.path.join(tmp_dir, "valid.bin")
        with open(fpath, 'wb') as f:
            f.write(content)
        sha = hashlib.sha256(content).hexdigest()
        assert updater.verify_file_integrity(fpath, sha) is True

    def test_invalid(self, tmp_dir):
        content = b"original"
        fpath = os.path.join(tmp_dir, "tampered.bin")
        with open(fpath, 'wb') as f:
            f.write(content)
        wrong_sha = "0" * 64
        assert updater.verify_file_integrity(fpath, wrong_sha) is False

    def test_case_insensitive(self, tmp_dir):
        content = b"case test"
        fpath = os.path.join(tmp_dir, "case.bin")
        with open(fpath, 'wb') as f:
            f.write(content)
        sha = hashlib.sha256(content).hexdigest().upper()
        assert updater.verify_file_integrity(fpath, sha) is True

    def test_missing_file(self):
        assert updater.verify_file_integrity("/nonexistent/file.zip", "abc") is False


class TestFetchManifest:
    @patch('updater.urlopen')
    def test_success(self, mock_urlopen):
        mock_resp = MagicMock()
        mock_resp.read.return_value = json.dumps(SAMPLE_MANIFEST).encode('utf-8')
        mock_resp.__enter__ = Mock(return_value=mock_resp)
        mock_resp.__exit__ = Mock(return_value=False)
        mock_urlopen.return_value = mock_resp

        manifest = updater.fetch_manifest("https://example.com/manifest.json")
        assert manifest.latest_version == "1.2.0"
        assert len(manifest.packages) == 3
        assert manifest.fetched_at is not None

    @patch('updater.urlopen')
    def test_network_error(self, mock_urlopen):
        from urllib.error import URLError
        mock_urlopen.side_effect = URLError("Connection refused")
        with pytest.raises(URLError):
            updater.fetch_manifest("https://unreachable.example.com/manifest.json")

    @patch('updater.urlopen')
    def test_invalid_json(self, mock_urlopen):
        mock_resp = MagicMock()
        mock_resp.read.return_value = b"not json"
        mock_resp.__enter__ = Mock(return_value=mock_resp)
        mock_resp.__exit__ = Mock(return_value=False)
        mock_urlopen.return_value = mock_resp
        with pytest.raises(json.JSONDecodeError):
            updater.fetch_manifest("https://example.com/manifest.json")

    @patch('updater.urlopen')
    def test_missing_required_field(self, mock_urlopen):
        bad_manifest = {"latest_version": "1.0.0"}
        mock_resp = MagicMock()
        mock_resp.read.return_value = json.dumps(bad_manifest).encode('utf-8')
        mock_resp.__enter__ = Mock(return_value=mock_resp)
        mock_resp.__exit__ = Mock(return_value=False)
        mock_urlopen.return_value = mock_resp
        with pytest.raises(KeyError):
            updater.fetch_manifest("https://example.com/manifest.json")

    @patch('updater.urlopen')
    def test_user_agent_header(self, mock_urlopen):
        mock_resp = MagicMock()
        mock_resp.read.return_value = json.dumps(SAMPLE_MANIFEST).encode('utf-8')
        mock_resp.__enter__ = Mock(return_value=mock_resp)
        mock_resp.__exit__ = Mock(return_value=False)
        mock_urlopen.return_value = mock_resp

        updater.fetch_manifest("https://example.com/manifest.json")

        call_args = mock_urlopen.call_args
        req = call_args[0][0]
        assert f'BankCheck/{updater.APP_VERSION}' in req.get_header('User-agent')


class TestCheckForUpdate:
    @patch('updater.fetch_manifest')
    def test_update_available(self, mock_fetch):
        mock_fetch.return_value = updater.Manifest.from_dict(SAMPLE_MANIFEST)
        available, manifest = updater.check_for_update(
            "https://example.com/manifest.json", current_version="1.0.0"
        )
        assert available is True
        assert manifest is not None
        assert manifest.latest_version == "1.2.0"

    @patch('updater.fetch_manifest')
    def test_already_up_to_date(self, mock_fetch):
        mock_fetch.return_value = updater.Manifest.from_dict(SAMPLE_MANIFEST)
        available, _ = updater.check_for_update(
            "https://example.com/manifest.json", current_version="1.2.0"
        )
        assert available is False

    @patch('updater.fetch_manifest')
    def test_newer_than_remote(self, mock_fetch):
        mock_fetch.return_value = updater.Manifest.from_dict(SAMPLE_MANIFEST)
        available, _ = updater.check_for_update(
            "https://example.com/manifest.json", current_version="2.0.0"
        )
        assert available is False

    @patch('updater.fetch_manifest')
    def test_fetch_failure(self, mock_fetch):
        mock_fetch.side_effect = Exception("Network error")
        available, manifest = updater.check_for_update(
            "https://example.com/manifest.json"
        )
        assert available is False
        assert manifest is None


class TestFindUpdatePackage:
    def test_incremental_match(self, sample_manifest):
        pkg = updater.find_update_package(sample_manifest, "1.1.0")
        assert pkg is not None
        assert pkg.version == "1.2.0"
        assert pkg.from_version == "1.1.0"

    def test_full_fallback(self, sample_manifest):
        pkg = updater.find_update_package(sample_manifest, "1.0.5")
        assert pkg is not None
        assert pkg.version == "1.2.0"
        assert pkg.type == "full"

    def test_no_update_needed(self, sample_manifest):
        pkg = updater.find_update_package(sample_manifest, "1.2.0")
        assert pkg is None

    def test_empty_packages(self):
        m = updater.Manifest(
            latest_version="2.0.0",
            min_supported_version="1.0.0",
            packages=[],
        )
        pkg = updater.find_update_package(m, "1.0.0")
        assert pkg is None


class TestDownloadPackage:
    @patch('updater.urlopen')
    def test_basic_download(self, mock_urlopen, tmp_dir):
        zip_data = _make_zip_bytes([("file.txt", "content")])
        mock_resp = MagicMock()
        mock_resp.read.side_effect = [zip_data, b""]
        mock_resp.__enter__ = Mock(return_value=mock_resp)
        mock_resp.__exit__ = Mock(return_value=False)
        mock_urlopen.return_value = mock_resp

        pkg = updater.PackageInfo(
            version="1.1.0",
            url="https://example.com/patch.zip",
            sha256="abc",
            size=len(zip_data),
        )
        result = updater.download_package(pkg, tmp_dir)
        assert os.path.isfile(result)
        assert result.endswith('.zip')

    @patch('updater.urlopen')
    def test_progress_callback(self, mock_urlopen, tmp_dir):
        chunk1 = b"A" * 100
        chunk2 = b"B" * 100
        mock_resp = MagicMock()
        mock_resp.read.side_effect = [chunk1, chunk2, b""]
        mock_resp.__enter__ = Mock(return_value=mock_resp)
        mock_resp.__exit__ = Mock(return_value=False)
        mock_urlopen.return_value = mock_resp

        progress_calls = []
        def on_progress(downloaded, total):
            progress_calls.append((downloaded, total))

        pkg = updater.PackageInfo(
            version="1.1.0",
            url="https://example.com/patch.zip",
            sha256="abc",
            size=200,
        )
        updater.download_package(pkg, tmp_dir, progress_callback=on_progress)
        assert len(progress_calls) >= 2
        assert progress_calls[0][0] == 100
        assert progress_calls[1][0] == 200

    @patch('updater.urlopen')
    def test_network_error_cleanup(self, mock_urlopen, tmp_dir):
        from urllib.error import URLError
        mock_urlopen.side_effect = URLError("Connection lost")
        pkg = updater.PackageInfo(
            version="1.1.0",
            url="https://example.com/patch.zip",
            sha256="abc",
            size=1024,
        )
        with pytest.raises(URLError):
            updater.download_package(pkg, tmp_dir)
        temp_file = os.path.join(tmp_dir, "patch.zip.downloading")
        assert not os.path.exists(temp_file)

    @patch('updater.urlopen')
    def test_resume_download(self, mock_urlopen, tmp_dir):
        existing_data = b"A" * 100
        temp_path = os.path.join(tmp_dir, "patch.zip.downloading")
        with open(temp_path, 'wb') as f:
            f.write(existing_data)

        new_data = b"B" * 100
        mock_resp = MagicMock()
        mock_resp.read.side_effect = [new_data, b""]
        mock_resp.__enter__ = Mock(return_value=mock_resp)
        mock_resp.__exit__ = Mock(return_value=False)
        mock_urlopen.return_value = mock_resp

        pkg = updater.PackageInfo(
            version="1.1.0",
            url="https://example.com/patch.zip",
            sha256="abc",
            size=200,
        )
        result = updater.download_package(pkg, tmp_dir, resume=True)
        assert os.path.isfile(result)
        with open(result, 'rb') as f:
            content = f.read()
        assert content == existing_data + new_data


class TestExtractPackage:
    def test_valid_zip(self, tmp_dir):
        zip_path = os.path.join(tmp_dir, "update.zip")
        _make_zip_file(zip_path, [
            ("new_file.txt", "new content"),
            ("subdir/nested.txt", "nested content"),
        ])
        extract_dir = os.path.join(tmp_dir, "extracted")
        result = updater.extract_package(zip_path, extract_dir)
        assert os.path.isfile(os.path.join(result, "new_file.txt"))
        assert os.path.isfile(os.path.join(result, "subdir", "nested.txt"))

    def test_corrupt_zip(self, tmp_dir):
        bad_zip = os.path.join(tmp_dir, "bad.zip")
        with open(bad_zip, 'wb') as f:
            f.write(b"not a zip file")
        extract_dir = os.path.join(tmp_dir, "extracted")
        with pytest.raises(zipfile.BadZipFile):
            updater.extract_package(bad_zip, extract_dir)

    def test_existing_extract_dir_cleanup(self, tmp_dir):
        zip_path = os.path.join(tmp_dir, "update.zip")
        _make_zip_file(zip_path, [("file.txt", "v2")])

        extract_dir = os.path.join(tmp_dir, "extracted")
        os.makedirs(extract_dir, exist_ok=True)
        old_file = os.path.join(extract_dir, "old.txt")
        with open(old_file, 'w') as f:
            f.write("old")

        updater.extract_package(zip_path, extract_dir)
        assert os.path.isfile(os.path.join(extract_dir, "file.txt"))


class TestBackupCurrentInstall:
    def test_backup_files(self, tmp_dir):
        app_dir = os.path.join(tmp_dir, "app")
        os.makedirs(app_dir)
        with open(os.path.join(app_dir, "bankcheck.exe"), 'w') as f:
            f.write("exe content")
        with open(os.path.join(app_dir, "config.json"), 'w') as f:
            f.write("{}")

        backup_dir = updater.backup_current_install(app_dir)
        assert os.path.isdir(backup_dir)
        assert os.path.isfile(os.path.join(backup_dir, "bankcheck.exe"))
        assert os.path.isfile(os.path.join(backup_dir, "config.json"))

    def test_backup_skips_special_dirs(self, tmp_dir):
        app_dir = os.path.join(tmp_dir, "app")
        os.makedirs(os.path.join(app_dir, "update_backup"))
        os.makedirs(os.path.join(app_dir, "update_temp"))
        os.makedirs(os.path.join(app_dir, "__pycache__"))
        with open(os.path.join(app_dir, "main.exe"), 'w') as f:
            f.write("main")

        backup_dir = updater.backup_current_install(app_dir)
        items = os.listdir(backup_dir)
        assert "update_backup" not in items
        assert "update_temp" not in items
        assert "__pycache__" not in items
        assert "main.exe" in items

    def test_backup_subdirectories(self, tmp_dir):
        app_dir = os.path.join(tmp_dir, "app")
        os.makedirs(os.path.join(app_dir, "data", "nested"), exist_ok=True)
        with open(os.path.join(app_dir, "data", "nested", "file.txt"), 'w') as f:
            f.write("nested")

        backup_dir = updater.backup_current_install(app_dir)
        assert os.path.isfile(os.path.join(backup_dir, "data", "nested", "file.txt"))


class TestApplyUpdate:
    def test_apply_files(self, tmp_dir):
        app_dir = os.path.join(tmp_dir, "app")
        os.makedirs(app_dir)
        with open(os.path.join(app_dir, "old.txt"), 'w') as f:
            f.write("old")

        extract_dir = os.path.join(tmp_dir, "extracted")
        os.makedirs(extract_dir)
        with open(os.path.join(extract_dir, "new.txt"), 'w') as f:
            f.write("new")
        with open(os.path.join(extract_dir, "old.txt"), 'w') as f:
            f.write("updated")

        result = updater.apply_update(extract_dir, app_dir, "1.1.0")
        assert result is True
        with open(os.path.join(app_dir, "new.txt")) as f:
            assert f.read() == "new"
        with open(os.path.join(app_dir, "old.txt")) as f:
            assert f.read() == "updated"

    def test_apply_with_subdirectories(self, tmp_dir):
        app_dir = os.path.join(tmp_dir, "app")
        os.makedirs(app_dir)
        extract_dir = os.path.join(tmp_dir, "extracted")
        os.makedirs(os.path.join(extract_dir, "lib"), exist_ok=True)
        with open(os.path.join(extract_dir, "lib", "core.dll"), 'w') as f:
            f.write("dll content")

        result = updater.apply_update(extract_dir, app_dir, "1.1.0")
        assert result is True
        assert os.path.isfile(os.path.join(app_dir, "lib", "core.dll"))

    def test_apply_failure_rolls_back_on_caller(self, tmp_dir):
        app_dir = os.path.join(tmp_dir, "app")
        os.makedirs(app_dir)
        extract_dir = os.path.join(tmp_dir, "extracted")
        os.makedirs(extract_dir)
        with open(os.path.join(extract_dir, "file.txt"), 'w') as f:
            f.write("data")

        with patch('updater.shutil.copy2', side_effect=PermissionError("access denied")):
            result = updater.apply_update(extract_dir, app_dir, "1.1.0")
        assert result is False


class TestRollback:
    def test_rollback_success(self, tmp_dir):
        app_dir = os.path.join(tmp_dir, "app")
        backup_dir = os.path.join(tmp_dir, "backup")

        os.makedirs(app_dir)
        with open(os.path.join(app_dir, "current.txt"), 'w') as f:
            f.write("current")

        os.makedirs(backup_dir)
        with open(os.path.join(backup_dir, "original.txt"), 'w') as f:
            f.write("original")

        result = updater.rollback(backup_dir, app_dir)
        assert result is True
        assert os.path.isfile(os.path.join(app_dir, "original.txt"))

    def test_rollback_failure(self, tmp_dir):
        result = updater.rollback("/nonexistent/backup", "/nonexistent/app")
        assert result is False


class TestCleanupOldBackups:
    def test_cleanup_keeps_latest(self, tmp_dir):
        backup_base = os.path.join(tmp_dir, "backups")
        for i in range(5):
            d = os.path.join(backup_base, f"backup_20240101_{i:06d}")
            os.makedirs(d)
            with open(os.path.join(d, "file.txt"), 'w') as f:
                f.write(str(i))

        updater.cleanup_old_backups(backup_base, keep=3)
        remaining = [d for d in os.listdir(backup_base) if d.startswith('backup_')]
        assert len(remaining) == 3

    def test_cleanup_no_dir(self):
        updater.cleanup_old_backups("/nonexistent/path", keep=3)

    def test_cleanup_fewer_than_keep(self, tmp_dir):
        backup_base = os.path.join(tmp_dir, "backups")
        os.makedirs(os.path.join(backup_base, "backup_20240101_000000"))

        updater.cleanup_old_backups(backup_base, keep=3)
        remaining = [d for d in os.listdir(backup_base) if d.startswith('backup_')]
        assert len(remaining) == 1


class TestUpdater:
    @patch('updater.fetch_manifest')
    def test_check_for_update_available(self, mock_fetch):
        mock_fetch.return_value = updater.Manifest.from_dict(SAMPLE_MANIFEST)
        u = updater.Updater("https://example.com/manifest.json", current_version="1.0.0")
        available, manifest = u.check_for_update()
        assert available is True
        assert manifest.latest_version == "1.2.0"

    @patch('updater.fetch_manifest')
    def test_check_for_update_not_needed(self, mock_fetch):
        mock_fetch.return_value = updater.Manifest.from_dict(SAMPLE_MANIFEST)
        u = updater.Updater("https://example.com/manifest.json", current_version="1.2.0")
        available, _ = u.check_for_update()
        assert available is False

    @patch('updater.fetch_manifest')
    def test_find_package(self, mock_fetch):
        mock_fetch.return_value = updater.Manifest.from_dict(SAMPLE_MANIFEST)
        u = updater.Updater("https://example.com/manifest.json", current_version="1.1.0")
        u.check_for_update()
        pkg = u.find_package()
        assert pkg is not None
        assert pkg.version == "1.2.0"

    def test_find_package_no_manifest(self):
        u = updater.Updater("https://example.com/manifest.json", current_version="1.0.0")
        assert u.find_package() is None

    @patch('updater.verify_file_integrity', return_value=True)
    @patch('updater.extract_package')
    @patch('updater.download_package')
    def test_download_and_verify_success(self, mock_download, mock_extract, mock_verify, tmp_dir):
        zip_path = os.path.join(tmp_dir, "downloaded.zip")
        mock_download.return_value = zip_path
        extract_dir = os.path.join(tmp_dir, "extracted")
        mock_extract.return_value = extract_dir

        u = updater.Updater("https://example.com/manifest.json",
                           app_dir=tmp_dir, current_version="1.0.0")
        pkg = updater.PackageInfo(
            version="1.1.0",
            url="https://example.com/patch.zip",
            sha256="abc123",
            size=1024,
        )
        ok, result = u.download_and_verify(pkg)
        assert ok is True

    @patch('updater.verify_file_integrity', return_value=False)
    @patch('updater.download_package')
    def test_download_and_verify_integrity_fail(self, mock_download, mock_verify, tmp_dir):
        zip_path = os.path.join(tmp_dir, "downloaded.zip")
        with open(zip_path, 'w') as f:
            f.write("data")
        mock_download.return_value = zip_path

        u = updater.Updater("https://example.com/manifest.json",
                           app_dir=tmp_dir, current_version="1.0.0")
        pkg = updater.PackageInfo(
            version="1.1.0",
            url="https://example.com/patch.zip",
            sha256="wrong",
            size=1024,
        )
        ok, msg = u.download_and_verify(pkg)
        assert ok is False
        assert "Integrity" in msg

    @patch('updater.download_package')
    def test_download_and_verify_download_fail(self, mock_download, tmp_dir):
        from urllib.error import URLError
        mock_download.side_effect = URLError("Connection lost")

        u = updater.Updater("https://example.com/manifest.json",
                           app_dir=tmp_dir, current_version="1.0.0")
        pkg = updater.PackageInfo(
            version="1.1.0",
            url="https://example.com/patch.zip",
            sha256="abc",
            size=1024,
        )
        ok, msg = u.download_and_verify(pkg)
        assert ok is False

    @patch('updater.backup_current_install')
    @patch('updater.apply_update')
    def test_perform_update_success(self, mock_apply, mock_backup, tmp_dir):
        backup_dir = os.path.join(tmp_dir, "backup")
        mock_backup.return_value = backup_dir
        mock_apply.return_value = True

        u = updater.Updater("https://example.com/manifest.json",
                           app_dir=tmp_dir, current_version="1.0.0")
        pkg = updater.PackageInfo(
            version="1.1.0",
            url="https://example.com/patch.zip",
            sha256="abc",
            size=1024,
        )
        result = u.perform_update(pkg, "/tmp/extracted")
        assert result.success is True
        assert result.old_version == "1.0.0"
        assert result.new_version == "1.1.0"
        assert result.rollback_available is True

    @patch('updater.backup_current_install')
    @patch('updater.apply_update')
    @patch('updater.rollback')
    def test_perform_update_rollback_on_failure(self, mock_rollback, mock_apply, mock_backup, tmp_dir):
        backup_dir = os.path.join(tmp_dir, "backup")
        mock_backup.return_value = backup_dir
        mock_apply.return_value = False
        mock_rollback.return_value = True

        u = updater.Updater("https://example.com/manifest.json",
                           app_dir=tmp_dir, current_version="1.0.0")
        pkg = updater.PackageInfo(
            version="1.1.0",
            url="https://example.com/patch.zip",
            sha256="abc",
            size=1024,
        )
        result = u.perform_update(pkg, "/tmp/extracted")
        assert result.success is False
        assert "rollback" in result.message.lower()

    @patch('updater.backup_current_install')
    def test_perform_update_backup_failure(self, mock_backup, tmp_dir):
        mock_backup.side_effect = OSError("Disk full")

        u = updater.Updater("https://example.com/manifest.json",
                           app_dir=tmp_dir, current_version="1.0.0")
        pkg = updater.PackageInfo(
            version="1.1.0",
            url="https://example.com/patch.zip",
            sha256="abc",
            size=1024,
        )
        result = u.perform_update(pkg, "/tmp/extracted")
        assert result.success is False
        assert "Backup failed" in result.message

    def test_cancel(self):
        u = updater.Updater("https://example.com/manifest.json")
        u.cancel()
        assert u._cancel_event.is_set()

    @patch('updater.verify_file_integrity', return_value=True)
    @patch('updater.extract_package')
    @patch('updater.download_package')
    def test_cancel_after_download(self, mock_download, mock_extract, mock_verify, tmp_dir):
        zip_path = os.path.join(tmp_dir, "downloaded.zip")
        mock_download.return_value = zip_path

        u = updater.Updater("https://example.com/manifest.json",
                           app_dir=tmp_dir, current_version="1.0.0")
        u.cancel()

        pkg = updater.PackageInfo(
            version="1.1.0",
            url="https://example.com/patch.zip",
            sha256="abc",
            size=1024,
        )
        ok, msg = u.download_and_verify(pkg)
        assert ok is False
        assert "Cancel" in msg

    def test_get_rollback_info_none(self, tmp_dir):
        u = updater.Updater("https://example.com/manifest.json", app_dir=tmp_dir)
        info = u.get_rollback_info()
        assert info is None

    def test_get_rollback_info_available(self, tmp_dir):
        backup_base = os.path.join(tmp_dir, "update_backup")
        backup_dir = os.path.join(backup_base, "backup_20240611_120000")
        os.makedirs(backup_dir)

        u = updater.Updater("https://example.com/manifest.json", app_dir=tmp_dir)
        info = u.get_rollback_info()
        assert info is not None
        assert info['available'] is True

    @patch('updater.rollback')
    def test_perform_rollback_success(self, mock_rollback, tmp_dir):
        backup_base = os.path.join(tmp_dir, "update_backup")
        backup_dir = os.path.join(backup_base, "backup_20240611_120000")
        os.makedirs(backup_dir)
        mock_rollback.return_value = True

        u = updater.Updater("https://example.com/manifest.json", app_dir=tmp_dir)
        result = u.perform_rollback()
        assert result.success is True

    @patch('updater.rollback')
    def test_perform_rollback_failure(self, mock_rollback, tmp_dir):
        backup_base = os.path.join(tmp_dir, "update_backup")
        backup_dir = os.path.join(backup_base, "backup_20240611_120000")
        os.makedirs(backup_dir)
        mock_rollback.return_value = False

        u = updater.Updater("https://example.com/manifest.json", app_dir=tmp_dir)
        result = u.perform_rollback()
        assert result.success is False


class TestRunFullUpdate:
    @patch('updater.Updater.perform_update')
    @patch('updater.Updater.download_and_verify')
    @patch('updater.Updater.find_package')
    @patch('updater.Updater.check_for_update')
    def test_full_update_success(self, mock_check, mock_find, mock_download, mock_perform):
        manifest = updater.Manifest.from_dict(SAMPLE_MANIFEST)
        mock_check.return_value = (True, manifest)
        pkg = updater.PackageInfo(
            version="1.1.0",
            url="https://example.com/patch.zip",
            sha256="abc",
            size=1024,
        )
        mock_find.return_value = pkg
        mock_download.return_value = (True, "/tmp/extracted")
        mock_perform.return_value = updater.UpdateResult(
            success=True,
            message="Updated 1.0.0 -> 1.1.0",
            old_version="1.0.0",
            new_version="1.1.0",
        )

        u = updater.Updater("https://example.com/manifest.json", current_version="1.0.0")
        result = u.run_full_update()
        assert result.success is True

    @patch('updater.Updater.check_for_update')
    def test_full_update_no_update_available(self, mock_check):
        mock_check.return_value = (False, None)
        u = updater.Updater("https://example.com/manifest.json", current_version="1.2.0")
        result = u.run_full_update()
        assert result.success is False
        assert "No update available" in result.message

    @patch('updater.Updater.find_package')
    @patch('updater.Updater.check_for_update')
    def test_full_update_no_compatible_package(self, mock_check, mock_find):
        manifest = updater.Manifest.from_dict(SAMPLE_MANIFEST)
        mock_check.return_value = (True, manifest)
        mock_find.return_value = None

        u = updater.Updater("https://example.com/manifest.json", current_version="1.0.0")
        result = u.run_full_update()
        assert result.success is False
        assert "No compatible package" in result.message

    @patch('updater.Updater.download_and_verify')
    @patch('updater.Updater.find_package')
    @patch('updater.Updater.check_for_update')
    def test_full_update_download_fail(self, mock_check, mock_find, mock_download):
        manifest = updater.Manifest.from_dict(SAMPLE_MANIFEST)
        mock_check.return_value = (True, manifest)
        pkg = updater.PackageInfo(
            version="1.1.0",
            url="https://example.com/patch.zip",
            sha256="abc",
            size=1024,
        )
        mock_find.return_value = pkg
        mock_download.return_value = (False, "Network error")

        u = updater.Updater("https://example.com/manifest.json", current_version="1.0.0")
        result = u.run_full_update()
        assert result.success is False
        assert "Download/verify failed" in result.message


class TestGetAppDir:
    def test_non_frozen(self):
        with patch.object(updater.sys, 'frozen', False, create=True):
            result = updater.get_app_dir()
            assert os.path.isabs(result)

    def test_frozen(self):
        with patch.object(updater.sys, 'frozen', True, create=True):
            with patch.object(updater.sys, 'executable', '/opt/bankcheck/bankcheck.exe'):
                result = updater.get_app_dir()
                assert result == '/opt/bankcheck'


class TestGetAppVersion:
    def test_returns_string(self):
        version = updater.get_app_version()
        assert isinstance(version, str)
        assert len(version) > 0


class TestUpdateResult:
    def test_success_result(self):
        r = updater.UpdateResult(
            success=True,
            message="OK",
            old_version="1.0.0",
            new_version="1.1.0",
            backup_path="/tmp/backup",
            rollback_available=True,
        )
        assert r.success is True
        assert r.new_version == "1.1.0"

    def test_failure_result(self):
        r = updater.UpdateResult(
            success=False,
            message="Failed",
        )
        assert r.success is False
        assert r.rollback_available is False


class TestEndToEndFlow:
    @patch('updater.restart_app')
    @patch('updater.urlopen')
    def test_e2e_update_flow(self, mock_urlopen, mock_restart, tmp_dir):
        app_dir = os.path.join(tmp_dir, "app")
        os.makedirs(app_dir)
        with open(os.path.join(app_dir, "bankcheck.exe"), 'w') as f:
            f.write("v1.0.0")
        with open(os.path.join(app_dir, "config.json"), 'w') as f:
            json.dump({"version": "1.0.0"}, f)

        zip_data = _make_zip_bytes([("bankcheck.exe", "v1.1.0")])
        zip_sha = _sha256_of_bytes(zip_data)

        manifest_data = {
            "latest_version": "1.1.0",
            "min_supported_version": "1.0.0",
            "release_notes": "Bug fix",
            "packages": [
                {
                    "version": "1.1.0",
                    "url": "https://updates.example.com/patch.zip",
                    "sha256": zip_sha,
                    "size": len(zip_data),
                    "type": "incremental",
                    "from_version": "1.0.0",
                }
            ],
        }

        mock_manifest_resp = MagicMock()
        mock_manifest_resp.read.return_value = json.dumps(manifest_data).encode('utf-8')
        mock_manifest_resp.__enter__ = Mock(return_value=mock_manifest_resp)
        mock_manifest_resp.__exit__ = Mock(return_value=False)

        mock_download_resp = MagicMock()
        mock_download_resp.read.side_effect = [zip_data, b""]
        mock_download_resp.__enter__ = Mock(return_value=mock_download_resp)
        mock_download_resp.__exit__ = Mock(return_value=False)

        mock_urlopen.side_effect = [mock_manifest_resp, mock_download_resp]

        u = updater.Updater("https://updates.example.com/manifest.json",
                           app_dir=app_dir, current_version="1.0.0")

        available, manifest = u.check_for_update()
        assert available is True

        pkg = u.find_package()
        assert pkg is not None

        ok, extract_dir = u.download_and_verify(pkg)
        assert ok is True

        result = u.perform_update(pkg, extract_dir)
        assert result.success is True
        assert result.old_version == "1.0.0"
        assert result.new_version == "1.1.0"

        with open(os.path.join(app_dir, "bankcheck.exe")) as f:
            assert f.read() == "v1.1.0"

        assert os.path.isdir(result.backup_path)
        with open(os.path.join(result.backup_path, "bankcheck.exe")) as f:
            assert f.read() == "v1.0.0"

    @patch('updater.restart_app')
    @patch('updater.urlopen')
    def test_e2e_rollback_flow(self, mock_urlopen, mock_restart, tmp_dir):
        app_dir = os.path.join(tmp_dir, "app")
        os.makedirs(app_dir)
        with open(os.path.join(app_dir, "bankcheck.exe"), 'w') as f:
            f.write("v1.0.0")

        zip_data = _make_zip_bytes([("bankcheck.exe", "v1.1.0")])
        zip_sha = _sha256_of_bytes(zip_data)

        manifest_data = {
            "latest_version": "1.1.0",
            "min_supported_version": "1.0.0",
            "packages": [
                {
                    "version": "1.1.0",
                    "url": "https://updates.example.com/patch.zip",
                    "sha256": zip_sha,
                    "size": len(zip_data),
                    "type": "incremental",
                    "from_version": "1.0.0",
                }
            ],
        }

        mock_manifest_resp = MagicMock()
        mock_manifest_resp.read.return_value = json.dumps(manifest_data).encode('utf-8')
        mock_manifest_resp.__enter__ = Mock(return_value=mock_manifest_resp)
        mock_manifest_resp.__exit__ = Mock(return_value=False)

        mock_download_resp = MagicMock()
        mock_download_resp.read.side_effect = [zip_data, b""]
        mock_download_resp.__enter__ = Mock(return_value=mock_download_resp)
        mock_download_resp.__exit__ = Mock(return_value=False)

        mock_urlopen.side_effect = [mock_manifest_resp, mock_download_resp]

        u = updater.Updater("https://updates.example.com/manifest.json",
                           app_dir=app_dir, current_version="1.0.0")
        available, _ = u.check_for_update()
        pkg = u.find_package()
        ok, extract_dir = u.download_and_verify(pkg)
        result = u.perform_update(pkg, extract_dir)
        assert result.success is True

        with open(os.path.join(app_dir, "bankcheck.exe")) as f:
            assert f.read() == "v1.1.0"

        rollback_result = u.perform_rollback()
        assert rollback_result.success is True

        with open(os.path.join(app_dir, "bankcheck.exe")) as f:
            assert f.read() == "v1.0.0"


class TestThreadSafety:
    def test_cancel_is_thread_safe(self):
        u = updater.Updater("https://example.com/manifest.json")
        t1 = threading.Thread(target=u.cancel)
        t2 = threading.Thread(target=u.cancel)
        t1.start()
        t2.start()
        t1.join()
        t2.join()
        assert u._cancel_event.is_set()
