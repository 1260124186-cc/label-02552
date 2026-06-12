# -*- coding: utf-8 -*-
"""
银企直连/网银导出目录对接模块
================================

功能：
1. 定义标准 inbox/outbox 目录结构约定
2. 监控 inbox 目录，自动发现银行流水文件
3. 检测文件稳定性和锁定状态，避免处理未完成下载的文件
4. 调用 bankcheck.run_pipeline 处理流水文件
5. 将处理结果归档到 outbox，错误文件移动到 error 目录
6. 支持定时轮询和单次运行两种模式
7. 与银行自动下载脚本衔接，形成端到端无人值守链路

目录约定：
    {root_dir}/
    ├── inbox/          # 待处理目录 - 银行下载脚本输出到此
    │   └── .lock/      # 锁定文件目录
    ├── outbox/         # 已完成目录 - 处理成功后归档
    │   └── {timestamp}/
    │       ├── 银行流水总表.xlsx
    │       ├── 流水检验报告.md
    │       └── original/
    ├── error/          # 错误目录 - 处理失败的文件
    └── processing/     # 处理中目录 - 临时

使用方式：
    # 单次运行处理当前 inbox 中的文件
    python bank_directory_connector.py run-once

    # 启动持续监控模式
    python bank_directory_connector.py watch

    # 作为模块调用
    from bank_directory_connector import BankDirectoryConnector
    connector = BankDirectoryConnector()
    connector.run_once()
"""

import os
import re
import sys
import time
import shutil
import logging
import threading
import subprocess
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from typing import List, Optional, Tuple, Dict, Any
from pathlib import Path

import yaml

BACKEND_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_CONFIG_PATH = os.path.join(BACKEND_DIR, 'bank_directories.yaml')

logger = logging.getLogger('bankcheck.directory_connector')


def _resolve_env_vars(value: Any) -> Any:
    """解析配置中的环境变量引用，如 ${VAR:-default}"""
    if isinstance(value, str):
        def replace_var(match):
            var_name = match.group(1)
            default_val = match.group(2) or ''
            return os.environ.get(var_name, default_val)

        pattern = r'\$\{([^}:]+)(?::-([^}]*))?\}'
        return re.sub(pattern, replace_var, value)
    elif isinstance(value, dict):
        return {k: _resolve_env_vars(v) for k, v in value.items()}
    elif isinstance(value, list):
        return [_resolve_env_vars(v) for v in value]
    return value


@dataclass
class DirectoryConfig:
    """目录配置"""
    root_dir: str
    poll_interval: int = 30
    file_stable_seconds: int = 5
    enable_lock_detection: bool = True
    lock_timeout_seconds: int = 3600
    archive_retention_days: int = 90
    processing_timeout_seconds: int = 1800


@dataclass
class BankDirectoryConfig:
    """单个银行的目录配置"""
    bank_name: str
    enabled: bool
    file_pattern: str
    directories: Dict[str, str] = field(default_factory=dict)
    download_script: Optional[Dict[str, Any]] = None


@dataclass
class ProcessingConfig:
    """处理策略配置"""
    incremental: bool = True
    keep_strategy: str = 'move_to_archive'
    archive_dir_name: str = 'original'
    generate_report: bool = True
    send_notification: bool = False
    notification_channels: List[str] = field(default_factory=lambda: ['console'])


@dataclass
class FileInfo:
    """待处理文件信息"""
    filepath: str
    bank_name: str
    size: int
    mtime: float
    is_stable: bool = False
    has_lock: bool = False
    matched_pattern: str = ''


@dataclass
class ProcessingResult:
    """目录处理结果"""
    success: bool
    processed_files: List[str] = field(default_factory=list)
    error_files: List[Tuple[str, str]] = field(default_factory=list)
    output_path: Optional[str] = None
    report_path: Optional[str] = None
    archive_dir: Optional[str] = None
    message: str = ''
    start_time: datetime = field(default_factory=datetime.now)
    end_time: Optional[datetime] = None


class BankDirectoryConnector:
    """银企直连目录对接器"""

    def __init__(self, config_path: Optional[str] = None, script_dir: Optional[str] = None):
        self.config_path = config_path or DEFAULT_CONFIG_PATH
        self.script_dir = script_dir or BACKEND_DIR
        self._config: Optional[Dict[str, Any]] = None
        self._directory_config: Optional[DirectoryConfig] = None
        self._bank_configs: Dict[str, BankDirectoryConfig] = {}
        self._processing_config: Optional[ProcessingConfig] = None
        self._stop_event = threading.Event()
        self._load_config()
        self._ensure_directories()

    def _load_config(self) -> None:
        """加载配置文件"""
        if not os.path.exists(self.config_path):
            raise FileNotFoundError(f"配置文件不存在: {self.config_path}")

        with open(self.config_path, 'r', encoding='utf-8') as f:
            raw_config = yaml.safe_load(f)

        self._config = _resolve_env_vars(raw_config)

        global_cfg = self._config.get('global', {})
        self._directory_config = DirectoryConfig(
            root_dir=os.path.abspath(global_cfg.get('root_dir', './bank_data')),
            poll_interval=global_cfg.get('poll_interval', 30),
            file_stable_seconds=global_cfg.get('file_stable_seconds', 5),
            enable_lock_detection=global_cfg.get('enable_lock_detection', True),
            lock_timeout_seconds=global_cfg.get('lock_timeout_seconds', 3600),
            archive_retention_days=global_cfg.get('archive_retention_days', 90),
            processing_timeout_seconds=global_cfg.get('processing_timeout_seconds', 1800),
        )

        for bank_cfg in self._config.get('banks', []):
            bank_name = bank_cfg['bank_name']
            self._bank_configs[bank_name] = BankDirectoryConfig(
                bank_name=bank_name,
                enabled=bank_cfg.get('enabled', True),
                file_pattern=bank_cfg.get('file_pattern', f'^{bank_name}_.*\\.xlsx$'),
                directories=bank_cfg.get('directories', {}),
                download_script=bank_cfg.get('download_script'),
            )

        proc_cfg = self._config.get('processing', {})
        self._processing_config = ProcessingConfig(
            incremental=proc_cfg.get('incremental', True),
            keep_strategy=proc_cfg.get('keep_strategy', 'move_to_archive'),
            archive_dir_name=proc_cfg.get('archive_dir_name', 'original'),
            generate_report=proc_cfg.get('generate_report', True),
            send_notification=proc_cfg.get('send_notification', False),
            notification_channels=proc_cfg.get('notification_channels', ['console']),
        )

        logger.info('配置加载完成，根目录: %s', self._directory_config.root_dir)
        logger.info('已配置银行: %s', ', '.join(self._bank_configs.keys()))

    def _ensure_directories(self) -> None:
        """确保所有必要的目录存在"""
        root = self._directory_config.root_dir
        directories = ['inbox', 'outbox', 'error', 'processing']

        for dir_name in directories:
            dir_path = os.path.join(root, dir_name)
            os.makedirs(dir_path, exist_ok=True)

        lock_dir = os.path.join(root, 'inbox', '.lock')
        os.makedirs(lock_dir, exist_ok=True)

        logger.info('目录结构已就绪: %s', root)

    def get_bank_dir(self, bank_name: str, dir_type: str) -> str:
        """获取指定银行的指定目录路径"""
        root = self._directory_config.root_dir
        bank_cfg = self._bank_configs.get(bank_name)

        if bank_cfg and bank_cfg.directories.get(dir_type):
            return os.path.join(root, bank_cfg.directories[dir_type])

        return os.path.join(root, dir_type)

    def _scan_inbox(self) -> List[FileInfo]:
        """扫描 inbox 目录，找出所有待处理的银行流水文件"""
        inbox_dir = os.path.join(self._directory_config.root_dir, 'inbox')
        files = []

        for filename in os.listdir(inbox_dir):
            filepath = os.path.join(inbox_dir, filename)

            if not os.path.isfile(filepath):
                continue

            if filename.startswith('.'):
                continue

            for bank_name, bank_cfg in self._bank_configs.items():
                if not bank_cfg.enabled:
                    continue

                if re.match(bank_cfg.file_pattern, filename):
                    stat = os.stat(filepath)
                    file_info = FileInfo(
                        filepath=filepath,
                        bank_name=bank_name,
                        size=stat.st_size,
                        mtime=stat.st_mtime,
                        matched_pattern=bank_cfg.file_pattern,
                    )
                    files.append(file_info)
                    break

        logger.info('扫描 inbox 发现 %d 个匹配文件', len(files))
        return files

    def _check_file_stable(self, file_info: FileInfo) -> bool:
        """检查文件是否稳定（大小不再变化）"""
        try:
            stat = os.stat(file_info.filepath)
            current_size = stat.st_size
            current_mtime = stat.st_mtime

            if current_size != file_info.size:
                return False

            age = time.time() - current_mtime
            return age >= self._directory_config.file_stable_seconds
        except OSError:
            return False

    def _check_lock_file(self, file_info: FileInfo) -> bool:
        """检查是否存在锁定文件"""
        if not self._directory_config.enable_lock_detection:
            return False

        inbox_dir = os.path.join(self._directory_config.root_dir, 'inbox')
        lock_dir = os.path.join(inbox_dir, '.lock')
        filename = os.path.basename(file_info.filepath)
        lock_file = os.path.join(lock_dir, f'{filename}.lock')

        if os.path.exists(lock_file):
            try:
                lock_mtime = os.stat(lock_file).st_mtime
                age = time.time() - lock_mtime
                if age > self._directory_config.lock_timeout_seconds:
                    logger.warning('锁定文件已超时，将忽略: %s', lock_file)
                    return False
            except OSError:
                pass
            return True

        return False

    def _validate_files(self, files: List[FileInfo]) -> List[FileInfo]:
        """验证文件，筛选出可处理的文件"""
        valid_files = []

        for file_info in files:
            file_info.has_lock = self._check_lock_file(file_info)
            if file_info.has_lock:
                logger.info('文件被锁定，跳过: %s', file_info.filepath)
                continue

            file_info.is_stable = self._check_file_stable(file_info)
            if not file_info.is_stable:
                logger.info('文件未稳定，跳过: %s', file_info.filepath)
                continue

            valid_files.append(file_info)

        logger.info('验证通过 %d 个文件，可进行处理', len(valid_files))
        return valid_files

    def _move_to_processing(self, file_info: FileInfo) -> str:
        """将文件移动到 processing 目录"""
        processing_dir = self.get_bank_dir(file_info.bank_name, 'processing')
        os.makedirs(processing_dir, exist_ok=True)

        filename = os.path.basename(file_info.filepath)
        dest_path = os.path.join(processing_dir, filename)

        shutil.move(file_info.filepath, dest_path)
        logger.info('已移动到处理目录: %s -> %s', file_info.filepath, dest_path)

        return dest_path

    def _move_to_error(self, filepath: str, error_msg: str) -> None:
        """将文件移动到 error 目录"""
        error_dir = os.path.join(self._directory_config.root_dir, 'error')
        os.makedirs(error_dir, exist_ok=True)

        filename = os.path.basename(filepath)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        dest_filename = f'{timestamp}_{filename}'
        dest_path = os.path.join(error_dir, dest_filename)

        shutil.move(filepath, dest_path)

        error_log_path = os.path.join(error_dir, f'{dest_filename}.error.txt')
        with open(error_log_path, 'w', encoding='utf-8') as f:
            f.write(f'处理时间: {datetime.now().isoformat()}\n')
            f.write(f'原文件: {filepath}\n')
            f.write(f'错误信息: {error_msg}\n')

        logger.error('已移动到错误目录: %s -> %s', filepath, dest_path)

    def _archive_results(self, processing_dir: str, pipeline_result: Any) -> str:
        """将处理结果归档到 outbox"""
        root = self._directory_config.root_dir
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        outbox_timestamp_dir = os.path.join(root, 'outbox', timestamp)
        os.makedirs(outbox_timestamp_dir, exist_ok=True)

        original_dir = os.path.join(outbox_timestamp_dir, self._processing_config.archive_dir_name)
        os.makedirs(original_dir, exist_ok=True)

        for filename in os.listdir(processing_dir):
            src_path = os.path.join(processing_dir, filename)
            if os.path.isfile(src_path):
                shutil.move(src_path, os.path.join(original_dir, filename))

        if hasattr(pipeline_result, 'output_path') and pipeline_result.output_path:
            output_filename = os.path.basename(pipeline_result.output_path)
            output_dest = os.path.join(outbox_timestamp_dir, output_filename)
            if os.path.exists(pipeline_result.output_path):
                shutil.copy2(pipeline_result.output_path, output_dest)

        if hasattr(pipeline_result, 'report_path') and pipeline_result.report_path:
            report_filename = os.path.basename(pipeline_result.report_path)
            report_dest = os.path.join(outbox_timestamp_dir, report_filename)
            if os.path.exists(pipeline_result.report_path):
                shutil.copy2(pipeline_result.report_path, report_dest)

        manifest_path = os.path.join(outbox_timestamp_dir, 'manifest.json')
        import json
        manifest = {
            'timestamp': timestamp,
            'processed_files': getattr(pipeline_result, 'processed_files', []),
            'error_files': getattr(pipeline_result, 'error_files', []),
            'total_records': len(getattr(pipeline_result, 'all_rows', [])),
            'pipeline_result': str(pipeline_result),
        }
        with open(manifest_path, 'w', encoding='utf-8') as f:
            json.dump(manifest, f, ensure_ascii=False, indent=2)

        logger.info('结果已归档到: %s', outbox_timestamp_dir)
        return outbox_timestamp_dir

    def _cleanup_old_archives(self) -> None:
        """清理过期的归档文件"""
        retention_days = self._directory_config.archive_retention_days
        if retention_days <= 0:
            return

        outbox_dir = os.path.join(self._directory_config.root_dir, 'outbox')
        cutoff_date = datetime.now() - timedelta(days=retention_days)

        for entry in os.listdir(outbox_dir):
            entry_path = os.path.join(outbox_dir, entry)
            if not os.path.isdir(entry_path):
                continue

            try:
                mtime = datetime.fromtimestamp(os.stat(entry_path).st_mtime)
                if mtime < cutoff_date:
                    shutil.rmtree(entry_path)
                    logger.info('已清理过期归档: %s', entry_path)
            except (OSError, ValueError):
                continue

    def run_once(self) -> ProcessingResult:
        """单次运行：处理当前 inbox 中的所有可用文件"""
        logger.info('========== 开始单次目录处理 ==========')
        result = ProcessingResult(success=False)

        try:
            files = self._scan_inbox()
            valid_files = self._validate_files(files)

            if not valid_files:
                result.success = True
                result.message = '没有可处理的文件'
                logger.info(result.message)
                return result

            bank_files: Dict[str, List[str]] = {}
            for file_info in valid_files:
                try:
                    processing_path = self._move_to_processing(file_info)
                    bank_files.setdefault(file_info.bank_name, []).append(processing_path)
                except Exception as e:
                    error_msg = f'移动文件失败: {e}'
                    logger.error(error_msg)
                    result.error_files.append((file_info.filepath, error_msg))

            if not bank_files:
                result.success = True
                result.message = '没有可处理的文件'
                return result

            all_processed = []
            all_errors = []
            pipeline_output = None

            for bank_name, processing_paths in bank_files.items():
                processing_dir = self.get_bank_dir(bank_name, 'processing')

                try:
                    import bankcheck
                    pipeline_result = bankcheck.run_pipeline_with_options(
                        processing_dir,
                        self.script_dir,
                        incremental=self._processing_config.incremental,
                        keep_strategy=self._processing_config.keep_strategy,
                    )

                    all_processed.extend(pipeline_result.processed_files)
                    all_errors.extend(pipeline_result.error_files)

                    if pipeline_result.output_path:
                        pipeline_output = pipeline_result.output_path

                    if getattr(pipeline_result, 'error_files', []):
                        for err_file, err_msg in pipeline_result.error_files:
                            self._move_to_error(err_file, err_msg)
                            all_errors.append((err_file, err_msg))

                    archive_dir = self._archive_results(processing_dir, pipeline_result)
                    result.archive_dir = archive_dir
                    result.output_path = pipeline_output

                    logger.info('银行 %s 处理完成', bank_name)

                except Exception as e:
                    error_msg = f'处理银行 {bank_name} 失败: {e}'
                    logger.exception(error_msg)
                    for path in processing_paths:
                        if os.path.exists(path):
                            self._move_to_error(path, error_msg)
                        all_errors.append((path, error_msg))

            result.processed_files = all_processed
            result.error_files = all_errors
            result.success = len(all_errors) == 0
            result.end_time = datetime.now()

            if result.success:
                result.message = f'成功处理 {len(all_processed)} 个文件'
            else:
                result.message = f'处理完成，{len(all_processed)} 个成功，{len(all_errors)} 个失败'

            self._cleanup_old_archives()

            logger.info('========== 目录处理完成 ==========')
            return result

        except Exception as e:
            logger.exception('目录处理发生异常')
            result.message = f'处理异常: {e}'
            result.success = False
            result.end_time = datetime.now()
            return result

    def watch(self, stop_on_first_empty: bool = False) -> None:
        """启动持续监控模式"""
        logger.info('========== 启动目录监控模式 ==========')
        logger.info('轮询间隔: %d 秒', self._directory_config.poll_interval)

        empty_count = 0

        while not self._stop_event.is_set():
            try:
                result = self.run_once()

                if stop_on_first_empty and not result.processed_files:
                    empty_count += 1
                    if empty_count >= 3:
                        logger.info('连续 3 次无文件可处理，退出监控')
                        break

                self._stop_event.wait(self._directory_config.poll_interval)

            except KeyboardInterrupt:
                logger.info('收到中断信号，正在退出...')
                break
            except Exception as e:
                logger.exception('监控循环发生异常: %s', e)
                self._stop_event.wait(self._directory_config.poll_interval)

        logger.info('========== 目录监控已停止 ==========')

    def stop(self) -> None:
        """停止监控"""
        logger.info('正在停止目录监控...')
        self._stop_event.set()

    def trigger_download(self, bank_name: str) -> Tuple[bool, str]:
        """触发指定银行的下载脚本"""
        bank_cfg = self._bank_configs.get(bank_name)
        if not bank_cfg:
            return False, f'银行 {bank_name} 未配置'

        if not bank_cfg.download_script:
            return False, f'银行 {bank_name} 未配置下载脚本'

        script_type = bank_cfg.download_script.get('type')
        template_name = bank_cfg.download_script.get('template')
        params = bank_cfg.download_script.get('params', {})

        if script_type == 'template':
            try:
                from bank_download_templates import run_download_template
                inbox_dir = os.path.join(self._directory_config.root_dir, 'inbox')
                success, message = run_download_template(
                    template_name,
                    output_dir=inbox_dir,
                    **params
                )
                return success, message
            except ImportError:
                return False, '下载脚本模板模块不可用'
            except Exception as e:
                return False, f'执行下载脚本失败: {e}'

        return False, f'不支持的脚本类型: {script_type}'

    def get_status(self) -> Dict[str, Any]:
        """获取当前状态"""
        root = self._directory_config.root_dir

        def count_files(directory: str) -> int:
            dir_path = os.path.join(root, directory)
            if not os.path.isdir(dir_path):
                return 0
            return len([f for f in os.listdir(dir_path)
                        if os.path.isfile(os.path.join(dir_path, f)) and not f.startswith('.')])

        inbox_files = self._scan_inbox()
        valid_files = self._validate_files(inbox_files)

        return {
            'root_dir': root,
            'inbox_count': count_files('inbox'),
            'outbox_count': count_files('outbox'),
            'error_count': count_files('error'),
            'pending_files': [os.path.basename(f.filepath) for f in inbox_files],
            'ready_files': [os.path.basename(f.filepath) for f in valid_files],
            'configured_banks': [
                {
                    'name': name,
                    'enabled': cfg.enabled,
                    'pattern': cfg.file_pattern,
                    'has_download_script': cfg.download_script is not None,
                }
                for name, cfg in self._bank_configs.items()
            ],
            'poll_interval': self._directory_config.poll_interval,
        }


def main():
    """命令行入口"""
    import argparse

    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S',
    )

    parser = argparse.ArgumentParser(description='银企直连目录对接工具')
    parser.add_argument('command', choices=['run-once', 'watch', 'status', 'download'],
                        help='执行命令')
    parser.add_argument('--config', default=DEFAULT_CONFIG_PATH,
                        help='配置文件路径')
    parser.add_argument('--bank', help='指定银行名称（用于 download 命令）')
    parser.add_argument('--stop-on-empty', action='store_true',
                        help='watch 模式下，无文件时自动退出')

    args = parser.parse_args()

    connector = BankDirectoryConnector(config_path=args.config)

    if args.command == 'run-once':
        result = connector.run_once()
        print(f'\n处理结果: {result.message}')
        print(f'成功文件数: {len(result.processed_files)}')
        print(f'失败文件数: {len(result.error_files)}')
        if result.output_path:
            print(f'输出文件: {result.output_path}')
        if result.archive_dir:
            print(f'归档目录: {result.archive_dir}')
        sys.exit(0 if result.success else 1)

    elif args.command == 'watch':
        try:
            connector.watch(stop_on_first_empty=args.stop_on_empty)
        except KeyboardInterrupt:
            connector.stop()
        sys.exit(0)

    elif args.command == 'status':
        status = connector.get_status()
        import json
        print(json.dumps(status, ensure_ascii=False, indent=2))
        sys.exit(0)

    elif args.command == 'download':
        if not args.bank:
            print('错误: 请使用 --bank 指定银行名称')
            sys.exit(1)
        success, message = connector.trigger_download(args.bank)
        print(f'{"成功" if success else "失败"}: {message}')
        sys.exit(0 if success else 1)


if __name__ == '__main__':
    main()
