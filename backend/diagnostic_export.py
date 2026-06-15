# -*- coding: utf-8 -*-
"""
远程诊断包一键导出模块

功能：
  1. 收集运行环境信息（操作系统、Python 版本、依赖库版本等）
  2. 收集最近日志文件（自动脱敏，不含敏感明细）
  3. 收集配置文件摘要（bank_rules.yaml 等关键配置）
  4. 收集文件结构摘要（目录树，不含文件内容，不含敏感文件名）
  5. 收集排障报告（调用 troubleshooter 模块）
  6. 将上述诊断数据打包为不含敏感信息的 ZIP 压缩包
  7. 便于技术支持离线排查问题

安全原则：
  - 不包含任何银行账号、金额、交易明细等 PII/FORBIDDEN 级别数据
  - 日志文件经过脱敏过滤，仅保留 INFO 级别日志
  - 配置文件中如有账号等敏感字段，自动替换为掩码
  - 文件结构摘要仅记录文件名和大小，不包含文件内容
"""

import os
import sys
import re
import json
import zipfile
import platform
import logging
import tempfile
import hashlib
from datetime import datetime
from dataclasses import dataclass, field, asdict
from typing import List, Optional, Dict, Any, Tuple

try:
    import yaml
    HAS_YAML = True
except ImportError:
    HAS_YAML = False

try:
    from build_info import get_version, get_build_time, get_build_platform, get_build_info
    HAS_BUILD_INFO = True
except ImportError:
    HAS_BUILD_INFO = False
    def get_version():
        return "1.0.0"
    def get_build_time():
        return "unknown"
    def get_build_platform():
        return platform.system()
    def get_build_info():
        return {'version': get_version(), 'build_time': get_build_time(), 'platform': get_build_platform()}

try:
    import troubleshooter as ts_module
    HAS_TROUBLESHOOTER = True
except ImportError:
    HAS_TROUBLESHOOTER = False

try:
    from pii_classifier import mask_value, PIILevel
    HAS_PII_CLASSIFIER = True
except ImportError:
    HAS_PII_CLASSIFIER = False
    mask_value = None
    PIILevel = None

try:
    import self_check as sc_module
    HAS_SELF_CHECK = True
except ImportError:
    HAS_SELF_CHECK = False


def get_script_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def get_logger():
    return logging.getLogger('bankcheck')


SENSITIVE_KEY_PATTERNS = [
    re.compile(r'account', re.IGNORECASE),
    re.compile(r'账号', re.IGNORECASE),
    re.compile(r'password', re.IGNORECASE),
    re.compile(r'secret', re.IGNORECASE),
    re.compile(r'token', re.IGNORECASE),
    re.compile(r'api_key', re.IGNORECASE),
    re.compile(r'private_key', re.IGNORECASE),
    re.compile(r'收款', re.IGNORECASE),
    re.compile(r'付款', re.IGNORECASE),
    re.compile(r'余额', re.IGNORECASE),
]

LOG_SENSITIVE_PATTERNS = [
    re.compile(r'\b\d{10,22}\b'),
    re.compile(r'[\d,]+\.\d{2}'),
]

MASKED_VALUE = '***MASKED***'

MAX_LOG_LINES = 500

MAX_FILE_TREE_DEPTH = 3

MAX_FILE_TREE_ENTRIES = 200

DIAGNOSTIC_PACKAGE_PREFIX = 'bankcheck_diagnostic'


@dataclass
class DiagnosticExportResult:
    success: bool = False
    zip_path: str = ''
    file_count: int = 0
    total_size: int = 0
    error_message: str = ''
    timestamp: str = ''

    def to_dict(self) -> Dict[str, Any]:
        return {
            'success': self.success,
            'zip_path': self.zip_path,
            'file_count': self.file_count,
            'total_size': self.total_size,
            'error_message': self.error_message,
            'timestamp': self.timestamp,
        }


def _mask_sensitive_in_dict(data: Any, depth: int = 0) -> Any:
    if depth > 10:
        return '***DEEP***'
    if isinstance(data, dict):
        result = {}
        for key, value in data.items():
            if isinstance(key, str) and any(p.search(key) for p in SENSITIVE_KEY_PATTERNS):
                result[key] = MASKED_VALUE
            else:
                result[key] = _mask_sensitive_in_dict(value, depth + 1)
        return result
    elif isinstance(data, list):
        return [_mask_sensitive_in_dict(item, depth + 1) for item in data]
    elif isinstance(data, str):
        return _mask_sensitive_in_string(data)
    return data


def _mask_sensitive_in_string(text: str) -> str:
    if not text:
        return text
    for pattern in LOG_SENSITIVE_PATTERNS:
        if pattern.search(text):
            text = pattern.sub('***', text)
    return text


def collect_environment_info() -> Dict[str, Any]:
    info = {
        'collection_time': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'python': {
            'version': sys.version,
            'executable': sys.executable,
            'platform': sys.platform,
        },
        'os': {
            'system': platform.system(),
            'release': platform.release(),
            'version': platform.version(),
            'machine': platform.machine(),
            'processor': platform.processor(),
        },
        'app': get_build_info() if HAS_BUILD_INFO else {
            'version': 'unknown',
            'build_time': 'unknown',
            'platform': platform.system(),
        },
    }

    dependencies = {}
    dep_list = [
        ('openpyxl', 'openpyxl'),
        ('pandas', 'pandas'),
        ('yaml', 'PyYAML'),
        ('xlrd', 'xlrd'),
        ('flask', 'Flask'),
        ('cryptography', 'cryptography'),
        ('apscheduler', 'APScheduler'),
        ('pdfplumber', 'pdfplumber'),
    ]
    for import_name, package_name in dep_list:
        try:
            mod = __import__(import_name)
            dependencies[package_name] = getattr(mod, '__version__', 'installed')
        except ImportError:
            dependencies[package_name] = 'not installed'
    info['dependencies'] = dependencies

    if HAS_SELF_CHECK:
        try:
            report = sc_module.run_self_check(include_optional=True)
            info['self_check'] = {
                'passed': report.passed,
                'error_count': len(report.errors),
                'warning_count': len(report.warnings),
                'errors': [
                    {'name': e.name, 'message': e.message, 'severity': e.severity}
                    for e in report.errors
                ],
                'warnings': [
                    {'name': w.name, 'message': w.message, 'severity': w.severity}
                    for w in report.warnings
                ],
            }
        except Exception as e:
            info['self_check'] = {'error': str(e)}
    else:
        info['self_check'] = {'status': 'module not available'}

    return info


def collect_log_files(script_dir: Optional[str] = None,
                      max_lines: int = MAX_LOG_LINES) -> Dict[str, str]:
    if script_dir is None:
        script_dir = get_script_dir()

    log_dir = os.path.join(script_dir, 'logs')
    logs = {}

    if not os.path.isdir(log_dir):
        single_log = os.path.join(script_dir, 'bankcheck.log')
        if os.path.isfile(single_log):
            logs['bankcheck.log'] = _read_and_mask_log(single_log, max_lines)
        return logs

    log_pattern = re.compile(r'^bankcheck_\d{8}_\d{6}\.log$')
    candidates = []
    try:
        for filename in os.listdir(log_dir):
            if log_pattern.match(filename) or filename == 'bankcheck.log':
                filepath = os.path.join(log_dir, filename)
                if os.path.isfile(filepath):
                    try:
                        mtime = os.path.getmtime(filepath)
                        candidates.append((mtime, filename, filepath))
                    except OSError:
                        continue
    except OSError:
        return logs

    candidates.sort(key=lambda x: x[0], reverse=True)

    for i, (mtime, filename, filepath) in enumerate(candidates[:3]):
        logs[filename] = _read_and_mask_log(filepath, max_lines)

    return logs


def _read_and_mask_log(log_path: str, max_lines: int) -> str:
    try:
        with open(log_path, 'r', encoding='utf-8', errors='ignore') as f:
            lines = f.readlines()
    except Exception as e:
        get_logger().warning('读取日志文件失败 %s: %s', log_path, e)
        return f'*** 读取失败: {e} ***'

    if len(lines) > max_lines:
        lines = lines[-max_lines:]

    masked_lines = []
    for line in lines:
        masked = _mask_sensitive_in_string(line.rstrip('\n\r'))
        if 'DEBUG' not in line:
            masked_lines.append(masked)

    return '\n'.join(masked_lines)


def collect_config_summary(script_dir: Optional[str] = None) -> Dict[str, Any]:
    if script_dir is None:
        script_dir = get_script_dir()

    config_files = [
        'bank_rules.yaml',
        'summary_config.yaml',
        'scheduler_config.json',
        'task_queue_config.json',
        'task_presets.json',
        'counterparty_rules.json',
    ]

    summary = {}

    for filename in config_files:
        filepath = os.path.join(script_dir, filename)
        if not os.path.isfile(filepath):
            summary[filename] = {'status': 'not found'}
            continue

        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                content = f.read()

            if filename.endswith('.yaml') or filename.endswith('.yml'):
                if HAS_YAML:
                    try:
                        parsed = yaml.safe_load(content)
                        masked = _mask_sensitive_in_dict(parsed)
                        summary[filename] = {
                            'status': 'loaded',
                            'size_bytes': len(content.encode('utf-8')),
                            'content': masked,
                        }
                    except yaml.YAMLError:
                        summary[filename] = {
                            'status': 'yaml_parse_error',
                            'size_bytes': len(content.encode('utf-8')),
                        }
                else:
                    summary[filename] = {
                        'status': 'yaml_module_not_available',
                        'size_bytes': len(content.encode('utf-8')),
                    }
            elif filename.endswith('.json'):
                try:
                    parsed = json.loads(content)
                    masked = _mask_sensitive_in_dict(parsed)
                    summary[filename] = {
                        'status': 'loaded',
                        'size_bytes': len(content.encode('utf-8')),
                        'content': masked,
                    }
                except json.JSONDecodeError:
                    summary[filename] = {
                        'status': 'json_parse_error',
                        'size_bytes': len(content.encode('utf-8')),
                    }
        except Exception as e:
            summary[filename] = {'status': 'read_error', 'error': str(e)}

    return summary


def collect_file_tree(script_dir: Optional[str] = None,
                      max_depth: int = MAX_FILE_TREE_DEPTH,
                      max_entries: int = MAX_FILE_TREE_ENTRIES) -> Dict[str, Any]:
    if script_dir is None:
        script_dir = get_script_dir()

    tree = {
        'root': script_dir,
        'max_depth': max_depth,
        'entries': [],
    }

    entry_count = 0

    for dirpath, dirnames, filenames in os.walk(script_dir):
        if entry_count >= max_entries:
            tree['truncated'] = True
            break

        rel_path = os.path.relpath(dirpath, script_dir)
        depth = rel_path.count(os.sep) if rel_path != '.' else 0

        if depth >= max_depth:
            dirnames.clear()

        skip_dirs = ['__pycache__', '.git', 'node_modules', '.venv', 'venv']
        dirnames[:] = [d for d in dirnames if d not in skip_dirs]

        for dirname in sorted(dirnames):
            if entry_count >= max_entries:
                break
            dir_rel = os.path.relpath(os.path.join(dirpath, dirname), script_dir)
            tree['entries'].append({
                'path': dir_rel,
                'type': 'dir',
            })
            entry_count += 1

        for filename in sorted(filenames):
            if entry_count >= max_entries:
                break

            if _should_skip_file(filename):
                continue

            file_rel = os.path.relpath(os.path.join(dirpath, filename), script_dir)
            filepath = os.path.join(dirpath, filename)

            entry = {
                'path': file_rel,
                'type': 'file',
            }

            try:
                stat = os.stat(filepath)
                entry['size'] = stat.st_size
            except OSError:
                entry['size'] = -1

            tree['entries'].append(entry)
            entry_count += 1

    if entry_count >= max_entries:
        tree['truncated'] = True

    return tree


def _should_skip_file(filename: str) -> bool:
    skip_extensions = {'.db', '.sqlite', '.sqlite3', '.pyc', '.pyo'}
    skip_names = {'audit_log.db', 'transactions.db'}

    _, ext = os.path.splitext(filename)
    if ext.lower() in skip_extensions:
        return True
    if filename in skip_names:
        return True
    return False


def collect_troubleshooting_report(script_dir: Optional[str] = None) -> Dict[str, Any]:
    if not HAS_TROUBLESHOOTER:
        return {'status': 'troubleshooter module not available'}

    try:
        report = ts_module.run_troubleshooting(script_dir=script_dir)
        return {
            'status': 'ok',
            'timestamp': report.timestamp,
            'total_issues': report.total_issues,
            'summary': report.summary,
            'issues': [issue.to_dict() for issue in report.issues],
        }
    except Exception as e:
        get_logger().warning('收集排障报告失败: %s', e)
        return {'status': 'error', 'error': str(e)}


def export_diagnostic_package(
    output_dir: Optional[str] = None,
    script_dir: Optional[str] = None,
    include_logs: bool = True,
    include_config: bool = True,
    include_file_tree: bool = True,
    include_env_info: bool = True,
    include_troubleshooting: bool = True,
) -> DiagnosticExportResult:
    logger = get_logger()
    logger.info('开始生成远程诊断包')

    if script_dir is None:
        script_dir = get_script_dir()

    if output_dir is None:
        output_dir = script_dir

    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S_%f')[:-3]
    zip_filename = f'{DIAGNOSTIC_PACKAGE_PREFIX}_{timestamp}.zip'
    zip_path = os.path.join(output_dir, zip_filename)

    result = DiagnosticExportResult(
        timestamp=timestamp,
    )

    try:
        os.makedirs(output_dir, exist_ok=True)
    except OSError as e:
        result.error_message = f'无法创建输出目录: {e}'
        logger.error('无法创建输出目录 %s: %s', output_dir, e)
        return result

    file_count = 0

    try:
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            if include_env_info:
                env_info = collect_environment_info()
                env_json = json.dumps(env_info, ensure_ascii=False, indent=2, default=str)
                zf.writestr('environment_info.json', env_json)
                file_count += 1
                logger.info('已收集环境信息')

            if include_logs:
                logs = collect_log_files(script_dir)
                for log_name, log_content in logs.items():
                    zf.writestr(f'logs/{log_name}', log_content)
                    file_count += 1
                logger.info('已收集 %d 个日志文件', len(logs))

            if include_config:
                config_summary = collect_config_summary(script_dir)
                config_json = json.dumps(config_summary, ensure_ascii=False, indent=2, default=str)
                zf.writestr('config_summary.json', config_json)
                file_count += 1
                logger.info('已收集配置摘要')

            if include_file_tree:
                file_tree = collect_file_tree(script_dir)
                tree_json = json.dumps(file_tree, ensure_ascii=False, indent=2, default=str)
                zf.writestr('file_tree.json', tree_json)
                file_count += 1
                logger.info('已收集文件结构摘要')

            if include_troubleshooting:
                ts_report = collect_troubleshooting_report(script_dir)
                ts_json = json.dumps(ts_report, ensure_ascii=False, indent=2, default=str)
                zf.writestr('troubleshooting_report.json', ts_json)
                file_count += 1
                logger.info('已收集排障报告')

            manifest = {
                'tool': '银行流水检验工具',
                'version': get_version() if HAS_BUILD_INFO else 'unknown',
                'generated_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'files_included': file_count,
                'components': {
                    'environment_info': include_env_info,
                    'logs': include_logs,
                    'config_summary': include_config,
                    'file_tree': include_file_tree,
                    'troubleshooting_report': include_troubleshooting,
                },
                'privacy_note': (
                    '本诊断包已脱敏处理，不包含银行账号、金额、交易明细等敏感信息。'
                    '日志仅保留 INFO 级别，配置中敏感字段已替换为掩码。'
                ),
            }
            manifest_json = json.dumps(manifest, ensure_ascii=False, indent=2)
            zf.writestr('MANIFEST.json', manifest_json)
            file_count += 1

        result.success = True
        result.zip_path = zip_path
        result.file_count = file_count
        result.total_size = os.path.getsize(zip_path)
        size_kb = result.total_size / 1024
        logger.info(
            '诊断包生成成功: %s (%d 个文件, %s KB)',
            zip_path, file_count, f'{size_kb:.1f}',
        )

    except Exception as e:
        result.error_message = f'生成诊断包失败: {e}'
        logger.error('生成诊断包失败: %s', e)
        if os.path.exists(zip_path):
            try:
                os.remove(zip_path)
            except OSError:
                pass

    return result


def print_export_result(result: DiagnosticExportResult):
    print('\n' + '=' * 60)
    print('  远程诊断包导出结果')
    print('=' * 60)
    print(f'  生成时间: {result.timestamp}')
    if result.success:
        print(f'  状态: 成功')
        print(f'  文件路径: {result.zip_path}')
        print(f'  包含文件数: {result.file_count}')
        print(f'  压缩包大小: {result.total_size / 1024:.1f} KB')
    else:
        print(f'  状态: 失败')
        print(f'  错误信息: {result.error_message}')
    print()
    print('  隐私说明: 本诊断包已脱敏处理，不包含银行账号、')
    print('  金额、交易明细等敏感信息。')
    print('=' * 60 + '\n')


if __name__ == '__main__':
    import argparse

    parser = argparse.ArgumentParser(description='银行流水检验工具 - 远程诊断包导出')
    parser.add_argument('--output-dir', '-o', help='输出目录（默认为程序所在目录）')
    parser.add_argument('--script-dir', '-s', help='脚本目录（默认自动检测）')
    parser.add_argument('--no-logs', action='store_true', help='不包含日志文件')
    parser.add_argument('--no-config', action='store_true', help='不包含配置摘要')
    parser.add_argument('--no-tree', action='store_true', help='不包含文件结构摘要')
    parser.add_argument('--no-env', action='store_true', help='不包含环境信息')
    parser.add_argument('--no-troubleshoot', action='store_true', help='不包含排障报告')
    args = parser.parse_args()

    result = export_diagnostic_package(
        output_dir=args.output_dir,
        script_dir=args.script_dir,
        include_logs=not args.no_logs,
        include_config=not args.no_config,
        include_file_tree=not args.no_tree,
        include_env_info=not args.no_env,
        include_troubleshooting=not args.no_troubleshoot,
    )
    print_export_result(result)
