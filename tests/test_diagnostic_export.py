#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
远程诊断包一键导出模块测试
"""

import os
import sys
import json
import shutil
import tempfile
import zipfile

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'backend'))

import diagnostic_export as de
from diagnostic_export import (
    DiagnosticExportResult,
    collect_environment_info,
    collect_log_files,
    collect_config_summary,
    collect_file_tree,
    collect_troubleshooting_report,
    export_diagnostic_package,
    _mask_sensitive_in_dict,
    _mask_sensitive_in_string,
    _should_skip_file,
    _read_and_mask_log,
    print_export_result,
    MASKED_VALUE,
)


@pytest.fixture
def temp_script_dir():
    d = tempfile.mkdtemp(prefix='diag_test_')
    yield d
    shutil.rmtree(d, ignore_errors=True)


@pytest.fixture
def temp_output_dir():
    d = tempfile.mkdtemp(prefix='diag_output_')
    yield d
    shutil.rmtree(d, ignore_errors=True)


@pytest.fixture
def script_dir_with_logs(temp_script_dir):
    log_dir = os.path.join(temp_script_dir, 'logs')
    os.makedirs(log_dir, exist_ok=True)
    for name, content in [
        ('bankcheck_20260614_120000.log', (
            '[2026-06-14 12:00:00] INFO - 开始处理\n'
            '[2026-06-14 12:00:01] WARNING - 未找到主体查找表\n'
            '[2026-06-14 12:00:02] DEBUG - 账号 6222021234567890123\n'
            '[2026-06-14 12:00:03] INFO - 处理完成\n'
        )),
        ('bankcheck_20260615_090000.log', (
            '[2026-06-15 09:00:00] INFO - 启动程序\n'
            '[2026-06-15 09:00:01] ERROR - 文件格式错误\n'
            '[2026-06-15 09:00:02] INFO - 结束\n'
        )),
    ]:
        with open(os.path.join(log_dir, name), 'w', encoding='utf-8') as f:
            f.write(content)
    return temp_script_dir


@pytest.fixture
def script_dir_with_configs(temp_script_dir):
    rules_yaml = """
banks:
  北京银行:
    prefix: 北京银行
    account:
      row: 2
      column: 2
    header_row: 3
    columns:
      交易日期: 2
      支出金额: 4
      收入金额: 5
      余额: 6
      账号: 8
"""
    with open(os.path.join(temp_script_dir, 'bank_rules.yaml'), 'w', encoding='utf-8') as f:
        f.write(rules_yaml)

    scheduler_config = {
        'enabled': True,
        'interval_minutes': 30,
        'account': '6222021234567890123',
        'password': 'secret123',
    }
    with open(os.path.join(temp_script_dir, 'scheduler_config.json'), 'w', encoding='utf-8') as f:
        json.dump(scheduler_config, f, ensure_ascii=False, indent=2)

    return temp_script_dir


@pytest.fixture
def script_dir_with_files(temp_script_dir):
    for subdir in ['logs', 'output', 'i18n/locales']:
        os.makedirs(os.path.join(temp_script_dir, subdir), exist_ok=True)
    for name in ['bankcheck.py', 'bank_rules.yaml', 'app.py', 'database.py']:
        with open(os.path.join(temp_script_dir, name), 'w', encoding='utf-8') as f:
            f.write('# test file\n')
    with open(os.path.join(temp_script_dir, 'audit_log.db'), 'w') as f:
        f.write('fake db')
    with open(os.path.join(temp_script_dir, 'transactions.db'), 'w') as f:
        f.write('fake db')
    return temp_script_dir


class TestMaskSensitiveInDict:
    def test_mask_account_key(self):
        data = {'银行账号': '6222021234567890123', 'name': 'test'}
        result = _mask_sensitive_in_dict(data)
        assert result['银行账号'] == MASKED_VALUE
        assert result['name'] == 'test'

    def test_mask_password_key(self):
        data = {'password': 'secret123', 'username': 'admin'}
        result = _mask_sensitive_in_dict(data)
        assert result['password'] == MASKED_VALUE
        assert result['username'] == 'admin'

    def test_mask_nested_dict(self):
        data = {'config': {'account': '123456', 'setting': 'value'}}
        result = _mask_sensitive_in_dict(data)
        assert result['config']['account'] == MASKED_VALUE
        assert result['config']['setting'] == 'value'

    def test_mask_in_list(self):
        data = [{'account': '123', 'name': 'a'}, {'account': '456', 'name': 'b'}]
        result = _mask_sensitive_in_dict(data)
        assert result[0]['account'] == MASKED_VALUE
        assert result[1]['account'] == MASKED_VALUE

    def test_no_masking_needed(self):
        data = {'name': 'test', 'version': '1.0.0', 'count': 42}
        result = _mask_sensitive_in_dict(data)
        assert result == data

    def test_max_depth(self):
        data = {'a': {'b': {'c': {'d': {'e': 'deep'}}}}}
        result = _mask_sensitive_in_dict(data)
        assert result is not None

    def test_mask_收款_key(self):
        data = {'收款': 1000.00, 'name': 'test'}
        result = _mask_sensitive_in_dict(data)
        assert result['收款'] == MASKED_VALUE
        assert result['name'] == 'test'

    def test_mask_token_key(self):
        data = {'api_key': 'sk-12345', 'base_url': 'http://example.com'}
        result = _mask_sensitive_in_dict(data)
        assert result['api_key'] == MASKED_VALUE
        assert result['base_url'] == 'http://example.com'


class TestMaskSensitiveInString:
    def test_mask_long_digit_sequence(self):
        text = '账号 6222021234567890123 处理完成'
        result = _mask_sensitive_in_string(text)
        assert '6222021234567890123' not in result
        assert '***' in result

    def test_mask_amount(self):
        text = '金额 1,234.56 已到账'
        result = _mask_sensitive_in_string(text)
        assert '1,234.56' not in result

    def test_no_masking_normal_text(self):
        text = '处理完成，共提取 100 条记录'
        result = _mask_sensitive_in_string(text)
        assert result == text

    def test_empty_string(self):
        assert _mask_sensitive_in_string('') == ''

    def test_multiple_sensitive_patterns(self):
        text = '账号 6222021234567890123 金额 5,000.00'
        result = _mask_sensitive_in_string(text)
        assert '6222021234567890123' not in result
        assert '5,000.00' not in result


class TestShouldSkipFile:
    def test_skip_db_files(self):
        assert _should_skip_file('audit_log.db') is True
        assert _should_skip_file('data.sqlite') is True

    def test_skip_pyc_files(self):
        assert _should_skip_file('__init__.pyc') is True
        assert _should_skip_file('module.pyo') is True

    def test_do_not_skip_normal_files(self):
        assert _should_skip_file('bank_rules.yaml') is False
        assert _should_skip_file('app.py') is False
        assert _should_skip_file('data.json') is False

    def test_skip_specific_db_names(self):
        assert _should_skip_file('transactions.db') is True


class TestCollectEnvironmentInfo:
    def test_basic_structure(self):
        info = collect_environment_info()
        assert 'collection_time' in info
        assert 'python' in info
        assert 'os' in info
        assert 'app' in info
        assert 'dependencies' in info
        assert 'self_check' in info

    def test_python_info(self):
        info = collect_environment_info()
        assert 'version' in info['python']
        assert 'executable' in info['python']
        assert 'platform' in info['python']

    def test_os_info(self):
        info = collect_environment_info()
        assert 'system' in info['os']
        assert 'release' in info['os']

    def test_dependencies_list(self):
        info = collect_environment_info()
        deps = info['dependencies']
        assert 'openpyxl' in deps
        assert 'pandas' in deps
        assert 'PyYAML' in deps

    def test_app_version(self):
        info = collect_environment_info()
        assert 'version' in info['app']

    def test_self_check_included(self):
        info = collect_environment_info()
        sc = info['self_check']
        assert 'passed' in sc or 'error' in sc or 'status' in sc


class TestCollectLogFiles:
    def test_collect_from_log_dir(self, script_dir_with_logs):
        logs = collect_log_files(script_dir_with_logs)
        assert len(logs) > 0

    def test_log_content_no_debug(self, script_dir_with_logs):
        logs = collect_log_files(script_dir_with_logs)
        for name, content in logs.items():
            lines = content.split('\n')
            for line in lines:
                if line.strip():
                    assert 'DEBUG' not in line

    def test_log_content_no_sensitive(self, script_dir_with_logs):
        logs = collect_log_files(script_dir_with_logs)
        for name, content in logs.items():
            assert '6222021234567890123' not in content

    def test_no_log_dir(self, temp_script_dir):
        logs = collect_log_files(temp_script_dir)
        assert isinstance(logs, dict)

    def test_single_log_file(self, temp_script_dir):
        with open(os.path.join(temp_script_dir, 'bankcheck.log'), 'w', encoding='utf-8') as f:
            f.write('[2026-06-15 09:00:00] INFO - 测试日志\n')
        logs = collect_log_files(temp_script_dir)
        assert 'bankcheck.log' in logs

    def test_max_three_logs(self, script_dir_with_logs):
        log_dir = os.path.join(script_dir_with_logs, 'logs')
        for i in range(3, 8):
            name = f'bankcheck_2026061{i}_090000.log'
            with open(os.path.join(log_dir, name), 'w', encoding='utf-8') as f:
                f.write('[2026-06-15 09:00:00] INFO - test\n')
        logs = collect_log_files(script_dir_with_logs)
        assert len(logs) <= 3


class TestReadAndMaskLog:
    def test_reads_file(self, tmp_path):
        log_file = tmp_path / 'test.log'
        log_file.write_text('[2026-06-15 09:00:00] INFO - 测试日志\n', encoding='utf-8')
        content = _read_and_mask_log(str(log_file), max_lines=100)
        assert '测试日志' in content

    def test_masks_sensitive(self, tmp_path):
        log_file = tmp_path / 'test.log'
        log_file.write_text(
            '[2026-06-15 09:00:00] INFO - 账号 6222021234567890123\n',
            encoding='utf-8',
        )
        content = _read_and_mask_log(str(log_file), max_lines=100)
        assert '6222021234567890123' not in content

    def test_filters_debug_lines(self, tmp_path):
        log_file = tmp_path / 'test.log'
        log_file.write_text(
            '[2026-06-15 09:00:00] INFO - normal line\n'
            '[2026-06-15 09:00:01] DEBUG - debug line\n'
            '[2026-06-15 09:00:02] WARNING - warning line\n',
            encoding='utf-8',
        )
        content = _read_and_mask_log(str(log_file), max_lines=100)
        assert 'normal line' in content
        assert 'debug line' not in content
        assert 'warning line' in content

    def test_truncates_long_file(self, tmp_path):
        log_file = tmp_path / 'long.log'
        lines = [f'[2026-06-15 09:00:{i:02d}] INFO - line {i}\n' for i in range(100)]
        log_file.write_text(''.join(lines), encoding='utf-8')
        content = _read_and_mask_log(str(log_file), max_lines=10)
        content_lines = [l for l in content.split('\n') if l.strip()]
        assert len(content_lines) <= 10

    def test_nonexistent_file(self):
        content = _read_and_mask_log('/nonexistent/path/test.log', max_lines=100)
        assert '读取失败' in content


class TestCollectConfigSummary:
    def test_collects_yaml_config(self, script_dir_with_configs):
        summary = collect_config_summary(script_dir_with_configs)
        assert 'bank_rules.yaml' in summary
        assert summary['bank_rules.yaml']['status'] == 'loaded'

    def test_collects_json_config(self, script_dir_with_configs):
        summary = collect_config_summary(script_dir_with_configs)
        assert 'scheduler_config.json' in summary
        assert summary['scheduler_config.json']['status'] == 'loaded'

    def test_masks_sensitive_in_config(self, script_dir_with_configs):
        summary = collect_config_summary(script_dir_with_configs)
        scheduler = summary['scheduler_config.json']
        if scheduler['status'] == 'loaded':
            content = scheduler['content']
            assert content.get('account') == MASKED_VALUE
            assert content.get('password') == MASKED_VALUE
            assert content.get('enabled') is True

    def test_missing_config(self, temp_script_dir):
        summary = collect_config_summary(temp_script_dir)
        for name in ['bank_rules.yaml', 'scheduler_config.json']:
            assert name in summary
            assert summary[name]['status'] == 'not found'

    def test_config_size(self, script_dir_with_configs):
        summary = collect_config_summary(script_dir_with_configs)
        for name, info in summary.items():
            if info['status'] == 'loaded':
                assert 'size_bytes' in info
                assert info['size_bytes'] > 0


class TestCollectFileTree:
    def test_collects_tree(self, script_dir_with_files):
        tree = collect_file_tree(script_dir_with_files)
        assert 'root' in tree
        assert 'entries' in tree
        assert len(tree['entries']) > 0

    def test_entry_types(self, script_dir_with_files):
        tree = collect_file_tree(script_dir_with_files)
        types = {e['type'] for e in tree['entries']}
        assert 'dir' in types
        assert 'file' in types

    def test_file_has_size(self, script_dir_with_files):
        tree = collect_file_tree(script_dir_with_files)
        file_entries = [e for e in tree['entries'] if e['type'] == 'file']
        for entry in file_entries:
            assert 'size' in entry

    def test_skips_db_files(self, script_dir_with_files):
        tree = collect_file_tree(script_dir_with_files)
        file_paths = [e['path'] for e in tree['entries'] if e['type'] == 'file']
        for path in file_paths:
            assert not path.endswith('.db')

    def test_skips_pycache(self, script_dir_with_files):
        pycache_dir = os.path.join(script_dir_with_files, '__pycache__')
        os.makedirs(pycache_dir, exist_ok=True)
        with open(os.path.join(pycache_dir, 'test.pyc'), 'w') as f:
            f.write('fake')
        tree = collect_file_tree(script_dir_with_files)
        paths = [e['path'] for e in tree['entries']]
        for p in paths:
            assert '__pycache__' not in p

    def test_max_depth(self, temp_script_dir):
        deep_dir = temp_script_dir
        for i in range(5):
            deep_dir = os.path.join(deep_dir, f'level{i}')
        os.makedirs(deep_dir, exist_ok=True)
        with open(os.path.join(deep_dir, 'deep_file.txt'), 'w') as f:
            f.write('deep')

        tree = collect_file_tree(temp_script_dir, max_depth=2)
        paths = [e['path'] for e in tree['entries']]
        has_deep = any('level3' in p for p in paths)
        assert not has_deep

    def test_max_entries(self, temp_script_dir):
        for i in range(50):
            with open(os.path.join(temp_script_dir, f'file_{i:03d}.txt'), 'w') as f:
                f.write('x')
        tree = collect_file_tree(temp_script_dir, max_entries=10)
        assert len(tree['entries']) <= 10
        if len(tree['entries']) == 10:
            assert tree.get('truncated') is True

    def test_empty_dir(self, temp_script_dir):
        tree = collect_file_tree(temp_script_dir)
        assert 'entries' in tree
        assert isinstance(tree['entries'], list)


class TestCollectTroubleshootingReport:
    def test_returns_dict(self):
        report = collect_troubleshooting_report()
        assert isinstance(report, dict)
        assert 'status' in report

    def test_status_value(self):
        report = collect_troubleshooting_report()
        assert report['status'] in ('ok', 'error', 'troubleshooter module not available')


class TestExportDiagnosticPackage:
    def test_basic_export(self, script_dir_with_logs, temp_output_dir):
        result = export_diagnostic_package(
            output_dir=temp_output_dir,
            script_dir=script_dir_with_logs,
        )
        assert result.success is True
        assert result.zip_path != ''
        assert result.file_count > 0
        assert result.total_size > 0
        assert os.path.exists(result.zip_path)

    def test_zip_is_valid(self, script_dir_with_logs, temp_output_dir):
        result = export_diagnostic_package(
            output_dir=temp_output_dir,
            script_dir=script_dir_with_logs,
        )
        assert zipfile.is_zipfile(result.zip_path)

    def test_zip_contains_manifest(self, script_dir_with_logs, temp_output_dir):
        result = export_diagnostic_package(
            output_dir=temp_output_dir,
            script_dir=script_dir_with_logs,
        )
        with zipfile.ZipFile(result.zip_path, 'r') as zf:
            assert 'MANIFEST.json' in zf.namelist()

    def test_manifest_content(self, script_dir_with_logs, temp_output_dir):
        result = export_diagnostic_package(
            output_dir=temp_output_dir,
            script_dir=script_dir_with_logs,
        )
        with zipfile.ZipFile(result.zip_path, 'r') as zf:
            manifest = json.loads(zf.read('MANIFEST.json'))
            assert 'tool' in manifest
            assert 'version' in manifest
            assert 'generated_at' in manifest
            assert 'files_included' in manifest
            assert 'privacy_note' in manifest
            assert 'components' in manifest

    def test_zip_contains_env_info(self, script_dir_with_logs, temp_output_dir):
        result = export_diagnostic_package(
            output_dir=temp_output_dir,
            script_dir=script_dir_with_logs,
            include_env_info=True,
        )
        with zipfile.ZipFile(result.zip_path, 'r') as zf:
            assert 'environment_info.json' in zf.namelist()
            env = json.loads(zf.read('environment_info.json'))
            assert 'python' in env
            assert 'os' in env

    def test_zip_contains_logs(self, script_dir_with_logs, temp_output_dir):
        result = export_diagnostic_package(
            output_dir=temp_output_dir,
            script_dir=script_dir_with_logs,
            include_logs=True,
        )
        with zipfile.ZipFile(result.zip_path, 'r') as zf:
            log_files = [n for n in zf.namelist() if n.startswith('logs/')]
            assert len(log_files) > 0

    def test_zip_contains_config(self, script_dir_with_configs, temp_output_dir):
        result = export_diagnostic_package(
            output_dir=temp_output_dir,
            script_dir=script_dir_with_configs,
            include_config=True,
        )
        with zipfile.ZipFile(result.zip_path, 'r') as zf:
            assert 'config_summary.json' in zf.namelist()

    def test_zip_contains_file_tree(self, script_dir_with_files, temp_output_dir):
        result = export_diagnostic_package(
            output_dir=temp_output_dir,
            script_dir=script_dir_with_files,
            include_file_tree=True,
        )
        with zipfile.ZipFile(result.zip_path, 'r') as zf:
            assert 'file_tree.json' in zf.namelist()
            tree = json.loads(zf.read('file_tree.json'))
            assert 'entries' in tree

    def test_exclude_components(self, script_dir_with_logs, temp_output_dir):
        result = export_diagnostic_package(
            output_dir=temp_output_dir,
            script_dir=script_dir_with_logs,
            include_logs=False,
            include_config=False,
            include_file_tree=False,
            include_env_info=False,
            include_troubleshooting=False,
        )
        with zipfile.ZipFile(result.zip_path, 'r') as zf:
            names = zf.namelist()
            assert 'environment_info.json' not in names
            assert 'config_summary.json' not in names
            assert 'file_tree.json' not in names
            log_entries = [n for n in names if n.startswith('logs/')]
            assert len(log_entries) == 0
            assert 'MANIFEST.json' in names

    def test_no_sesitive_data_in_zip(self, script_dir_with_configs, temp_output_dir):
        result = export_diagnostic_package(
            output_dir=temp_output_dir,
            script_dir=script_dir_with_configs,
        )
        assert result.success is True

        with zipfile.ZipFile(result.zip_path, 'r') as zf:
            for name in zf.namelist():
                content = zf.read(name).decode('utf-8', errors='ignore')
                assert '6222021234567890123' not in content, f'Sensitive account found in {name}'
                assert 'secret123' not in content, f'Sensitive password found in {name}'

    def test_output_dir_creation(self, script_dir_with_logs, temp_output_dir):
        nested_dir = os.path.join(temp_output_dir, 'nested', 'sub', 'dir')
        result = export_diagnostic_package(
            output_dir=nested_dir,
            script_dir=script_dir_with_logs,
        )
        assert result.success is True
        assert os.path.exists(result.zip_path)

    def test_invalid_output_dir(self, script_dir_with_logs):
        result = export_diagnostic_package(
            output_dir='/nonexistent_root_no_perm/dir',
            script_dir=script_dir_with_logs,
        )
        assert result.success is False
        assert result.error_message != ''

    def test_default_script_dir(self, temp_output_dir):
        result = export_diagnostic_package(
            output_dir=temp_output_dir,
            script_dir=None,
        )
        assert isinstance(result, DiagnosticExportResult)

    def test_default_output_dir(self, script_dir_with_logs):
        result = export_diagnostic_package(
            output_dir=None,
            script_dir=script_dir_with_logs,
        )
        assert isinstance(result, DiagnosticExportResult)
        if result.success:
            assert script_dir_with_logs in result.zip_path


class TestDiagnosticExportResult:
    def test_default_values(self):
        result = DiagnosticExportResult()
        assert result.success is False
        assert result.zip_path == ''
        assert result.file_count == 0
        assert result.total_size == 0
        assert result.error_message == ''
        assert result.timestamp == ''

    def test_to_dict(self):
        result = DiagnosticExportResult(
            success=True,
            zip_path='/path/to/diag.zip',
            file_count=5,
            total_size=1024,
            timestamp='20260615_120000',
        )
        d = result.to_dict()
        assert d['success'] is True
        assert d['zip_path'] == '/path/to/diag.zip'
        assert d['file_count'] == 5
        assert d['total_size'] == 1024

    def test_result_with_error(self):
        result = DiagnosticExportResult(
            success=False,
            error_message='磁盘空间不足',
        )
        assert result.success is False
        assert result.error_message == '磁盘空间不足'


class TestPrintExportResult:
    def test_print_success(self, capsys):
        result = DiagnosticExportResult(
            success=True,
            zip_path='/tmp/diag.zip',
            file_count=6,
            total_size=2048,
            timestamp='20260615_120000',
        )
        print_export_result(result)
        captured = capsys.readouterr()
        assert '成功' in captured.out
        assert '/tmp/diag.zip' in captured.out
        assert '6' in captured.out
        assert '隐私说明' in captured.out

    def test_print_failure(self, capsys):
        result = DiagnosticExportResult(
            success=False,
            error_message='磁盘空间不足',
            timestamp='20260615_120000',
        )
        print_export_result(result)
        captured = capsys.readouterr()
        assert '失败' in captured.out
        assert '磁盘空间不足' in captured.out


class TestIntegration:
    def test_full_export_workflow(self, temp_script_dir, temp_output_dir):
        log_dir = os.path.join(temp_script_dir, 'logs')
        os.makedirs(log_dir, exist_ok=True)
        with open(os.path.join(log_dir, 'bankcheck_20260615_090000.log'), 'w', encoding='utf-8') as f:
            f.write('[2026-06-15 09:00:00] INFO - 开始处理\n')
            f.write('[2026-06-15 09:00:01] DEBUG - 账号 6222021234567890123\n')
            f.write('[2026-06-15 09:00:02] WARNING - 文件格式错误\n')

        with open(os.path.join(temp_script_dir, 'bank_rules.yaml'), 'w', encoding='utf-8') as f:
            f.write('banks:\n  test_bank:\n    prefix: 测试银行\n    account:\n      row: 2\n')

        result = export_diagnostic_package(
            output_dir=temp_output_dir,
            script_dir=temp_script_dir,
        )

        assert result.success is True
        assert os.path.exists(result.zip_path)
        assert zipfile.is_zipfile(result.zip_path)

        with zipfile.ZipFile(result.zip_path, 'r') as zf:
            names = zf.namelist()
            assert 'MANIFEST.json' in names
            assert 'environment_info.json' in names
            assert 'config_summary.json' in names
            assert 'file_tree.json' in names
            assert 'troubleshooting_report.json' in names

            log_entries = [n for n in names if n.startswith('logs/')]
            assert len(log_entries) > 0

            for name in names:
                content = zf.read(name).decode('utf-8', errors='ignore')
                assert '6222021234567890123' not in content

            manifest = json.loads(zf.read('MANIFEST.json'))
            assert manifest['components']['environment_info'] is True
            assert manifest['components']['logs'] is True
            assert manifest['components']['config_summary'] is True
            assert manifest['components']['file_tree'] is True
            assert manifest['components']['troubleshooting_report'] is True

    def test_export_empty_script_dir(self, temp_script_dir, temp_output_dir):
        result = export_diagnostic_package(
            output_dir=temp_output_dir,
            script_dir=temp_script_dir,
        )
        assert result.success is True
        assert os.path.exists(result.zip_path)

    def test_export_no_duplicates_on_retry(self, script_dir_with_logs, temp_output_dir):
        result1 = export_diagnostic_package(
            output_dir=temp_output_dir,
            script_dir=script_dir_with_logs,
        )
        result2 = export_diagnostic_package(
            output_dir=temp_output_dir,
            script_dir=script_dir_with_logs,
        )
        assert result1.success is True
        assert result2.success is True
        assert result1.zip_path != result2.zip_path
        assert os.path.exists(result1.zip_path)
        assert os.path.exists(result2.zip_path)


if __name__ == '__main__':
    pytest.main([__file__, '-v'])
