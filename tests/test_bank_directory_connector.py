# -*- coding: utf-8 -*-
"""
银企直连目录对接模块单元测试
"""

import os
import sys
import time
import shutil
import tempfile
import pytest
from datetime import datetime

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'backend'))

from conftest import _create_beijing_bank_excel, _create_east_asia_bank_excel, _create_lookup_table


class TestDirectoryConfig:
    """测试配置加载"""

    def test_load_default_config(self, tmp_dir, monkeypatch):
        """测试加载默认配置"""
        from bank_directory_connector import BankDirectoryConnector, _resolve_env_vars

        config_path = os.path.join(tmp_dir, 'test_config.yaml')
        shutil.copy(
            os.path.join(os.path.dirname(__file__), '..', 'backend', 'bank_directories.yaml'),
            config_path
        )

        monkeypatch.setenv('BANK_ROOT', tmp_dir)
        connector = BankDirectoryConnector(config_path=config_path)

        assert connector._directory_config.root_dir == tmp_dir
        assert connector._directory_config.poll_interval == 30
        assert connector._directory_config.file_stable_seconds == 5
        assert '北京银行' in connector._bank_configs
        assert connector._bank_configs['北京银行'].enabled is True

    def test_resolve_env_vars(self, monkeypatch):
        """测试环境变量解析"""
        from bank_directory_connector import _resolve_env_vars

        monkeypatch.setenv('TEST_VAR', 'test_value')
        assert _resolve_env_vars('${TEST_VAR}') == 'test_value'
        assert _resolve_env_vars('${NON_EXISTENT:-default}') == 'default'
        assert _resolve_env_vars('prefix_${TEST_VAR}_suffix') == 'prefix_test_value_suffix'
        assert _resolve_env_vars({'key': '${TEST_VAR}'}) == {'key': 'test_value'}
        assert _resolve_env_vars(['${TEST_VAR}', 'other']) == ['test_value', 'other']

    def test_ensure_directories(self, tmp_dir, monkeypatch):
        """测试目录结构创建"""
        from bank_directory_connector import BankDirectoryConnector

        config_path = os.path.join(tmp_dir, 'test_config.yaml')
        shutil.copy(
            os.path.join(os.path.dirname(__file__), '..', 'backend', 'bank_directories.yaml'),
            config_path
        )

        bank_root = os.path.join(tmp_dir, 'bank_data')
        monkeypatch.setenv('BANK_ROOT', bank_root)
        connector = BankDirectoryConnector(config_path=config_path)

        root = connector._directory_config.root_dir
        assert os.path.isdir(os.path.join(root, 'inbox'))
        assert os.path.isdir(os.path.join(root, 'outbox'))
        assert os.path.isdir(os.path.join(root, 'error'))
        assert os.path.isdir(os.path.join(root, 'processing'))
        assert os.path.isdir(os.path.join(root, 'inbox', '.lock'))


class TestFileScanning:
    """测试文件扫描和验证"""

    def _create_connector(self, tmp_dir, monkeypatch):
        """创建测试用的连接器"""
        from bank_directory_connector import BankDirectoryConnector

        config_path = os.path.join(tmp_dir, 'test_config.yaml')
        shutil.copy(
            os.path.join(os.path.dirname(__file__), '..', 'backend', 'bank_directories.yaml'),
            config_path
        )

        bank_root = os.path.join(tmp_dir, 'bank_data')
        monkeypatch.setenv('BANK_ROOT', bank_root)
        return BankDirectoryConnector(config_path=config_path)

    def test_scan_inbox(self, tmp_dir, monkeypatch):
        """测试扫描 inbox 目录"""
        connector = self._create_connector(tmp_dir, monkeypatch)
        inbox_dir = os.path.join(connector._directory_config.root_dir, 'inbox')

        _create_beijing_bank_excel(os.path.join(inbox_dir, '北京银行_20260612_093000.xlsx'))
        _create_east_asia_bank_excel(os.path.join(inbox_dir, '东亚银行_20260612_093001.xlsx'))

        with open(os.path.join(inbox_dir, 'other_file.txt'), 'w') as f:
            f.write('not an excel')

        with open(os.path.join(inbox_dir, '.hidden.xlsx'), 'w') as f:
            f.write('hidden')

        files = connector._scan_inbox()
        assert len(files) == 2
        bank_names = {f.bank_name for f in files}
        assert '北京银行' in bank_names
        assert '东亚银行' in bank_names

    def test_check_file_stable(self, tmp_dir, monkeypatch):
        """测试文件稳定性检测"""
        connector = self._create_connector(tmp_dir, monkeypatch)
        inbox_dir = os.path.join(connector._directory_config.root_dir, 'inbox')

        filepath = os.path.join(inbox_dir, '北京银行_test.xlsx')
        _create_beijing_bank_excel(filepath)

        from bank_directory_connector import FileInfo
        stat = os.stat(filepath)
        file_info = FileInfo(
            filepath=filepath,
            bank_name='北京银行',
            size=stat.st_size,
            mtime=stat.st_mtime,
        )

        connector._directory_config.file_stable_seconds = 0
        assert connector._check_file_stable(file_info) is True

        connector._directory_config.file_stable_seconds = 100
        assert connector._check_file_stable(file_info) is False

    def test_lock_file_detection(self, tmp_dir, monkeypatch):
        """测试锁定文件检测"""
        connector = self._create_connector(tmp_dir, monkeypatch)
        inbox_dir = os.path.join(connector._directory_config.root_dir, 'inbox')
        lock_dir = os.path.join(inbox_dir, '.lock')

        filepath = os.path.join(inbox_dir, '北京银行_test.xlsx')
        _create_beijing_bank_excel(filepath)

        from bank_directory_connector import FileInfo
        stat = os.stat(filepath)
        file_info = FileInfo(
            filepath=filepath,
            bank_name='北京银行',
            size=stat.st_size,
            mtime=stat.st_mtime,
        )

        assert connector._check_lock_file(file_info) is False

        lock_file = os.path.join(lock_dir, '北京银行_test.xlsx.lock')
        with open(lock_file, 'w') as f:
            f.write('locked')

        assert connector._check_lock_file(file_info) is True

    def test_validate_files(self, tmp_dir, monkeypatch):
        """测试文件验证"""
        connector = self._create_connector(tmp_dir, monkeypatch)
        inbox_dir = os.path.join(connector._directory_config.root_dir, 'inbox')
        lock_dir = os.path.join(inbox_dir, '.lock')

        f1 = os.path.join(inbox_dir, '北京银行_1.xlsx')
        _create_beijing_bank_excel(f1)

        f2 = os.path.join(inbox_dir, '北京银行_2.xlsx')
        _create_beijing_bank_excel(f2)
        with open(os.path.join(lock_dir, '北京银行_2.xlsx.lock'), 'w') as f:
            f.write('locked')

        connector._directory_config.file_stable_seconds = 0

        files = connector._scan_inbox()
        valid_files = connector._validate_files(files)

        assert len(valid_files) == 1
        assert '北京银行_1.xlsx' in valid_files[0].filepath
        assert valid_files[0].has_lock is False
        assert valid_files[0].is_stable is True


class TestFileOperations:
    """测试文件操作"""

    def _create_connector(self, tmp_dir, monkeypatch):
        from bank_directory_connector import BankDirectoryConnector

        config_path = os.path.join(tmp_dir, 'test_config.yaml')
        shutil.copy(
            os.path.join(os.path.dirname(__file__), '..', 'backend', 'bank_directories.yaml'),
            config_path
        )

        bank_root = os.path.join(tmp_dir, 'bank_data')
        monkeypatch.setenv('BANK_ROOT', bank_root)
        return BankDirectoryConnector(config_path=config_path)

    def test_move_to_processing(self, tmp_dir, monkeypatch):
        """测试移动文件到处理目录"""
        connector = self._create_connector(tmp_dir, monkeypatch)
        inbox_dir = os.path.join(connector._directory_config.root_dir, 'inbox')

        filepath = os.path.join(inbox_dir, '北京银行_test.xlsx')
        _create_beijing_bank_excel(filepath)

        from bank_directory_connector import FileInfo
        stat = os.stat(filepath)
        file_info = FileInfo(
            filepath=filepath,
            bank_name='北京银行',
            size=stat.st_size,
            mtime=stat.st_mtime,
        )

        processing_path = connector._move_to_processing(file_info)

        assert not os.path.exists(filepath)
        assert os.path.exists(processing_path)
        assert 'processing' in processing_path

    def test_move_to_error(self, tmp_dir, monkeypatch):
        """测试移动文件到错误目录"""
        connector = self._create_connector(tmp_dir, monkeypatch)
        processing_dir = os.path.join(connector._directory_config.root_dir, 'processing')
        os.makedirs(processing_dir, exist_ok=True)

        filepath = os.path.join(processing_dir, '北京银行_test.xlsx')
        _create_beijing_bank_excel(filepath)

        connector._move_to_error(filepath, '测试错误')

        error_dir = os.path.join(connector._directory_config.root_dir, 'error')
        error_files = os.listdir(error_dir)
        assert len(error_files) >= 2
        assert any('北京银行_test' in f for f in error_files)
        assert any(f.endswith('.error.txt') for f in error_files)

        error_txt = [f for f in error_files if f.endswith('.error.txt')][0]
        with open(os.path.join(error_dir, error_txt), 'r', encoding='utf-8') as f:
            content = f.read()
            assert '测试错误' in content


class TestRunOnce:
    """测试单次运行"""

    def _create_connector(self, tmp_dir, monkeypatch):
        from bank_directory_connector import BankDirectoryConnector

        config_path = os.path.join(tmp_dir, 'test_config.yaml')
        shutil.copy(
            os.path.join(os.path.dirname(__file__), '..', 'backend', 'bank_directories.yaml'),
            config_path
        )

        script_dir = os.path.join(tmp_dir, 'script')
        os.makedirs(script_dir, exist_ok=True)
        _create_lookup_table(os.path.join(script_dir, '主体查找表.xlsx'))

        bank_root = os.path.join(tmp_dir, 'bank_data')
        monkeypatch.setenv('BANK_ROOT', bank_root)
        return BankDirectoryConnector(config_path=config_path, script_dir=script_dir)

    def test_run_once_no_files(self, tmp_dir, monkeypatch):
        """测试无文件时的运行"""
        connector = self._create_connector(tmp_dir, monkeypatch)
        result = connector.run_once()

        assert result.success is True
        assert '没有可处理的文件' in result.message
        assert len(result.processed_files) == 0

    def test_run_once_with_valid_files(self, tmp_dir, monkeypatch):
        """测试处理有效文件"""
        connector = self._create_connector(tmp_dir, monkeypatch)
        inbox_dir = os.path.join(connector._directory_config.root_dir, 'inbox')

        connector._directory_config.file_stable_seconds = 0

        _create_beijing_bank_excel(os.path.join(inbox_dir, '北京银行_20260612_093000.xlsx'))

        result = connector.run_once()

        assert result.success is True
        assert len(result.processed_files) >= 1
        assert result.archive_dir is not None
        assert os.path.isdir(result.archive_dir)

        outbox_dir = os.path.join(connector._directory_config.root_dir, 'outbox')
        assert len(os.listdir(outbox_dir)) >= 1

    def test_run_once_with_error_files(self, tmp_dir, monkeypatch):
        """测试处理包含错误文件的情况"""
        connector = self._create_connector(tmp_dir, monkeypatch)
        inbox_dir = os.path.join(connector._directory_config.root_dir, 'inbox')

        connector._directory_config.file_stable_seconds = 0

        _create_beijing_bank_excel(os.path.join(inbox_dir, '北京银行_20260612_093000.xlsx'))

        corrupt_path = os.path.join(inbox_dir, '北京银行_corrupt.xlsx')
        with open(corrupt_path, 'wb') as f:
            f.write(b'not a valid excel')

        result = connector.run_once()

        assert len(result.error_files) >= 1
        error_dir = os.path.join(connector._directory_config.root_dir, 'error')
        assert len(os.listdir(error_dir)) >= 2


class TestArchiveResults:
    """测试结果归档"""

    def _create_connector(self, tmp_dir, monkeypatch):
        from bank_directory_connector import BankDirectoryConnector

        config_path = os.path.join(tmp_dir, 'test_config.yaml')
        shutil.copy(
            os.path.join(os.path.dirname(__file__), '..', 'backend', 'bank_directories.yaml'),
            config_path
        )

        bank_root = os.path.join(tmp_dir, 'bank_data')
        monkeypatch.setenv('BANK_ROOT', bank_root)
        return BankDirectoryConnector(config_path=config_path)

    def test_archive_results(self, tmp_dir, monkeypatch):
        """测试归档处理结果"""
        connector = self._create_connector(tmp_dir, monkeypatch)
        processing_dir = os.path.join(connector._directory_config.root_dir, 'processing')
        os.makedirs(processing_dir, exist_ok=True)

        _create_beijing_bank_excel(os.path.join(processing_dir, '北京银行_test.xlsx'))

        class MockPipelineResult:
            def __init__(self):
                self.output_path = os.path.join(tmp_dir, '银行流水总表.xlsx')
                self.report_path = None
                self.processed_files = ['北京银行_test.xlsx']
                self.error_files = []
                self.all_rows = [{'银行': '北京银行', '金额': 1000}]

            def __str__(self):
                return 'MockPipelineResult'

        mock_result = MockPipelineResult()
        with open(mock_result.output_path, 'w') as f:
            f.write('mock output')

        archive_dir = connector._archive_results(processing_dir, mock_result)

        assert os.path.isdir(archive_dir)
        assert 'outbox' in archive_dir

        original_dir = os.path.join(archive_dir, 'original')
        assert os.path.isdir(original_dir)
        assert os.path.exists(os.path.join(original_dir, '北京银行_test.xlsx'))
        assert os.path.exists(os.path.join(archive_dir, '银行流水总表.xlsx'))
        assert os.path.exists(os.path.join(archive_dir, 'manifest.json'))


class TestDownloadTemplates:
    """测试下载脚本模板"""

    def test_lock_file_context_manager(self, tmp_dir):
        """测试锁定文件上下文管理器"""
        from bank_download_templates import LockFile

        output_dir = os.path.join(tmp_dir, 'inbox')
        os.makedirs(output_dir, exist_ok=True)

        with LockFile(output_dir, 'test.xlsx'):
            lock_file = os.path.join(output_dir, '.lock', 'test.xlsx.lock')
            assert os.path.exists(lock_file)

        assert not os.path.exists(lock_file)

    def test_generate_filename(self):
        """测试生成文件名"""
        from bank_download_templates import _generate_filename

        filename = _generate_filename('北京银行')
        assert filename.startswith('北京银行_')
        assert filename.endswith('.xlsx')
        parts = filename.replace('.xlsx', '').split('_')
        assert len(parts) == 3

    def test_get_date_range(self):
        """测试日期范围解析"""
        from bank_download_templates import _get_date_range

        from datetime import datetime, timedelta
        today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)

        start, end = _get_date_range('today')
        assert start == today
        assert end == today

        start, end = _get_date_range('yesterday')
        yesterday = today - timedelta(days=1)
        assert start == yesterday
        assert end == yesterday

        start, end = _get_date_range('last_7_days')
        assert start == today - timedelta(days=7)
        assert end == today - timedelta(days=1)

    def test_run_download_template_mock(self, tmp_dir):
        """测试模拟下载模板"""
        from bank_download_templates import run_download_template

        output_dir = os.path.join(tmp_dir, 'inbox')
        os.makedirs(output_dir, exist_ok=True)

        success, message = run_download_template(
            'mock_beijing_bank',
            output_dir=output_dir,
        )

        assert success is True
        assert '成功' in message

        files = os.listdir(output_dir)
        excel_files = [f for f in files if f.endswith('.xlsx') and not f.startswith('.')]
        assert len(excel_files) == 1
        assert excel_files[0].startswith('北京银行_')

    def test_run_download_template_not_found(self, tmp_dir):
        """测试不存在的模板"""
        from bank_download_templates import run_download_template

        output_dir = os.path.join(tmp_dir, 'inbox')
        os.makedirs(output_dir, exist_ok=True)

        success, message = run_download_template(
            'non_existent_template',
            output_dir=output_dir,
        )

        assert success is False
        assert '未找到' in message


class TestBankcheckIntegration:
    """测试与 bankcheck 的集成"""

    def test_run_directory_pipeline(self, tmp_dir, monkeypatch):
        """测试 bankcheck 中的目录对接集成函数"""
        import bankcheck
        import yaml

        script_dir = os.path.join(tmp_dir, 'script')
        os.makedirs(script_dir, exist_ok=True)
        _create_lookup_table(os.path.join(script_dir, '主体查找表.xlsx'))

        config_path = os.path.join(tmp_dir, 'test_config.yaml')
        shutil.copy(
            os.path.join(os.path.dirname(__file__), '..', 'backend', 'bank_directories.yaml'),
            config_path
        )

        with open(config_path, 'r', encoding='utf-8') as f:
            config_data = yaml.safe_load(f)
        config_data['global']['file_stable_seconds'] = 0
        with open(config_path, 'w', encoding='utf-8') as f:
            yaml.dump(config_data, f, allow_unicode=True)

        bank_root = os.path.join(tmp_dir, 'bank_data')
        monkeypatch.setenv('BANK_ROOT', bank_root)
        inbox_dir = os.path.join(bank_root, 'inbox')
        os.makedirs(inbox_dir, exist_ok=True)
        os.makedirs(os.path.join(inbox_dir, '.lock'), exist_ok=True)

        _create_beijing_bank_excel(os.path.join(inbox_dir, '北京银行_20260612_093000.xlsx'))

        result = bankcheck.run_directory_pipeline(
            script_dir=script_dir,
            config_path=config_path,
            incremental=False,
        )

        assert result.success is True
        assert len(result.processed_files) >= 1

    def test_get_directory_status(self, tmp_dir, monkeypatch):
        """测试获取目录状态"""
        import bankcheck

        config_path = os.path.join(tmp_dir, 'test_config.yaml')
        shutil.copy(
            os.path.join(os.path.dirname(__file__), '..', 'backend', 'bank_directories.yaml'),
            config_path
        )

        bank_root = os.path.join(tmp_dir, 'bank_data')
        monkeypatch.setenv('BANK_ROOT', bank_root)

        status = bankcheck.get_directory_status(config_path=config_path)

        assert status is not None
        assert 'inbox_count' in status
        assert 'outbox_count' in status
        assert 'error_count' in status
        assert 'configured_banks' in status
        assert len(status['configured_banks']) >= 5


class TestCronAndSystemd:
    """测试 cron 和 systemd 脚本生成"""

    def test_create_cron_script(self):
        """测试生成 crontab 脚本"""
        from bank_download_templates import create_cron_script

        script = create_cron_script(
            bank_name='北京银行',
            template_name='beijing_bank',
            schedule='0 9 * * 1-5',
            config_path='/path/to/bank_directories.yaml',
            log_file='/var/log/bank_download.log',
            bank_root='/data/bank',
        )

        assert '北京银行' in script
        assert '0 9 * * 1-5' in script
        assert 'bank_directory_connector.py download --bank 北京银行' in script
        assert 'BANK_ROOT=/data/bank' in script

    def test_create_systemd_service(self):
        """测试生成 systemd service 文件"""
        from bank_download_templates import create_systemd_service

        service = create_systemd_service(
            service_name='bank-watcher',
            config_path='/opt/bankcheck/bank_directories.yaml',
            bank_root='/data/bank',
            user='bankuser',
        )

        assert '[Unit]' in service
        assert '[Service]' in service
        assert '[Install]' in service
        assert 'bank_directory_connector.py watch' in service
        assert 'User=bankuser' in service
        assert 'BANK_ROOT=/data/bank' in service
