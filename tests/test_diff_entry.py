import os
import sys
import tempfile
import shutil
from io import StringIO
from unittest import mock

import pandas as pd
import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'backend'))

import bankcheck
from conftest import _create_beijing_bank_excel, _create_lookup_table


SUMMARY_COLUMNS = [
    '唯一id', '银行', '银行账号', '主体', '交易日期',
    '付款', '收款', '摘要', '对方户名', '余额', '交易流水号',
]


def _make_summary_data(rows):
    return pd.DataFrame(rows, columns=SUMMARY_COLUMNS)


def _write_summary(df, path):
    df.to_excel(path, index=False, engine='openpyxl')
    return path


class TestCLIAskFile:
    def test_valid_file(self, tmp_path):
        test_file = tmp_path / 'test.xlsx'
        test_file.touch()
        with mock.patch('builtins.input', return_value=str(test_file)):
            result = bankcheck.cli_askfile('请选择文件')
            assert result == str(test_file)

    def test_invalid_file(self, tmp_path):
        with mock.patch('builtins.input', return_value=str(tmp_path / 'nonexistent.xlsx')):
            result = bankcheck.cli_askfile('请选择文件')
            assert result == ''

    def test_strips_quotes(self, tmp_path):
        test_file = tmp_path / 'test.xlsx'
        test_file.touch()
        with mock.patch('builtins.input', return_value=f'"{test_file}"'):
            result = bankcheck.cli_askfile('请选择文件')
            assert result == str(test_file)


class TestCLIAskMode:
    def test_default_is_pipeline(self):
        with mock.patch('builtins.input', return_value=''):
            assert bankcheck.cli_askmode() == 'pipeline'

    def test_choice_1_is_pipeline(self):
        with mock.patch('builtins.input', return_value='1'):
            assert bankcheck.cli_askmode() == 'pipeline'

    def test_choice_2_is_diff(self):
        with mock.patch('builtins.input', return_value='2'):
            assert bankcheck.cli_askmode() == 'diff'

    def test_invalid_choice_defaults_to_pipeline(self):
        with mock.patch('builtins.input', return_value='invalid'):
            assert bankcheck.cli_askmode() == 'pipeline'


class TestGUIAskMode:
    def test_yes_returns_pipeline(self):
        with mock.patch.object(bankcheck, 'tk', create=True), \
             mock.patch.object(bankcheck, 'messagebox', create=True), \
             mock.patch.object(bankcheck.messagebox, 'askyesnocancel', return_value=True):
            assert bankcheck.gui_askmode() == 'pipeline'

    def test_no_returns_diff(self):
        with mock.patch.object(bankcheck, 'tk', create=True), \
             mock.patch.object(bankcheck, 'messagebox', create=True), \
             mock.patch.object(bankcheck.messagebox, 'askyesnocancel', return_value=False):
            assert bankcheck.gui_askmode() == 'diff'

    def test_cancel_returns_export(self):
        with mock.patch.object(bankcheck, 'tk', create=True), \
             mock.patch.object(bankcheck, 'messagebox', create=True), \
             mock.patch.object(bankcheck.messagebox, 'askyesnocancel', return_value=None):
            assert bankcheck.gui_askmode() == 'export'


class TestRunPipelineFlow:
    def test_calls_run_pipeline(self, tmp_path):
        script_dir = tmp_path / 'script'
        script_dir.mkdir()
        source_folder = tmp_path / '流水'
        source_folder.mkdir()
        _create_beijing_bank_excel(source_folder / '北京银行_流水.xlsx')
        _create_lookup_table(script_dir / '主体查找表.xlsx')

        with mock.patch.object(bankcheck, 'ask_directory', return_value=str(source_folder)), \
             mock.patch.object(bankcheck, 'ask_incremental_mode', return_value=True), \
             mock.patch.object(bankcheck, 'show_info') as mock_show:
            bankcheck.run_pipeline_flow(str(script_dir))
            assert mock_show.called
            assert '处理完成' in mock_show.call_args[0][1]

    def test_no_folder_selected(self, tmp_path):
        script_dir = tmp_path / 'script'
        script_dir.mkdir()
        with mock.patch.object(bankcheck, 'ask_directory', return_value=''), \
             mock.patch.object(bankcheck, 'show_info') as mock_show:
            bankcheck.run_pipeline_flow(str(script_dir))
            assert mock_show.called
            assert '未选择文件夹' in mock_show.call_args[0][1]


class TestRunDiffFlow:
    def test_runs_diff_successfully(self, tmp_path):
        script_dir = tmp_path / 'script'
        script_dir.mkdir()

        rows_old = [
            {
                '唯一id': '1', '银行': '北京银行', '银行账号': '01090312345678901',
                '主体': 'A公司', '交易日期': '2024-01-05', '付款': -50000.0, '收款': None,
                '摘要': '付款', '对方户名': '供应商', '余额': 1500000, '交易流水号': 'BJ001',
            },
            {
                '唯一id': '2', '银行': '北京银行', '银行账号': '01090312345678901',
                '主体': 'A公司', '交易日期': '2024-01-10', '付款': None, '收款': 80000.0,
                '摘要': '收款', '对方户名': '客户', '余额': 1580000, '交易流水号': 'BJ002',
            },
        ]
        rows_new = [
            {
                '唯一id': '1', '银行': '北京银行', '银行账号': '01090312345678901',
                '主体': 'A公司', '交易日期': '2024-01-05', '付款': -50000.0, '收款': None,
                '摘要': '付款', '对方户名': '供应商', '余额': 1500000, '交易流水号': 'BJ001',
            },
            {
                '唯一id': '3', '银行': '东亚银行', '银行账号': '38812345678',
                '主体': 'B公司', '交易日期': '2024-01-15', '付款': None, '收款': 30000.0,
                '摘要': '服务费', '对方户名': '客户C', '余额': 545000, '交易流水号': 'EA003',
            },
        ]

        old_path = tmp_path / 'old.xlsx'
        new_path = tmp_path / 'new.xlsx'
        _write_summary(_make_summary_data(rows_old), str(old_path))
        _write_summary(_make_summary_data(rows_new), str(new_path))

        with mock.patch.object(bankcheck, 'ask_file', side_effect=[str(old_path), str(new_path)]), \
             mock.patch.object(bankcheck, 'show_info') as mock_show:
            bankcheck.run_diff_flow(str(script_dir))
            assert mock_show.called
            title = mock_show.call_args[0][0]
            msg = mock_show.call_args[0][1]
            assert '发现差异' in title
            assert '新增交易：1' in msg
            assert '删除交易：1' in msg

    def test_no_old_file_selected(self, tmp_path):
        script_dir = tmp_path / 'script'
        script_dir.mkdir()
        with mock.patch.object(bankcheck, 'ask_file', return_value=''), \
             mock.patch.object(bankcheck, 'show_info') as mock_show:
            bankcheck.run_diff_flow(str(script_dir))
            assert mock_show.called
            assert '未选择旧批次文件' in mock_show.call_args[0][1]

    def test_no_new_file_selected(self, tmp_path):
        script_dir = tmp_path / 'script'
        script_dir.mkdir()
        old_path = tmp_path / 'old.xlsx'
        old_path.touch()
        with mock.patch.object(bankcheck, 'ask_file', side_effect=[str(old_path), '']), \
             mock.patch.object(bankcheck, 'show_info') as mock_show:
            bankcheck.run_diff_flow(str(script_dir))
            assert mock_show.called
            assert '未选择新批次文件' in mock_show.call_args[0][1]

    def test_shows_no_diff_when_identical(self, tmp_path):
        script_dir = tmp_path / 'script'
        script_dir.mkdir()

        rows = [
            {
                '唯一id': '1', '银行': '北京银行', '银行账号': '01090312345678901',
                '主体': 'A公司', '交易日期': '2024-01-05', '付款': -50000.0, '收款': None,
                '摘要': '付款', '对方户名': '供应商', '余额': 1500000, '交易流水号': 'BJ001',
            },
        ]

        old_path = tmp_path / 'old.xlsx'
        new_path = tmp_path / 'new.xlsx'
        _write_summary(_make_summary_data(rows), str(old_path))
        _write_summary(_make_summary_data(rows), str(new_path))

        with mock.patch.object(bankcheck, 'ask_file', side_effect=[str(old_path), str(new_path)]), \
             mock.patch.object(bankcheck, 'show_info') as mock_show:
            bankcheck.run_diff_flow(str(script_dir))
            assert mock_show.called
            title = mock_show.call_args[0][0]
            assert '无差异' in title

    def test_handles_file_not_found(self, tmp_path):
        script_dir = tmp_path / 'script'
        script_dir.mkdir()
        old_path = tmp_path / 'old.xlsx'
        new_path = tmp_path / 'nonexistent.xlsx'
        old_path.touch()

        with mock.patch.object(bankcheck, 'ask_file', side_effect=[str(old_path), str(new_path)]), \
             mock.patch.object(bankcheck, 'show_warning') as mock_warn:
            bankcheck.run_diff_flow(str(script_dir))
            assert mock_warn.called
            assert '不存在' in mock_warn.call_args[0][1]
