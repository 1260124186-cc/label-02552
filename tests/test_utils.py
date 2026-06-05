import os
import tempfile

import openpyxl
import pytest

import bankcheck
from bankcheck import open_workbook_compat, cleanup_temp_file


class TestOpenWorkbookCompatXlsx:
    def test_open_xlsx(self, tmp_dir):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws['A1'] = 'test'
        path = os.path.join(tmp_dir, 'test.xlsx')
        wb.save(path)
        wb.close()

        result_wb, tmp_path = open_workbook_compat(path)
        assert tmp_path is None
        assert result_wb.active['A1'].value == 'test'
        result_wb.close()

    def test_no_temp_file_for_xlsx(self, tmp_dir):
        wb = openpyxl.Workbook()
        path = os.path.join(tmp_dir, 'test.xlsx')
        wb.save(path)
        wb.close()

        _, tmp_path = open_workbook_compat(path)
        assert tmp_path is None


class TestCleanupTempFile:
    def test_cleanup_existing_file(self, tmp_dir):
        tmp_path = os.path.join(tmp_dir, 'temp_file.xlsx')
        with open(tmp_path, 'w') as f:
            f.write('test')
        assert os.path.exists(tmp_path)
        cleanup_temp_file(tmp_path)
        assert not os.path.exists(tmp_path)

    def test_cleanup_nonexistent_file(self):
        cleanup_temp_file('/nonexistent/file.xlsx')

    def test_cleanup_none(self):
        cleanup_temp_file(None)


class TestGetScriptDir:
    def test_returns_directory(self):
        result = bankcheck.get_script_dir()
        assert os.path.isdir(result)

    def test_returns_absolute_path(self):
        result = bankcheck.get_script_dir()
        assert os.path.isabs(result)


class TestSetupLogging:
    def test_returns_logger(self):
        logger = bankcheck.setup_logging()
        assert logger is not None
        assert logger.name == 'bankcheck'

    def test_logger_has_handlers(self):
        logger = bankcheck.setup_logging()
        assert len(logger.handlers) >= 1

    def test_get_logger(self):
        logger = bankcheck.get_logger()
        assert logger.name == 'bankcheck'


class TestCLIFunctions:
    def test_cli_showinfo(self, capsys):
        bankcheck.cli_showinfo('Title', 'Message')
        captured = capsys.readouterr()
        assert 'Title' in captured.out
        assert 'Message' in captured.out

    def test_cli_showwarning(self, capsys):
        bankcheck.cli_showwarning('Title', 'Message')
        captured = capsys.readouterr()
        assert 'Title' in captured.out
        assert 'Message' in captured.out

    def test_cli_askdirectory_valid(self, tmp_dir, monkeypatch):
        monkeypatch.setattr('builtins.input', lambda _: tmp_dir)
        result = bankcheck.cli_askdirectory()
        assert result == tmp_dir

    def test_cli_askdirectory_invalid(self, monkeypatch):
        monkeypatch.setattr('builtins.input', lambda _: '/nonexistent/path')
        result = bankcheck.cli_askdirectory()
        assert result == ''

    def test_cli_askdirectory_empty(self, monkeypatch):
        monkeypatch.setattr('builtins.input', lambda _: '')
        result = bankcheck.cli_askdirectory()
        assert result == ''
