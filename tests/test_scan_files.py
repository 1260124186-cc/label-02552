import os

import pytest

from bankcheck import scan_excel_files


class TestScanExcelFiles:
    def test_finds_xlsx_files(self, tmp_dir):
        open(os.path.join(tmp_dir, 'test.xlsx'), 'w').close()
        result = scan_excel_files(tmp_dir)
        assert len(result) == 1
        assert result[0].endswith('test.xlsx')

    def test_finds_xls_files(self, tmp_dir):
        open(os.path.join(tmp_dir, 'test.xls'), 'w').close()
        result = scan_excel_files(tmp_dir)
        assert len(result) == 1
        assert result[0].endswith('test.xls')

    def test_finds_both_xlsx_and_xls(self, tmp_dir):
        open(os.path.join(tmp_dir, 'a.xlsx'), 'w').close()
        open(os.path.join(tmp_dir, 'b.xls'), 'w').close()
        result = scan_excel_files(tmp_dir)
        assert len(result) == 2

    def test_excludes_temp_files(self, tmp_dir):
        open(os.path.join(tmp_dir, '~$temp.xlsx'), 'w').close()
        open(os.path.join(tmp_dir, 'real.xlsx'), 'w').close()
        result = scan_excel_files(tmp_dir)
        assert len(result) == 1
        assert 'real.xlsx' in result[0]

    def test_excludes_non_excel_files(self, tmp_dir):
        open(os.path.join(tmp_dir, 'data.csv'), 'w').close()
        open(os.path.join(tmp_dir, 'notes.txt'), 'w').close()
        open(os.path.join(tmp_dir, 'report.pdf'), 'w').close()
        result = scan_excel_files(tmp_dir)
        assert len(result) == 0

    def test_recursive_scan(self, tmp_dir):
        sub = os.path.join(tmp_dir, 'subdir')
        os.makedirs(sub)
        open(os.path.join(tmp_dir, 'root.xlsx'), 'w').close()
        open(os.path.join(sub, 'nested.xlsx'), 'w').close()
        result = scan_excel_files(tmp_dir)
        assert len(result) == 2

    def test_empty_folder(self, tmp_dir):
        result = scan_excel_files(tmp_dir)
        assert len(result) == 0

    def test_case_insensitive_extension(self, tmp_dir):
        open(os.path.join(tmp_dir, 'test.XLSX'), 'w').close()
        open(os.path.join(tmp_dir, 'test2.Xls'), 'w').close()
        result = scan_excel_files(tmp_dir)
        assert len(result) == 2
