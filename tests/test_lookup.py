import os

import pytest

from conftest import _create_lookup_table
from bankcheck import find_lookup_file, get_subject


class TestFindLookupFile:
    def test_exact_match_xlsx(self, tmp_dir):
        path = os.path.join(tmp_dir, '主体查找表.xlsx')
        _create_lookup_table(path)
        assert find_lookup_file(tmp_dir) == path

    def test_exact_match_xls(self, tmp_dir):
        path = os.path.join(tmp_dir, '主体查找表.xls')
        _create_lookup_table(path)
        result = find_lookup_file(tmp_dir)
        assert result == path

    def test_fallback_single_excel(self, tmp_dir):
        path = os.path.join(tmp_dir, '映射表.xlsx')
        _create_lookup_table(path)
        result = find_lookup_file(tmp_dir)
        assert result == path

    def test_no_excel_files(self, tmp_dir):
        result = find_lookup_file(tmp_dir)
        assert result is None

    def test_multiple_excel_files_ambiguous(self, tmp_dir):
        _create_lookup_table(os.path.join(tmp_dir, '表1.xlsx'))
        _create_lookup_table(os.path.join(tmp_dir, '表2.xlsx'))
        result = find_lookup_file(tmp_dir)
        assert result is None

    def test_excludes_output_table(self, tmp_dir):
        output_path = os.path.join(tmp_dir, '银行流水总表.xlsx')
        _create_lookup_table(output_path)
        lookup_path = os.path.join(tmp_dir, '主体查找表.xlsx')
        _create_lookup_table(lookup_path)
        result = find_lookup_file(tmp_dir)
        assert result == lookup_path

    def test_excludes_temp_files(self, tmp_dir):
        open(os.path.join(tmp_dir, '~$temp.xlsx'), 'w').close()
        path = os.path.join(tmp_dir, '映射表.xlsx')
        _create_lookup_table(path)
        result = find_lookup_file(tmp_dir)
        assert result == path


class TestGetSubject:
    def test_found(self, tmp_dir):
        path = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        result = get_subject('01090312345678901', path)
        assert result == '北京XX科技有限公司'

    def test_found_east_asia(self, tmp_dir):
        path = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        result = get_subject('38812345678', path)
        assert result == '上海YY贸易有限公司'

    def test_not_found(self, tmp_dir):
        path = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        result = get_subject('99999999999', path)
        assert result == ''

    def test_none_account(self, tmp_dir):
        path = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        result = get_subject(None, path)
        assert result == ''

    def test_empty_account(self, tmp_dir):
        path = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        result = get_subject('', path)
        assert result == ''

    def test_none_lookup_file(self):
        result = get_subject('01090312345678901', None)
        assert result == ''

    def test_nonexistent_lookup_file(self):
        result = get_subject('01090312345678901', '/nonexistent/path.xlsx')
        assert result == ''

    def test_account_with_spaces(self, tmp_dir):
        path = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'), [
            ('测试公司', ' 12345 '),
        ])
        result = get_subject('12345', path)
        assert result == '测试公司'
