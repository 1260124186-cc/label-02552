import os

import pytest

from conftest import _create_lookup_table
from bankcheck import find_lookup_file, get_subject, _normalize_account_str, _account_key


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

    def test_int_account_matches_string_in_lookup(self, tmp_dir):
        path = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'), [
            ('北京XX科技有限公司', '01090312345678901'),
        ])
        result = get_subject(1090312345678901, path)
        assert result == '北京XX科技有限公司'

    def test_float_account_matches_string_in_lookup(self, tmp_dir):
        path = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'), [
            ('北京XX科技有限公司', '01090312345678901'),
        ])
        result = get_subject(1090312345678901.0, path)
        assert result == '北京XX科技有限公司'

    def test_string_account_matches_int_in_lookup(self, tmp_dir):
        path = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'), [
            ('北京XX科技有限公司', 1090312345678901),
        ])
        result = get_subject('01090312345678901', path)
        assert result == '北京XX科技有限公司'

    def test_string_account_matches_float_in_lookup(self, tmp_dir):
        path = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'), [
            ('北京XX科技有限公司', 1090312345678901.0),
        ])
        result = get_subject('01090312345678901', path)
        assert result == '北京XX科技有限公司'

    def test_int_account_matches_int_in_lookup(self, tmp_dir):
        path = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'), [
            ('上海YY贸易有限公司', 38812345678),
        ])
        result = get_subject(38812345678, path)
        assert result == '上海YY贸易有限公司'

    def test_float_account_matches_float_in_lookup(self, tmp_dir):
        path = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'), [
            ('上海YY贸易有限公司', 38812345678.0),
        ])
        result = get_subject(38812345678.0, path)
        assert result == '上海YY贸易有限公司'

    def test_string_with_leading_zero_matches_int(self, tmp_dir):
        path = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'), [
            ('测试公司A', 123456789),
        ])
        result = get_subject('0123456789', path)
        assert result == '测试公司A'

    def test_int_matches_string_with_leading_zero(self, tmp_dir):
        path = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'), [
            ('测试公司B', '0123456789'),
        ])
        result = get_subject(123456789, path)
        assert result == '测试公司B'

    def test_string_dot_zero_matches_int_in_lookup(self, tmp_dir):
        path = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'), [
            ('测试公司C', 38812345678),
        ])
        result = get_subject('38812345678.0', path)
        assert result == '测试公司C'

    def test_no_false_positive_on_different_accounts(self, tmp_dir):
        path = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'), [
            ('公司A', '1234567890'),
            ('公司B', '01234567890'),
        ])
        result = get_subject('01234567890', path)
        assert result == '公司A' or result == '公司B'


class TestNormalizeAccountStr:
    def test_none(self):
        assert _normalize_account_str(None) == ''

    def test_int(self):
        assert _normalize_account_str(123456789) == '123456789'

    def test_float_whole(self):
        assert _normalize_account_str(123456789.0) == '123456789'

    def test_float_with_decimals(self):
        assert _normalize_account_str(123.45) == '123.45'

    def test_float_nan(self):
        assert _normalize_account_str(float('nan')) == ''

    def test_string_digits(self):
        assert _normalize_account_str('01090312345678901') == '01090312345678901'

    def test_string_dot_zero(self):
        assert _normalize_account_str('38812345678.0') == '38812345678'

    def test_string_with_spaces(self):
        assert _normalize_account_str('  12345  ') == '12345'

    def test_empty_string(self):
        assert _normalize_account_str('') == ''

    def test_large_int(self):
        assert _normalize_account_str(1090312345678901) == '1090312345678901'


class TestAccountKey:
    def test_string_with_leading_zero(self):
        assert _account_key('01090312345678901') == '1090312345678901'

    def test_int_without_leading_zero(self):
        assert _account_key(1090312345678901) == '1090312345678901'

    def test_float_without_leading_zero(self):
        assert _account_key(1090312345678901.0) == '1090312345678901'

    def test_both_equal(self):
        assert _account_key('01090312345678901') == _account_key(1090312345678901)

    def test_zero_account(self):
        assert _account_key(0) == '0'

    def test_zero_string(self):
        assert _account_key('0') == '0'

    def test_none(self):
        assert _account_key(None) == '0'
