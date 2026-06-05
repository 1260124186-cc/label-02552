import os

import pytest

from bankcheck import identify_bank


class TestIdentifyBank:
    def test_beijing_bank_prefix(self):
        assert identify_bank('/some/path/北京银行_流水.xlsx') == '北京银行'

    def test_beijing_bank_prefix_with_spaces(self):
        assert identify_bank('/some/path/北京银行流水.xlsx') == '北京银行'

    def test_east_asia_bank_prefix(self):
        assert identify_bank('/some/path/东亚银行_流水.xlsx') == '东亚银行'

    def test_east_asia_bank_prefix_with_spaces(self):
        assert identify_bank('/some/path/东亚银行流水.xlsx') == '东亚银行'

    def test_unknown_bank(self):
        assert identify_bank('/some/path/招商银行_流水.xlsx') is None

    def test_no_bank_name(self):
        assert identify_bank('/some/path/流水文件.xlsx') is None

    def test_bank_name_in_middle(self):
        assert identify_bank('/some/path/2024年北京银行流水.xlsx') is None

    def test_empty_filename(self):
        assert identify_bank('/some/path/.xlsx') is None

    def test_beijing_bank_not_east_asia(self):
        result = identify_bank('/some/path/北京银行东亚银行_流水.xlsx')
        assert result == '北京银行'

    def test_case_sensitivity(self):
        assert identify_bank('/some/path/东亚银行.xlsx') == '东亚银行'
