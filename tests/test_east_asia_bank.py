import os

import pytest

from conftest import _create_east_asia_bank_excel, _create_lookup_table
from bankcheck import process_east_asia_bank


class TestProcessEastAsiaBank:
    def test_basic_extraction(self, tmp_dir):
        filepath = _create_east_asia_bank_excel(os.path.join(tmp_dir, '东亚银行_流水.xlsx'))
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = process_east_asia_bank(filepath, lookup)
        assert len(rows) == 2

    def test_bank_name(self, tmp_dir):
        filepath = _create_east_asia_bank_excel(os.path.join(tmp_dir, '东亚银行_流水.xlsx'))
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = process_east_asia_bank(filepath, lookup)
        assert rows[0]['银行'] == '东亚银行'
        assert rows[1]['银行'] == '东亚银行'

    def test_bank_account(self, tmp_dir):
        filepath = _create_east_asia_bank_excel(os.path.join(tmp_dir, '东亚银行_流水.xlsx'))
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = process_east_asia_bank(filepath, lookup)
        assert rows[0]['银行账号'] == '38812345678'

    def test_subject_lookup(self, tmp_dir):
        filepath = _create_east_asia_bank_excel(os.path.join(tmp_dir, '东亚银行_流水.xlsx'))
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = process_east_asia_bank(filepath, lookup)
        assert rows[0]['主体'] == '上海YY贸易有限公司'

    def test_payment_negative(self, tmp_dir):
        filepath = _create_east_asia_bank_excel(os.path.join(tmp_dir, '东亚银行_流水.xlsx'))
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = process_east_asia_bank(filepath, lookup)
        assert rows[0]['付款'] == -20000.0

    def test_receipt_positive(self, tmp_dir):
        filepath = _create_east_asia_bank_excel(os.path.join(tmp_dir, '东亚银行_流水.xlsx'))
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = process_east_asia_bank(filepath, lookup)
        assert rows[1]['收款'] == 35000.0

    def test_no_payment_gives_none(self, tmp_dir):
        filepath = _create_east_asia_bank_excel(os.path.join(tmp_dir, '东亚银行_流水.xlsx'))
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = process_east_asia_bank(filepath, lookup)
        assert rows[1]['付款'] is None

    def test_no_receipt_gives_none(self, tmp_dir):
        filepath = _create_east_asia_bank_excel(os.path.join(tmp_dir, '东亚银行_流水.xlsx'))
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = process_east_asia_bank(filepath, lookup)
        assert rows[0]['收款'] is None

    def test_trade_date(self, tmp_dir):
        filepath = _create_east_asia_bank_excel(os.path.join(tmp_dir, '东亚银行_流水.xlsx'))
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = process_east_asia_bank(filepath, lookup)
        assert rows[0]['交易日期'] == '2024-01-03'
        assert rows[1]['交易日期'] == '2024-01-08'

    def test_summary(self, tmp_dir):
        filepath = _create_east_asia_bank_excel(os.path.join(tmp_dir, '东亚银行_流水.xlsx'))
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = process_east_asia_bank(filepath, lookup)
        assert rows[0]['摘要'] == '向 张三 付款-货款'

    def test_counterpart_same_as_summary(self, tmp_dir):
        filepath = _create_east_asia_bank_excel(os.path.join(tmp_dir, '东亚银行_流水.xlsx'))
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = process_east_asia_bank(filepath, lookup)
        assert rows[0]['对方户名'] == rows[0]['摘要']

    def test_balance(self, tmp_dir):
        filepath = _create_east_asia_bank_excel(os.path.join(tmp_dir, '东亚银行_流水.xlsx'))
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = process_east_asia_bank(filepath, lookup)
        assert rows[0]['余额'] == 480000
        assert rows[1]['余额'] == 515000

    def test_transaction_id(self, tmp_dir):
        filepath = _create_east_asia_bank_excel(os.path.join(tmp_dir, '东亚银行_流水.xlsx'))
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = process_east_asia_bank(filepath, lookup)
        assert rows[0]['交易流水号'] == 'EA20240103001'
        assert rows[1]['交易流水号'] == 'EA20240108002'

    def test_unique_id_present(self, tmp_dir):
        filepath = _create_east_asia_bank_excel(os.path.join(tmp_dir, '东亚银行_流水.xlsx'))
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = process_east_asia_bank(filepath, lookup)
        assert rows[0]['唯一id'] != rows[1]['唯一id']

    def test_no_lookup_file(self, tmp_dir):
        filepath = _create_east_asia_bank_excel(os.path.join(tmp_dir, '东亚银行_流水.xlsx'))
        rows = process_east_asia_bank(filepath, None)
        assert rows[0]['主体'] == ''

    def test_all_columns_present(self, tmp_dir):
        filepath = _create_east_asia_bank_excel(os.path.join(tmp_dir, '东亚银行_流水.xlsx'))
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = process_east_asia_bank(filepath, lookup)
        expected_keys = {'唯一id', '银行', '银行账号', '主体', '交易日期', '付款', '收款', '摘要', '对方户名', '余额', '交易流水号'}
        assert set(rows[0].keys()) == expected_keys

    def test_skip_empty_date_rows(self, tmp_dir):
        rows_data = [
            ['2024-01-03', '09:30:00', 'CNY', 20000, None, 100, 0, 0, 480000, '转账', 'EA001', '付款'],
            [None, None, None, None, None, None, None, None, None, None, None, None],
            ['2024-01-08', '14:15:00', 'CNY', None, 35000, 0, 0, 0, 515000, '转账', 'EA002', '收款'],
        ]
        filepath = _create_east_asia_bank_excel(
            os.path.join(tmp_dir, '东亚银行_流水.xlsx'), rows=rows_data
        )
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = process_east_asia_bank(filepath, lookup)
        assert len(rows) == 2

    def test_empty_account(self, tmp_dir):
        filepath = _create_east_asia_bank_excel(
            os.path.join(tmp_dir, '东亚银行_流水.xlsx'), account=None
        )
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = process_east_asia_bank(filepath, lookup)
        assert len(rows) == 2
