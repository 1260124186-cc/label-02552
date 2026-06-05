import os

import pytest

from conftest import _create_beijing_bank_excel, _create_lookup_table
from bankcheck import process_beijing_bank


class TestProcessBeijingBank:
    def test_basic_extraction(self, tmp_dir):
        filepath = _create_beijing_bank_excel(os.path.join(tmp_dir, '北京银行_流水.xlsx'))
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = process_beijing_bank(filepath, lookup)
        assert len(rows) == 2

    def test_bank_name(self, tmp_dir):
        filepath = _create_beijing_bank_excel(os.path.join(tmp_dir, '北京银行_流水.xlsx'))
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = process_beijing_bank(filepath, lookup)
        assert rows[0]['银行'] == '北京银行'
        assert rows[1]['银行'] == '北京银行'

    def test_bank_account(self, tmp_dir):
        filepath = _create_beijing_bank_excel(os.path.join(tmp_dir, '北京银行_流水.xlsx'))
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = process_beijing_bank(filepath, lookup)
        assert rows[0]['银行账号'] == '01090312345678901'
        assert rows[1]['银行账号'] == '01090312345678901'

    def test_subject_lookup(self, tmp_dir):
        filepath = _create_beijing_bank_excel(os.path.join(tmp_dir, '北京银行_流水.xlsx'))
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = process_beijing_bank(filepath, lookup)
        assert rows[0]['主体'] == '北京XX科技有限公司'

    def test_payment_negative(self, tmp_dir):
        filepath = _create_beijing_bank_excel(os.path.join(tmp_dir, '北京银行_流水.xlsx'))
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = process_beijing_bank(filepath, lookup)
        assert rows[0]['付款'] == -50000.0

    def test_receipt_positive(self, tmp_dir):
        filepath = _create_beijing_bank_excel(os.path.join(tmp_dir, '北京银行_流水.xlsx'))
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = process_beijing_bank(filepath, lookup)
        assert rows[1]['收款'] == 80000.0

    def test_no_payment_gives_none(self, tmp_dir):
        filepath = _create_beijing_bank_excel(os.path.join(tmp_dir, '北京银行_流水.xlsx'))
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = process_beijing_bank(filepath, lookup)
        assert rows[1]['付款'] is None

    def test_no_receipt_gives_none(self, tmp_dir):
        filepath = _create_beijing_bank_excel(os.path.join(tmp_dir, '北京银行_流水.xlsx'))
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = process_beijing_bank(filepath, lookup)
        assert rows[0]['收款'] is None

    def test_trade_date(self, tmp_dir):
        filepath = _create_beijing_bank_excel(os.path.join(tmp_dir, '北京银行_流水.xlsx'))
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = process_beijing_bank(filepath, lookup)
        assert rows[0]['交易日期'] == '2024-01-05'
        assert rows[1]['交易日期'] == '2024-01-10'

    def test_summary(self, tmp_dir):
        filepath = _create_beijing_bank_excel(os.path.join(tmp_dir, '北京银行_流水.xlsx'))
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = process_beijing_bank(filepath, lookup)
        assert rows[0]['摘要'] == '采购付款'
        assert rows[1]['摘要'] == '销售收款'

    def test_counterpart(self, tmp_dir):
        filepath = _create_beijing_bank_excel(os.path.join(tmp_dir, '北京银行_流水.xlsx'))
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = process_beijing_bank(filepath, lookup)
        assert rows[0]['对方户名'] == '供应商A公司'
        assert rows[1]['对方户名'] == '客户B公司'

    def test_balance(self, tmp_dir):
        filepath = _create_beijing_bank_excel(os.path.join(tmp_dir, '北京银行_流水.xlsx'))
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = process_beijing_bank(filepath, lookup)
        assert rows[0]['余额'] == 1500000
        assert rows[1]['余额'] == 1580000

    def test_transaction_id(self, tmp_dir):
        filepath = _create_beijing_bank_excel(os.path.join(tmp_dir, '北京银行_流水.xlsx'))
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = process_beijing_bank(filepath, lookup)
        assert rows[0]['交易流水号'] == 'BJ20240105001'
        assert rows[1]['交易流水号'] == 'BJ20240110002'

    def test_unique_id_present(self, tmp_dir):
        filepath = _create_beijing_bank_excel(os.path.join(tmp_dir, '北京银行_流水.xlsx'))
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = process_beijing_bank(filepath, lookup)
        assert rows[0]['唯一id'] != rows[1]['唯一id']
        assert len(rows[0]['唯一id']) > 0

    def test_no_lookup_file(self, tmp_dir):
        filepath = _create_beijing_bank_excel(os.path.join(tmp_dir, '北京银行_流水.xlsx'))
        rows = process_beijing_bank(filepath, None)
        assert rows[0]['主体'] == ''

    def test_empty_account(self, tmp_dir):
        filepath = _create_beijing_bank_excel(
            os.path.join(tmp_dir, '北京银行_流水.xlsx'), account=None
        )
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = process_beijing_bank(filepath, lookup)
        assert len(rows) == 2
        assert rows[0]['主体'] == ''

    def test_skip_empty_date_rows(self, tmp_dir):
        rows_data = [
            [1, '2024-01-05', 'CNY', 50000, None, 1500000, '供应商A', '622001', '工商银行', '转账', '001', '付款', None, None, None, 'BJ001'],
            [None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None],
            [2, '2024-01-10', 'CNY', None, 80000, 1580000, '客户B', '622002', '建设银行', '转账', '002', '收款', None, None, None, 'BJ002'],
        ]
        filepath = _create_beijing_bank_excel(
            os.path.join(tmp_dir, '北京银行_流水.xlsx'), rows=rows_data
        )
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = process_beijing_bank(filepath, lookup)
        assert len(rows) == 2

    def test_all_columns_present(self, tmp_dir):
        filepath = _create_beijing_bank_excel(os.path.join(tmp_dir, '北京银行_流水.xlsx'))
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = process_beijing_bank(filepath, lookup)
        expected_keys = {'唯一id', '银行', '银行账号', '主体', '交易日期', '付款', '收款', '摘要', '对方户名', '余额', '交易流水号'}
        assert set(rows[0].keys()) == expected_keys
