import os

import pytest

from conftest import BANK_TEST_CONFIGS, _create_bank_excel, _create_lookup_table
import bankcheck


_BANK_NAMES = list(BANK_TEST_CONFIGS.keys())


def _get_processor(bank_name):
    return bankcheck.BANK_PROCESSORS.get(bank_name)


@pytest.mark.parametrize('bank_name', _BANK_NAMES, ids=_BANK_NAMES)
class TestBankProcessor:
    def test_basic_extraction(self, tmp_dir, bank_name):
        cfg = BANK_TEST_CONFIGS[bank_name]
        filepath = _create_bank_excel(os.path.join(tmp_dir, f'{bank_name}_流水.xlsx'), bank_name)
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = _get_processor(bank_name)(filepath, lookup)
        assert len(rows) == cfg['expected']['row_count']

    def test_bank_name(self, tmp_dir, bank_name):
        cfg = BANK_TEST_CONFIGS[bank_name]
        filepath = _create_bank_excel(os.path.join(tmp_dir, f'{bank_name}_流水.xlsx'), bank_name)
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = _get_processor(bank_name)(filepath, lookup)
        for row in rows:
            assert row['银行'] == cfg['expected']['bank_name']

    def test_bank_account(self, tmp_dir, bank_name):
        cfg = BANK_TEST_CONFIGS[bank_name]
        filepath = _create_bank_excel(os.path.join(tmp_dir, f'{bank_name}_流水.xlsx'), bank_name)
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = _get_processor(bank_name)(filepath, lookup)
        for row in rows:
            assert row['银行账号'] == cfg['expected']['account']

    def test_subject_lookup(self, tmp_dir, bank_name):
        cfg = BANK_TEST_CONFIGS[bank_name]
        filepath = _create_bank_excel(os.path.join(tmp_dir, f'{bank_name}_流水.xlsx'), bank_name)
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = _get_processor(bank_name)(filepath, lookup)
        assert rows[0]['主体'] == cfg['expected']['subject']

    def test_payment_negative(self, tmp_dir, bank_name):
        cfg = BANK_TEST_CONFIGS[bank_name]
        filepath = _create_bank_excel(os.path.join(tmp_dir, f'{bank_name}_流水.xlsx'), bank_name)
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = _get_processor(bank_name)(filepath, lookup)
        assert rows[0]['付款'] == cfg['expected']['first_payment']

    def test_receipt_positive(self, tmp_dir, bank_name):
        cfg = BANK_TEST_CONFIGS[bank_name]
        filepath = _create_bank_excel(os.path.join(tmp_dir, f'{bank_name}_流水.xlsx'), bank_name)
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = _get_processor(bank_name)(filepath, lookup)
        assert rows[1]['收款'] == cfg['expected']['second_receipt']

    def test_no_payment_gives_none(self, tmp_dir, bank_name):
        cfg = BANK_TEST_CONFIGS[bank_name]
        filepath = _create_bank_excel(os.path.join(tmp_dir, f'{bank_name}_流水.xlsx'), bank_name)
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = _get_processor(bank_name)(filepath, lookup)
        assert rows[1]['付款'] == cfg['expected']['second_payment']

    def test_no_receipt_gives_none(self, tmp_dir, bank_name):
        cfg = BANK_TEST_CONFIGS[bank_name]
        filepath = _create_bank_excel(os.path.join(tmp_dir, f'{bank_name}_流水.xlsx'), bank_name)
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = _get_processor(bank_name)(filepath, lookup)
        assert rows[0]['收款'] == cfg['expected']['first_receipt']

    def test_trade_date(self, tmp_dir, bank_name):
        cfg = BANK_TEST_CONFIGS[bank_name]
        filepath = _create_bank_excel(os.path.join(tmp_dir, f'{bank_name}_流水.xlsx'), bank_name)
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = _get_processor(bank_name)(filepath, lookup)
        expected_dates = cfg['expected']['trade_dates']
        for i, expected in enumerate(expected_dates):
            assert rows[i]['交易日期'] == expected

    def test_summary(self, tmp_dir, bank_name):
        cfg = BANK_TEST_CONFIGS[bank_name]
        filepath = _create_bank_excel(os.path.join(tmp_dir, f'{bank_name}_流水.xlsx'), bank_name)
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = _get_processor(bank_name)(filepath, lookup)
        assert rows[0]['摘要'] == cfg['expected']['first_summary']

    def test_counterpart(self, tmp_dir, bank_name):
        cfg = BANK_TEST_CONFIGS[bank_name]
        filepath = _create_bank_excel(os.path.join(tmp_dir, f'{bank_name}_流水.xlsx'), bank_name)
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = _get_processor(bank_name)(filepath, lookup)
        exp = cfg['expected']
        if exp.get('counterpart_same_as_summary'):
            assert rows[0]['对方户名'] == rows[0]['摘要']
        else:
            assert rows[0]['对方户名'] == exp['first_counterpart']
            assert rows[1]['对方户名'] == exp.get('second_counterpart')

    def test_balance(self, tmp_dir, bank_name):
        cfg = BANK_TEST_CONFIGS[bank_name]
        filepath = _create_bank_excel(os.path.join(tmp_dir, f'{bank_name}_流水.xlsx'), bank_name)
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = _get_processor(bank_name)(filepath, lookup)
        assert rows[0]['余额'] == cfg['expected']['first_balance']
        assert rows[1]['余额'] == cfg['expected']['second_balance']

    def test_transaction_id(self, tmp_dir, bank_name):
        cfg = BANK_TEST_CONFIGS[bank_name]
        filepath = _create_bank_excel(os.path.join(tmp_dir, f'{bank_name}_流水.xlsx'), bank_name)
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = _get_processor(bank_name)(filepath, lookup)
        expected_ids = cfg['expected']['transaction_ids']
        for i, expected in enumerate(expected_ids):
            assert rows[i]['交易流水号'] == expected

    def test_unique_id_present(self, tmp_dir, bank_name):
        filepath = _create_bank_excel(os.path.join(tmp_dir, f'{bank_name}_流水.xlsx'), bank_name)
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = _get_processor(bank_name)(filepath, lookup)
        assert rows[0]['唯一id'] != rows[1]['唯一id']
        assert len(rows[0]['唯一id']) > 0

    def test_no_lookup_file(self, tmp_dir, bank_name):
        filepath = _create_bank_excel(os.path.join(tmp_dir, f'{bank_name}_流水.xlsx'), bank_name)
        rows = _get_processor(bank_name)(filepath, None)
        assert rows[0]['主体'] == ''

    def test_empty_account(self, tmp_dir, bank_name):
        filepath = _create_bank_excel(
            os.path.join(tmp_dir, f'{bank_name}_流水.xlsx'), bank_name, account=None
        )
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = _get_processor(bank_name)(filepath, lookup)
        cfg = BANK_TEST_CONFIGS[bank_name]
        assert len(rows) == cfg['expected']['row_count']
        assert rows[0]['主体'] == ''

    def test_skip_empty_date_rows(self, tmp_dir, bank_name):
        cfg = BANK_TEST_CONFIGS[bank_name]
        filepath = _create_bank_excel(
            os.path.join(tmp_dir, f'{bank_name}_流水.xlsx'), bank_name,
            rows=cfg['skip_empty_date_rows'],
        )
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = _get_processor(bank_name)(filepath, lookup)
        assert len(rows) == 2

    def test_all_columns_present(self, tmp_dir, bank_name):
        cfg = BANK_TEST_CONFIGS[bank_name]
        filepath = _create_bank_excel(os.path.join(tmp_dir, f'{bank_name}_流水.xlsx'), bank_name)
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = _get_processor(bank_name)(filepath, lookup)
        assert set(rows[0].keys()) == cfg['expected']['expected_keys']
