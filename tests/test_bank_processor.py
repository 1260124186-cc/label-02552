import os

import pytest

from conftest import BANK_TEST_CONFIGS, _create_bank_excel, _create_lookup_table_all_banks
import bankcheck


_BANK_NAMES = list(BANK_TEST_CONFIGS.keys())


def _get_processor(bank_name):
    return bankcheck.BANK_PROCESSORS.get(bank_name)


@pytest.mark.parametrize('bank_name', _BANK_NAMES, ids=_BANK_NAMES)
class TestBankProcessor:
    def test_basic_extraction(self, tmp_dir, bank_name):
        cfg = BANK_TEST_CONFIGS[bank_name]
        filepath = _create_bank_excel(os.path.join(tmp_dir, f'{bank_name}_流水.xlsx'), bank_name)
        lookup = _create_lookup_table_all_banks(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = _get_processor(bank_name)(filepath, lookup)
        assert len(rows) == cfg['expected']['row_count']

    def test_bank_name(self, tmp_dir, bank_name):
        cfg = BANK_TEST_CONFIGS[bank_name]
        filepath = _create_bank_excel(os.path.join(tmp_dir, f'{bank_name}_流水.xlsx'), bank_name)
        lookup = _create_lookup_table_all_banks(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = _get_processor(bank_name)(filepath, lookup)
        for row in rows:
            assert row['银行'] == cfg['expected']['bank_name']

    def test_bank_account(self, tmp_dir, bank_name):
        cfg = BANK_TEST_CONFIGS[bank_name]
        filepath = _create_bank_excel(os.path.join(tmp_dir, f'{bank_name}_流水.xlsx'), bank_name)
        lookup = _create_lookup_table_all_banks(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = _get_processor(bank_name)(filepath, lookup)
        for row in rows:
            assert row['银行账号'] == cfg['expected']['account']

    def test_subject_lookup(self, tmp_dir, bank_name):
        cfg = BANK_TEST_CONFIGS[bank_name]
        filepath = _create_bank_excel(os.path.join(tmp_dir, f'{bank_name}_流水.xlsx'), bank_name)
        lookup = _create_lookup_table_all_banks(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = _get_processor(bank_name)(filepath, lookup)
        assert rows[0]['主体'] == cfg['expected']['subject']

    def test_payment_negative(self, tmp_dir, bank_name):
        cfg = BANK_TEST_CONFIGS[bank_name]
        filepath = _create_bank_excel(os.path.join(tmp_dir, f'{bank_name}_流水.xlsx'), bank_name)
        lookup = _create_lookup_table_all_banks(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = _get_processor(bank_name)(filepath, lookup)
        assert rows[0]['付款'] == cfg['expected']['first_payment']

    def test_receipt_positive(self, tmp_dir, bank_name):
        cfg = BANK_TEST_CONFIGS[bank_name]
        filepath = _create_bank_excel(os.path.join(tmp_dir, f'{bank_name}_流水.xlsx'), bank_name)
        lookup = _create_lookup_table_all_banks(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = _get_processor(bank_name)(filepath, lookup)
        assert rows[1]['收款'] == cfg['expected']['second_receipt']

    def test_no_payment_gives_none(self, tmp_dir, bank_name):
        cfg = BANK_TEST_CONFIGS[bank_name]
        filepath = _create_bank_excel(os.path.join(tmp_dir, f'{bank_name}_流水.xlsx'), bank_name)
        lookup = _create_lookup_table_all_banks(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = _get_processor(bank_name)(filepath, lookup)
        assert rows[1]['付款'] == cfg['expected']['second_payment']

    def test_no_receipt_gives_none(self, tmp_dir, bank_name):
        cfg = BANK_TEST_CONFIGS[bank_name]
        filepath = _create_bank_excel(os.path.join(tmp_dir, f'{bank_name}_流水.xlsx'), bank_name)
        lookup = _create_lookup_table_all_banks(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = _get_processor(bank_name)(filepath, lookup)
        assert rows[0]['收款'] == cfg['expected']['first_receipt']

    def test_trade_date(self, tmp_dir, bank_name):
        cfg = BANK_TEST_CONFIGS[bank_name]
        filepath = _create_bank_excel(os.path.join(tmp_dir, f'{bank_name}_流水.xlsx'), bank_name)
        lookup = _create_lookup_table_all_banks(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = _get_processor(bank_name)(filepath, lookup)
        expected_dates = cfg['expected']['trade_dates']
        for i, expected in enumerate(expected_dates):
            assert rows[i]['交易日期'] == expected

    def test_summary(self, tmp_dir, bank_name):
        cfg = BANK_TEST_CONFIGS[bank_name]
        filepath = _create_bank_excel(os.path.join(tmp_dir, f'{bank_name}_流水.xlsx'), bank_name)
        lookup = _create_lookup_table_all_banks(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = _get_processor(bank_name)(filepath, lookup)
        assert rows[0]['摘要'] == cfg['expected']['first_summary']

    def test_counterpart(self, tmp_dir, bank_name):
        cfg = BANK_TEST_CONFIGS[bank_name]
        filepath = _create_bank_excel(os.path.join(tmp_dir, f'{bank_name}_流水.xlsx'), bank_name)
        lookup = _create_lookup_table_all_banks(os.path.join(tmp_dir, '主体查找表.xlsx'))
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
        lookup = _create_lookup_table_all_banks(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = _get_processor(bank_name)(filepath, lookup)
        assert rows[0]['余额'] == cfg['expected']['first_balance']
        assert rows[1]['余额'] == cfg['expected']['second_balance']

    def test_transaction_id(self, tmp_dir, bank_name):
        cfg = BANK_TEST_CONFIGS[bank_name]
        filepath = _create_bank_excel(os.path.join(tmp_dir, f'{bank_name}_流水.xlsx'), bank_name)
        lookup = _create_lookup_table_all_banks(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = _get_processor(bank_name)(filepath, lookup)
        expected_ids = cfg['expected']['transaction_ids']
        for i, expected in enumerate(expected_ids):
            assert rows[i]['交易流水号'] == expected

    def test_unique_id_present(self, tmp_dir, bank_name):
        filepath = _create_bank_excel(os.path.join(tmp_dir, f'{bank_name}_流水.xlsx'), bank_name)
        lookup = _create_lookup_table_all_banks(os.path.join(tmp_dir, '主体查找表.xlsx'))
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
        lookup = _create_lookup_table_all_banks(os.path.join(tmp_dir, '主体查找表.xlsx'))
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
        lookup = _create_lookup_table_all_banks(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = _get_processor(bank_name)(filepath, lookup)
        assert len(rows) == 2

    def test_all_columns_present(self, tmp_dir, bank_name):
        cfg = BANK_TEST_CONFIGS[bank_name]
        filepath = _create_bank_excel(os.path.join(tmp_dir, f'{bank_name}_流水.xlsx'), bank_name)
        lookup = _create_lookup_table_all_banks(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = _get_processor(bank_name)(filepath, lookup)
        assert set(rows[0].keys()) == cfg['expected']['expected_keys']


class TestTraceabilityFields:
    """溯源字段单元测试：验证来源文件名、来源相对路径、处理时间字段正确性"""

    def test_source_filename_without_base_dir(self, tmp_dir):
        """无 base_dir 时，来源文件名正确，来源相对路径为绝对路径"""
        bank_name = '北京银行'
        filename = '北京银行_202401_流水.xlsx'
        filepath = _create_bank_excel(os.path.join(tmp_dir, filename), bank_name)
        lookup = _create_lookup_table_all_banks(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = _get_processor(bank_name)(filepath, lookup)

        assert len(rows) > 0
        for row in rows:
            assert row['来源文件名'] == filename
            assert row['来源相对路径'] == filepath
            assert '处理时间' in row
            assert len(row['处理时间']) > 0

    def test_source_filename_with_base_dir(self, tmp_dir):
        """有 base_dir 时，来源相对路径为相对路径"""
        bank_name = '北京银行'
        filename = '北京银行_202401_流水.xlsx'
        sub_dir = os.path.join(tmp_dir, '2024年1月')
        os.makedirs(sub_dir, exist_ok=True)
        filepath = _create_bank_excel(os.path.join(sub_dir, filename), bank_name)
        lookup = _create_lookup_table_all_banks(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = _get_processor(bank_name)(filepath, lookup, base_dir=tmp_dir)

        assert len(rows) > 0
        expected_rel = os.path.join('2024年1月', filename)
        for row in rows:
            assert row['来源文件名'] == filename
            assert row['来源相对路径'] == expected_rel
            assert row['来源相对路径'] != filepath

    def test_processed_at_format(self, tmp_dir):
        """处理时间格式应为 YYYY-MM-DD HH:MM:SS"""
        from datetime import datetime
        bank_name = '北京银行'
        filepath = _create_bank_excel(os.path.join(tmp_dir, '北京银行_流水.xlsx'), bank_name)
        lookup = _create_lookup_table_all_banks(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = _get_processor(bank_name)(filepath, lookup)

        for row in rows:
            processed_at = row['处理时间']
            assert len(processed_at) == 19
            try:
                datetime.strptime(processed_at, '%Y-%m-%d %H:%M:%S')
            except ValueError:
                pytest.fail(f'处理时间格式错误: {processed_at}')

    def test_all_records_same_traceability(self, tmp_dir):
        """同一文件的所有记录应具有相同的来源文件名和相对路径"""
        bank_name = '招商银行'
        filename = '招行_工资卡流水.xlsx'
        filepath = _create_bank_excel(os.path.join(tmp_dir, filename), bank_name)
        lookup = _create_lookup_table_all_banks(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = _get_processor(bank_name)(filepath, lookup, base_dir=tmp_dir)

        assert len(rows) >= 2
        first_filename = rows[0]['来源文件名']
        first_relpath = rows[0]['来源相对路径']
        for row in rows[1:]:
            assert row['来源文件名'] == first_filename
            assert row['来源相对路径'] == first_relpath

    def test_build_traceability_fields_helper(self, tmp_dir):
        """build_traceability_fields 辅助函数正确性测试"""
        filename = 'test.xlsx'
        sub_dir = os.path.join(tmp_dir, 'subdir')
        os.makedirs(sub_dir, exist_ok=True)
        filepath = os.path.join(sub_dir, filename)

        fields_no_base = bankcheck.build_traceability_fields(filepath)
        assert fields_no_base['来源文件名'] == filename
        assert fields_no_base['来源相对路径'] == filepath
        assert '处理时间' in fields_no_base

        fields_with_base = bankcheck.build_traceability_fields(filepath, base_dir=tmp_dir)
        assert fields_with_base['来源文件名'] == filename
        assert fields_with_base['来源相对路径'] == os.path.join('subdir', filename)
        assert fields_with_base['来源相对路径'] != filepath

    def test_add_traceability_to_records_helper(self, tmp_dir):
        """add_traceability_to_records 辅助函数正确性测试"""
        records = [
            {'唯一id': '1', '银行': '测试银行'},
            {'唯一id': '2', '银行': '测试银行'},
        ]
        filepath = os.path.join(tmp_dir, 'subdir', 'test.xlsx')
        result = bankcheck.add_traceability_to_records(records, filepath, base_dir=tmp_dir)

        assert result is records
        for r in records:
            assert r['来源文件名'] == 'test.xlsx'
            assert r['来源相对路径'] == os.path.join('subdir', 'test.xlsx')
            assert '处理时间' in r

    def test_standard_columns_includes_traceability(self):
        """STANDARD_COLUMNS 应包含溯源字段"""
        for col in ('来源文件名', '来源相对路径', '处理时间'):
            assert col in bankcheck.STANDARD_COLUMNS

    def test_summary_columns_includes_traceability(self, tmp_dir):
        """get_summary_columns 应返回包含溯源字段的列列表"""
        sample_records = [
            {'唯一id': '1', '银行': '测试银行', '银行账号': '123',
             '来源文件名': 'a.xlsx', '来源相对路径': 'a.xlsx', '处理时间': '2024-01-01 00:00:00'}
        ]
        columns = bankcheck.get_summary_columns(sample_records)
        for col in ('来源文件名', '来源相对路径', '处理时间'):
            assert col in columns
