import os
from collections import Counter

import openpyxl
import pandas as pd
import pytest

from conftest import (
    BANK_TEST_CONFIGS,
    _create_bank_excel,
    _create_lookup_table_all_banks,
)
import bankcheck


_ALL_BANKS = list(BANK_TEST_CONFIGS.keys())


class TestBankRegistration:
    def test_all_five_banks_registered(self):
        for bank_name in _ALL_BANKS:
            assert bank_name in bankcheck.BANK_PREFIXES, f"{bank_name} 不在 BANK_PREFIXES 中"
            assert bank_name in bankcheck.BANK_PROCESSORS, f"{bank_name} 不在 BANK_PROCESSORS 中"

    def test_processor_count_matches_prefixes(self):
        assert len(bankcheck.BANK_PREFIXES) == len(bankcheck.BANK_PROCESSORS)

    def test_backward_alias_icbc(self):
        assert bankcheck.process_icbc_bank is not None
        assert bankcheck.process_icbc_bank is bankcheck.process_industrial_commercial_bank

    def test_backward_alias_ccb(self):
        assert bankcheck.process_ccb_bank is not None
        assert bankcheck.process_ccb_bank is bankcheck.process_construction_bank

    def test_backward_alias_cmb(self):
        assert bankcheck.process_cmb_bank is not None
        assert bankcheck.process_cmb_bank is bankcheck.process_merchants_bank

    def test_processor_callable(self):
        for bank_name in _ALL_BANKS:
            assert callable(bankcheck.BANK_PROCESSORS[bank_name])


class TestMultiBankIdentify:
    @pytest.mark.parametrize('bank_name', _ALL_BANKS, ids=_ALL_BANKS)
    def test_identify_by_prefix(self, tmp_dir, bank_name):
        filepath = _create_bank_excel(
            os.path.join(tmp_dir, f'{bank_name}_2024年1月流水.xlsx'), bank_name
        )
        identified = bankcheck.identify_bank(filepath)
        assert identified == bank_name, f"前缀匹配失败: {bank_name}"

    @pytest.mark.parametrize('bank_name', _ALL_BANKS, ids=_ALL_BANKS)
    def test_identify_by_filename_contains(self, tmp_dir, bank_name):
        filepath = _create_bank_excel(
            os.path.join(tmp_dir, f'2024年1月{bank_name}流水明细.xlsx'), bank_name
        )
        identified = bankcheck.identify_bank(filepath)
        assert identified == bank_name, f"文件名包含匹配失败: {bank_name}"

    @pytest.mark.parametrize('bank_name', _ALL_BANKS, ids=_ALL_BANKS)
    def test_identify_with_separators_in_name(self, tmp_dir, bank_name):
        filepath = _create_bank_excel(
            os.path.join(tmp_dir, f'{bank_name} - 2024-01.xlsx'), bank_name
        )
        identified = bankcheck.identify_bank(filepath)
        assert identified == bank_name, f"含分隔符匹配失败: {bank_name}"


class TestMultiBankProcessing:
    @pytest.mark.parametrize('bank_name', _ALL_BANKS, ids=_ALL_BANKS)
    def test_single_bank_produces_records(self, tmp_dir, bank_name):
        cfg = BANK_TEST_CONFIGS[bank_name]
        filepath = _create_bank_excel(
            os.path.join(tmp_dir, f'{bank_name}_流水.xlsx'), bank_name
        )
        lookup = _create_lookup_table_all_banks(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = bankcheck.BANK_PROCESSORS[bank_name](filepath, lookup)
        assert len(rows) == cfg['expected']['row_count']
        for row in rows:
            assert row['银行'] == cfg['expected']['bank_name']

    def test_process_all_banks_together(self, tmp_dir):
        lookup = _create_lookup_table_all_banks(os.path.join(tmp_dir, '主体查找表.xlsx'))
        all_rows = []
        per_bank_counts = {}

        for bank_name in _ALL_BANKS:
            filepath = _create_bank_excel(
                os.path.join(tmp_dir, f'{bank_name}_流水.xlsx'), bank_name
            )
            rows = bankcheck.BANK_PROCESSORS[bank_name](filepath, lookup)
            per_bank_counts[bank_name] = len(rows)
            all_rows.extend(rows)

        expected_total = sum(
            BANK_TEST_CONFIGS[b]['expected']['row_count'] for b in _ALL_BANKS
        )
        assert len(all_rows) == expected_total

        actual_counts = Counter(row['银行'] for row in all_rows)
        for bank_name in _ALL_BANKS:
            expected = BANK_TEST_CONFIGS[bank_name]['expected']['row_count']
            assert actual_counts[bank_name] == expected, \
                f"{bank_name} 记录数不匹配: 期望 {expected}, 实际 {actual_counts[bank_name]}"

    def test_all_records_have_standard_columns(self, tmp_dir):
        lookup = _create_lookup_table_all_banks(os.path.join(tmp_dir, '主体查找表.xlsx'))
        standard_cols = {
            '唯一id', '银行', '银行账号', '主体', '交易日期',
            '付款', '收款', '摘要', '对方户名', '余额', '交易流水号',
        }

        for bank_name in _ALL_BANKS:
            filepath = _create_bank_excel(
                os.path.join(tmp_dir, f'{bank_name}_流水.xlsx'), bank_name
            )
            rows = bankcheck.BANK_PROCESSORS[bank_name](filepath, lookup)
            for i, row in enumerate(rows):
                missing = standard_cols - set(row.keys())
                assert not missing, \
                    f"{bank_name} 第 {i + 1} 条记录缺少字段: {missing}"

    def test_all_unique_ids_are_distinct(self, tmp_dir):
        lookup = _create_lookup_table_all_banks(os.path.join(tmp_dir, '主体查找表.xlsx'))
        all_ids = set()

        for bank_name in _ALL_BANKS:
            filepath = _create_bank_excel(
                os.path.join(tmp_dir, f'{bank_name}_流水.xlsx'), bank_name
            )
            rows = bankcheck.BANK_PROCESSORS[bank_name](filepath, lookup)
            for row in rows:
                uid = row['唯一id']
                assert uid not in all_ids, f"重复的唯一id: {uid} ({bank_name})"
                all_ids.add(uid)


class TestUnifiedSummaryExport:
    def test_generate_unified_summary_excel(self, tmp_dir):
        lookup = _create_lookup_table_all_banks(os.path.join(tmp_dir, '主体查找表.xlsx'))
        all_rows = []

        for bank_name in _ALL_BANKS:
            filepath = _create_bank_excel(
                os.path.join(tmp_dir, f'{bank_name}_流水.xlsx'), bank_name
            )
            rows = bankcheck.BANK_PROCESSORS[bank_name](filepath, lookup)
            all_rows.extend(rows)

        columns = bankcheck.get_summary_columns(all_rows, lookup)
        output_path = os.path.join(tmp_dir, '银行流水总表.xlsx')
        df = pd.DataFrame(all_rows, columns=columns)
        df.to_excel(output_path, index=False, engine='openpyxl')

        assert os.path.exists(output_path)

        df_read = pd.read_excel(output_path, engine='openpyxl')
        assert len(df_read) == len(all_rows)

        expected_total = sum(
            BANK_TEST_CONFIGS[b]['expected']['row_count'] for b in _ALL_BANKS
        )
        assert len(df_read) == expected_total

        bank_counts = df_read['银行'].value_counts().to_dict()
        for bank_name in _ALL_BANKS:
            expected = BANK_TEST_CONFIGS[bank_name]['expected']['row_count']
            assert bank_counts.get(bank_name, 0) == expected

    def test_summary_columns_contain_all_standard(self, tmp_dir):
        lookup = _create_lookup_table_all_banks(os.path.join(tmp_dir, '主体查找表.xlsx'))
        all_rows = []
        for bank_name in _ALL_BANKS:
            filepath = _create_bank_excel(
                os.path.join(tmp_dir, f'{bank_name}_流水.xlsx'), bank_name
            )
            rows = bankcheck.BANK_PROCESSORS[bank_name](filepath, lookup)
            all_rows.extend(rows)

        columns = bankcheck.get_summary_columns(all_rows, lookup)
        standard = [
            '唯一id', '银行', '银行账号', '主体', '交易日期',
            '付款', '收款', '摘要', '对方户名', '余额', '交易流水号',
        ]
        for col in standard:
            assert col in columns, f"总表缺少标准列: {col}"


class TestBackwardAliasDirectCall:
    def test_process_icbc_bank_direct(self, tmp_dir):
        filepath = _create_bank_excel(
            os.path.join(tmp_dir, '工商银行_流水.xlsx'), '工商银行'
        )
        lookup = _create_lookup_table_all_banks(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = bankcheck.process_icbc_bank(filepath, lookup)
        assert len(rows) == 2
        assert rows[0]['银行'] == '工商银行'

    def test_process_ccb_bank_direct(self, tmp_dir):
        filepath = _create_bank_excel(
            os.path.join(tmp_dir, '建设银行_流水.xlsx'), '建设银行'
        )
        lookup = _create_lookup_table_all_banks(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = bankcheck.process_ccb_bank(filepath, lookup)
        assert len(rows) == 2
        assert rows[0]['银行'] == '建设银行'

    def test_process_cmb_bank_direct(self, tmp_dir):
        filepath = _create_bank_excel(
            os.path.join(tmp_dir, '招商银行_流水.xlsx'), '招商银行'
        )
        lookup = _create_lookup_table_all_banks(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = bankcheck.process_cmb_bank(filepath, lookup)
        assert len(rows) == 2
        assert rows[0]['银行'] == '招商银行'


class TestIdentifyByContent:
    @pytest.mark.parametrize('bank_name', _ALL_BANKS, ids=_ALL_BANKS)
    def test_content_unique_cell_identified(self, tmp_dir, bank_name):
        filepath = _create_bank_excel(
            os.path.join(tmp_dir, '无银行名_流水.xlsx'), bank_name
        )
        identified = bankcheck._identify_bank_by_content(filepath)
        assert identified == bank_name, \
            f"{bank_name} 内容识别失败: account_cell 应唯一匹配"


class TestPaymentSignConsistency:
    @pytest.mark.parametrize('bank_name', _ALL_BANKS, ids=_ALL_BANKS)
    def test_payment_is_negative(self, tmp_dir, bank_name):
        cfg = BANK_TEST_CONFIGS[bank_name]
        if cfg['expected'].get('first_payment') is None:
            pytest.skip(f"{bank_name} 第一条记录非付款")
        filepath = _create_bank_excel(
            os.path.join(tmp_dir, f'{bank_name}_流水.xlsx'), bank_name
        )
        lookup = _create_lookup_table_all_banks(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = bankcheck.BANK_PROCESSORS[bank_name](filepath, lookup)
        assert rows[0]['付款'] < 0, f"{bank_name} 付款应为负数"

    @pytest.mark.parametrize('bank_name', _ALL_BANKS, ids=_ALL_BANKS)
    def test_receipt_is_positive(self, tmp_dir, bank_name):
        cfg = BANK_TEST_CONFIGS[bank_name]
        if cfg['expected'].get('second_receipt') is None:
            pytest.skip(f"{bank_name} 第二条记录非收款")
        filepath = _create_bank_excel(
            os.path.join(tmp_dir, f'{bank_name}_流水.xlsx'), bank_name
        )
        lookup = _create_lookup_table_all_banks(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = bankcheck.BANK_PROCESSORS[bank_name](filepath, lookup)
        assert rows[1]['收款'] > 0, f"{bank_name} 收款应为正数"


class TestSubjectLookupAllBanks:
    @pytest.mark.parametrize('bank_name', _ALL_BANKS, ids=_ALL_BANKS)
    def test_subject_resolved_correctly(self, tmp_dir, bank_name):
        cfg = BANK_TEST_CONFIGS[bank_name]
        filepath = _create_bank_excel(
            os.path.join(tmp_dir, f'{bank_name}_流水.xlsx'), bank_name
        )
        lookup = _create_lookup_table_all_banks(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = bankcheck.BANK_PROCESSORS[bank_name](filepath, lookup)
        for row in rows:
            assert row['主体'] == cfg['expected']['subject'], \
                f"{bank_name} 主体不匹配: 期望 {cfg['expected']['subject']}, 实际 {row['主体']}"
