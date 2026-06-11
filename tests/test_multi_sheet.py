import os
import re

import pytest

from conftest import (
    _create_beijing_bank_excel,
    _create_beijing_bank_multi_sheet_excel,
    _create_beijing_bank_excel_with_extra_sheets,
    _create_lookup_table,
)
from bankcheck import GenericBankParser, BankRule, process_beijing_bank


class TestMultiSheetBasicExtraction:
    def test_multi_sheet_merges_records(self, tmp_dir):
        filepath = _create_beijing_bank_multi_sheet_excel(
            os.path.join(tmp_dir, '北京银行_流水.xlsx'))
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = process_beijing_bank(filepath, lookup)
        assert len(rows) == 4

    def test_multi_sheet_preserves_bank_name(self, tmp_dir):
        filepath = _create_beijing_bank_multi_sheet_excel(
            os.path.join(tmp_dir, '北京银行_流水.xlsx'))
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = process_beijing_bank(filepath, lookup)
        for row in rows:
            assert row['银行'] == '北京银行'

    def test_multi_sheet_preserves_account(self, tmp_dir):
        filepath = _create_beijing_bank_multi_sheet_excel(
            os.path.join(tmp_dir, '北京银行_流水.xlsx'))
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = process_beijing_bank(filepath, lookup)
        for row in rows:
            assert row['银行账号'] == '01090312345678901'

    def test_multi_sheet_preserves_subject(self, tmp_dir):
        filepath = _create_beijing_bank_multi_sheet_excel(
            os.path.join(tmp_dir, '北京银行_流水.xlsx'))
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = process_beijing_bank(filepath, lookup)
        for row in rows:
            assert row['主体'] == '北京XX科技有限公司'

    def test_multi_sheet_dates_from_both_sheets(self, tmp_dir):
        filepath = _create_beijing_bank_multi_sheet_excel(
            os.path.join(tmp_dir, '北京银行_流水.xlsx'))
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = process_beijing_bank(filepath, lookup)
        dates = [row['交易日期'] for row in rows]
        assert '2024-01-05' in dates
        assert '2024-01-10' in dates
        assert '2024-02-03' in dates
        assert '2024-02-15' in dates

    def test_multi_sheet_payment_receipt_values(self, tmp_dir):
        filepath = _create_beijing_bank_multi_sheet_excel(
            os.path.join(tmp_dir, '北京银行_流水.xlsx'))
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = process_beijing_bank(filepath, lookup)
        payments = [row['付款'] for row in rows if row['付款'] is not None]
        receipts = [row['收款'] for row in rows if row['收款'] is not None]
        assert len(payments) == 2
        assert len(receipts) == 2
        assert -50000.0 in payments
        assert -30000.0 in payments
        assert 80000.0 in receipts
        assert 60000.0 in receipts

    def test_multi_sheet_unique_ids(self, tmp_dir):
        filepath = _create_beijing_bank_multi_sheet_excel(
            os.path.join(tmp_dir, '北京银行_流水.xlsx'))
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = process_beijing_bank(filepath, lookup)
        ids = [row['唯一id'] for row in rows]
        assert len(set(ids)) == 4

    def test_multi_sheet_transaction_ids(self, tmp_dir):
        filepath = _create_beijing_bank_multi_sheet_excel(
            os.path.join(tmp_dir, '北京银行_流水.xlsx'))
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = process_beijing_bank(filepath, lookup)
        tx_ids = [row['交易流水号'] for row in rows]
        assert 'BJ20240105001' in tx_ids
        assert 'BJ20240110002' in tx_ids
        assert 'BJ20240203003' in tx_ids
        assert 'BJ20240215004' in tx_ids


class TestMultiSheetDifferentAccounts:
    def test_different_accounts_per_sheet(self, tmp_dir):
        sheets = [
            {
                'title': '账号1',
                'account': '01090312345678901',
                'rows': [
                    [1, '2024-01-05', 'CNY', 50000, None, 1500000, '供应商A', '622001', '工商银行', '转账', '001', '采购', None, None, None, 'BJ001'],
                ],
            },
            {
                'title': '账号2',
                'account': '01090399999999999',
                'rows': [
                    [1, '2024-02-01', 'CNY', 20000, None, 800000, '供应商B', '622002', '建设银行', '转账', '002', '采购', None, None, None, 'BJ002'],
                ],
            },
        ]
        filepath = _create_beijing_bank_multi_sheet_excel(
            os.path.join(tmp_dir, '北京银行_流水.xlsx'), sheets=sheets)
        mappings = [
            ('公司A', '01090312345678901'),
            ('公司B', '01090399999999999'),
        ]
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'),
                                       mappings=mappings)
        rows = process_beijing_bank(filepath, lookup)
        assert len(rows) == 2
        accounts = {row['银行账号'] for row in rows}
        assert accounts == {'01090312345678901', '01090399999999999'}
        subjects = {row['主体'] for row in rows}
        assert subjects == {'公司A', '公司B'}


class TestMultiSheetWithNonDataSheets:
    def test_extra_non_data_sheets_produce_no_records(self, tmp_dir):
        filepath = _create_beijing_bank_excel_with_extra_sheets(
            os.path.join(tmp_dir, '北京银行_流水.xlsx'))
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = process_beijing_bank(filepath, lookup)
        assert len(rows) == 1
        assert rows[0]['交易日期'] == '2024-01-05'

    def test_extra_sheets_do_not_duplicate_records(self, tmp_dir):
        filepath = _create_beijing_bank_excel_with_extra_sheets(
            os.path.join(tmp_dir, '北京银行_流水.xlsx'))
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = process_beijing_bank(filepath, lookup)
        tx_ids = [row['交易流水号'] for row in rows]
        assert tx_ids.count('BJ20240105001') == 1


class TestMultiSheetSkipSheetsConfig:
    def test_skip_sheets_excludes_named_sheet(self, tmp_dir):
        filepath = _create_beijing_bank_multi_sheet_excel(
            os.path.join(tmp_dir, '北京银行_流水.xlsx'))
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rule = BankRule(
            bank_name='北京银行',
            account_cell='B2',
            start_row=4,
            columns={
                'trade_date': 2, 'payment': 4, 'receipt': 5,
                'summary': 12, 'counterpart': 7, 'balance': 6,
                'transaction_id': 16,
            },
            payment_sign='negative',
            skip_sheets=['2月'],
        )
        parser = GenericBankParser(rule)
        rows = parser.parse(filepath, lookup)
        assert len(rows) == 2
        dates = [row['交易日期'] for row in rows]
        assert all(d.startswith('2024-01') for d in dates)

    def test_skip_sheets_empty_list_processes_all(self, tmp_dir):
        filepath = _create_beijing_bank_multi_sheet_excel(
            os.path.join(tmp_dir, '北京银行_流水.xlsx'))
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rule = BankRule(
            bank_name='北京银行',
            account_cell='B2',
            start_row=4,
            columns={
                'trade_date': 2, 'payment': 4, 'receipt': 5,
                'summary': 12, 'counterpart': 7, 'balance': 6,
                'transaction_id': 16,
            },
            payment_sign='negative',
            skip_sheets=[],
        )
        parser = GenericBankParser(rule)
        rows = parser.parse(filepath, lookup)
        assert len(rows) == 4

    def test_skip_sheets_none_processes_all(self, tmp_dir):
        filepath = _create_beijing_bank_multi_sheet_excel(
            os.path.join(tmp_dir, '北京银行_流水.xlsx'))
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rule = BankRule(
            bank_name='北京银行',
            account_cell='B2',
            start_row=4,
            columns={
                'trade_date': 2, 'payment': 4, 'receipt': 5,
                'summary': 12, 'counterpart': 7, 'balance': 6,
                'transaction_id': 16,
            },
            payment_sign='negative',
            skip_sheets=None,
        )
        parser = GenericBankParser(rule)
        rows = parser.parse(filepath, lookup)
        assert len(rows) == 4

    def test_skip_all_sheets_returns_empty(self, tmp_dir):
        filepath = _create_beijing_bank_multi_sheet_excel(
            os.path.join(tmp_dir, '北京银行_流水.xlsx'))
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rule = BankRule(
            bank_name='北京银行',
            account_cell='B2',
            start_row=4,
            columns={
                'trade_date': 2, 'payment': 4, 'receipt': 5,
                'summary': 12, 'counterpart': 7, 'balance': 6,
                'transaction_id': 16,
            },
            payment_sign='negative',
            skip_sheets=['1月', '2月'],
        )
        parser = GenericBankParser(rule)
        rows = parser.parse(filepath, lookup)
        assert len(rows) == 0


class TestMultiSheetEmptySheet:
    def test_empty_sheet_contributes_no_records(self, tmp_dir):
        sheets = [
            {
                'title': '有数据',
                'account': '01090312345678901',
                'rows': [
                    [1, '2024-01-05', 'CNY', 50000, None, 1500000, '供应商A', '622001', '工商银行', '转账', '001', '采购', None, None, None, 'BJ001'],
                ],
            },
            {
                'title': '空表',
                'account': '01090312345678901',
                'rows': [],
            },
        ]
        filepath = _create_beijing_bank_multi_sheet_excel(
            os.path.join(tmp_dir, '北京银行_流水.xlsx'), sheets=sheets)
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = process_beijing_bank(filepath, lookup)
        assert len(rows) == 1
        assert rows[0]['交易日期'] == '2024-01-05'

    def test_all_empty_sheets_returns_empty(self, tmp_dir):
        sheets = [
            {'title': '空表1', 'account': '01090312345678901', 'rows': []},
            {'title': '空表2', 'account': '01090312345678901', 'rows': []},
        ]
        filepath = _create_beijing_bank_multi_sheet_excel(
            os.path.join(tmp_dir, '北京银行_流水.xlsx'), sheets=sheets)
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = process_beijing_bank(filepath, lookup)
        assert len(rows) == 0


class TestMultiSheetMissingAccount:
    def test_sheet_with_missing_account_still_extracts(self, tmp_dir):
        sheets = [
            {
                'title': '有账号',
                'account': '01090312345678901',
                'rows': [
                    [1, '2024-01-05', 'CNY', 50000, None, 1500000, '供应商A', '622001', '工商银行', '转账', '001', '采购', None, None, None, 'BJ001'],
                ],
            },
            {
                'title': '无账号',
                'account': None,
                'rows': [
                    [1, '2024-02-01', 'CNY', 20000, None, 800000, '供应商B', '622002', '建设银行', '转账', '002', '采购', None, None, None, 'BJ002'],
                ],
            },
        ]
        filepath = _create_beijing_bank_multi_sheet_excel(
            os.path.join(tmp_dir, '北京银行_流水.xlsx'), sheets=sheets)
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = process_beijing_bank(filepath, lookup)
        assert len(rows) == 2
        accounts = [row['银行账号'] for row in rows]
        assert accounts[0] == '01090312345678901'
        assert accounts[1] is None


class TestSingleSheetBackwardCompat:
    def test_single_sheet_still_works(self, tmp_dir):
        filepath = _create_beijing_bank_excel(
            os.path.join(tmp_dir, '北京银行_流水.xlsx'))
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = process_beijing_bank(filepath, lookup)
        assert len(rows) == 2
        assert rows[0]['银行'] == '北京银行'
        assert rows[0]['付款'] == -50000.0
        assert rows[1]['收款'] == 80000.0

    def test_single_sheet_standard_columns(self, tmp_dir):
        filepath = _create_beijing_bank_excel(
            os.path.join(tmp_dir, '北京银行_流水.xlsx'))
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = process_beijing_bank(filepath, lookup)
        expected_keys = {'唯一id', '银行', '银行账号', '主体', '交易日期',
                         '付款', '收款', '摘要', '对方户名', '余额', '交易流水号'}
        assert set(rows[0].keys()) == expected_keys


class TestParseSheetMethod:
    def test_parse_sheet_returns_records(self, tmp_dir):
        filepath = _create_beijing_bank_excel(
            os.path.join(tmp_dir, '北京银行_流水.xlsx'))
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rule = BankRule(
            bank_name='北京银行',
            account_cell='B2',
            start_row=4,
            columns={
                'trade_date': 2, 'payment': 4, 'receipt': 5,
                'summary': 12, 'counterpart': 7, 'balance': 6,
                'transaction_id': 16,
            },
            payment_sign='negative',
        )
        parser = GenericBankParser(rule)
        import openpyxl
        wb = openpyxl.load_workbook(filepath, data_only=True)
        ws = wb.active
        rows = parser._parse_sheet(ws, filepath, ws.title, lookup)
        wb.close()
        assert len(rows) == 2
        assert rows[0]['银行'] == '北京银行'

    def test_parse_sheet_empty_sheet_returns_empty(self, tmp_dir):
        rule = BankRule(
            bank_name='北京银行',
            account_cell='B2',
            start_row=4,
            columns={
                'trade_date': 2, 'payment': 4, 'receipt': 5,
                'summary': 12, 'counterpart': 7, 'balance': 6,
                'transaction_id': 16,
            },
            payment_sign='negative',
        )
        parser = GenericBankParser(rule)
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = '空表'
        rows = parser._parse_sheet(ws, '/fake/path.xlsx', '空表', None)
        wb.close()
        assert len(rows) == 0


class TestGenericBankParserThreeSheets:
    def test_three_sheets_merged(self, tmp_dir):
        sheets = [
            {
                'title': '1月',
                'account': '01090312345678901',
                'rows': [
                    [1, '2024-01-05', 'CNY', 10000, None, 1490000, 'A', '1', 'ICBC', 'T', '1', 'P1', None, None, None, 'T1'],
                ],
            },
            {
                'title': '2月',
                'account': '01090312345678901',
                'rows': [
                    [1, '2024-02-01', 'CNY', 20000, None, 1470000, 'B', '2', 'CCB', 'T', '2', 'P2', None, None, None, 'T2'],
                ],
            },
            {
                'title': '3月',
                'account': '01090312345678901',
                'rows': [
                    [1, '2024-03-01', 'CNY', None, 30000, 1500000, 'C', '3', 'BOC', 'T', '3', 'P3', None, None, None, 'T3'],
                ],
            },
        ]
        filepath = _create_beijing_bank_multi_sheet_excel(
            os.path.join(tmp_dir, '北京银行_流水.xlsx'), sheets=sheets)
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rows = process_beijing_bank(filepath, lookup)
        assert len(rows) == 3
        dates = sorted([row['交易日期'] for row in rows])
        assert dates == ['2024-01-05', '2024-02-01', '2024-03-01']

    def test_three_sheets_skip_middle(self, tmp_dir):
        sheets = [
            {
                'title': '1月',
                'account': '01090312345678901',
                'rows': [
                    [1, '2024-01-05', 'CNY', 10000, None, 1490000, 'A', '1', 'ICBC', 'T', '1', 'P1', None, None, None, 'T1'],
                ],
            },
            {
                'title': '2月',
                'account': '01090312345678901',
                'rows': [
                    [1, '2024-02-01', 'CNY', 20000, None, 1470000, 'B', '2', 'CCB', 'T', '2', 'P2', None, None, None, 'T2'],
                ],
            },
            {
                'title': '3月',
                'account': '01090312345678901',
                'rows': [
                    [1, '2024-03-01', 'CNY', None, 30000, 1500000, 'C', '3', 'BOC', 'T', '3', 'P3', None, None, None, 'T3'],
                ],
            },
        ]
        filepath = _create_beijing_bank_multi_sheet_excel(
            os.path.join(tmp_dir, '北京银行_流水.xlsx'), sheets=sheets)
        lookup = _create_lookup_table(os.path.join(tmp_dir, '主体查找表.xlsx'))
        rule = BankRule(
            bank_name='北京银行',
            account_cell='B2',
            start_row=4,
            columns={
                'trade_date': 2, 'payment': 4, 'receipt': 5,
                'summary': 12, 'counterpart': 7, 'balance': 6,
                'transaction_id': 16,
            },
            payment_sign='negative',
            skip_sheets=['2月'],
        )
        parser = GenericBankParser(rule)
        rows = parser.parse(filepath, lookup)
        assert len(rows) == 2
        dates = sorted([row['交易日期'] for row in rows])
        assert dates == ['2024-01-05', '2024-03-01']
