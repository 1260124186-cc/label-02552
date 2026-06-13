import os
import sys
import tempfile
import shutil

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'backend'))

import bankcheck
from conftest import (
    _create_beijing_bank_excel,
    _create_east_asia_bank_excel,
    _create_lookup_table,
)


def _create_mock_pdf_like_text():
    return """北京银行
账号 01090312345678901
交易明细
序号\t交易日期\t币种\t支出金额\t收入金额\t余额\t对方户名\t对方账号\t对方行名\t凭证种类\t凭证号码\t摘要\t备注1\t备注2\t备注3\t交易流水号
1\t2024-01-05\tCNY\t50,000.00\t\t1,500,000.00\t供应商A公司\t622001234\t工商银行\t转账\t001\t采购付款\t\t\t\tBJ20240105001
2\t2024-01-10\tCNY\t\t80,000.00\t1,580,000.00\t客户B公司\t622005678\t建设银行\t转账\t002\t销售收款\t\t\t\tBJ20240110002
"""


class TestPdfParserImports:
    def test_import_pdf_bank_parser(self):
        try:
            from pdf_bank_parser import (
                PdfParseResult,
                parse_pdf_bank_statement,
                process_pdf_file,
                scan_pdf_files,
                scan_bank_files,
                is_pdf_file,
                detect_header_row,
                map_columns_by_header,
                convert_table_to_rows,
                _parse_amount,
                _parse_date,
            )
            assert True
        except ImportError as e:
            pytest.fail(f'Failed to import pdf_bank_parser: {e}')

    def test_is_pdf_file(self):
        from pdf_bank_parser import is_pdf_file
        assert is_pdf_file('test.pdf')
        assert is_pdf_file('/path/to/流水.PDF')
        assert not is_pdf_file('test.xlsx')
        assert not is_pdf_file('test.xls')
        assert not is_pdf_file(None)


class TestAmountParsing:
    def test_parse_amount_basic(self):
        from pdf_bank_parser import _parse_amount
        assert _parse_amount('1234.56') == 1234.56
        assert _parse_amount('-5000.00') == -5000.00
        assert _parse_amount('0.00') == 0.0

    def test_parse_amount_thousands_separator(self):
        from pdf_bank_parser import _parse_amount
        assert _parse_amount('1,500,000.00') == 1500000.00
        assert _parse_amount('50,000.00') == 50000.00
        assert _parse_amount('1,234,567.89') == 1234567.89

    def test_parse_amount_currency_symbols(self):
        from pdf_bank_parser import _parse_amount
        assert _parse_amount('¥1234.56') == 1234.56
        assert _parse_amount('￥5,000.00') == 5000.00
        assert _parse_amount('CNY 10000') == 10000.0

    def test_parse_amount_parentheses_negative(self):
        from pdf_bank_parser import _parse_amount
        assert _parse_amount('(5000.00)') == -5000.00
        assert _parse_amount('(1,234.56)') == -1234.56

    def test_parse_amount_empty(self):
        from pdf_bank_parser import _parse_amount
        assert _parse_amount(None) is None
        assert _parse_amount('') is None
        assert _parse_amount('-') is None
        assert _parse_amount('--') is None


class TestDateParsing:
    def test_parse_date_standard_formats(self):
        from pdf_bank_parser import _parse_date
        assert _parse_date('2024-01-05') == '2024-01-05'
        assert _parse_date('2024/01/05') == '2024-01-05'
        assert _parse_date('2024年1月5日') == '2024-01-05'
        assert _parse_date('2024.01.05') == '2024-01-05'

    def test_parse_date_compact(self):
        from pdf_bank_parser import _parse_date
        assert _parse_date('20240105') == '20240105'

    def test_parse_date_us_format(self):
        from pdf_bank_parser import _parse_date
        assert _parse_date('01/05/2024') == '2024-01-05'

    def test_parse_date_empty(self):
        from pdf_bank_parser import _parse_date
        assert _parse_date(None) is None
        assert _parse_date('') is None


class TestHeaderDetection:
    def test_detect_header_row_standard(self):
        from pdf_bank_parser import detect_header_row
        table = [
            ['北京银行交易明细', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
            ['', '账号：01090312345678901', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
            ['序号', '交易日期', '币种', '支出金额', '收入金额', '余额', '对方户名', '对方账号',
             '对方行名', '凭证种类', '凭证号码', '摘要', '备注1', '备注2', '备注3', '交易流水号'],
            ['1', '2024-01-05', 'CNY', '50000', '', '1500000', '供应商A', '', '', '转账', '001', '采购', '', '', '', 'BJ001'],
        ]
        assert detect_header_row(table) == 2

    def test_map_columns_beijing_bank(self):
        from pdf_bank_parser import map_columns_by_header
        header = ['序号', '交易日期', '币种', '支出金额', '收入金额', '余额',
                  '对方户名', '对方账号', '对方行名', '凭证种类', '凭证号码',
                  '摘要', '备注1', '备注2', '备注3', '交易流水号']
        mapping = map_columns_by_header(header)
        assert mapping['trade_date'] == 1
        assert mapping['payment'] == 3
        assert mapping['receipt'] == 4
        assert mapping['balance'] == 5
        assert mapping['counterpart'] == 6
        assert mapping['summary'] == 11
        assert mapping['transaction_id'] == 15

    def test_map_columns_east_asia_bank(self):
        from pdf_bank_parser import map_columns_by_header
        header = ['交易日期', '交易时间', '币种', '支出金额', '收入金额',
                  '手续费', '利息', '税费', '余额', '交易类型', '交易流水号', '交易描述/对方']
        mapping = map_columns_by_header(header)
        assert mapping['trade_date'] == 0
        assert mapping['payment'] == 3
        assert mapping['receipt'] == 4
        assert mapping['balance'] == 8
        assert mapping['transaction_id'] == 10
        assert 'summary' in mapping or 'counterpart' in mapping


class TestTableToRowsConversion:
    def test_convert_beijing_bank_table(self, tmp_dir):
        from pdf_bank_parser import convert_table_to_rows

        lookup_file = os.path.join(tmp_dir, '主体查找表.xlsx')
        _create_lookup_table(lookup_file)
        lookup_data = bankcheck.load_lookup_table(lookup_file)

        table = [
            ['序号', '交易日期', '币种', '支出金额', '收入金额', '余额', '对方户名',
             '对方账号', '对方行名', '凭证种类', '凭证号码', '摘要', '备注1',
             '备注2', '备注3', '交易流水号'],
            ['1', '2024-01-05', 'CNY', '50,000.00', '', '1,500,000.00', '供应商A公司',
             '622001234', '工商银行', '转账', '001', '采购付款', '', '', '', 'BJ20240105001'],
            ['2', '2024-01-10', 'CNY', '', '80,000.00', '1,580,000.00', '客户B公司',
             '622005678', '建设银行', '转账', '002', '销售收款', '', '', '', 'BJ20240110002'],
        ]

        rows = convert_table_to_rows(
            table, '北京银行', '01090312345678901', lookup_data,
        )

        assert len(rows) == 2
        assert rows[0]['银行'] == '北京银行'
        assert rows[0]['银行账号'] == '01090312345678901'
        assert rows[0]['交易日期'] == '2024-01-05'
        assert rows[0]['付款'] == -50000.0
        assert rows[0]['收款'] is None
        assert rows[0]['余额'] == 1500000.0
        assert rows[0]['摘要'] == '采购付款'
        assert rows[0]['对方户名'] == '供应商A公司'
        assert rows[0]['交易流水号'] == 'BJ20240105001'

        assert rows[1]['收款'] == 80000.0
        assert rows[1]['付款'] is None
        assert rows[1]['余额'] == 1580000.0

    def test_convert_empty_table(self, tmp_dir):
        from pdf_bank_parser import convert_table_to_rows
        lookup_file = os.path.join(tmp_dir, '主体查找表.xlsx')
        _create_lookup_table(lookup_file)
        lookup_data = bankcheck.load_lookup_table(lookup_file)
        assert convert_table_to_rows([], '北京银行', '', lookup_data) == []
        assert convert_table_to_rows([['header1', 'header2']], '北京银行', '', lookup_data) == []


class TestTextParsing:
    def test_parse_text_to_rows(self, tmp_dir):
        from pdf_bank_parser import parse_text_to_rows

        lookup_file = os.path.join(tmp_dir, '主体查找表.xlsx')
        _create_lookup_table(lookup_file)
        lookup_data = bankcheck.load_lookup_table(lookup_file)

        text = _create_mock_pdf_like_text()
        rows = parse_text_to_rows(text, '北京银行', '01090312345678901', lookup_data)

        assert len(rows) >= 1
        for r in rows:
            assert r['银行'] == '北京银行'
            assert r['银行账号'] == '01090312345678901'
            assert r['交易日期'] is not None


class TestScanPdfFiles:
    def test_scan_pdf_files_empty(self, tmp_dir):
        from pdf_bank_parser import scan_pdf_files
        result = scan_pdf_files(tmp_dir)
        assert result == []

    def test_scan_pdf_files_mixed(self, tmp_dir):
        from pdf_bank_parser import scan_pdf_files

        pdf1 = os.path.join(tmp_dir, '北京银行_流水.pdf')
        pdf2 = os.path.join(tmp_dir, 'subdir', '东亚银行.PDF')
        excel1 = os.path.join(tmp_dir, '其他文件.xlsx')

        os.makedirs(os.path.dirname(pdf2), exist_ok=True)
        for f in [pdf1, pdf2, excel1]:
            with open(f, 'w') as fp:
                fp.write('mock')

        results = scan_pdf_files(tmp_dir)
        assert len(results) == 2
        basenames = [os.path.basename(r) for r in results]
        assert '北京银行_流水.pdf' in basenames
        assert '东亚银行.PDF' in basenames

    def test_scan_bank_files(self, tmp_dir):
        from pdf_bank_parser import scan_bank_files

        pdf1 = os.path.join(tmp_dir, '北京银行_流水.pdf')
        excel1 = os.path.join(tmp_dir, '东亚银行_流水.xlsx')
        for f in [pdf1, excel1]:
            with open(f, 'w') as fp:
                fp.write('mock')

        results = scan_bank_files(tmp_dir)
        assert len(results) == 2
        basenames = [os.path.basename(r) for r in results]
        assert '北京银行_流水.pdf' in basenames
        assert '东亚银行_流水.xlsx' in basenames


class TestBankKeywords:
    def test_bank_keywords_defined(self):
        from pdf_bank_parser import BANK_KEYWORDS
        assert '北京银行' in BANK_KEYWORDS
        assert '东亚银行' in BANK_KEYWORDS
        assert '工商银行' in BANK_KEYWORDS
        assert '建设银行' in BANK_KEYWORDS
        assert '招商银行' in BANK_KEYWORDS
        assert isinstance(BANK_KEYWORDS['北京银行'], list)
        assert len(BANK_KEYWORDS['北京银行']) > 0


class TestAccountBlockDetection:
    def test_detect_single_account_block(self):
        from pdf_bank_parser import detect_account_blocks_in_tables
        tables = [[
            ['', '账号：01090312345678901', '', ''],
            ['交易日期', '支出金额', '收入金额', '余额'],
            ['2024-01-05', '50000', '', '1500000'],
            ['2024-01-10', '', '80000', '1580000'],
        ]]
        blocks = detect_account_blocks_in_tables(tables)
        assert len(blocks) == 1
        assert blocks[0]['account'] == '01090312345678901'
        assert blocks[0]['table_idx'] == 0

    def test_detect_multi_account_blocks(self):
        from pdf_bank_parser import detect_account_blocks_in_tables
        tables = [[
            ['', '账号：01090312345678901', '', ''],
            ['交易日期', '支出金额', '收入金额', '余额'],
            ['2024-01-05', '50000', '', '1500000'],
            ['', '账号：01090399999999999', '', ''],
            ['交易日期', '支出金额', '收入金额', '余额'],
            ['2024-02-01', '20000', '', '800000'],
        ]]
        blocks = detect_account_blocks_in_tables(tables)
        accounts = [b['account'] for b in blocks]
        assert '01090312345678901' in accounts
        assert '01090399999999999' in accounts


class TestPdfIntegrationWithPipeline:
    def test_run_pipeline_mixed_excel_pdf_with_empty_pdf(self, tmp_dir):
        script_dir = os.path.join(tmp_dir, 'script')
        os.makedirs(script_dir, exist_ok=True)
        source_folder = os.path.join(tmp_dir, '流水文件夹')
        os.makedirs(source_folder, exist_ok=True)

        _create_beijing_bank_excel(os.path.join(source_folder, '北京银行_流水.xlsx'))

        fake_pdf = os.path.join(source_folder, '东亚银行_流水.pdf')
        with open(fake_pdf, 'wb') as f:
            f.write(b'%PDF-1.4 fake pdf content (not actually valid)')

        _create_lookup_table(os.path.join(script_dir, '主体查找表.xlsx'))

        result = bankcheck.run_pipeline(source_folder, script_dir)

        assert len(result.processed_files) + len(result.unprocessed_files) + len(result.error_files) >= 2
        basenames = [os.path.basename(f) for f in result.processed_files]
        if '北京银行_流水.xlsx' in basenames:
            assert any('北京银行' == r.get('银行') for r in result.all_rows)

    def test_process_pdf_file_invalid_file(self, tmp_dir):
        from pdf_bank_parser import process_pdf_file

        lookup_file = os.path.join(tmp_dir, '主体查找表.xlsx')
        _create_lookup_table(lookup_file)
        lookup_data = bankcheck.load_lookup_table(lookup_file)

        nonexistent = os.path.join(tmp_dir, '不存在.pdf')
        rows = process_pdf_file(nonexistent, lookup_data)
        assert rows == []

    def test_process_pdf_file_non_pdf(self, tmp_dir):
        from pdf_bank_parser import process_pdf_file

        lookup_file = os.path.join(tmp_dir, '主体查找表.xlsx')
        _create_lookup_table(lookup_file)
        lookup_data = bankcheck.load_lookup_table(lookup_file)

        not_pdf = os.path.join(tmp_dir, 'test.txt')
        with open(not_pdf, 'w') as f:
            f.write('not a pdf')
        rows = process_pdf_file(not_pdf, lookup_data)
        assert rows == []


class TestPdfParseResult:
    def test_result_dataclass_defaults(self):
        from pdf_bank_parser import PdfParseResult
        r = PdfParseResult()
        assert r.rows == []
        assert r.bank_name is None
        assert r.account is None
        assert r.page_count == 0
        assert r.method == 'table'
        assert r.raw_tables == []
        assert r.error is None

    def test_result_dataclass_custom(self):
        from pdf_bank_parser import PdfParseResult
        r = PdfParseResult(
            rows=[{'银行': '测试'}],
            bank_name='测试银行',
            account='123456',
            page_count=3,
            method='ocr',
            error=None,
        )
        assert len(r.rows) == 1
        assert r.bank_name == '测试银行'
        assert r.account == '123456'
        assert r.page_count == 3
        assert r.method == 'ocr'


class TestColumnKeywords:
    def test_column_keywords_required_fields(self):
        from pdf_bank_parser import COLUMN_KEYWORDS
        for required in ['trade_date', 'payment', 'receipt', 'balance', 'summary', 'counterpart']:
            assert required in COLUMN_KEYWORDS
            assert isinstance(COLUMN_KEYWORDS[required], list)
            assert len(COLUMN_KEYWORDS[required]) > 0
