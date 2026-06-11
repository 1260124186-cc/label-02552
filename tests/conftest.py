import os
import sys
import shutil
import tempfile

import openpyxl
import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'backend'))

import bankcheck


@pytest.fixture(autouse=True)
def init_logging():
    bankcheck.setup_logging()


@pytest.fixture
def tmp_dir():
    d = tempfile.mkdtemp(prefix='bankcheck_test_')
    yield d
    shutil.rmtree(d, ignore_errors=True)


def _create_beijing_bank_excel(path, account='01090312345678901', rows=None):
    if rows is None:
        rows = [
            [1, '2024-01-05', 'CNY', 50000, None, 1500000, '供应商A公司', '622001234', '工商银行', '转账', '001', '采购付款', None, None, None, 'BJ20240105001'],
            [2, '2024-01-10', 'CNY', None, 80000, 1580000, '客户B公司', '622005678', '建设银行', '转账', '002', '销售收款', None, None, None, 'BJ20240110002'],
        ]
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '交易明细'
    ws['A1'] = '北京银行交易明细'
    ws['B2'] = account
    headers = ['序号', '交易日期', '币种', '支出金额', '收入金额', '余额', '对方户名', '对方账号', '对方行名', '凭证种类', '凭证号码', '摘要', '备注1', '备注2', '备注3', '交易流水号']
    for c, h in enumerate(headers, 1):
        ws.cell(row=3, column=c, value=h)
    for i, row_data in enumerate(rows):
        for j, val in enumerate(row_data):
            ws.cell(row=4 + i, column=j + 1, value=val)
    wb.save(path)
    wb.close()
    return path


def _create_beijing_bank_multi_sheet_excel(path, sheets=None):
    if sheets is None:
        sheets = [
            {
                'title': '1月',
                'account': '01090312345678901',
                'rows': [
                    [1, '2024-01-05', 'CNY', 50000, None, 1500000, '供应商A公司', '622001234', '工商银行', '转账', '001', '采购付款', None, None, None, 'BJ20240105001'],
                    [2, '2024-01-10', 'CNY', None, 80000, 1580000, '客户B公司', '622005678', '建设银行', '转账', '002', '销售收款', None, None, None, 'BJ20240110002'],
                ],
            },
            {
                'title': '2月',
                'account': '01090312345678901',
                'rows': [
                    [1, '2024-02-03', 'CNY', 30000, None, 1550000, '供应商C公司', '622009999', '农业银行', '转账', '003', '材料采购', None, None, None, 'BJ20240203003'],
                    [2, '2024-02-15', 'CNY', None, 60000, 1610000, '客户D公司', '622008888', '中国银行', '转账', '004', '服务收款', None, None, None, 'BJ20240215004'],
                ],
            },
        ]
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    headers = ['序号', '交易日期', '币种', '支出金额', '收入金额', '余额', '对方户名', '对方账号', '对方行名', '凭证种类', '凭证号码', '摘要', '备注1', '备注2', '备注3', '交易流水号']
    for sheet_data in sheets:
        ws = wb.create_sheet(title=sheet_data['title'])
        ws['A1'] = '北京银行交易明细'
        ws['B2'] = sheet_data.get('account')
        for c, h in enumerate(headers, 1):
            ws.cell(row=3, column=c, value=h)
        for i, row_data in enumerate(sheet_data.get('rows', [])):
            for j, val in enumerate(row_data):
                ws.cell(row=4 + i, column=j + 1, value=val)
    wb.save(path)
    wb.close()
    return path


def _create_east_asia_bank_excel(path, account='38812345678', rows=None):
    if rows is None:
        rows = [
            ['2024-01-03', '09:30:00', 'CNY', 20000, None, 100, 0, 0, 480000, '转账', 'EA20240103001', '向 张三 付款-货款'],
            ['2024-01-08', '14:15:00', 'CNY', None, 35000, 0, 0, 0, 515000, '转账', 'EA20240108002', '收到 李四 汇款-服务费'],
        ]
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '交易明细'
    ws['A1'] = '账号'
    ws['B1'] = account
    ws['A2'] = '东亚银行交易明细'
    ws['A3'] = '查询期间：2024-01-01 至 2024-01-31'
    headers = ['交易日期', '交易时间', '币种', '支出金额', '收入金额', '手续费', '利息', '税费', '余额', '交易类型', '交易流水号', '交易描述/对方']
    for c, h in enumerate(headers, 1):
        ws.cell(row=4, column=c, value=h)
    for i, row_data in enumerate(rows):
        for j, val in enumerate(row_data):
            ws.cell(row=5 + i, column=j + 1, value=val)
    wb.save(path)
    wb.close()
    return path


def _create_beijing_bank_excel_with_extra_sheets(path, data_rows=None, extra_sheet_titles=None):
    if data_rows is None:
        data_rows = [
            [1, '2024-01-05', 'CNY', 50000, None, 1500000, '供应商A公司', '622001234', '工商银行', '转账', '001', '采购付款', None, None, None, 'BJ20240105001'],
        ]
    if extra_sheet_titles is None:
        extra_sheet_titles = ['说明', '汇总']
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '交易明细'
    ws['A1'] = '北京银行交易明细'
    ws['B2'] = '01090312345678901'
    headers = ['序号', '交易日期', '币种', '支出金额', '收入金额', '余额', '对方户名', '对方账号', '对方行名', '凭证种类', '凭证号码', '摘要', '备注1', '备注2', '备注3', '交易流水号']
    for c, h in enumerate(headers, 1):
        ws.cell(row=3, column=c, value=h)
    for i, row_data in enumerate(data_rows):
        for j, val in enumerate(row_data):
            ws.cell(row=4 + i, column=j + 1, value=val)
    for title in extra_sheet_titles:
        extra_ws = wb.create_sheet(title=title)
        extra_ws['A1'] = '此表为说明页，无流水数据'
    wb.save(path)
    wb.close()
    return path


def _create_lookup_table(path, mappings=None):
    if mappings is None:
        mappings = [
            ('北京XX科技有限公司', '01090312345678901'),
            ('上海YY贸易有限公司', '38812345678'),
        ]
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '主体映射'
    ws['A1'] = '主体名称'
    ws['B1'] = '银行账号'
    for i, (subject, account) in enumerate(mappings, 2):
        ws.cell(row=i, column=1, value=subject)
        ws.cell(row=i, column=2, value=account)
    wb.save(path)
    wb.close()
    return path
