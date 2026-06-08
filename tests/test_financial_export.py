"""
财务软件对接导出模块单元测试
"""
import os
import sys
import tempfile
import shutil
from datetime import datetime

import openpyxl
import pandas as pd
import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'backend'))

import bankcheck


@pytest.fixture
def tmp_dir():
    d = tempfile.mkdtemp(prefix='financial_export_test_')
    yield d
    shutil.rmtree(d, ignore_errors=True)


def _create_test_total_table(path, records=None):
    """创建测试用总表"""
    if records is None:
        records = [
            {
                '唯一id': 'TEST001',
                '银行': '北京银行',
                '银行账号': '01090312345678901',
                '主体': '北京XX科技有限公司',
                '交易日期': '2024-01-05',
                '付款': -50000.0,
                '收款': None,
                '摘要': '采购付款-办公设备',
                '对方户名': '供应商A公司',
                '余额': 1500000.0,
                '交易流水号': 'BJ20240105001',
            },
            {
                '唯一id': 'TEST002',
                '银行': '北京银行',
                '银行账号': '01090312345678901',
                '主体': '北京XX科技有限公司',
                '交易日期': '2024-01-10',
                '付款': None,
                '收款': 80000.0,
                '摘要': '销售收款-产品销售',
                '对方户名': '客户B公司',
                '余额': 1580000.0,
                '交易流水号': 'BJ20240110002',
            },
            {
                '唯一id': 'TEST003',
                '银行': '东亚银行',
                '银行账号': '38812345678',
                '主体': '上海YY贸易有限公司',
                '交易日期': '2024-01-15',
                '付款': -1500.0,
                '收款': None,
                '摘要': '手续费',
                '对方户名': '银行手续费',
                '余额': 513500.0,
                '交易流水号': 'EA20240115003',
            },
            {
                '唯一id': 'TEST004',
                '银行': '东亚银行',
                '银行账号': '38812345678',
                '主体': '上海YY贸易有限公司',
                '交易日期': '2024-01-20',
                '付款': -3000.0,
                '收款': None,
                '摘要': '差旅费报销',
                '对方户名': '员工张三',
                '余额': 510500.0,
                '交易流水号': 'EA20240120004',
            },
        ]
    
    columns = ['唯一id', '银行', '银行账号', '主体', '交易日期', 
               '付款', '收款', '摘要', '对方户名', '余额', '交易流水号']
    df = pd.DataFrame(records, columns=columns)
    df.to_excel(path, index=False, engine='openpyxl')
    return path


class TestDataClasses:
    """测试数据类"""
    
    def test_standard_transaction_creation(self):
        """测试 StandardTransaction 创建"""
        txn = bankcheck.StandardTransaction(
            transaction_date=datetime(2024, 1, 5),
            voucher_number='记-202401-0001',
            summary='采购付款',
            subject_code='1002',
            subject_name='银行存款',
            debit_amount=50000.0,
            credit_amount=0.0,
            bank_account='01090312345678901',
            bank_name='北京银行',
        )
        assert txn.summary == '采购付款'
        assert txn.subject_code == '1002'
        assert txn.debit_amount == 50000.0
        assert txn.transaction_date == datetime(2024, 1, 5)
    
    def test_standard_voucher_creation(self):
        """测试 StandardVoucher 创建"""
        entry1 = bankcheck.StandardTransaction(
            summary='采购付款',
            subject_code='1002',
            subject_name='银行存款',
            debit_amount=50000.0,
        )
        entry2 = bankcheck.StandardTransaction(
            summary='采购付款',
            subject_code='2202',
            subject_name='应付账款',
            credit_amount=50000.0,
        )
        voucher = bankcheck.StandardVoucher(
            voucher_date=datetime(2024, 1, 5),
            voucher_number='记-202401-0001',
            entries=[entry1, entry2],
        )
        assert len(voucher.entries) == 2
        assert voucher.entries[0].debit_amount == 50000.0
        assert voucher.entries[1].credit_amount == 50000.0


class TestUtilityFunctions:
    """测试工具函数"""
    
    def test_normalize_date_string(self):
        """测试日期规范化 - 字符串"""
        assert bankcheck._normalize_date('2024-01-05') == datetime(2024, 1, 5)
        assert bankcheck._normalize_date('2024/01/05') == datetime(2024, 1, 5)
        assert bankcheck._normalize_date('2024.01.05') == datetime(2024, 1, 5)
        assert bankcheck._normalize_date('20240105') == datetime(2024, 1, 5)
        assert bankcheck._normalize_date('2024-01-05 10:30:00') == datetime(2024, 1, 5, 10, 30, 0)
    
    def test_normalize_date_datetime(self):
        """测试日期规范化 - datetime"""
        dt = datetime(2024, 1, 5, 10, 30)
        assert bankcheck._normalize_date(dt) == dt
    
    def test_normalize_date_none(self):
        """测试日期规范化 - None"""
        assert bankcheck._normalize_date(None) is None
    
    def test_normalize_date_invalid(self):
        """测试日期规范化 - 无效格式"""
        assert bankcheck._normalize_date('invalid_date') is None
    
    def test_format_date_for_export(self):
        """测试日期格式化导出"""
        dt = datetime(2024, 1, 5)
        assert bankcheck._format_date_for_export(dt) == '2024-01-05'
        assert bankcheck._format_date_for_export(dt, '%Y/%m/%d') == '2024/01/05'
        assert bankcheck._format_date_for_export(None) == ''
    
    def test_format_amount(self):
        """测试金额格式化"""
        assert bankcheck._format_amount(100.123) == 100.12
        assert bankcheck._format_amount(100.126) == 100.13
        assert bankcheck._format_amount(None) == 0.0
        assert bankcheck._format_amount('100.5') == 100.5
        assert bankcheck._format_amount('invalid') == 0.0


class TestStandardConversion:
    """测试总表到标准化凭证的转换"""
    
    def test_load_total_table(self, tmp_dir):
        """测试加载总表"""
        total_path = os.path.join(tmp_dir, '总表.xlsx')
        _create_test_total_table(total_path)
        
        records = bankcheck.load_total_table(total_path)
        assert len(records) == 4
        assert records[0]['交易流水号'] == 'BJ20240105001'
        assert pd.isna(records[0]['收款']) or records[0]['收款'] is None
        assert float(records[1]['收款']) == 80000.0
    
    def test_load_total_table_not_exists(self, tmp_dir):
        """测试加载不存在的总表"""
        records = bankcheck.load_total_table(os.path.join(tmp_dir, 'not_exists.xlsx'))
        assert records == []
    
    def test_total_to_standard_transactions(self, tmp_dir):
        """测试总表转标准化凭证"""
        total_path = os.path.join(tmp_dir, '总表.xlsx')
        _create_test_total_table(total_path)
        records = bankcheck.load_total_table(total_path)
        
        vouchers = bankcheck.total_to_standard_transactions(records, operator='张三')
        
        assert len(vouchers) == 4
        assert vouchers[0].prepared_by == '张三'
        assert vouchers[0].voucher_type == '记'
        
        for voucher in vouchers:
            assert len(voucher.entries) == 2
            total_debit = sum(e.debit_amount for e in voucher.entries)
            total_credit = sum(e.credit_amount for e in voucher.entries)
            assert abs(total_debit - total_credit) < 0.01
    
    def test_payment_transaction_conversion(self, tmp_dir):
        """测试付款交易转换"""
        records = [{
            '交易日期': '2024-01-05',
            '付款': -50000.0,
            '收款': None,
            '摘要': '采购付款',
            '对方户名': '供应商A',
            '银行账号': '12345',
            '银行': '北京银行',
            '主体': 'XX公司',
            '余额': 1000000.0,
            '交易流水号': 'TXN001',
        }]
        
        vouchers = bankcheck.total_to_standard_transactions(records)
        assert len(vouchers) == 1
        
        bank_entry = vouchers[0].entries[0]
        other_entry = vouchers[0].entries[1]
        
        assert bank_entry.subject_name == '银行存款'
        assert bank_entry.credit_amount == 50000.0
        assert bank_entry.direction == '贷'
        
        assert other_entry.debit_amount == 50000.0
        assert other_entry.subject_code == '2202'
        assert other_entry.subject_name == '应付账款'
    
    def test_receipt_transaction_conversion(self, tmp_dir):
        """测试收款交易转换"""
        records = [{
            '交易日期': '2024-01-10',
            '付款': None,
            '收款': 80000.0,
            '摘要': '销售收款-主营业务收入',
            '对方户名': '客户B',
            '银行账号': '12345',
            '银行': '北京银行',
            '主体': 'XX公司',
            '余额': 1080000.0,
            '交易流水号': 'TXN002',
        }]
        
        vouchers = bankcheck.total_to_standard_transactions(records)
        assert len(vouchers) == 1
        
        bank_entry = vouchers[0].entries[0]
        other_entry = vouchers[0].entries[1]
        
        assert bank_entry.subject_name == '银行存款'
        assert bank_entry.debit_amount == 80000.0
        assert bank_entry.direction == '借'
        
        assert other_entry.credit_amount == 80000.0
        assert other_entry.subject_code == '6001'
        assert other_entry.subject_name == '主营业务收入'
    
    def test_expense_type_recognition(self):
        """测试费用类型智能识别"""
        test_cases = [
            ('差旅费报销', '6602', '管理费用'),
            ('办公费', '6602', '管理费用'),
            ('业务招待费', '6602', '管理费用'),
            ('广告费', '6601', '销售费用'),
            ('推广费', '6601', '销售费用'),
            ('手续费', '6603', '财务费用'),
            ('利息支出', '6603', '财务费用'),
            ('主营业务成本', '6401', '主营业务成本'),
        ]
        
        for summary, expected_code, expected_name in test_cases:
            records = [{
                '交易日期': '2024-01-01',
                '付款': -1000.0,
                '收款': None,
                '摘要': summary,
                '对方户名': '测试',
                '银行账号': '123',
                '银行': '测试银行',
                '主体': '测试公司',
                '余额': 10000.0,
                '交易流水号': 'TEST',
            }]
            vouchers = bankcheck.total_to_standard_transactions(records)
            other_entry = vouchers[0].entries[1]
            assert other_entry.subject_code == expected_code, f"摘要: {summary}"
            assert other_entry.subject_name == expected_name, f"摘要: {summary}"
    
    def test_empty_amount_skipped(self):
        """测试无金额的记录被跳过"""
        records = [{
            '交易日期': '2024-01-01',
            '付款': None,
            '收款': None,
            '摘要': '无金额记录',
            '对方户名': '测试',
            '银行账号': '123',
            '银行': '测试银行',
            '主体': '测试公司',
            '余额': 10000.0,
            '交易流水号': 'TEST',
        }]
        vouchers = bankcheck.total_to_standard_transactions(records)
        assert len(vouchers) == 0


class TestYonyouExport:
    """测试用友凭证导出"""
    
    def test_export_yonyou_voucher(self, tmp_dir):
        """测试导出用友凭证"""
        records = [
            {
                '交易日期': '2024-01-05',
                '付款': -50000.0,
                '收款': None,
                '摘要': '采购付款',
                '对方户名': '供应商A',
                '银行账号': '12345',
                '银行': '北京银行',
                '主体': 'XX公司',
                '余额': 1000000.0,
                '交易流水号': 'TXN001',
            },
            {
                '交易日期': '2024-01-10',
                '付款': None,
                '收款': 80000.0,
                '摘要': '销售收款',
                '对方户名': '客户B',
                '银行账号': '12345',
                '银行': '北京银行',
                '主体': 'XX公司',
                '余额': 1080000.0,
                '交易流水号': 'TXN002',
            },
        ]
        
        vouchers = bankcheck.total_to_standard_transactions(records, operator='张三')
        output_path = os.path.join(tmp_dir, '用友凭证.xlsx')
        
        result = bankcheck.export_yonyou_voucher(vouchers, output_path)
        
        assert result == output_path
        assert os.path.exists(output_path)
        
        df = pd.read_excel(output_path, engine='openpyxl')
        assert len(df) == 4
        
        expected_columns = [
            '凭证类别字', '凭证编号', '凭证日期', '附单据数',
            '制单人', '审核人', '记账人',
            '摘要', '科目编码', '科目名称',
            '借方金额', '贷方金额',
            '部门编码', '部门名称',
            '个人编码', '个人名称',
            '客户编码', '客户名称',
            '供应商编码', '供应商名称',
            '项目编码', '项目名称',
            '银行账号', '票据号',
        ]
        assert list(df.columns) == expected_columns
        
        assert df['凭证类别字'].iloc[0] == '记'
        assert df['制单人'].iloc[0] == '张三'
        assert str(df['科目编码'].iloc[0]) == '1002'
        assert df['科目名称'].iloc[0] == '银行存款'
        assert float(df['贷方金额'].iloc[0]) == 50000.0
        assert float(df['借方金额'].iloc[1]) == 50000.0
    
    def test_export_yonyou_empty_vouchers(self, tmp_dir):
        """测试导出空凭证列表"""
        output_path = os.path.join(tmp_dir, '空.xlsx')
        result = bankcheck.export_yonyou_voucher([], output_path)
        assert result is None


class TestKingdeeExport:
    """测试金蝶凭证导出"""
    
    def test_export_kingdee_voucher(self, tmp_dir):
        """测试导出金蝶凭证"""
        records = [{
            '交易日期': '2024-01-05',
            '付款': -50000.0,
            '收款': None,
            '摘要': '采购付款',
            '对方户名': '供应商A',
            '银行账号': '12345',
            '银行': '北京银行',
            '主体': 'XX公司',
            '余额': 1000000.0,
            '交易流水号': 'TXN001',
        }]
        
        vouchers = bankcheck.total_to_standard_transactions(records, operator='李四')
        output_path = os.path.join(tmp_dir, '金蝶凭证.xlsx')
        
        result = bankcheck.export_kingdee_voucher(vouchers, output_path)
        
        assert result == output_path
        assert os.path.exists(output_path)
        
        df = pd.read_excel(output_path, engine='openpyxl')
        assert len(df) == 2
        
        expected_columns = [
            '凭证字号', '凭证日期', '附件数',
            '制单人', '审核人', '过账人',
            '摘要', '科目代码', '科目名称',
            '借方金额', '贷方金额',
            '核算项目类别', '核算项目代码', '核算项目名称',
            '币别', '汇率', '原币金额',
            '结算方式', '结算号', '业务日期',
        ]
        assert list(df.columns) == expected_columns
        
        assert df['制单人'].iloc[0] == '李四'
        assert str(df['科目代码'].iloc[0]) == '1002'
        assert df['科目名称'].iloc[0] == '银行存款'
        assert df['币别'].iloc[0] == '人民币'
        assert float(df['汇率'].iloc[0]) == 1.0
        assert df['结算方式'].iloc[0] == '银行转账'
        assert df['结算号'].iloc[0] == 'TXN001'


class TestBankJournalExport:
    """测试银行日记账导出"""
    
    def test_export_bank_journal(self, tmp_dir):
        """测试导出银行日记账"""
        records = [
            {
                '交易日期': '2024-01-05',
                '付款': -50000.0,
                '收款': None,
                '摘要': '采购付款',
                '对方户名': '供应商A',
                '银行账号': '12345',
                '银行': '北京银行',
                '主体': 'XX公司',
                '余额': 1000000.0,
                '交易流水号': 'TXN001',
            },
            {
                '交易日期': '2024-01-10',
                '付款': None,
                '收款': 80000.0,
                '摘要': '销售收款',
                '对方户名': '客户B',
                '银行账号': '12345',
                '银行': '北京银行',
                '主体': 'XX公司',
                '余额': 1080000.0,
                '交易流水号': 'TXN002',
            },
        ]
        
        vouchers = bankcheck.total_to_standard_transactions(records)
        output_path = os.path.join(tmp_dir, '日记账.xlsx')
        
        result = bankcheck.export_bank_journal(vouchers, output_path)
        
        assert result == output_path
        assert os.path.exists(output_path)
        
        df = pd.read_excel(output_path, engine='openpyxl')
        assert len(df) == 2
        
        expected_columns = [
            '日期', '凭证号', '摘要', '对方科目',
            '借方金额', '贷方金额', '方向', '余额',
            '银行账号', '开户银行', '核算主体',
            '对方户名', '交易流水号', '备注',
        ]
        assert list(df.columns) == expected_columns
        
        assert df['方向'].iloc[0] == '贷'
        assert df['贷方金额'].iloc[0] == 50000.0
        assert df['借方金额'].iloc[0] == 0.0
        
        assert df['方向'].iloc[1] == '借'
        assert df['借方金额'].iloc[1] == 80000.0
        assert df['贷方金额'].iloc[1] == 0.0
        
        assert df['余额'].iloc[1] == 1080000.0
        assert df['开户银行'].iloc[0] == '北京银行'
        assert df['交易流水号'].iloc[0] == 'TXN001'


class TestExportEntry:
    """测试统一导出入口"""
    
    def test_export_financial_template_yonyou(self, tmp_dir):
        """测试用友模板导出入口"""
        total_path = os.path.join(tmp_dir, '总表.xlsx')
        _create_test_total_table(total_path)
        
        result = bankcheck.export_financial_template(
            total_path=total_path,
            template_type='yonyou_voucher',
            output_dir=tmp_dir,
            operator='测试员',
        )
        
        assert result['success'] is True
        assert result['template_type'] == 'yonyou_voucher'
        assert result['template_name'] == '用友凭证导入模板'
        assert result['voucher_count'] == 4
        assert result['entry_count'] == 8
        assert os.path.exists(result['output_path'])
        assert '用友凭证导入' in result['output_path']
    
    def test_export_financial_template_kingdee(self, tmp_dir):
        """测试金蝶模板导出入口"""
        total_path = os.path.join(tmp_dir, '总表.xlsx')
        _create_test_total_table(total_path)
        
        result = bankcheck.export_financial_template(
            total_path=total_path,
            template_type='kingdee_voucher',
            output_dir=tmp_dir,
        )
        
        assert result['success'] is True
        assert '金蝶凭证导入' in result['output_path']
    
    def test_export_financial_template_bank_journal(self, tmp_dir):
        """测试银行日记账模板导出入口"""
        total_path = os.path.join(tmp_dir, '总表.xlsx')
        _create_test_total_table(total_path)
        
        result = bankcheck.export_financial_template(
            total_path=total_path,
            template_type='bank_journal',
            output_dir=tmp_dir,
        )
        
        assert result['success'] is True
        assert '银行日记账' in result['output_path']
    
    def test_export_financial_template_invalid_type(self, tmp_dir):
        """测试无效模板类型"""
        total_path = os.path.join(tmp_dir, '总表.xlsx')
        _create_test_total_table(total_path)
        
        with pytest.raises(ValueError, match='不支持的导出模板类型'):
            bankcheck.export_financial_template(
                total_path=total_path,
                template_type='invalid_type',
                output_dir=tmp_dir,
            )
    
    def test_export_financial_template_empty_total(self, tmp_dir):
        """测试空总表导出"""
        empty_total = os.path.join(tmp_dir, '空总表.xlsx')
        df = pd.DataFrame(columns=['唯一id', '银行', '银行账号'])
        df.to_excel(empty_total, index=False, engine='openpyxl')
        
        result = bankcheck.export_financial_template(
            total_path=empty_total,
            template_type='yonyou_voucher',
            output_dir=tmp_dir,
        )
        
        assert result['success'] is False
        assert result['error'] == '总表无数据'
    
    def test_export_financial_template_output_dir_default(self, tmp_dir):
        """测试默认输出目录"""
        total_path = os.path.join(tmp_dir, '总表.xlsx')
        _create_test_total_table(total_path)
        
        result = bankcheck.export_financial_template(
            total_path=total_path,
            template_type='yonyou_voucher',
        )
        
        assert result['success'] is True
        assert os.path.dirname(result['output_path']) == tmp_dir


class TestExportTemplates:
    """测试模板配置"""
    
    def test_export_templates_defined(self):
        """测试导出模板定义"""
        templates = bankcheck.FINANCIAL_EXPORT_TEMPLATES
        assert 'yonyou_voucher' in templates
        assert 'kingdee_voucher' in templates
        assert 'bank_journal' in templates
        
        assert templates['yonyou_voucher']['name'] == '用友凭证导入模板'
        assert templates['kingdee_voucher']['name'] == '金蝶凭证导入模板'
        assert templates['bank_journal']['name'] == '银行日记账模板'
    
    def test_default_account_mapping(self):
        """测试默认科目映射"""
        mapping = bankcheck.DEFAULT_ACCOUNT_MAPPING
        assert mapping['bank_deposit']['code'] == '1002'
        assert mapping['bank_deposit']['name'] == '银行存款'
        assert mapping['accounts_receivable']['code'] == '1122'
        assert mapping['accounts_payable']['code'] == '2202'
        assert mapping['operating_revenue']['code'] == '6001'
        assert mapping['management_expense']['code'] == '6602'


class TestAccountMapping:
    """测试自定义科目映射"""
    
    def test_custom_account_mapping(self, tmp_dir):
        """测试使用自定义科目映射"""
        custom_mapping = {
            'bank_deposit': {'code': '100201', 'name': '银行存款-北京银行'},
            'accounts_receivable': {'code': '112201', 'name': '应收账款-客户'},
            'accounts_payable': {'code': '220201', 'name': '应付账款-供应商'},
            'operating_revenue': {'code': '600101', 'name': '主营业务收入-产品'},
            'operating_cost': {'code': '640101', 'name': '主营业务成本-产品'},
            'management_expense': {'code': '660201', 'name': '管理费用-办公'},
            'sales_expense': {'code': '660101', 'name': '销售费用-广告'},
            'financial_expense': {'code': '660301', 'name': '财务费用-手续费'},
            'cash': {'code': '1001', 'name': '库存现金'},
        }
        
        records = [{
            '交易日期': '2024-01-05',
            '付款': -50000.0,
            '收款': None,
            '摘要': '采购付款',
            '对方户名': '供应商A',
            '银行账号': '12345',
            '银行': '北京银行',
            '主体': 'XX公司',
            '余额': 1000000.0,
            '交易流水号': 'TXN001',
        }]
        
        vouchers = bankcheck.total_to_standard_transactions(records, account_mapping=custom_mapping)
        
        assert vouchers[0].entries[0].subject_code == '100201'
        assert vouchers[0].entries[0].subject_name == '银行存款-北京银行'
        assert vouchers[0].entries[1].subject_code == '220201'
        assert vouchers[0].entries[1].subject_name == '应付账款-供应商'
