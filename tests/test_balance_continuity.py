"""
余额连续性校验模块单元测试

覆盖场景：
1. 连续支出场景（付款为负数，余额递减）- 无误报
2. 连续收款场景（收款为正数，余额递增）- 无误报
3. 正负付款混用场景（收付款交替）- 无误报
4. 真实异常断裂场景 - 正确识别
5. 多账号混合场景 - 账号间独立校验
6. 边界场景（空数据、单条记录、无有效日期）
7. 乱序数据排序后校验
8. 跨月数据校验
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
from bankcheck import (
    check_balance_continuity,
    export_balance_check_result,
    generate_balance_check_from_records,
    BalanceCheckResult,
    BalanceBreakRecord,
)


@pytest.fixture
def tmp_dir():
    d = tempfile.mkdtemp(prefix='balance_check_test_')
    yield d
    shutil.rmtree(d, ignore_errors=True)


def _record(unique_id, account, date, payment=None, receipt=None, balance=0.0,
            bank='北京银行', subject='北京XX科技有限公司',
            summary='', counterpart='', txn_id=''):
    """构造标准化交易记录，付款约定为负数（与北京银行 payment_sign=negative 一致）"""
    return {
        '唯一id': unique_id,
        '银行': bank,
        '银行账号': account,
        '主体': subject,
        '交易日期': date,
        '付款': payment,
        '收款': receipt,
        '摘要': summary,
        '对方户名': counterpart,
        '余额': balance,
        '交易流水号': txn_id or unique_id,
    }


class TestContinuousPayment:
    """连续支出场景：多笔付款（负数），余额逐笔递减 - 无误报"""

    def test_five_consecutive_payments_no_break(self):
        """连续5笔支出，付款为负数，余额正确递减"""
        records = [
            _record('T001', '62220001', '2024-01-01', balance=100000.0, txn_id='T001'),
            _record('T002', '62220001', '2024-01-02', payment=-1000.0, balance=99000.0, txn_id='T002'),
            _record('T003', '62220001', '2024-01-03', payment=-2500.50, balance=96499.50, txn_id='T003'),
            _record('T004', '62220001', '2024-01-04', payment=-500.0, balance=95999.50, txn_id='T004'),
            _record('T005', '62220001', '2024-01-05', payment=-10000.0, balance=85999.50, txn_id='T005'),
        ]
        result = check_balance_continuity(records)
        assert result.break_count == 0
        assert len(result.break_records) == 0
        assert len(result.accounts_with_breaks) == 0
        assert result.checked_accounts == 1

    def test_large_amount_consecutive_payments(self):
        """大金额连续支出，确保浮点精度处理正确"""
        records = [
            _record('L001', '62220002', '2024-03-01', balance=10000000.0, txn_id='L001'),
            _record('L002', '62220002', '2024-03-02', payment=-1234567.89, balance=8765432.11, txn_id='L002'),
            _record('L003', '62220002', '2024-03-03', payment=-987654.32, balance=7777777.79, txn_id='L003'),
            _record('L004', '62220002', '2024-03-04', payment=-5555555.55, balance=2222222.24, txn_id='L004'),
        ]
        result = check_balance_continuity(records, tolerance=0.01)
        assert result.break_count == 0
        assert len(result.break_records) == 0

    def test_payment_and_none_receipt(self):
        """付款有值，收款为 None 的场景"""
        records = [
            _record('P001', '62220003', '2024-05-01', balance=50000.0, txn_id='P001'),
            _record('P002', '62220003', '2024-05-02', payment=-3000.0, receipt=None, balance=47000.0, txn_id='P002'),
            _record('P003', '62220003', '2024-05-03', payment=-500.0, receipt=None, balance=46500.0, txn_id='P003'),
        ]
        result = check_balance_continuity(records)
        assert result.break_count == 0


class TestContinuousReceipt:
    """连续收款场景：多笔收款（正数），余额逐笔递增 - 无误报"""

    def test_five_consecutive_receipts_no_break(self):
        """连续5笔收款，余额正确递增"""
        records = [
            _record('R001', '62220101', '2024-01-10', balance=50000.0, txn_id='R001'),
            _record('R002', '62220101', '2024-01-11', receipt=20000.0, balance=70000.0, txn_id='R002'),
            _record('R003', '62220101', '2024-01-12', receipt=5500.25, balance=75500.25, txn_id='R003'),
            _record('R004', '62220101', '2024-01-13', receipt=100000.0, balance=175500.25, txn_id='R004'),
            _record('R005', '62220101', '2024-01-14', receipt=3333.33, balance=178833.58, txn_id='R005'),
        ]
        result = check_balance_continuity(records)
        assert result.break_count == 0
        assert len(result.break_records) == 0

    def test_receipt_and_none_payment(self):
        """收款有值，付款为 None 的场景"""
        records = [
            _record('R101', '62220102', '2024-06-01', balance=10000.0, txn_id='R101'),
            _record('R102', '62220102', '2024-06-02', payment=None, receipt=8888.88, balance=18888.88, txn_id='R102'),
            _record('R103', '62220102', '2024-06-03', payment=None, receipt=9999.99, balance=28888.87, txn_id='R103'),
        ]
        result = check_balance_continuity(records)
        assert result.break_count == 0


class TestMixedPaymentReceipt:
    """正负付款混用场景：收付款交替 - 无误报"""

    def test_alternating_payment_receipt(self):
        """收支交替的日常经营场景"""
        records = [
            _record('M001', '62220201', '2024-02-01', balance=200000.0, txn_id='M001'),
            _record('M002', '62220201', '2024-02-02', payment=-15000.0, balance=185000.0, summary='采购付款', txn_id='M002'),
            _record('M003', '62220201', '2024-02-03', receipt=50000.0, balance=235000.0, summary='销售收款', txn_id='M003'),
            _record('M004', '62220201', '2024-02-04', payment=-8000.0, balance=227000.0, summary='办公用品', txn_id='M004'),
            _record('M005', '62220201', '2024-02-05', payment=-3000.0, balance=224000.0, summary='差旅费', txn_id='M005'),
            _record('M006', '62220201', '2024-02-06', receipt=120000.0, balance=344000.0, summary='客户回款', txn_id='M006'),
            _record('M007', '62220201', '2024-02-07', payment=-45000.0, balance=299000.0, summary='工资发放', txn_id='M007'),
            _record('M008', '62220201', '2024-02-08', receipt=15000.0, balance=314000.0, summary='服务费收入', txn_id='M008'),
        ]
        result = check_balance_continuity(records)
        assert result.break_count == 0
        assert len(result.break_records) == 0
        assert result.checked_accounts == 1

    def test_multiple_small_transactions(self):
        """多笔小额收支混合（如手续费、利息）"""
        records = [
            _record('MS01', '62220202', '2024-04-01', balance=10000.0, txn_id='MS01'),
            _record('MS02', '62220202', '2024-04-02', payment=-10.0, balance=9990.0, summary='短信服务费', txn_id='MS02'),
            _record('MS03', '62220202', '2024-04-03', receipt=5.50, balance=9995.50, summary='存款利息', txn_id='MS03'),
            _record('MS04', '62220202', '2024-04-04', payment=-25.0, balance=9970.50, summary='网银转账费', txn_id='MS04'),
            _record('MS05', '62220202', '2024-04-05', receipt=3000.0, balance=12970.50, summary='小额收款', txn_id='MS05'),
            _record('MS06', '62220202', '2024-04-06', payment=-100.0, balance=12870.50, summary='工本费', txn_id='MS06'),
        ]
        result = check_balance_continuity(records)
        assert result.break_count == 0

    def test_cross_month_mixed(self):
        """跨月收付款混合场景"""
        records = [
            _record('CM01', '62220203', '2024-01-30', balance=80000.0, txn_id='CM01'),
            _record('CM02', '62220203', '2024-01-31', payment=-12000.0, balance=68000.0, txn_id='CM02'),
            _record('CM03', '62220203', '2024-02-01', receipt=25000.0, balance=93000.0, txn_id='CM03'),
            _record('CM04', '62220203', '2024-02-28', payment=-8000.0, balance=85000.0, txn_id='CM04'),
            _record('CM05', '62220203', '2024-03-01', receipt=45000.0, balance=130000.0, txn_id='CM05'),
            _record('CM06', '62220203', '2024-03-15', payment=-20000.0, balance=110000.0, txn_id='CM06'),
        ]
        result = check_balance_continuity(records)
        assert result.break_count == 0


class TestGenuineBreakDetection:
    """真实异常断裂场景 - 正确识别异常"""

    def test_single_break_detected(self):
        """单笔余额断裂，应检测出1笔异常"""
        records = [
            _record('G001', '62220301', '2024-01-01', balance=100000.0, txn_id='G001'),
            _record('G002', '62220301', '2024-01-02', payment=-5000.0, balance=95000.0, txn_id='G002'),
            _record('G003', '62220301', '2024-01-03', receipt=20000.0, balance=100000.0, txn_id='G003'),
        ]
        result = check_balance_continuity(records)
        assert result.break_count == 1
        assert len(result.break_records) == 1
        assert '62220301' in result.accounts_with_breaks

        br = result.break_records[0]
        assert br.prev_balance == 95000.0
        assert br.receipt == 20000.0
        assert br.payment == 0.0
        assert br.expected_balance == 115000.0
        assert br.actual_balance == 100000.0
        assert br.diff_amount == -15000.0

    def test_multiple_breaks_in_same_account(self):
        """同一账号多笔断裂"""
        records = [
            _record('G101', '62220302', '2024-05-01', balance=50000.0, txn_id='G101'),
            _record('G102', '62220302', '2024-05-02', payment=-1000.0, balance=49000.0, txn_id='G102'),
            _record('G103', '62220302', '2024-05-03', receipt=5000.0, balance=50000.0, txn_id='G103'),
            _record('G104', '62220302', '2024-05-04', payment=-2000.0, balance=50000.0, txn_id='G104'),
        ]
        result = check_balance_continuity(records)
        assert result.break_count == 2
        assert len(result.accounts_with_breaks) == 1

    def test_break_at_boundary_tolerance(self):
        """容差边界测试：差异刚好等于容差时不算断裂"""
        records = [
            _record('G201', '62220303', '2024-06-01', balance=1000.0, txn_id='G201'),
            _record('G202', '62220303', '2024-06-02', payment=-100.0, balance=900.009, txn_id='G202'),
        ]
        result = check_balance_continuity(records, tolerance=0.01)
        assert result.break_count == 0

    def test_break_exceeds_tolerance(self):
        """差异超过容差时算断裂"""
        records = [
            _record('G211', '62220304', '2024-06-01', balance=1000.0, txn_id='G211'),
            _record('G212', '62220304', '2024-06-02', payment=-100.0, balance=900.02, txn_id='G212'),
        ]
        result = check_balance_continuity(records, tolerance=0.01)
        assert result.break_count == 1


class TestMultiAccount:
    """多账号混合场景 - 账号间独立校验"""

    def test_two_accounts_one_break(self):
        """两个账号，其中一个有断裂，另一个正常"""
        records = [
            _record('MA01', 'ACC001', '2024-01-01', balance=10000.0, txn_id='MA01'),
            _record('MA02', 'ACC001', '2024-01-02', payment=-1000.0, balance=9000.0, txn_id='MA02'),
            _record('MA03', 'ACC001', '2024-01-03', receipt=2000.0, balance=11000.0, txn_id='MA03'),
            _record('MA04', 'ACC002', '2024-01-01', balance=50000.0, txn_id='MA04'),
            _record('MA05', 'ACC002', '2024-01-02', payment=-5000.0, balance=45000.0, txn_id='MA05'),
            _record('MA06', 'ACC002', '2024-01-03', receipt=10000.0, balance=50000.0, txn_id='MA06'),
        ]
        result = check_balance_continuity(records)
        assert result.checked_accounts == 2
        assert result.break_count == 1
        assert 'ACC002' in result.accounts_with_breaks
        assert 'ACC001' not in result.accounts_with_breaks

        br = result.break_records[0]
        assert br.bank_account == 'ACC002'
        assert br.expected_balance == 55000.0
        assert br.actual_balance == 50000.0

    def test_three_accounts_all_clean(self):
        """三个账号全部正常"""
        records = []
        for i, acc in enumerate(['ACC101', 'ACC102', 'ACC103']):
            base = 100000.0 * (i + 1)
            records.append(_record(f'MA{i}1', acc, '2024-07-01', balance=base, txn_id=f'MA{i}1'))
            records.append(_record(f'MA{i}2', acc, '2024-07-02', payment=-1000.0, balance=base - 1000.0, txn_id=f'MA{i}2'))
            records.append(_record(f'MA{i}3', acc, '2024-07-03', receipt=3000.0, balance=base + 2000.0, txn_id=f'MA{i}3'))
        result = check_balance_continuity(records)
        assert result.total_accounts == 3
        assert result.checked_accounts == 3
        assert result.break_count == 0


class TestBoundaryCases:
    """边界场景测试"""

    def test_empty_records(self):
        """空数据"""
        result = check_balance_continuity([])
        assert result.total_accounts == 0
        assert result.break_count == 0
        assert result.break_records == []

    def test_single_record(self):
        """单条记录无法校验（无前后对比）"""
        records = [
            _record('S001', 'SACC01', '2024-01-01', balance=50000.0, txn_id='S001'),
        ]
        result = check_balance_continuity(records)
        assert result.checked_accounts == 1
        assert result.break_count == 0

    def test_records_without_account(self):
        """无账号的记录应被忽略"""
        records = [
            {'唯一id': 'X001', '银行': '北京银行', '银行账号': '', '交易日期': '2024-01-01', '余额': 1000.0, '交易流水号': 'X001'},
            {'唯一id': 'X002', '银行': '北京银行', '交易日期': '2024-01-02', '余额': 2000.0, '交易流水号': 'X002'},
        ]
        result = check_balance_continuity(records)
        assert result.total_accounts == 0

    def test_account_without_valid_dates(self):
        """账号无有效日期，应跳过"""
        records = [
            _record('N001', 'NACC01', None, balance=10000.0, txn_id='N001'),
            _record('N002', 'NACC01', '', balance=9000.0, txn_id='N002'),
        ]
        result = check_balance_continuity(records)
        assert result.skipped_accounts == 1
        assert result.checked_accounts == 0

    def test_zero_amounts(self):
        """付款和收款都是0或空的情况"""
        records = [
            _record('Z001', 'ZACC01', '2024-08-01', balance=50000.0, txn_id='Z001'),
            _record('Z002', 'ZACC01', '2024-08-02', payment=0.0, receipt=0.0, balance=50000.0, txn_id='Z002'),
            _record('Z003', 'ZACC01', '2024-08-03', payment=None, receipt=None, balance=50000.0, txn_id='Z003'),
            _record('Z004', 'ZACC01', '2024-08-04', receipt=1000.0, balance=51000.0, txn_id='Z004'),
        ]
        result = check_balance_continuity(records)
        assert result.break_count == 0


class TestOutOfOrderSorting:
    """乱序数据排序后校验"""

    def test_out_of_order_dates_sorted_correctly(self):
        """日期乱序的记录应先排序再校验"""
        records = [
            _record('O003', 'OACC01', '2024-01-03', receipt=5000.0, balance=15000.0, txn_id='O003'),
            _record('O001', 'OACC01', '2024-01-01', balance=10000.0, txn_id='O001'),
            _record('O004', 'OACC01', '2024-01-04', payment=-2000.0, balance=13000.0, txn_id='O004'),
            _record('O002', 'OACC01', '2024-01-02', payment=-0.0, balance=10000.0, txn_id='O002'),
        ]
        result = check_balance_continuity(records)
        assert result.break_count == 0

    def test_same_date_sorted_by_txn_id(self):
        """同日期的记录按交易流水号排序"""
        records = [
            _record('D001', 'DACC01', '2024-09-01', balance=10000.0, txn_id='TXN-003'),
            _record('D002', 'DACC01', '2024-09-01', payment=-1000.0, balance=9000.0, txn_id='TXN-001'),
            _record('D003', 'DACC01', '2024-09-01', receipt=5000.0, balance=14000.0, txn_id='TXN-002'),
        ]
        result = check_balance_continuity(records)
        assert result.break_count == 1


class TestExportFunction:
    """导出功能测试"""

    def test_export_clean_result(self, tmp_dir):
        """导出正常校验结果"""
        records = [
            _record('E001', 'EACC01', '2024-01-01', balance=50000.0, txn_id='E001'),
            _record('E002', 'EACC01', '2024-01-02', payment=-5000.0, balance=45000.0, txn_id='E002'),
            _record('E003', 'EACC01', '2024-01-03', receipt=10000.0, balance=55000.0, txn_id='E003'),
        ]
        result = check_balance_continuity(records)
        output = os.path.join(tmp_dir, 'clean_check.xlsx')
        export_balance_check_result(result, output, {'来源': '测试'})
        assert os.path.exists(output)
        assert os.path.getsize(output) > 0

    def test_export_with_breaks(self, tmp_dir):
        """导出含异常的校验结果"""
        records = [
            _record('E101', 'EACC02', '2024-01-01', balance=50000.0, txn_id='E101'),
            _record('E102', 'EACC02', '2024-01-02', payment=-5000.0, balance=40000.0, txn_id='E102'),
            _record('E103', 'EACC02', '2024-01-03', receipt=10000.0, balance=50000.0, txn_id='E103'),
        ]
        result = check_balance_continuity(records)
        output = os.path.join(tmp_dir, 'break_check.xlsx')
        path = export_balance_check_result(result, output)
        assert os.path.exists(path)

        xl = pd.ExcelFile(path, engine='openpyxl')
        sheets = xl.sheet_names
        assert '校验总览' in sheets
        assert '异常明细' in sheets
        assert '异常账号清单' in sheets

        detail_df = pd.read_excel(path, sheet_name='异常明细', engine='openpyxl')
        assert len(detail_df) == 1

    def test_generate_from_records(self, tmp_dir):
        """从记录列表生成报告"""
        records = [
            _record('E201', 'EACC03', '2024-02-01', balance=30000.0, txn_id='E201'),
            _record('E202', 'EACC03', '2024-02-02', payment=-3000.0, balance=27000.0, txn_id='E202'),
            _record('E203', 'EACC03', '2024-02-03', receipt=8000.0, balance=35000.0, txn_id='E203'),
        ]
        path = generate_balance_check_from_records(records, output_dir=tmp_dir)
        assert path is not None
        assert os.path.exists(path)


class TestBeijingBankRealScenario:
    """北京银行真实业务场景综合测试"""

    def test_full_month_operations(self):
        """模拟北京银行某账号一个月的完整流水（连续支出+收款+手续费）"""
        account = '01090312345678901'
        bank = '北京银行'
        subject = '北京XX科技有限公司'

        records = [
            _record('BJ001', account, '2024-03-01',
                    balance=500000.0, bank=bank, subject=subject,
                    summary='期初余额', txn_id='BJ240301001'),
            _record('BJ002', account, '2024-03-02', payment=-85000.0,
                    balance=415000.0, bank=bank, subject=subject,
                    summary='采购付款-原材料A', counterpart='供应商甲公司',
                    txn_id='BJ240302001'),
            _record('BJ003', account, '2024-03-03', payment=-120000.0,
                    balance=295000.0, bank=bank, subject=subject,
                    summary='采购付款-原材料B', counterpart='供应商乙公司',
                    txn_id='BJ240303001'),
            _record('BJ004', account, '2024-03-05', receipt=200000.0,
                    balance=495000.0, bank=bank, subject=subject,
                    summary='销售收款-产品A', counterpart='客户丙公司',
                    txn_id='BJ240305001'),
            _record('BJ005', account, '2024-03-06', payment=-25.0,
                    balance=494975.0, bank=bank, subject=subject,
                    summary='网银手续费', counterpart='北京银行',
                    txn_id='BJ240306001'),
            _record('BJ006', account, '2024-03-08', payment=-15000.0,
                    balance=479975.0, bank=bank, subject=subject,
                    summary='差旅费报销', counterpart='员工张三',
                    txn_id='BJ240308001'),
            _record('BJ007', account, '2024-03-10', receipt=350000.0,
                    balance=829975.0, bank=bank, subject=subject,
                    summary='销售收款-产品B', counterpart='客户丁公司',
                    txn_id='BJ240310001'),
            _record('BJ008', account, '2024-03-12', payment=-80000.0,
                    balance=749975.0, bank=bank, subject=subject,
                    summary='采购付款-办公设备', counterpart='供应商戊公司',
                    txn_id='BJ240312001'),
            _record('BJ009', account, '2024-03-15', payment=-180000.0,
                    balance=569975.0, bank=bank, subject=subject,
                    summary='工资发放（3月）', counterpart='员工工资代发',
                    txn_id='BJ240315001'),
            _record('BJ010', account, '2024-03-15', payment=-45000.0,
                    balance=524975.0, bank=bank, subject=subject,
                    summary='社保公积金缴纳', counterpart='北京市社保局',
                    txn_id='BJ240315002'),
            _record('BJ011', account, '2024-03-18', receipt=120000.0,
                    balance=644975.0, bank=bank, subject=subject,
                    summary='服务收入-咨询费', counterpart='客户己公司',
                    txn_id='BJ240318001'),
            _record('BJ012', account, '2024-03-20', payment=-3000.0,
                    balance=641975.0, bank=bank, subject=subject,
                    summary='办公用品采购', counterpart='超市',
                    txn_id='BJ240320001'),
            _record('BJ013', account, '2024-03-21', receipt=412.50,
                    balance=642387.50, bank=bank, subject=subject,
                    summary='季度存款利息', counterpart='北京银行',
                    txn_id='BJ240321001'),
            _record('BJ014', account, '2024-03-25', payment=-280000.0,
                    balance=362387.50, bank=bank, subject=subject,
                    summary='采购付款-大额订单', counterpart='供应商庚公司',
                    txn_id='BJ240325001'),
            _record('BJ015', account, '2024-03-28', receipt=500000.0,
                    balance=862387.50, bank=bank, subject=subject,
                    summary='股东投资款', counterpart='股东A',
                    txn_id='BJ240328001'),
            _record('BJ016', account, '2024-03-30', payment=-100.0,
                    balance=862287.50, bank=bank, subject=subject,
                    summary='账户管理费', counterpart='北京银行',
                    txn_id='BJ240330001'),
            _record('BJ017', account, '2024-03-31', receipt=180000.0,
                    balance=1042287.50, bank=bank, subject=subject,
                    summary='销售回款-月结', counterpart='客户辛公司',
                    txn_id='BJ240331001'),
        ]

        result = check_balance_continuity(records)
        assert result.break_count == 0, (
            f'北京银行月流水场景不应有误报，'
            f'但检测到 {result.break_count} 笔断裂。'
            f'断裂记录: {[(br.transaction_date, br.diff_amount) for br in result.break_records]}'
        )
        assert len(result.break_records) == 0
        assert len(result.accounts_with_breaks) == 0

    def test_intraday_multiple_txns(self):
        """同日多笔交易的连续性校验"""
        account = '01090398765432109'
        records = [
            _record('IN001', account, '2024-04-15', balance=100000.0, txn_id='T090001'),
            _record('IN002', account, '2024-04-15', payment=-5000.0, balance=95000.0, txn_id='T090500'),
            _record('IN003', account, '2024-04-15', receipt=30000.0, balance=125000.0, txn_id='T103000'),
            _record('IN004', account, '2024-04-15', payment=-800.0, balance=124200.0, txn_id='T110000'),
            _record('IN005', account, '2024-04-15', payment=-12000.0, balance=112200.0, txn_id='T140000'),
            _record('IN006', account, '2024-04-15', receipt=50000.0, balance=162200.0, txn_id='T153000'),
            _record('IN007', account, '2024-04-15', payment=-200.0, balance=162000.0, txn_id='T170000'),
        ]
        result = check_balance_continuity(records)
        assert result.break_count == 0


class TestFormulaCorrectness:
    """公式正确性专项验证

    核心验证：付款存储为负数时，公式必须是 prev + receipt + payment
    如果使用 prev + receipt - payment 则会导致连续支出场景误报
    """

    def test_formula_with_negative_payment_should_use_plus(self):
        """
        付款为负数场景的公式验证：

        初始余额: 1000
        T1: 付款 -100 → 余额应为 900
        T2: 付款 -200 → 余额应为 700

        正确公式: prev + receipt + payment
          T1: 1000 + 0 + (-100) = 900 ✓
          T2:  900 + 0 + (-200) = 700 ✓

        错误公式: prev + receipt - payment
          T1: 1000 + 0 - (-100) = 1100 ✗ (误报)
          T2:  900 + 0 - (-200) = 1100 ✗ (误报)
        """
        records = [
            _record('F001', 'FORMULA01', '2024-01-01', balance=1000.0, txn_id='F001'),
            _record('F002', 'FORMULA01', '2024-01-02', payment=-100.0, balance=900.0, txn_id='F002'),
            _record('F003', 'FORMULA01', '2024-01-03', payment=-200.0, balance=700.0, txn_id='F003'),
        ]
        result = check_balance_continuity(records)
        assert result.break_count == 0, (
            '公式验证失败：连续支出场景不应有误报。'
            '请确认使用的公式是 prev_balance + receipt + payment。'
        )

    def test_formula_with_positive_receipt(self):
        """
        收款场景的公式验证：

        初始余额: 1000
        T1: 收款 +500 → 余额应为 1500
        T2: 收款 +300 → 余额应为 1800

        公式: prev + receipt + payment
          T1: 1000 + 500 + 0 = 1500 ✓
          T2: 1500 + 300 + 0 = 1800 ✓
        """
        records = [
            _record('F011', 'FORMULA02', '2024-01-01', balance=1000.0, txn_id='F011'),
            _record('F012', 'FORMULA02', '2024-01-02', receipt=500.0, balance=1500.0, txn_id='F012'),
            _record('F013', 'FORMULA02', '2024-01-03', receipt=300.0, balance=1800.0, txn_id='F013'),
        ]
        result = check_balance_continuity(records)
        assert result.break_count == 0

    def test_formula_mixed_scenario(self):
        """
        混合场景公式验证：

        初始余额: 5000
        T1: 付款 -500  → 4500
        T2: 收款 +2000 → 6500
        T3: 付款 -1000 → 5500
        T4: 收款 +500  → 6000

        公式: prev + receipt + payment
          T1: 5000 +    0 + (-500) = 4500 ✓
          T2: 4500 + 2000 +    0   = 6500 ✓
          T3: 6500 +    0 + (-1000) = 5500 ✓
          T4: 5500 +  500 +    0   = 6000 ✓
        """
        records = [
            _record('F021', 'FORMULA03', '2024-01-01', balance=5000.0, txn_id='F021'),
            _record('F022', 'FORMULA03', '2024-01-02', payment=-500.0, balance=4500.0, txn_id='F022'),
            _record('F023', 'FORMULA03', '2024-01-03', receipt=2000.0, balance=6500.0, txn_id='F023'),
            _record('F024', 'FORMULA03', '2024-01-04', payment=-1000.0, balance=5500.0, txn_id='F024'),
            _record('F025', 'FORMULA03', '2024-01-05', receipt=500.0, balance=6000.0, txn_id='F025'),
        ]
        result = check_balance_continuity(records)
        assert result.break_count == 0, (
            '混合场景公式验证失败。'
            f'断裂详情: {[(br.transaction_id, br.diff_amount) for br in result.break_records]}'
        )
