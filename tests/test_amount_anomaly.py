# -*- coding: utf-8 -*-
"""
金额异常检测模块测试
"""

import os
import sys
import tempfile
import shutil

import pytest
import openpyxl
import pandas as pd

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'backend'))

import bankcheck
from bankcheck import (
    AmountAnomalyType,
    AmountAnomalyConfig,
    detect_amount_anomalies,
    apply_amount_anomaly_detection,
    ANOMALY_FLAG_COLUMN,
    ANOMALY_DETAIL_COLUMN,
    STANDARD_COLUMNS,
)


@pytest.fixture(autouse=True)
def init_logging():
    import logging
    logging.basicConfig(level=logging.WARNING)


class TestAmountAnomalyType:
    """金额异常类型常量测试"""

    def test_anomaly_types_defined(self):
        """测试所有异常类型都已定义"""
        assert hasattr(AmountAnomalyType, 'LARGE_AMOUNT')
        assert hasattr(AmountAnomalyType, 'ZERO_WITH_COUNTERPARTY')
        assert hasattr(AmountAnomalyType, 'NEGATIVE_RECEIPT')
        assert hasattr(AmountAnomalyType, 'POSITIVE_PAYMENT')
        assert hasattr(AmountAnomalyType, 'BALANCE_NEGATIVE')
        assert hasattr(AmountAnomalyType, 'BOTH_PAYMENT_AND_RECEIPT')
        assert hasattr(AmountAnomalyType, 'NEITHER_PAYMENT_NOR_RECEIPT')

    def test_labels_complete(self):
        """测试所有异常类型都有对应的标签"""
        all_types = [
            AmountAnomalyType.LARGE_AMOUNT,
            AmountAnomalyType.ZERO_WITH_COUNTERPARTY,
            AmountAnomalyType.NEGATIVE_RECEIPT,
            AmountAnomalyType.POSITIVE_PAYMENT,
            AmountAnomalyType.BALANCE_NEGATIVE,
            AmountAnomalyType.BOTH_PAYMENT_AND_RECEIPT,
            AmountAnomalyType.NEITHER_PAYMENT_NOR_RECEIPT,
        ]
        for t in all_types:
            assert t in AmountAnomalyType.LABELS
            assert AmountAnomalyType.LABELS[t]
            assert t in AmountAnomalyType.RISK_LEVELS


class TestAmountAnomalyConfig:
    """金额异常检测配置测试"""

    def test_default_config(self):
        """测试默认配置"""
        config = AmountAnomalyConfig.default()
        assert config.single_amount_threshold == 500000.0
        assert config.enable_zero_with_counterparty is True
        assert config.enable_negative_receipt is True
        assert config.enable_positive_payment is True
        assert config.enable_negative_balance is True
        assert config.enable_both_amounts is True
        assert config.enable_no_amounts is True

    def test_to_dict(self):
        """测试配置转换为字典"""
        config = AmountAnomalyConfig.default()
        d = config.to_dict()
        assert d['single_amount_threshold'] == 500000.0
        assert d['enable_zero_with_counterparty'] is True


class TestDetectAmountAnomalies:
    """单条记录异常检测测试"""

    def test_normal_payment(self):
        """测试正常付款记录"""
        record = {
            '付款': -1000.0,
            '收款': None,
            '余额': 9000.0,
            '对方户名': '测试公司',
        }
        has_anomaly, anomalies, descriptions = detect_amount_anomalies(record)
        assert has_anomaly is False
        assert anomalies == []
        assert descriptions == []

    def test_normal_receipt(self):
        """测试正常收款记录"""
        record = {
            '付款': None,
            '收款': 5000.0,
            '余额': 14000.0,
            '对方户名': '客户公司',
        }
        has_anomaly, anomalies, descriptions = detect_amount_anomalies(record)
        assert has_anomaly is False

    def test_large_amount_payment(self):
        """测试大额付款"""
        record = {
            '付款': -600000.0,
            '收款': None,
            '余额': 400000.0,
            '对方户名': '大供应商',
        }
        has_anomaly, anomalies, descriptions = detect_amount_anomalies(record)
        assert has_anomaly is True
        assert AmountAnomalyType.LARGE_AMOUNT in anomalies
        assert any('600,000.00' in d for d in descriptions)

    def test_large_amount_receipt(self):
        """测试大额收款"""
        record = {
            '付款': None,
            '收款': 1000000.0,
            '余额': 1500000.0,
            '对方户名': '大客户',
        }
        has_anomaly, anomalies, descriptions = detect_amount_anomalies(record)
        assert has_anomaly is True
        assert AmountAnomalyType.LARGE_AMOUNT in anomalies

    def test_custom_threshold(self):
        """测试自定义阈值"""
        config = AmountAnomalyConfig(single_amount_threshold=10000.0)
        record = {
            '付款': -15000.0,
            '收款': None,
            '余额': 5000.0,
            '对方户名': '供应商',
        }
        has_anomaly, anomalies, _ = detect_amount_anomalies(record, config)
        assert has_anomaly is True
        assert AmountAnomalyType.LARGE_AMOUNT in anomalies

    def test_zero_with_counterparty_both_zero(self):
        """测试金额为0但有对方户名（两者都为0）"""
        record = {
            '付款': 0,
            '收款': 0,
            '余额': 10000.0,
            '对方户名': '测试公司',
        }
        has_anomaly, anomalies, descriptions = detect_amount_anomalies(record)
        assert has_anomaly is True
        assert AmountAnomalyType.ZERO_WITH_COUNTERPARTY in anomalies
        assert any('测试公司' in d for d in descriptions)

    def test_zero_with_counterparty_both_none(self):
        """测试金额为None但有对方户名"""
        record = {
            '付款': None,
            '收款': None,
            '余额': 10000.0,
            '对方户名': '测试公司',
        }
        has_anomaly, anomalies, _ = detect_amount_anomalies(record)
        assert has_anomaly is True
        assert AmountAnomalyType.ZERO_WITH_COUNTERPARTY in anomalies
        assert AmountAnomalyType.NEITHER_PAYMENT_NOR_RECEIPT in anomalies

    def test_zero_with_counterparty_mixed(self):
        """测试一个为0一个为None但有对方户名"""
        record = {
            '付款': 0,
            '收款': None,
            '余额': 10000.0,
            '对方户名': '测试公司',
        }
        has_anomaly, anomalies, _ = detect_amount_anomalies(record)
        assert has_anomaly is True
        assert AmountAnomalyType.ZERO_WITH_COUNTERPARTY in anomalies

    def test_no_counterparty_zero_amount(self):
        """测试金额为0但没有对方户名"""
        record = {
            '付款': 0,
            '收款': 0,
            '余额': 10000.0,
            '对方户名': None,
        }
        has_anomaly, anomalies, _ = detect_amount_anomalies(record)
        assert has_anomaly is True
        assert AmountAnomalyType.ZERO_WITH_COUNTERPARTY not in anomalies
        assert AmountAnomalyType.NEITHER_PAYMENT_NOR_RECEIPT in anomalies

    def test_empty_counterparty_zero_amount(self):
        """测试金额为0但对方户名为空字符串"""
        record = {
            '付款': 0,
            '收款': 0,
            '余额': 10000.0,
            '对方户名': '',
        }
        has_anomaly, anomalies, _ = detect_amount_anomalies(record)
        assert has_anomaly is True
        assert AmountAnomalyType.ZERO_WITH_COUNTERPARTY not in anomalies

    def test_negative_receipt(self):
        """测试收款金额为负数"""
        record = {
            '付款': None,
            '收款': -500.0,
            '余额': 9500.0,
            '对方户名': None,
        }
        has_anomaly, anomalies, _ = detect_amount_anomalies(record)
        assert has_anomaly is True
        assert AmountAnomalyType.NEGATIVE_RECEIPT in anomalies

    def test_positive_payment(self):
        """测试付款金额为正数"""
        record = {
            '付款': 1000.0,
            '收款': None,
            '余额': 11000.0,
            '对方户名': None,
        }
        has_anomaly, anomalies, _ = detect_amount_anomalies(record)
        assert has_anomaly is True
        assert AmountAnomalyType.POSITIVE_PAYMENT in anomalies

    def test_negative_balance(self):
        """测试余额为负数"""
        record = {
            '付款': -15000.0,
            '收款': None,
            '余额': -5000.0,
            '对方户名': None,
        }
        has_anomaly, anomalies, _ = detect_amount_anomalies(record)
        assert has_anomaly is True
        assert AmountAnomalyType.BALANCE_NEGATIVE in anomalies

    def test_both_payment_and_receipt(self):
        """测试付款和收款同时有值"""
        record = {
            '付款': -500.0,
            '收款': 1000.0,
            '余额': 10500.0,
            '对方户名': None,
        }
        has_anomaly, anomalies, _ = detect_amount_anomalies(record)
        assert has_anomaly is True
        assert AmountAnomalyType.BOTH_PAYMENT_AND_RECEIPT in anomalies

    def test_neither_payment_nor_receipt(self):
        """测试付款和收款均无值"""
        record = {
            '付款': None,
            '收款': None,
            '余额': 10000.0,
            '对方户名': None,
        }
        has_anomaly, anomalies, _ = detect_amount_anomalies(record)
        assert has_anomaly is True
        assert AmountAnomalyType.NEITHER_PAYMENT_NOR_RECEIPT in anomalies

    def test_neither_payment_nor_receipt_both_zero(self):
        """测试付款和收款均为0"""
        record = {
            '付款': 0,
            '收款': 0,
            '余额': 10000.0,
            '对方户名': None,
        }
        has_anomaly, anomalies, _ = detect_amount_anomalies(record)
        assert has_anomaly is True
        assert AmountAnomalyType.NEITHER_PAYMENT_NOR_RECEIPT in anomalies

    def test_multiple_anomalies(self):
        """测试一条记录有多个异常"""
        record = {
            '付款': 600000.0,
            '收款': -100.0,
            '余额': -5000.0,
            '对方户名': '测试公司',
        }
        has_anomaly, anomalies, descriptions = detect_amount_anomalies(record)
        assert has_anomaly is True
        assert len(anomalies) >= 3
        assert AmountAnomalyType.LARGE_AMOUNT in anomalies
        assert AmountAnomalyType.POSITIVE_PAYMENT in anomalies
        assert AmountAnomalyType.NEGATIVE_RECEIPT in anomalies
        assert AmountAnomalyType.BALANCE_NEGATIVE in anomalies
        assert AmountAnomalyType.BOTH_PAYMENT_AND_RECEIPT in anomalies
        assert len(descriptions) == len(anomalies)

    def test_disable_detection(self):
        """测试禁用某些检测项"""
        config = AmountAnomalyConfig(
            single_amount_threshold=500000.0,
            enable_negative_receipt=False,
            enable_positive_payment=False,
        )
        record = {
            '付款': 1000.0,
            '收款': -500.0,
            '余额': 10500.0,
            '对方户名': None,
        }
        has_anomaly, anomalies, _ = detect_amount_anomalies(record, config)
        assert has_anomaly is True
        assert AmountAnomalyType.POSITIVE_PAYMENT not in anomalies
        assert AmountAnomalyType.NEGATIVE_RECEIPT not in anomalies
        assert AmountAnomalyType.BOTH_PAYMENT_AND_RECEIPT in anomalies

    def test_string_amounts(self):
        """测试字符串格式的金额"""
        record = {
            '付款': '-1,000.00',
            '收款': None,
            '余额': '9,000.00',
            '对方户名': '测试公司',
        }
        has_anomaly, anomalies, _ = detect_amount_anomalies(record)
        assert has_anomaly is False

    def test_zero_threshold(self):
        """测试阈值为0时禁用大额检测"""
        config = AmountAnomalyConfig(single_amount_threshold=0)
        record = {
            '付款': -1000000.0,
            '收款': None,
            '余额': 0.0,
            '对方户名': None,
        }
        has_anomaly, anomalies, _ = detect_amount_anomalies(record, config)
        assert AmountAnomalyType.LARGE_AMOUNT not in anomalies


class TestApplyAmountAnomalyDetection:
    """批量异常检测测试"""

    def test_empty_records(self):
        """测试空记录列表"""
        records, summary = apply_amount_anomaly_detection([])
        assert records == []
        assert summary['total_records'] == 0
        assert summary['anomaly_count'] == 0

    def test_normal_records(self):
        """测试正常记录"""
        records = [
            {'付款': -1000.0, '收款': None, '余额': 9000.0, '对方户名': '公司A'},
            {'付款': None, '收款': 2000.0, '余额': 11000.0, '对方户名': '公司B'},
            {'付款': -500.0, '收款': None, '余额': 10500.0, '对方户名': '公司C'},
        ]
        records, summary = apply_amount_anomaly_detection(records)
        assert summary['total_records'] == 3
        assert summary['anomaly_count'] == 0
        for r in records:
            assert ANOMALY_FLAG_COLUMN in r
            assert r[ANOMALY_FLAG_COLUMN] == ''
            assert ANOMALY_DETAIL_COLUMN in r
            assert r[ANOMALY_DETAIL_COLUMN] == ''

    def test_mixed_records(self):
        """测试混合正常和异常记录"""
        records = [
            {'付款': -1000.0, '收款': None, '余额': 9000.0, '对方户名': '公司A'},
            {'付款': -600000.0, '收款': None, '余额': 400000.0, '对方户名': '公司B'},
            {'付款': 0, '收款': 0, '余额': 400000.0, '对方户名': '公司C'},
            {'付款': None, '收款': -500.0, '余额': 399500.0, '对方户名': None},
        ]
        records, summary = apply_amount_anomaly_detection(records)
        assert summary['total_records'] == 4
        assert summary['anomaly_count'] == 3
        assert summary['anomaly_rate'] == 0.75

        assert records[0][ANOMALY_FLAG_COLUMN] == ''
        assert AmountAnomalyType.LARGE_AMOUNT in summary['anomaly_types']
        assert AmountAnomalyType.ZERO_WITH_COUNTERPARTY in summary['anomaly_types']
        assert AmountAnomalyType.NEGATIVE_RECEIPT in summary['anomaly_types']

    def test_anomaly_flag_format(self):
        """测试异常标记格式"""
        records = [
            {'付款': -600000.0, '收款': None, '余额': 400000.0, '对方户名': '公司A'},
        ]
        records, _ = apply_amount_anomaly_detection(records)
        flag = records[0][ANOMALY_FLAG_COLUMN]
        assert '[高]' in flag
        assert '单笔金额超过阈值' in flag

    def test_risk_levels_in_flag(self):
        """测试风险等级在标记中正确显示"""
        records = [
            {'付款': None, '收款': None, '余额': 10000.0, '对方户名': None},
        ]
        records, _ = apply_amount_anomaly_detection(records)
        flag = records[0][ANOMALY_FLAG_COLUMN]
        assert '[低]' in flag
        assert '付款和收款均无值' in flag


class TestStandardColumns:
    """标准列定义测试"""

    def test_anomaly_columns_in_standard(self):
        """测试异常标记列在标准列中"""
        assert ANOMALY_FLAG_COLUMN in STANDARD_COLUMNS
        assert ANOMALY_DETAIL_COLUMN in STANDARD_COLUMNS

    def test_anomaly_columns_order(self):
        """测试异常标记列的位置"""
        idx_flag = STANDARD_COLUMNS.index(ANOMALY_FLAG_COLUMN)
        idx_detail = STANDARD_COLUMNS.index(ANOMALY_DETAIL_COLUMN)
        idx_transaction_id = STANDARD_COLUMNS.index('交易流水号')
        assert idx_flag > idx_transaction_id
        assert idx_detail > idx_flag


class TestIntegrationWithSummary:
    """与总表输出的集成测试"""

    def test_summary_columns_include_anomaly(self, tmp_path):
        """测试总表列包含异常标记列"""
        records = [
            {
                '唯一id': 'TEST001',
                '银行': '测试银行',
                '银行账号': '123456',
                '主体': '测试主体',
                '交易日期': '2024-01-01',
                '付款': -1000.0,
                '收款': None,
                '摘要': '测试',
                '对方户名': '测试公司',
                '余额': 9000.0,
                '交易流水号': 'TXN001',
            },
            {
                '唯一id': 'TEST002',
                '银行': '测试银行',
                '银行账号': '123456',
                '主体': '测试主体',
                '交易日期': '2024-01-02',
                '付款': -600000.0,
                '收款': None,
                '摘要': '大额测试',
                '对方户名': '大公司',
                '余额': 400000.0,
                '交易流水号': 'TXN002',
            },
        ]

        records, _ = apply_amount_anomaly_detection(records)
        columns = bankcheck.get_summary_columns(records)

        assert ANOMALY_FLAG_COLUMN in columns
        assert ANOMALY_DETAIL_COLUMN in columns

        df = pd.DataFrame(records, columns=columns)
        output_path = tmp_path / 'test_summary.xlsx'
        df.to_excel(output_path, index=False, engine='openpyxl')

        assert output_path.exists()

        df_read = pd.read_excel(output_path, engine='openpyxl')
        assert ANOMALY_FLAG_COLUMN in df_read.columns
        assert ANOMALY_DETAIL_COLUMN in df_read.columns
        assert df_read.iloc[0][ANOMALY_FLAG_COLUMN] == '' or pd.isna(df_read.iloc[0][ANOMALY_FLAG_COLUMN])
        assert '单笔金额超过阈值' in str(df_read.iloc[1][ANOMALY_FLAG_COLUMN])


class TestRiskLevelAssignment:
    """风险等级分配测试"""

    def test_high_risk_levels(self):
        """测试高风险异常类型"""
        high_risk_types = [
            AmountAnomalyType.LARGE_AMOUNT,
            AmountAnomalyType.NEGATIVE_RECEIPT,
            AmountAnomalyType.POSITIVE_PAYMENT,
        ]
        for t in high_risk_types:
            assert AmountAnomalyType.RISK_LEVELS[t] == 'high'

    def test_medium_risk_levels(self):
        """测试中风险异常类型"""
        medium_risk_types = [
            AmountAnomalyType.ZERO_WITH_COUNTERPARTY,
            AmountAnomalyType.BALANCE_NEGATIVE,
            AmountAnomalyType.BOTH_PAYMENT_AND_RECEIPT,
        ]
        for t in medium_risk_types:
            assert AmountAnomalyType.RISK_LEVELS[t] == 'medium'

    def test_low_risk_levels(self):
        """测试低风险异常类型"""
        assert AmountAnomalyType.RISK_LEVELS[AmountAnomalyType.NEITHER_PAYMENT_NOR_RECEIPT] == 'low'


if __name__ == '__main__':
    pytest.main([__file__, '-v'])
