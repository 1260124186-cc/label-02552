"""
PII 分级与脱敏模块单元测试

验证：
  1. 字段分级分类正确性
  2. 各级别脱敏函数输出合法性
  3. 字典/列表递归脱敏
  4. PIILogFilter 集成效果
  5. 导出字段白名单功能
"""
import os
import sys
import io
import logging
import tempfile
import shutil

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'backend'))

from pii_classifier import (
    PIILevel,
    classify_field,
    mask_value,
    sanitize_dict_for_log,
    sanitize_for_log,
    PIILogFilter,
    build_safe_log_context,
    setup_pii_aware_logging,
    is_field_forbidden,
    is_field_debug_only,
    is_field_info_safe,
    get_export_field_whitelist,
)


class TestFieldClassification:
    """字段分级分类测试"""

    def test_info_safe_fields(self):
        """INFO_SAFE 级字段"""
        assert classify_field('银行') == PIILevel.INFO_SAFE
        assert classify_field('银行名称') == PIILevel.INFO_SAFE
        assert classify_field('记录数') == PIILevel.INFO_SAFE
        assert classify_field('状态') == PIILevel.INFO_SAFE
        assert classify_field('模式') == PIILevel.INFO_SAFE
        assert classify_field('sheet_name') == PIILevel.INFO_SAFE
        assert classify_field('币种') == PIILevel.INFO_SAFE
        assert classify_field('黑白名单标签') == PIILevel.INFO_SAFE
        assert classify_field('命中规则名称') == PIILevel.INFO_SAFE

    def test_debug_only_fields(self):
        """DEBUG_ONLY 级字段"""
        assert classify_field('唯一id') == PIILevel.DEBUG_ONLY
        assert classify_field('主体') == PIILevel.DEBUG_ONLY
        assert classify_field('主体名称') == PIILevel.DEBUG_ONLY
        assert classify_field('摘要') == PIILevel.DEBUG_ONLY
        assert classify_field('对方户名') == PIILevel.DEBUG_ONLY
        assert classify_field('交易描述') == PIILevel.DEBUG_ONLY
        assert classify_field('备注') == PIILevel.DEBUG_ONLY
        assert classify_field('subject') == PIILevel.DEBUG_ONLY
        assert classify_field('counterparty') == PIILevel.DEBUG_ONLY
        assert classify_field('workflow_id') == PIILevel.DEBUG_ONLY

    def test_forbidden_fields(self):
        """FORBIDDEN 级字段（禁止落盘）"""
        assert classify_field('银行账号') == PIILevel.FORBIDDEN
        assert classify_field('对方账号') == PIILevel.FORBIDDEN
        assert classify_field('付款') == PIILevel.FORBIDDEN
        assert classify_field('支出金额') == PIILevel.FORBIDDEN
        assert classify_field('收款') == PIILevel.FORBIDDEN
        assert classify_field('收入金额') == PIILevel.FORBIDDEN
        assert classify_field('余额') == PIILevel.FORBIDDEN
        assert classify_field('交易流水号') == PIILevel.FORBIDDEN
        assert classify_field('匹配键') == PIILevel.FORBIDDEN
        assert classify_field('交易日期') == PIILevel.FORBIDDEN
        assert classify_field('交易时间') == PIILevel.FORBIDDEN
        assert classify_field('附件路径') == PIILevel.FORBIDDEN
        assert classify_field('account') == PIILevel.FORBIDDEN
        assert classify_field('balance') == PIILevel.FORBIDDEN
        assert classify_field('payment') == PIILevel.FORBIDDEN
        assert classify_field('amount') == PIILevel.FORBIDDEN

    def test_keyword_inference_forbidden(self):
        """关键词模糊推断：FORBIDDEN"""
        assert classify_field('商户账号') == PIILevel.FORBIDDEN
        assert classify_field('到账金额') == PIILevel.FORBIDDEN
        assert classify_field('内部流水号') == PIILevel.FORBIDDEN

    def test_keyword_inference_debug(self):
        """关键词模糊推断：DEBUG_ONLY"""
        assert classify_field('交易对手名称') == PIILevel.DEBUG_ONLY
        assert classify_field('交易摘要文本') == PIILevel.DEBUG_ONLY

    def test_unknown_default_safe(self):
        """未知字段默认 INFO_SAFE"""
        assert classify_field('完全不认识的字段') == PIILevel.INFO_SAFE
        assert classify_field('') == PIILevel.INFO_SAFE
        assert classify_field(None) == PIILevel.INFO_SAFE

    def test_is_field_helpers(self):
        """便捷判断函数"""
        assert is_field_forbidden('银行账号')
        assert not is_field_forbidden('银行')
        assert is_field_debug_only('主体')
        assert not is_field_debug_only('银行账号')
        assert is_field_info_safe('记录数')
        assert not is_field_info_safe('摘要')


class TestMaskValue:
    """字段值脱敏测试"""

    def test_bank_account_mask(self):
        """银行账号脱敏：前3后3"""
        result = mask_value('银行账号', '01090312345678901')
        assert result.startswith('010')
        assert result.endswith('901')
        assert '*' in result
        assert '12345678' not in result

        short = mask_value('银行账号', '123456')
        assert short == '******'

        empty = mask_value('银行账号', '')
        assert empty == ''

    def test_subject_name_mask(self):
        """主体/对方户名脱敏：首尾各1"""
        result = mask_value('主体', '北京XX科技有限公司')
        assert result.startswith('北')
        assert result.endswith('司')
        assert '*' in result
        assert 'XX科技' not in result

        short = mask_value('对方户名', 'AB')
        assert short == 'A*'

        single = mask_value('对方户名', '张')
        assert single == '*'

    def test_summary_mask(self):
        """摘要/文本脱敏：保留前2字"""
        result = mask_value('摘要', '采购办公设备货款')
        assert result.startswith('采购')
        assert result.endswith('***')
        assert '办公' not in result

    def test_transaction_id_mask(self):
        """交易流水号脱敏：仅前4位"""
        result = mask_value('交易流水号', 'BJ20240105001')
        assert result.startswith('BJ20')
        assert '*' in result
        assert '240105' not in result

    def test_unique_id_mask(self):
        """唯一ID脱敏：前8位"""
        result = mask_value('唯一id', 'TEST001-UNIQUE-ID-0001')
        assert result.startswith('TEST001-')
        assert result.endswith('****')

    def test_amount_mask_ranges(self):
        """金额脱敏：显示区间而非精确值"""
        assert mask_value('付款', 500) == '[<1千]'
        assert mask_value('收款', 5000) == '[1千-1万]'
        assert mask_value('余额', 50000) == '[1万-10万]'
        assert mask_value('支出金额', 500000) == '[10万-100万]'
        assert mask_value('收入金额', 5000000) == '[100万-1千万]'
        assert mask_value('付款', 50000000) == '[>=1千万]'
        assert mask_value('余额', 0) == '[0]'
        assert mask_value('付款', -50000) == '[1万-10万]'

    def test_date_mask(self):
        """日期脱敏：仅年月"""
        result = mask_value('交易日期', '2024-01-05')
        assert result == '2024年01月'
        assert '05' not in result

        result2 = mask_value('交易日期', '2024/12/31')
        assert result2 == '2024年12月'

    def test_path_mask(self):
        """路径脱敏：仅保留文件名"""
        result = mask_value('附件路径', '/var/data/secret/receipt_001.pdf')
        assert result == 'receipt_001.pdf'
        assert '/var/data/secret/' not in result

    def test_info_safe_unchanged(self):
        """INFO_SAFE 字段不做修改"""
        assert mask_value('银行', '北京银行') == '北京银行'
        assert mask_value('记录数', 100) == 100
        assert mask_value('状态', 'success') == 'success'

    def test_none_values(self):
        """None 值保持原样"""
        assert mask_value('银行账号', None) is None
        assert mask_value('余额', None) is None

    def test_debug_vs_info_strictness(self):
        """DEBUG 与 INFO 目标级别的严格度差异
        FORBIDDEN 字段在两级均必须脱敏（不输出精确值）
        """
        debug_val = mask_value('余额', 1500000.0, target_level=PIILevel.DEBUG_ONLY)
        info_val = mask_value('余额', 1500000.0, target_level=PIILevel.INFO_SAFE)
        assert debug_val == info_val
        assert '1500000' not in str(debug_val)
        assert '1500000' not in str(info_val)


class TestSanitizeDict:
    """字典结构递归脱敏测试"""

    def test_flat_dict(self):
        """扁平字典"""
        data = {
            '银行': '北京银行',
            '银行账号': '01090312345678901',
            '主体': '北京XX科技有限公司',
            '余额': 1500000.0,
            '记录数': 100,
        }
        sanitized = sanitize_dict_for_log(data)

        assert sanitized['银行'] == '北京银行'
        assert sanitized['记录数'] == 100
        assert '01090312345678901' not in sanitized['银行账号']
        assert '*' in sanitized['银行账号']
        assert '北京XX科技有限公司' not in sanitized['主体']
        assert 1500000.0 != sanitized['余额']

    def test_nested_dict(self):
        """嵌套字典"""
        data = {
            'header': {
                '银行': '北京银行',
                '银行账号': '01090312345678901',
            },
            'row': {
                '主体': '北京XX科技有限公司',
                '对方户名': '供应商A公司',
                '余额': 1500000,
            },
            '统计': {'记录数': 100},
        }
        sanitized = sanitize_dict_for_log(data)

        assert sanitized['header']['银行'] == '北京银行'
        assert '01090312345678901' not in sanitized['header']['银行账号']
        assert '北京XX科技有限公司' not in sanitized['row']['主体']
        assert sanitized['统计']['记录数'] == 100

    def test_list_of_dicts(self):
        """列表中包含字典"""
        data = {
            'rows': [
                {'银行': '北京银行', '余额': 1000, '银行账号': 'AAA'},
                {'银行': '东亚银行', '余额': 2000, '银行账号': 'BBB'},
            ],
        }
        sanitized = sanitize_dict_for_log(data)

        assert sanitized['rows'][0]['银行'] == '北京银行'
        assert sanitized['rows'][1]['银行'] == '东亚银行'
        assert 1000 != sanitized['rows'][0]['余额']
        assert 'AAA' not in sanitized['rows'][0]['银行账号']

    def test_allowed_fields_bypass(self):
        """显式允许字段绕过脱敏"""
        data = {
            '银行账号': '01090312345678901',
            '主体': '北京XX科技有限公司',
        }
        sanitized = sanitize_dict_for_log(
            data,
            allowed_fields=['银行账号'],
        )
        assert sanitized['银行账号'] == '01090312345678901'
        assert '北京XX科技有限公司' not in sanitized['主体']


class TestSanitizeForLogMessage:
    """日志消息文本脱敏测试"""

    def test_message_extra_tuple(self):
        """消息+extra 联合脱敏"""
        msg = '处理完成'
        extra = {'银行账号': '01090312345678901', '记录数': 50}
        msg_out, extra_out = sanitize_for_log(msg, extra)

        assert msg_out == '处理完成'
        assert extra_out['记录数'] == 50
        assert '01090312345678901' not in extra_out['银行账号']

    def test_message_pattern_inline_account(self):
        """消息文本中内联的账号脱敏"""
        msg = '检测到银行账号: 01090312345678901 异常'
        msg_out, _ = sanitize_for_log(msg)
        assert '01090312345678901' not in msg_out


class TestBuildSafeLogContext:
    """安全日志上下文构建器测试"""

    def test_standard_fields_unchanged(self):
        ctx = build_safe_log_context(
            record_count=100,
            bank_name='北京银行',
            sheet_name='交易明细',
            status='success',
            month='2024-01',
        )
        assert ctx['记录数'] == 100
        assert ctx['银行'] == '北京银行'
        assert ctx['工作表'] == '交易明细'
        assert ctx['状态'] == 'success'
        assert ctx['月份'] == '2024-01'

    def test_extra_fields_auto_sanitized(self):
        ctx = build_safe_log_context(
            record_count=50,
            银行账号='01090312345678901',
            余额=1500000,
        )
        assert ctx['记录数'] == 50
        assert '01090312345678901' not in ctx['银行账号']
        assert ctx['余额'] != 1500000


class TestPIILogFilter:
    """日志过滤器集成测试"""

    def _make_record(self, levelno, msg, args=None):
        record = logging.LogRecord(
            name='test',
            level=levelno,
            pathname='test.py',
            lineno=1,
            msg=msg,
            args=args,
            exc_info=None,
        )
        return record

    def test_info_level_strict_masking(self):
        """INFO 级日志：严格脱敏"""
        record = self._make_record(
            logging.INFO,
            '处理账号 %s 余额 %s',
            ('01090312345678901', 1500000),
        )
        pii_filter = PIILogFilter()
        result = pii_filter.filter(record)

        assert result is True
        assert '01090312345678901' not in str(record.args)

    def test_debug_level_partial_masking(self):
        """DEBUG 级日志：FORBIDDEN 字段仍需脱敏（不输出精确值）"""
        record = self._make_record(
            logging.DEBUG,
            '调试: 余额=%s',
            (1500000,),
        )
        pii_filter = PIILogFilter()
        result = pii_filter.filter(record)

        assert result is True
        assert '1500000' not in str(record.args)

    def test_dict_args_sanitized(self):
        """字典类型 args 脱敏"""
        record = self._make_record(
            logging.WARNING,
            '警告',
            {'银行账号': '01090312345678901', '主体': '敏感公司'},
        )
        pii_filter = PIILogFilter()
        result = pii_filter.filter(record)

        assert result is True
        assert '01090312345678901' not in str(record.args)


class TestSetupPIIAwareLogging:
    """PII 感知日志系统初始化测试"""

    @pytest.fixture
    def tmp_dir(self):
        d = tempfile.mkdtemp(prefix='pii_log_test_')
        yield d
        shutil.rmtree(d, ignore_errors=True)

    def test_setup_returns_logger(self, tmp_dir):
        log_file = os.path.join(tmp_dir, 'test.log')
        logger = setup_pii_aware_logging(
            logger_name='test_pii_logger',
            log_file=log_file,
        )
        assert isinstance(logger, logging.Logger)
        assert logger.name == 'test_pii_logger'
        assert len(logger.handlers) >= 2

    def test_handlers_have_pii_filter(self, tmp_dir):
        log_file = os.path.join(tmp_dir, 'test.log')
        logger = setup_pii_aware_logging(
            logger_name='test_pii_filter_check',
            log_file=log_file,
        )
        for h in logger.handlers:
            has_pii = any(isinstance(f, PIILogFilter) for f in h.filters)
            assert has_pii, f"Handler {type(h).__name__} 缺少 PIILogFilter"

    def test_info_log_output_masked(self, tmp_dir):
        """INFO 输出到文件时内容被脱敏"""
        log_file = os.path.join(tmp_dir, 'test_output.log')
        logger = setup_pii_aware_logging(
            logger_name='test_pii_output',
            log_file=log_file,
            console_level=logging.WARNING,
            file_level=logging.INFO,
        )

        logger.info(
            '处理完成: 银行账号=%s, 余额=%s, 记录数=%s',
            '01090312345678901',
            1500000.0,
            100,
        )

        for h in logger.handlers:
            h.flush()

        with open(log_file, 'r', encoding='utf-8') as f:
            content = f.read()

        assert '01090312345678901' not in content
        assert '1500000.0' not in content
        assert '100' in content


class TestExportFieldWhitelist:
    """导出字段白名单测试"""

    def test_debug_level_whitelist_excludes_forbidden(self):
        """DEBUG 级白名单不包含 FORBIDDEN 字段"""
        whitelist = get_export_field_whitelist(PIILevel.DEBUG_ONLY)
        assert '银行' in whitelist
        assert '主体' in whitelist
        assert '银行账号' not in whitelist
        assert '余额' not in whitelist
        assert '付款' not in whitelist
        assert '交易流水号' not in whitelist

    def test_info_level_whitelist_strictest(self):
        """INFO 级白名单仅包含 INFO_SAFE"""
        whitelist = get_export_field_whitelist(PIILevel.INFO_SAFE)
        assert '银行' in whitelist
        assert '记录数' in whitelist
        assert '主体' not in whitelist
        assert '摘要' not in whitelist
        assert '银行账号' not in whitelist
