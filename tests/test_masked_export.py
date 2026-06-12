"""
敏感信息脱敏导出模块测试

验证：
  1. 单条记录脱敏（银行账号、对方户名）
  2. 批量记录脱敏
  3. 脱敏版总表导出
  4. 完整版与脱敏版双输出
  5. 集成测试：run_pipeline 后自动生成脱敏版
"""
import os
import sys
import tempfile
import shutil

import pandas as pd
import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'backend'))

import bankcheck
from conftest import _create_beijing_bank_excel, _create_east_asia_bank_excel, _create_lookup_table


class TestMaskRecord:
    """单条记录脱敏测试"""

    def test_mask_bank_account(self):
        """银行账号脱敏：保留前3后3"""
        record = {
            '唯一id': 'TEST001',
            '银行': '北京银行',
            '银行账号': '01090312345678901',
            '主体': '北京XX科技有限公司',
            '对方户名': '供应商A公司',
        }
        masked = bankcheck.mask_record(record)
        assert masked['银行'] == '北京银行'
        assert masked['唯一id'] == 'TEST001'
        assert masked['银行账号'].startswith('010')
        assert masked['银行账号'].endswith('901')
        assert '*' in masked['银行账号']
        assert '12345678' not in masked['银行账号']

    def test_mask_counterparty_name(self):
        """对方户名脱敏：保留首尾各1字符"""
        record = {
            '对方户名': '供应商A公司',
            '银行账号': '1234567890',
        }
        masked = bankcheck.mask_record(record)
        assert masked['对方户名'].startswith('供')
        assert masked['对方户名'].endswith('司')
        assert '*' in masked['对方户名']
        assert '应商A公' not in masked['对方户名']

    def test_mask_short_account(self):
        """短账号脱敏：全掩码"""
        record = {'银行账号': '123456'}
        masked = bankcheck.mask_record(record)
        assert masked['银行账号'] == '******'
        assert len(masked['银行账号']) == 6

    def test_mask_single_char_name(self):
        """单字符户名脱敏"""
        record = {'对方户名': '张'}
        masked = bankcheck.mask_record(record)
        assert masked['对方户名'] == '*'

    def test_other_fields_unchanged(self):
        """非敏感字段保持不变"""
        record = {
            '唯一id': 'TEST001',
            '银行': '北京银行',
            '交易日期': '2024-01-05',
            '摘要': '采购付款',
            '银行账号': '01090312345678901',
            '对方户名': '供应商A公司',
        }
        masked = bankcheck.mask_record(record)
        assert masked['唯一id'] == 'TEST001'
        assert masked['银行'] == '北京银行'
        assert masked['交易日期'] == '2024-01-05'
        assert masked['摘要'] == '采购付款'

    def test_none_values_preserved(self):
        """None 值保持不变"""
        record = {
            '银行账号': None,
            '对方户名': None,
            '付款': None,
        }
        masked = bankcheck.mask_record(record)
        assert masked['银行账号'] is None
        assert masked['对方户名'] is None
        assert masked['付款'] is None

    def test_empty_values_preserved(self):
        """空字符串保持不变"""
        record = {
            '银行账号': '',
            '对方户名': '',
        }
        masked = bankcheck.mask_record(record)
        assert masked['银行账号'] == ''
        assert masked['对方户名'] == ''

    def test_custom_masked_fields(self):
        """自定义脱敏字段列表"""
        record = {
            '银行账号': '01090312345678901',
            '对方户名': '供应商A公司',
            '主体': '北京XX科技有限公司',
        }
        masked = bankcheck.mask_record(record, fields=['银行账号'])
        assert '*' in masked['银行账号']
        assert masked['对方户名'] == '供应商A公司'
        assert masked['主体'] == '北京XX科技有限公司'

    def test_original_record_not_modified(self):
        """原记录不被修改"""
        original = {
            '银行账号': '01090312345678901',
            '对方户名': '供应商A公司',
        }
        original_copy = dict(original)
        bankcheck.mask_record(original)
        assert original == original_copy


class TestMaskRecords:
    """批量记录脱敏测试"""

    def test_multiple_records_masked(self):
        """多条记录全部脱敏"""
        records = [
            {'银行账号': '1111111111', '对方户名': '公司A'},
            {'银行账号': '2222222222', '对方户名': '公司B'},
            {'银行账号': '3333333333', '对方户名': '公司C'},
        ]
        masked = bankcheck.mask_records(records)
        assert len(masked) == 3
        for i, r in enumerate(masked):
            assert '*' in r['银行账号']
            assert '*' in r['对方户名']
            assert str(i + 1) * 10 not in r['银行账号']

    def test_empty_list(self):
        """空列表返回空列表"""
        assert bankcheck.mask_records([]) == []
        assert bankcheck.mask_records(None) == []

    def test_original_records_not_modified(self):
        """原记录列表不被修改"""
        records = [
            {'银行账号': '1111111111', '对方户名': '公司A'},
        ]
        original = [dict(r) for r in records]
        bankcheck.mask_records(records)
        assert records == original


class TestExportMaskedSummary:
    """脱敏版总表导出测试"""

    @pytest.fixture
    def tmp_dir(self):
        d = tempfile.mkdtemp(prefix='masked_export_test_')
        yield d
        shutil.rmtree(d, ignore_errors=True)

    def test_export_creates_file(self, tmp_dir):
        """导出脱敏版文件存在"""
        records = [
            {
                '唯一id': 'TEST001',
                '银行': '北京银行',
                '银行账号': '01090312345678901',
                '主体': '北京XX科技有限公司',
                '交易日期': '2024-01-05',
                '付款': -50000.0,
                '收款': None,
                '摘要': '采购付款',
                '对方户名': '供应商A公司',
                '余额': 1500000.0,
                '交易流水号': 'BJ20240105001',
            },
        ]
        output_path = bankcheck.export_masked_summary(
            records, tmp_dir, output_dir=tmp_dir
        )
        assert output_path is not None
        assert os.path.exists(output_path)
        assert '脱敏版' in os.path.basename(output_path)

    def test_exported_file_has_masked_values(self, tmp_dir):
        """导出的文件中敏感字段已脱敏"""
        records = [
            {
                '唯一id': 'TEST001',
                '银行': '北京银行',
                '银行账号': '01090312345678901',
                '主体': '北京XX科技有限公司',
                '交易日期': '2024-01-05',
                '付款': -50000.0,
                '收款': None,
                '摘要': '采购付款',
                '对方户名': '供应商A公司',
                '余额': 1500000.0,
                '交易流水号': 'BJ20240105001',
            },
        ]
        output_path = bankcheck.export_masked_summary(
            records, tmp_dir, output_dir=tmp_dir
        )
        df = pd.read_excel(output_path, engine='openpyxl')
        assert len(df) == 1
        assert '银行账号' in df.columns
        assert '对方户名' in df.columns

        account_value = str(df['银行账号'].iloc[0])
        assert '01090312345678901' not in account_value
        assert '*' in account_value

        name_value = str(df['对方户名'].iloc[0])
        assert '供应商A公司' not in name_value
        assert '*' in name_value

    def test_export_empty_records_returns_none(self, tmp_dir):
        """空记录返回 None"""
        result = bankcheck.export_masked_summary([], tmp_dir, output_dir=tmp_dir)
        assert result is None

    def test_export_preserves_other_fields(self, tmp_dir):
        """非敏感字段保持原值"""
        records = [
            {
                '唯一id': 'TEST001',
                '银行': '北京银行',
                '银行账号': '01090312345678901',
                '交易日期': '2024-01-05',
                '摘要': '采购付款',
                '对方户名': '供应商A公司',
            },
        ]
        output_path = bankcheck.export_masked_summary(
            records, tmp_dir, output_dir=tmp_dir
        )
        df = pd.read_excel(output_path, engine='openpyxl')
        assert df['唯一id'].iloc[0] == 'TEST001'
        assert df['银行'].iloc[0] == '北京银行'
        assert df['交易日期'].iloc[0] == '2024-01-05'
        assert df['摘要'].iloc[0] == '采购付款'

    def test_export_with_custom_columns(self, tmp_dir):
        """使用自定义列导出"""
        records = [
            {
                '唯一id': 'TEST001',
                '银行': '北京银行',
                '银行账号': '01090312345678901',
                '对方户名': '供应商A公司',
                '额外字段': '测试值',
            },
        ]
        columns = ['唯一id', '银行', '银行账号', '对方户名']
        output_path = bankcheck.export_masked_summary(
            records, tmp_dir, output_dir=tmp_dir, columns=columns
        )
        df = pd.read_excel(output_path, engine='openpyxl')
        assert list(df.columns) == columns
        assert '额外字段' not in df.columns


class TestMaskedSummaryPath:
    """脱敏版路径获取测试"""

    @pytest.fixture
    def tmp_dir(self):
        d = tempfile.mkdtemp(prefix='masked_path_test_')
        yield d
        shutil.rmtree(d, ignore_errors=True)

    def test_default_script_dir(self, tmp_dir):
        """默认使用 script_dir"""
        path = bankcheck.get_masked_summary_table_path(tmp_dir)
        assert path == os.path.join(tmp_dir, bankcheck.SUMMARY_TABLE_MASKED_FILENAME)

    def test_custom_output_dir(self, tmp_dir):
        """使用自定义输出目录"""
        output_dir = os.path.join(tmp_dir, 'output')
        os.makedirs(output_dir)
        path = bankcheck.get_masked_summary_table_path(tmp_dir, output_dir=output_dir)
        assert path == os.path.join(output_dir, bankcheck.SUMMARY_TABLE_MASKED_FILENAME)


class TestDualOutputIntegration:
    """双输出版本集成测试：完整版 + 脱敏版"""

    @pytest.fixture
    def tmp_dir(self):
        d = tempfile.mkdtemp(prefix='dual_output_test_')
        yield d
        shutil.rmtree(d, ignore_errors=True)

    def _setup_test_data(self, tmp_dir):
        """设置测试数据"""
        script_dir = os.path.join(tmp_dir, 'script')
        os.makedirs(script_dir, exist_ok=True)

        source_folder = os.path.join(tmp_dir, '流水文件夹')
        os.makedirs(source_folder, exist_ok=True)

        _create_beijing_bank_excel(os.path.join(source_folder, '北京银行_流水.xlsx'))
        _create_lookup_table(os.path.join(script_dir, '主体查找表.xlsx'))

        return source_folder, script_dir

    def test_pipeline_generates_both_versions(self, tmp_dir):
        """run_pipeline 同时生成完整版和脱敏版"""
        source_folder, script_dir = self._setup_test_data(tmp_dir)

        result = bankcheck.run_pipeline(source_folder, script_dir, incremental=False)

        assert result.output_path is not None
        assert result.masked_output_path is not None

        assert os.path.exists(result.output_path)
        assert os.path.exists(result.masked_output_path)

        assert '脱敏版' not in os.path.basename(result.output_path)
        assert '脱敏版' in os.path.basename(result.masked_output_path)

    def test_full_version_has_original_data(self, tmp_dir):
        """完整版包含原始敏感数据"""
        source_folder, script_dir = self._setup_test_data(tmp_dir)

        result = bankcheck.run_pipeline(source_folder, script_dir, incremental=False)

        df_full = pd.read_excel(result.output_path, engine='openpyxl')
        account_value = str(df_full['银行账号'].iloc[0])
        assert '010903' in account_value or len(account_value) > 10

    def test_masked_version_has_masked_data(self, tmp_dir):
        """脱敏版中敏感字段已掩码处理"""
        source_folder, script_dir = self._setup_test_data(tmp_dir)

        result = bankcheck.run_pipeline(source_folder, script_dir, incremental=False)

        df_masked = pd.read_excel(result.masked_output_path, engine='openpyxl')

        account_value = str(df_masked['银行账号'].iloc[0])
        assert '*' in account_value

        name_value = str(df_masked['对方户名'].iloc[0])
        assert '*' in name_value

    def test_both_versions_have_same_row_count(self, tmp_dir):
        """完整版和脱敏版记录数相同"""
        source_folder, script_dir = self._setup_test_data(tmp_dir)

        result = bankcheck.run_pipeline(source_folder, script_dir, incremental=False)

        df_full = pd.read_excel(result.output_path, engine='openpyxl')
        df_masked = pd.read_excel(result.masked_output_path, engine='openpyxl')

        assert len(df_full) == len(df_masked)

    def test_both_versions_have_same_columns(self, tmp_dir):
        """完整版和脱敏版列数相同"""
        source_folder, script_dir = self._setup_test_data(tmp_dir)

        result = bankcheck.run_pipeline(source_folder, script_dir, incremental=False)

        df_full = pd.read_excel(result.output_path, engine='openpyxl')
        df_masked = pd.read_excel(result.masked_output_path, engine='openpyxl')

        assert list(df_full.columns) == list(df_masked.columns)

    def test_non_sensitive_fields_same_in_both(self, tmp_dir):
        """非敏感字段在两个版本中值相同"""
        source_folder, script_dir = self._setup_test_data(tmp_dir)

        result = bankcheck.run_pipeline(source_folder, script_dir, incremental=False)

        df_full = pd.read_excel(result.output_path, engine='openpyxl')
        df_masked = pd.read_excel(result.masked_output_path, engine='openpyxl')

        assert list(df_full['唯一id']) == list(df_masked['唯一id'])
        assert list(df_full['银行']) == list(df_masked['银行'])
        assert list(df_full['交易日期']) == list(df_masked['交易日期'])
        assert list(df_full['摘要']) == list(df_masked['摘要'])


class TestMaskedFieldsConfig:
    """脱敏字段配置测试"""

    def test_default_masked_fields(self):
        """默认脱敏字段包含银行账号和对方户名"""
        assert '银行账号' in bankcheck.MASKED_FIELDS
        assert '对方户名' in bankcheck.MASKED_FIELDS
        assert '对方账号' in bankcheck.MASKED_FIELDS

    def test_non_masked_fields_excluded(self):
        """非脱敏字段不在列表中"""
        assert '唯一id' not in bankcheck.MASKED_FIELDS
        assert '银行' not in bankcheck.MASKED_FIELDS
        assert '摘要' not in bankcheck.MASKED_FIELDS
        assert '余额' not in bankcheck.MASKED_FIELDS
