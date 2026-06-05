import os
import shutil

import openpyxl
import pandas as pd
import pytest

from conftest import _create_beijing_bank_excel, _create_east_asia_bank_excel, _create_lookup_table
import bankcheck


class TestRunPipeline:
    def _setup_folder(self, tmp_dir, script_dir, files=None):
        source_folder = os.path.join(tmp_dir, '流水文件夹')
        os.makedirs(source_folder, exist_ok=True)

        if files is None:
            files = ['北京银行', '东亚银行']

        for bank in files:
            if bank == '北京银行':
                _create_beijing_bank_excel(os.path.join(source_folder, '北京银行_流水.xlsx'))
            elif bank == '东亚银行':
                _create_east_asia_bank_excel(os.path.join(source_folder, '东亚银行_流水.xlsx'))
            elif bank == '未知':
                wb = openpyxl.Workbook()
                ws = wb.active
                ws['A1'] = '未知银行数据'
                wb.save(os.path.join(source_folder, '未知银行_流水.xlsx'))
                wb.close()

        _create_lookup_table(os.path.join(script_dir, '主体查找表.xlsx'))
        return source_folder

    def test_pipeline_beijing(self, tmp_dir):
        script_dir = os.path.join(tmp_dir, 'script')
        os.makedirs(script_dir, exist_ok=True)
        source = self._setup_folder(tmp_dir, script_dir, ['北京银行'])

        result = bankcheck.run_pipeline(source, script_dir)

        assert len(result.all_rows) == 2
        assert result.all_rows[0]['银行'] == '北京银行'
        assert result.all_rows[0]['主体'] == '北京XX科技有限公司'
        assert result.all_rows[0]['付款'] == -50000.0
        assert len(result.processed_files) == 1
        assert not result.folder_empty
        assert result.output_path is not None

    def test_pipeline_east_asia(self, tmp_dir):
        script_dir = os.path.join(tmp_dir, 'script')
        os.makedirs(script_dir, exist_ok=True)
        source = self._setup_folder(tmp_dir, script_dir, ['东亚银行'])

        result = bankcheck.run_pipeline(source, script_dir)

        assert len(result.all_rows) == 2
        assert result.all_rows[0]['银行'] == '东亚银行'
        assert result.all_rows[0]['主体'] == '上海YY贸易有限公司'

    def test_pipeline_both_banks(self, tmp_dir):
        script_dir = os.path.join(tmp_dir, 'script')
        os.makedirs(script_dir, exist_ok=True)
        source = self._setup_folder(tmp_dir, script_dir, ['北京银行', '东亚银行'])

        result = bankcheck.run_pipeline(source, script_dir)

        assert len(result.all_rows) == 4
        beijing_rows = [r for r in result.all_rows if r['银行'] == '北京银行']
        east_asia_rows = [r for r in result.all_rows if r['银行'] == '东亚银行']
        assert len(beijing_rows) == 2
        assert len(east_asia_rows) == 2

    def test_output_table(self, tmp_dir):
        script_dir = os.path.join(tmp_dir, 'script')
        os.makedirs(script_dir, exist_ok=True)
        source = self._setup_folder(tmp_dir, script_dir, ['北京银行', '东亚银行'])

        result = bankcheck.run_pipeline(source, script_dir)

        assert result.output_path is not None
        assert os.path.exists(result.output_path)
        df_read = pd.read_excel(result.output_path, engine='openpyxl')
        assert len(df_read) == 4
        columns = ['唯一id', '银行', '银行账号', '主体', '交易日期', '付款', '收款', '摘要', '对方户名', '余额', '交易流水号']
        assert list(df_read.columns) == columns

    def test_unprocessed_files_kept(self, tmp_dir):
        script_dir = os.path.join(tmp_dir, 'script')
        os.makedirs(script_dir, exist_ok=True)
        source = self._setup_folder(tmp_dir, script_dir, ['北京银行', '未知'])

        result = bankcheck.run_pipeline(source, script_dir)

        assert len(result.unprocessed_files) == 1
        assert '未知银行' in os.path.basename(result.unprocessed_files[0])

    def test_processed_files_deleted(self, tmp_dir):
        script_dir = os.path.join(tmp_dir, 'script')
        os.makedirs(script_dir, exist_ok=True)
        source = self._setup_folder(tmp_dir, script_dir, ['北京银行', '未知'])

        result = bankcheck.run_pipeline(source, script_dir)

        new_folder = source + '＋检验版'
        remaining = bankcheck.scan_excel_files(new_folder)
        assert len(remaining) == 1
        assert '未知银行' in os.path.basename(remaining[0])

    def test_original_folder_untouched(self, tmp_dir):
        script_dir = os.path.join(tmp_dir, 'script')
        os.makedirs(script_dir, exist_ok=True)
        source = self._setup_folder(tmp_dir, script_dir, ['北京银行'])

        original_count = len(os.listdir(source))
        bankcheck.run_pipeline(source, script_dir)

        assert len(os.listdir(source)) == original_count
        assert os.path.exists(os.path.join(source, '北京银行_流水.xlsx'))

    def test_empty_folder(self, tmp_dir):
        script_dir = os.path.join(tmp_dir, 'script')
        os.makedirs(script_dir, exist_ok=True)
        source = os.path.join(tmp_dir, '空文件夹')
        os.makedirs(source, exist_ok=True)
        _create_lookup_table(os.path.join(script_dir, '主体查找表.xlsx'))

        result = bankcheck.run_pipeline(source, script_dir)

        assert result.folder_empty is True
        assert len(result.all_rows) == 0
        assert result.output_path is None

    def test_lookup_missing(self, tmp_dir):
        script_dir = os.path.join(tmp_dir, 'script')
        os.makedirs(script_dir, exist_ok=True)
        source = self._setup_folder(tmp_dir, script_dir, ['北京银行'])
        os.remove(os.path.join(script_dir, '主体查找表.xlsx'))

        result = bankcheck.run_pipeline(source, script_dir)

        assert result.lookup_missing is True
        assert result.all_rows[0]['主体'] == ''

    def test_error_files_kept_after_pipeline(self, tmp_dir):
        script_dir = os.path.join(tmp_dir, 'script')
        os.makedirs(script_dir, exist_ok=True)
        source_folder = os.path.join(tmp_dir, '流水文件夹')
        os.makedirs(source_folder, exist_ok=True)

        corrupt_path = os.path.join(source_folder, '北京银行_坏.xlsx')
        with open(corrupt_path, 'wb') as f:
            f.write(b'not a valid excel file')

        _create_lookup_table(os.path.join(script_dir, '主体查找表.xlsx'))

        result = bankcheck.run_pipeline(source_folder, script_dir)

        assert len(result.error_files) == 1
        error_filepath = result.error_files[0][0]
        assert '北京银行_坏' in os.path.basename(error_filepath)

        new_folder = source_folder + '＋检验版'
        remaining = bankcheck.scan_excel_files(new_folder)
        assert len(remaining) == 1
        assert '北京银行_坏' in os.path.basename(remaining[0])

    def test_mixed_success_error_unprocessed(self, tmp_dir):
        script_dir = os.path.join(tmp_dir, 'script')
        os.makedirs(script_dir, exist_ok=True)
        source_folder = os.path.join(tmp_dir, '流水文件夹')
        os.makedirs(source_folder, exist_ok=True)

        _create_beijing_bank_excel(os.path.join(source_folder, '北京银行_正常.xlsx'))
        corrupt_path = os.path.join(source_folder, '北京银行_坏.xlsx')
        with open(corrupt_path, 'wb') as f:
            f.write(b'corrupt')
        wb = openpyxl.Workbook()
        ws = wb.active
        ws['A1'] = 'unknown'
        wb.save(os.path.join(source_folder, '未知银行_流水.xlsx'))
        wb.close()

        _create_lookup_table(os.path.join(script_dir, '主体查找表.xlsx'))

        result = bankcheck.run_pipeline(source_folder, script_dir)

        assert len(result.processed_files) == 1
        assert len(result.error_files) == 1
        assert len(result.unprocessed_files) == 1

        new_folder = source_folder + '＋检验版'
        remaining = bankcheck.scan_excel_files(new_folder)
        remaining_names = [os.path.basename(f) for f in remaining]
        assert '北京银行_坏.xlsx' in remaining_names
        assert '未知银行_流水.xlsx' in remaining_names
        assert '北京银行_正常.xlsx' not in remaining_names

    def test_bank_processors_registry(self):
        assert '北京银行' in bankcheck.BANK_PROCESSORS
        assert '东亚银行' in bankcheck.BANK_PROCESSORS
        assert bankcheck.BANK_PROCESSORS['北京银行'] == bankcheck.process_beijing_bank
        assert bankcheck.BANK_PROCESSORS['东亚银行'] == bankcheck.process_east_asia_bank

    def test_sample_files(self):
        samples_dir = os.path.join(os.path.dirname(__file__), '..', 'samples')
        if not os.path.isdir(samples_dir):
            pytest.skip('samples directory not found')

        script_dir = os.path.join(os.path.dirname(__file__), '..', 'backend')
        lookup_file = bankcheck.find_lookup_file(script_dir)
        if not lookup_file:
            pytest.skip('主体查找表 not found in backend directory')

        beijing_sample = os.path.join(samples_dir, '北京银行_示例流水.xlsx')
        east_asia_sample = os.path.join(samples_dir, '东亚银行_示例流水.xlsx')

        if os.path.exists(beijing_sample):
            rows = bankcheck.process_beijing_bank(beijing_sample, lookup_file)
            assert len(rows) > 0
            assert rows[0]['银行'] == '北京银行'

        if os.path.exists(east_asia_sample):
            rows = bankcheck.process_east_asia_bank(east_asia_sample, lookup_file)
            assert len(rows) > 0
            assert rows[0]['银行'] == '东亚银行'


class TestFormatResultMessage:
    def test_folder_empty(self):
        result = bankcheck.ProcessingResult(folder_empty=True)
        msg = bankcheck.format_result_message(result)
        assert msg == '文件夹中未发现任何 Excel 文件。'

    def test_with_records(self):
        result = bankcheck.ProcessingResult(
            all_rows=[{'银行': '北京银行'}, {'银行': '东亚银行'}],
            processed_files=['/path/a.xlsx', '/path/b.xlsx'],
            output_path='/path/to/银行流水总表.xlsx',
        )
        msg = bankcheck.format_result_message(result)
        assert '处理完成！' in msg
        assert '已处理文件数：2' in msg
        assert '提取记录数：2' in msg
        assert '/path/to/银行流水总表.xlsx' in msg

    def test_no_records(self):
        result = bankcheck.ProcessingResult()
        msg = bankcheck.format_result_message(result)
        assert '未提取到任何银行流水记录。' in msg

    def test_unprocessed_files(self):
        result = bankcheck.ProcessingResult(
            all_rows=[{'银行': '北京银行'}],
            processed_files=['/path/a.xlsx'],
            unprocessed_files=['/path/未知银行_流水.xlsx'],
            output_path='/path/to/银行流水总表.xlsx',
        )
        msg = bankcheck.format_result_message(result)
        assert '无法识别的文件（1 个' in msg
        assert '未知银行_流水.xlsx' in msg

    def test_error_files(self):
        result = bankcheck.ProcessingResult(
            error_files=[('/path/北京银行_坏.xlsx', 'Bad file')],
        )
        msg = bankcheck.format_result_message(result)
        assert '处理出错的文件（1 个' in msg
        assert '北京银行_坏.xlsx' in msg
        assert 'Bad file' in msg

    def test_error_files_message_says_preserved(self):
        result = bankcheck.ProcessingResult(
            error_files=[('/path/北京银行_坏.xlsx', 'File is corrupt')],
        )
        msg = bankcheck.format_result_message(result)
        assert '处理出错的文件' in msg
        assert '已保留' in msg
        assert '北京银行_坏.xlsx' in msg

    def test_unprocessed_and_error(self):
        result = bankcheck.ProcessingResult(
            unprocessed_files=['/path/未知.xlsx'],
            error_files=[('/path/北京银行_坏.xlsx', 'Error')],
        )
        msg = bankcheck.format_result_message(result)
        assert '无法识别的文件' in msg
        assert '处理出错的文件' in msg


class TestDeleteProcessedFiles:
    def test_deletes_non_kept_files(self, tmp_dir):
        f1 = os.path.join(tmp_dir, 'a.xlsx')
        f2 = os.path.join(tmp_dir, 'b.xlsx')
        open(f1, 'w').close()
        open(f2, 'w').close()

        bankcheck.delete_processed_files([f1, f2], {f2})

        assert not os.path.exists(f1)
        assert os.path.exists(f2)

    def test_keeps_all_when_all_in_keep_set(self, tmp_dir):
        f1 = os.path.join(tmp_dir, 'a.xlsx')
        f2 = os.path.join(tmp_dir, 'b.xlsx')
        open(f1, 'w').close()
        open(f2, 'w').close()

        bankcheck.delete_processed_files([f1, f2], {f1, f2})

        assert os.path.exists(f1)
        assert os.path.exists(f2)

    def test_deletes_all_when_keep_set_empty(self, tmp_dir):
        f1 = os.path.join(tmp_dir, 'a.xlsx')
        f2 = os.path.join(tmp_dir, 'b.xlsx')
        open(f1, 'w').close()
        open(f2, 'w').close()

        bankcheck.delete_processed_files([f1, f2], set())

        assert not os.path.exists(f1)
        assert not os.path.exists(f2)

    def test_empty_file_list(self, tmp_dir):
        bankcheck.delete_processed_files([], set())

    def test_nonexistent_file_handled_gracefully(self, tmp_dir):
        nonexistent = os.path.join(tmp_dir, 'missing.xlsx')
        bankcheck.delete_processed_files([nonexistent], set())

    def test_keeps_error_files_in_keep_set(self, tmp_dir):
        f1 = os.path.join(tmp_dir, '北京银行_正常.xlsx')
        f2 = os.path.join(tmp_dir, '北京银行_坏.xlsx')
        f3 = os.path.join(tmp_dir, '未知.xlsx')
        open(f1, 'w').close()
        open(f2, 'w').close()
        open(f3, 'w').close()

        error_file_paths = {f2}
        unprocessed = {f3}
        bankcheck.delete_processed_files([f1, f2, f3], unprocessed | error_file_paths)

        assert not os.path.exists(f1)
        assert os.path.exists(f2)
        assert os.path.exists(f3)
