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

    def _create_files(self, tmp_dir):
        f_success = os.path.join(tmp_dir, '北京银行_正常.xlsx')
        f_error = os.path.join(tmp_dir, '北京银行_坏.xlsx')
        f_unprocessed = os.path.join(tmp_dir, '未知.xlsx')
        for f in [f_success, f_error, f_unprocessed]:
            with open(f, 'w') as fp:
                fp.write('test')
        return f_success, f_error, f_unprocessed

    def test_strategy_keep_unprocessed_deletes_only_success_files(self, tmp_dir):
        f_success, f_error, f_unprocessed = self._create_files(tmp_dir)
        excel_files = [f_success, f_error, f_unprocessed]
        processed_files = [f_success]
        error_files = [(f_error, 'parse error')]
        unprocessed_files = [f_unprocessed]

        bankcheck.delete_processed_files(
            excel_files, processed_files, error_files, unprocessed_files,
            strategy='keep_unprocessed'
        )

        assert not os.path.exists(f_success)
        assert os.path.exists(f_error)
        assert os.path.exists(f_unprocessed)

    def test_strategy_keep_unprocessed_is_default(self, tmp_dir):
        f_success, f_error, f_unprocessed = self._create_files(tmp_dir)
        excel_files = [f_success, f_error, f_unprocessed]
        processed_files = [f_success]
        error_files = [(f_error, 'parse error')]
        unprocessed_files = [f_unprocessed]

        bankcheck.delete_processed_files(
            excel_files, processed_files, error_files, unprocessed_files
        )

        assert not os.path.exists(f_success)
        assert os.path.exists(f_error)
        assert os.path.exists(f_unprocessed)

    def test_strategy_keep_all_preserves_all_files(self, tmp_dir):
        f_success, f_error, f_unprocessed = self._create_files(tmp_dir)
        excel_files = [f_success, f_error, f_unprocessed]
        processed_files = [f_success]
        error_files = [(f_error, 'parse error')]
        unprocessed_files = [f_unprocessed]

        bankcheck.delete_processed_files(
            excel_files, processed_files, error_files, unprocessed_files,
            strategy='keep_all'
        )

        assert os.path.exists(f_success)
        assert os.path.exists(f_error)
        assert os.path.exists(f_unprocessed)

    def test_strategy_delete_all_removes_everything(self, tmp_dir):
        f_success, f_error, f_unprocessed = self._create_files(tmp_dir)
        excel_files = [f_success, f_error, f_unprocessed]
        processed_files = [f_success]
        error_files = [(f_error, 'parse error')]
        unprocessed_files = [f_unprocessed]

        bankcheck.delete_processed_files(
            excel_files, processed_files, error_files, unprocessed_files,
            strategy='delete_all'
        )

        assert not os.path.exists(f_success)
        assert not os.path.exists(f_error)
        assert not os.path.exists(f_unprocessed)

    def test_strategy_move_to_archive_moves_success_files(self, tmp_dir):
        f_success, f_error, f_unprocessed = self._create_files(tmp_dir)
        excel_files = [f_success, f_error, f_unprocessed]
        processed_files = [f_success]
        error_files = [(f_error, 'parse error')]
        unprocessed_files = [f_unprocessed]

        bankcheck.delete_processed_files(
            excel_files, processed_files, error_files, unprocessed_files,
            strategy='move_to_archive'
        )

        assert not os.path.exists(f_success)
        assert os.path.exists(f_error)
        assert os.path.exists(f_unprocessed)

        archive_dir = os.path.join(tmp_dir, '已处理归档')
        assert os.path.isdir(archive_dir)
        archived_files = os.listdir(archive_dir)
        assert len(archived_files) == 1
        assert '北京银行_正常.xlsx' in archived_files
        assert os.path.exists(os.path.join(archive_dir, '北京银行_正常.xlsx'))

    def test_strategy_move_to_archive_custom_dir_name(self, tmp_dir):
        f_success, f_error, f_unprocessed = self._create_files(tmp_dir)
        excel_files = [f_success, f_error, f_unprocessed]
        processed_files = [f_success]
        error_files = [(f_error, 'parse error')]
        unprocessed_files = [f_unprocessed]

        bankcheck.delete_processed_files(
            excel_files, processed_files, error_files, unprocessed_files,
            strategy='move_to_archive',
            archive_dir_name='processed_backup'
        )

        archive_dir = os.path.join(tmp_dir, 'processed_backup')
        assert os.path.isdir(archive_dir)
        assert len(os.listdir(archive_dir)) == 1

    def test_strategy_move_to_archive_handles_name_conflicts(self, tmp_dir):
        f_success1 = os.path.join(tmp_dir, '北京银行_正常.xlsx')
        f_success2 = os.path.join(tmp_dir, '北京银行_正常_dup.xlsx')
        for f in [f_success1, f_success2]:
            with open(f, 'w') as fp:
                fp.write('test')

        archive_dir = os.path.join(tmp_dir, '已处理归档')
        os.makedirs(archive_dir, exist_ok=True)
        with open(os.path.join(archive_dir, '北京银行_正常.xlsx'), 'w') as fp:
            fp.write('existing')

        excel_files = [f_success1, f_success2]
        processed_files = [f_success1, f_success2]
        error_files = []
        unprocessed_files = []

        bankcheck.delete_processed_files(
            excel_files, processed_files, error_files, unprocessed_files,
            strategy='move_to_archive'
        )

        archived = sorted(os.listdir(archive_dir))
        assert len(archived) >= 3
        assert '北京银行_正常.xlsx' in archived
        assert '北京银行_正常_1.xlsx' in archived

    def test_strategy_move_to_archive_no_processed_files(self, tmp_dir):
        f_error = os.path.join(tmp_dir, '北京银行_坏.xlsx')
        f_unprocessed = os.path.join(tmp_dir, '未知.xlsx')
        for f in [f_error, f_unprocessed]:
            with open(f, 'w') as fp:
                fp.write('test')

        excel_files = [f_error, f_unprocessed]
        processed_files = []
        error_files = [(f_error, 'parse error')]
        unprocessed_files = [f_unprocessed]

        bankcheck.delete_processed_files(
            excel_files, processed_files, error_files, unprocessed_files,
            strategy='move_to_archive'
        )

        assert os.path.exists(f_error)
        assert os.path.exists(f_unprocessed)

    def test_unknown_strategy_falls_back_to_keep_unprocessed(self, tmp_dir):
        f_success, f_error, f_unprocessed = self._create_files(tmp_dir)
        excel_files = [f_success, f_error, f_unprocessed]
        processed_files = [f_success]
        error_files = [(f_error, 'parse error')]
        unprocessed_files = [f_unprocessed]

        bankcheck.delete_processed_files(
            excel_files, processed_files, error_files, unprocessed_files,
            strategy='invalid_strategy_xyz'
        )

        assert not os.path.exists(f_success)
        assert os.path.exists(f_error)
        assert os.path.exists(f_unprocessed)

    def test_empty_file_list_all_strategies(self, tmp_dir):
        for strategy in ['keep_all', 'keep_unprocessed', 'delete_all', 'move_to_archive']:
            bankcheck.delete_processed_files(
                [], [], [], [], strategy=strategy
            )

    def test_nonexistent_file_handled_gracefully(self, tmp_dir):
        nonexistent = os.path.join(tmp_dir, 'missing.xlsx')
        for strategy in ['keep_unprocessed', 'delete_all']:
            bankcheck.delete_processed_files(
                [nonexistent], [nonexistent], [], [], strategy=strategy
            )

    def test_strategy_keep_all_with_empty_processed(self, tmp_dir):
        f_error, f_unprocessed = (
            os.path.join(tmp_dir, '坏.xlsx'),
            os.path.join(tmp_dir, '未知.xlsx'),
        )
        for f in [f_error, f_unprocessed]:
            with open(f, 'w') as fp:
                fp.write('test')

        bankcheck.delete_processed_files(
            [f_error, f_unprocessed], [], [(f_error, 'err')], [f_unprocessed],
            strategy='keep_all'
        )

        assert os.path.exists(f_error)
        assert os.path.exists(f_unprocessed)

    def test_strategy_move_to_archive_preserves_error_and_unprocessed(self, tmp_dir):
        f_success, f_error, f_unprocessed = self._create_files(tmp_dir)
        excel_files = [f_success, f_error, f_unprocessed]
        processed_files = [f_success]
        error_files = [(f_error, 'parse error')]
        unprocessed_files = [f_unprocessed]

        bankcheck.delete_processed_files(
            excel_files, processed_files, error_files, unprocessed_files,
            strategy='move_to_archive'
        )

        assert os.path.exists(f_error), 'error file should remain in original location'
        assert os.path.exists(f_unprocessed), 'unprocessed file should remain in original location'
        archive_dir = os.path.join(tmp_dir, '已处理归档')
        archived = os.listdir(archive_dir)
        assert os.path.basename(f_error) not in archived
        assert os.path.basename(f_unprocessed) not in archived

    def test_multiple_success_files_all_moved_to_archive(self, tmp_dir):
        files = []
        for i in range(3):
            f = os.path.join(tmp_dir, f'银行{i}_流水.xlsx')
            with open(f, 'w') as fp:
                fp.write(f'data{i}')
            files.append(f)

        bankcheck.delete_processed_files(
            files, files, [], [], strategy='move_to_archive'
        )

        archive_dir = os.path.join(tmp_dir, '已处理归档')
        archived = os.listdir(archive_dir)
        assert len(archived) == 3
        for f in files:
            assert not os.path.exists(f)
            assert os.path.basename(f) in archived


class TestKeepStrategiesConfig:

    def test_keep_strategies_dict_has_expected_keys(self):
        expected_keys = {'keep_unprocessed', 'keep_all', 'delete_all', 'move_to_archive'}
        assert set(bankcheck.KEEP_STRATEGIES.keys()) == expected_keys

    def test_keep_strategies_descriptions_are_non_empty(self):
        for key, desc in bankcheck.KEEP_STRATEGIES.items():
            assert isinstance(desc, str)
            assert len(desc.strip()) > 0, f'Description for {key} should not be empty'

    def test_default_keep_strategy_is_keep_unprocessed(self):
        import inspect
        sig = inspect.signature(bankcheck.run_pipeline)
        assert sig.parameters['keep_strategy'].default == 'keep_unprocessed'

        sig2 = inspect.signature(bankcheck.delete_processed_files)
        assert sig2.parameters['strategy'].default == 'keep_unprocessed'

        sig3 = inspect.signature(bankcheck.run_pipeline_with_options)
        assert sig3.parameters['keep_strategy'].default == 'keep_unprocessed'


class TestCliAskKeepStrategy:

    def test_default_choice_is_keep_unprocessed(self, monkeypatch):
        monkeypatch.setattr('builtins.input', lambda _: '')
        result = bankcheck.cli_ask_keep_strategy()
        assert result == 'keep_unprocessed'

    def test_choice_1_keep_unprocessed(self, monkeypatch):
        monkeypatch.setattr('builtins.input', lambda _: '1')
        result = bankcheck.cli_ask_keep_strategy()
        assert result == 'keep_unprocessed'

    def test_choice_2_keep_all(self, monkeypatch):
        monkeypatch.setattr('builtins.input', lambda _: '2')
        result = bankcheck.cli_ask_keep_strategy()
        assert result == 'keep_all'

    def test_choice_3_delete_all(self, monkeypatch):
        monkeypatch.setattr('builtins.input', lambda _: '3')
        result = bankcheck.cli_ask_keep_strategy()
        assert result == 'delete_all'

    def test_choice_4_move_to_archive(self, monkeypatch):
        monkeypatch.setattr('builtins.input', lambda _: '4')
        result = bankcheck.cli_ask_keep_strategy()
        assert result == 'move_to_archive'

    def test_invalid_choice_falls_back_to_default(self, monkeypatch):
        monkeypatch.setattr('builtins.input', lambda _: '999')
        result = bankcheck.cli_ask_keep_strategy()
        assert result == 'keep_unprocessed'

    def test_non_numeric_choice_falls_back_to_default(self, monkeypatch):
        monkeypatch.setattr('builtins.input', lambda _: 'abc')
        result = bankcheck.cli_ask_keep_strategy()
        assert result == 'keep_unprocessed'


class TestRunPipelineKeepStrategyIntegration:

    def _setup_folder(self, tmp_dir, script_dir):
        source_folder = os.path.join(tmp_dir, '流水文件夹')
        os.makedirs(source_folder, exist_ok=True)

        _create_beijing_bank_excel = bankcheck.__dict__.get('_create_beijing_bank_excel')
        if _create_beijing_bank_excel is None:
            from conftest import _create_beijing_bank_excel

        _create_beijing_bank_excel(os.path.join(source_folder, '北京银行_流水.xlsx'))

        wb = openpyxl.Workbook()
        ws = wb.active
        ws['A1'] = '未知银行数据'
        wb.save(os.path.join(source_folder, '未知银行_流水.xlsx'))
        wb.close()

        from conftest import _create_lookup_table
        _create_lookup_table(os.path.join(script_dir, '主体查找表.xlsx'))
        return source_folder

    def test_pipeline_keep_all_strategy_preserves_all_files(self, tmp_dir):
        script_dir = os.path.join(tmp_dir, 'script')
        os.makedirs(script_dir, exist_ok=True)
        source = self._setup_folder(tmp_dir, script_dir)

        result = bankcheck.run_pipeline(source, script_dir, keep_strategy='keep_all')

        assert len(result.processed_files) == 1
        assert len(result.unprocessed_files) == 1

        new_folder = source + '＋检验版'
        remaining = bankcheck.scan_excel_files(new_folder)
        assert len(remaining) == 2
        names = [os.path.basename(f) for f in remaining]
        assert '北京银行_流水.xlsx' in names
        assert '未知银行_流水.xlsx' in names

    def test_pipeline_move_to_archive_strategy_moves_success_files(self, tmp_dir):
        script_dir = os.path.join(tmp_dir, 'script')
        os.makedirs(script_dir, exist_ok=True)
        source = self._setup_folder(tmp_dir, script_dir)

        result = bankcheck.run_pipeline(source, script_dir, keep_strategy='move_to_archive')

        assert len(result.processed_files) == 1
        assert len(result.unprocessed_files) == 1

        new_folder = source + '＋检验版'
        success_path = os.path.join(new_folder, '北京银行_流水.xlsx')
        unprocessed_path = os.path.join(new_folder, '未知银行_流水.xlsx')
        assert not os.path.exists(success_path), '成功文件应从原位置移除'
        assert os.path.exists(unprocessed_path), '未处理文件应保留在原位置'

        archive_dir = os.path.join(new_folder, '已处理归档')
        assert os.path.isdir(archive_dir)
        archived = os.listdir(archive_dir)
        assert '北京银行_流水.xlsx' in archived
        assert os.path.exists(os.path.join(archive_dir, '北京银行_流水.xlsx'))

    def test_pipeline_delete_all_strategy_removes_everything(self, tmp_dir):
        script_dir = os.path.join(tmp_dir, 'script')
        os.makedirs(script_dir, exist_ok=True)
        source = self._setup_folder(tmp_dir, script_dir)

        result = bankcheck.run_pipeline(source, script_dir, keep_strategy='delete_all')

        assert len(result.processed_files) == 1
        assert len(result.unprocessed_files) == 1

        new_folder = source + '＋检验版'
        remaining = bankcheck.scan_excel_files(new_folder)
        assert len(remaining) == 0

    def test_pipeline_default_keep_unprocessed_deletes_only_success(self, tmp_dir):
        script_dir = os.path.join(tmp_dir, 'script')
        os.makedirs(script_dir, exist_ok=True)
        source = self._setup_folder(tmp_dir, script_dir)

        result = bankcheck.run_pipeline(source, script_dir)

        new_folder = source + '＋检验版'
        remaining = bankcheck.scan_excel_files(new_folder)
        remaining_names = [os.path.basename(f) for f in remaining]
        assert len(remaining) == 1
        assert '未知银行_流水.xlsx' in remaining_names
        assert '北京银行_流水.xlsx' not in remaining_names
