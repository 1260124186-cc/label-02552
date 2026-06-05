import os
import shutil

import openpyxl
import pandas as pd
import pytest

from conftest import _create_beijing_bank_excel, _create_east_asia_bank_excel, _create_lookup_table
import bankcheck


class TestMainIntegration:
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

    def test_full_pipeline_beijing(self, tmp_dir):
        script_dir = os.path.join(tmp_dir, 'script')
        os.makedirs(script_dir, exist_ok=True)
        source = self._setup_folder(tmp_dir, script_dir, ['北京银行'])

        new_folder = source + '＋检验版'
        shutil.copytree(source, new_folder)

        lookup_file = bankcheck.find_lookup_file(script_dir)
        excel_files = bankcheck.scan_excel_files(new_folder)
        assert len(excel_files) == 1

        all_rows = []
        processed = []
        unprocessed = []
        for fp in excel_files:
            bank = bankcheck.identify_bank(fp)
            if bank and bank in bankcheck.BANK_PROCESSORS:
                rows = bankcheck.BANK_PROCESSORS[bank](fp, lookup_file)
                all_rows.extend(rows)
                processed.append(fp)

        assert len(all_rows) == 2
        assert all_rows[0]['银行'] == '北京银行'
        assert all_rows[0]['主体'] == '北京XX科技有限公司'
        assert all_rows[0]['付款'] == -50000.0

    def test_full_pipeline_east_asia(self, tmp_dir):
        script_dir = os.path.join(tmp_dir, 'script')
        os.makedirs(script_dir, exist_ok=True)
        source = self._setup_folder(tmp_dir, script_dir, ['东亚银行'])

        new_folder = source + '＋检验版'
        shutil.copytree(source, new_folder)

        lookup_file = bankcheck.find_lookup_file(script_dir)
        excel_files = bankcheck.scan_excel_files(new_folder)

        all_rows = []
        for fp in excel_files:
            bank_name = bankcheck.identify_bank(fp)
            if bank_name and bank_name in bankcheck.BANK_PROCESSORS:
                rows = bankcheck.BANK_PROCESSORS[bank_name](fp, lookup_file)
                all_rows.extend(rows)

        assert len(all_rows) == 2
        assert all_rows[0]['银行'] == '东亚银行'
        assert all_rows[0]['主体'] == '上海YY贸易有限公司'

    def test_full_pipeline_both_banks(self, tmp_dir):
        script_dir = os.path.join(tmp_dir, 'script')
        os.makedirs(script_dir, exist_ok=True)
        source = self._setup_folder(tmp_dir, script_dir, ['北京银行', '东亚银行'])

        new_folder = source + '＋检验版'
        shutil.copytree(source, new_folder)

        lookup_file = bankcheck.find_lookup_file(script_dir)
        excel_files = bankcheck.scan_excel_files(new_folder)

        all_rows = []
        for fp in excel_files:
            bank_name = bankcheck.identify_bank(fp)
            if bank_name and bank_name in bankcheck.BANK_PROCESSORS:
                rows = bankcheck.BANK_PROCESSORS[bank_name](fp, lookup_file)
                all_rows.extend(rows)

        assert len(all_rows) == 4
        beijing_rows = [r for r in all_rows if r['银行'] == '北京银行']
        east_asia_rows = [r for r in all_rows if r['银行'] == '东亚银行']
        assert len(beijing_rows) == 2
        assert len(east_asia_rows) == 2

    def test_output_table(self, tmp_dir):
        script_dir = os.path.join(tmp_dir, 'script')
        os.makedirs(script_dir, exist_ok=True)
        source = self._setup_folder(tmp_dir, script_dir, ['北京银行', '东亚银行'])

        new_folder = source + '＋检验版'
        shutil.copytree(source, new_folder)

        lookup_file = bankcheck.find_lookup_file(script_dir)
        excel_files = bankcheck.scan_excel_files(new_folder)

        all_rows = []
        for fp in excel_files:
            bank_name = bankcheck.identify_bank(fp)
            if bank_name and bank_name in bankcheck.BANK_PROCESSORS:
                rows = bankcheck.BANK_PROCESSORS[bank_name](fp, lookup_file)
                all_rows.extend(rows)

        columns = ['唯一id', '银行', '银行账号', '主体', '交易日期', '付款', '收款', '摘要', '对方户名', '余额', '交易流水号']
        df = pd.DataFrame(all_rows, columns=columns)
        output_path = os.path.join(script_dir, '银行流水总表.xlsx')
        df.to_excel(output_path, index=False, engine='openpyxl')

        assert os.path.exists(output_path)
        df_read = pd.read_excel(output_path, engine='openpyxl')
        assert len(df_read) == 4
        assert list(df_read.columns) == columns

    def test_unprocessed_files_kept(self, tmp_dir):
        script_dir = os.path.join(tmp_dir, 'script')
        os.makedirs(script_dir, exist_ok=True)
        source = self._setup_folder(tmp_dir, script_dir, ['北京银行', '未知'])

        new_folder = source + '＋检验版'
        shutil.copytree(source, new_folder)

        excel_files = bankcheck.scan_excel_files(new_folder)
        unprocessed = []
        for fp in excel_files:
            bank_name = bankcheck.identify_bank(fp)
            if not bank_name or bank_name not in bankcheck.BANK_PROCESSORS:
                unprocessed.append(fp)

        assert len(unprocessed) == 1
        assert '未知银行' in os.path.basename(unprocessed[0])

    def test_processed_files_deleted(self, tmp_dir):
        script_dir = os.path.join(tmp_dir, 'script')
        os.makedirs(script_dir, exist_ok=True)
        source = self._setup_folder(tmp_dir, script_dir, ['北京银行', '未知'])

        new_folder = source + '＋检验版'
        shutil.copytree(source, new_folder)

        lookup_file = bankcheck.find_lookup_file(script_dir)
        excel_files = bankcheck.scan_excel_files(new_folder)

        processed = []
        unprocessed = []
        for fp in excel_files:
            bank_name = bankcheck.identify_bank(fp)
            if bank_name and bank_name in bankcheck.BANK_PROCESSORS:
                processed.append(fp)
            else:
                unprocessed.append(fp)

        for fp in excel_files:
            if fp not in unprocessed:
                os.remove(fp)

        remaining = bankcheck.scan_excel_files(new_folder)
        assert len(remaining) == 1
        assert '未知银行' in os.path.basename(remaining[0])

    def test_original_folder_untouched(self, tmp_dir):
        script_dir = os.path.join(tmp_dir, 'script')
        os.makedirs(script_dir, exist_ok=True)
        source = self._setup_folder(tmp_dir, script_dir, ['北京银行'])

        original_count = len(os.listdir(source))

        new_folder = source + '＋检验版'
        shutil.copytree(source, new_folder)

        assert len(os.listdir(source)) == original_count
        assert os.path.exists(os.path.join(source, '北京银行_流水.xlsx'))

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
