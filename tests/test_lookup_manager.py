import os
import sys
import shutil
import tempfile

import pytest
import openpyxl

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'backend'))

import lookup_manager as lm


@pytest.fixture
def tmp_dir():
    d = tempfile.mkdtemp(prefix='lookup_test_')
    yield d
    shutil.rmtree(d, ignore_errors=True)


@pytest.fixture
def sample_lookup_file(tmp_dir):
    path = os.path.join(tmp_dir, '主体查找表.xlsx')
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '主体映射'
    ws.cell(row=1, column=1, value='主体名称')
    ws.cell(row=1, column=2, value='银行账号')
    ws.cell(row=2, column=1, value='北京XX科技有限公司')
    ws.cell(row=2, column=2, value='01090312345678901')
    ws.cell(row=3, column=1, value='上海YY贸易有限公司')
    ws.cell(row=3, column=2, value='38812345678')
    ws.cell(row=4, column=1, value='广州ZZ实业公司')
    ws.cell(row=4, column=2, value='6222021234567890123')
    wb.save(path)
    wb.close()
    return path


class TestFindLookupFile:
    def test_find_exact_xlsx(self, tmp_dir):
        path = os.path.join(tmp_dir, '主体查找表.xlsx')
        openpyxl.Workbook().save(path)
        result = lm.find_lookup_file(tmp_dir)
        assert result == path

    def test_find_exact_xls(self, tmp_dir):
        path = os.path.join(tmp_dir, '主体查找表.xls')
        open(path, 'a').close()
        result = lm.find_lookup_file(tmp_dir)
        assert result == path

    def test_not_found(self, tmp_dir):
        result = lm.find_lookup_file(tmp_dir)
        assert result is None


class TestReadLookupEntries:
    def test_read_entries(self, sample_lookup_file):
        entries = lm.read_lookup_entries(sample_lookup_file)
        assert len(entries) == 3
        assert entries[0].subject == '北京XX科技有限公司'
        assert entries[0].account == '01090312345678901'
        assert entries[1].subject == '上海YY贸易有限公司'
        assert entries[1].account == '38812345678'

    def test_read_nonexistent_file(self):
        entries = lm.read_lookup_entries('/nonexistent/path.xlsx')
        assert len(entries) == 0

    def test_read_empty_file(self, tmp_dir):
        path = os.path.join(tmp_dir, 'empty.xlsx')
        wb = openpyxl.Workbook()
        wb.save(path)
        wb.close()
        entries = lm.read_lookup_entries(path)
        assert len(entries) == 0

    def test_read_with_header_only(self, tmp_dir):
        path = os.path.join(tmp_dir, 'header_only.xlsx')
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.cell(row=1, column=1, value='主体名称')
        ws.cell(row=1, column=2, value='银行账号')
        wb.save(path)
        wb.close()
        entries = lm.read_lookup_entries(path)
        assert len(entries) == 0

    def test_read_normalizes_account(self, tmp_dir):
        path = os.path.join(tmp_dir, 'normalize_test.xlsx')
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.cell(row=1, column=1, value='主体名称')
        ws.cell(row=1, column=2, value='银行账号')
        ws.cell(row=2, column=1, value='测试公司')
        ws.cell(row=2, column=2, value=12345.0)
        ws.cell(row=3, column=1, value='测试公司2')
        ws.cell(row=3, column=2, value='  67890  ')
        wb.save(path)
        wb.close()
        entries = lm.read_lookup_entries(path)
        assert len(entries) == 2
        assert entries[0].account == '12345'
        assert entries[1].account == '67890'


class TestSaveLookupEntries:
    def test_save_entries(self, tmp_dir):
        path = os.path.join(tmp_dir, 'test_save.xlsx')
        entries = [
            lm.LookupEntry(subject='公司A', account='111'),
            lm.LookupEntry(subject='公司B', account='222'),
        ]
        success = lm.save_lookup_entries(entries, path)
        assert success is True
        assert os.path.exists(path)

        saved_entries = lm.read_lookup_entries(path)
        assert len(saved_entries) == 2
        assert saved_entries[0].subject == '公司A'
        assert saved_entries[0].account == '111'

    def test_save_overwrites_existing(self, sample_lookup_file):
        original_entries = lm.read_lookup_entries(sample_lookup_file)
        assert len(original_entries) == 3

        new_entries = [
            lm.LookupEntry(subject='新公司', account='999'),
        ]
        success = lm.save_lookup_entries(new_entries, sample_lookup_file)
        assert success is True

        saved_entries = lm.read_lookup_entries(sample_lookup_file)
        assert len(saved_entries) == 1
        assert saved_entries[0].subject == '新公司'


class TestGetEntryByAccount:
    def test_get_found(self, sample_lookup_file):
        entry = lm.get_entry_by_account('01090312345678901', sample_lookup_file)
        assert entry is not None
        assert entry.subject == '北京XX科技有限公司'

    def test_get_not_found(self, sample_lookup_file):
        entry = lm.get_entry_by_account('999999999', sample_lookup_file)
        assert entry is None

    def test_get_with_normalization(self, sample_lookup_file):
        entry = lm.get_entry_by_account(1090312345678901, sample_lookup_file)
        assert entry is not None
        assert entry.subject == '北京XX科技有限公司'

    def test_get_with_leading_zero(self, sample_lookup_file):
        entry = lm.get_entry_by_account('001090312345678901', sample_lookup_file)
        assert entry is not None
        assert entry.subject == '北京XX科技有限公司'


class TestAddEntry:
    def test_add_success(self, sample_lookup_file):
        success, msg = lm.add_entry('新公司', '999999999', sample_lookup_file)
        assert success is True
        assert '添加成功' in msg

        entries = lm.read_lookup_entries(sample_lookup_file)
        assert len(entries) == 4
        assert any(e.account == '999999999' for e in entries)

    def test_add_duplicate(self, sample_lookup_file):
        success, msg = lm.add_entry('重复公司', '01090312345678901', sample_lookup_file)
        assert success is False
        assert '已存在' in msg

        entries = lm.read_lookup_entries(sample_lookup_file)
        assert len(entries) == 3

    def test_add_empty_subject(self, sample_lookup_file):
        success, msg = lm.add_entry('', '12345', sample_lookup_file)
        assert success is False
        assert '不能为空' in msg

    def test_add_empty_account(self, sample_lookup_file):
        success, msg = lm.add_entry('测试公司', '', sample_lookup_file)
        assert success is False
        assert '不能为空' in msg

    def test_add_normalizes_account(self, sample_lookup_file):
        success, msg = lm.add_entry('测试公司', '  88888  ', sample_lookup_file)
        assert success is True

        entry = lm.get_entry_by_account('88888', sample_lookup_file)
        assert entry is not None
        assert entry.account == '88888'


class TestUpdateEntry:
    def test_update_success(self, sample_lookup_file):
        success, msg = lm.update_entry(
            '01090312345678901',
            '更新后的公司名',
            '01090312345678901',
            sample_lookup_file
        )
        assert success is True
        assert '更新成功' in msg

        entry = lm.get_entry_by_account('01090312345678901', sample_lookup_file)
        assert entry.subject == '更新后的公司名'

    def test_update_account(self, sample_lookup_file):
        success, msg = lm.update_entry(
            '01090312345678901',
            '北京XX科技有限公司',
            '111111111111',
            sample_lookup_file
        )
        assert success is True

        old_entry = lm.get_entry_by_account('01090312345678901', sample_lookup_file)
        assert old_entry is None

        new_entry = lm.get_entry_by_account('111111111111', sample_lookup_file)
        assert new_entry is not None
        assert new_entry.subject == '北京XX科技有限公司'

    def test_update_not_found(self, sample_lookup_file):
        success, msg = lm.update_entry(
            '999999999',
            '测试公司',
            '999999999',
            sample_lookup_file
        )
        assert success is False
        assert '未找到' in msg

    def test_update_conflict(self, sample_lookup_file):
        success, msg = lm.update_entry(
            '01090312345678901',
            '测试公司',
            '38812345678',
            sample_lookup_file
        )
        assert success is False
        assert '已被使用' in msg


class TestDeleteEntry:
    def test_delete_success(self, sample_lookup_file):
        success, msg = lm.delete_entry('01090312345678901', sample_lookup_file)
        assert success is True
        assert '成功删除' in msg

        entries = lm.read_lookup_entries(sample_lookup_file)
        assert len(entries) == 2
        assert all(e.account != '01090312345678901' for e in entries)

    def test_delete_not_found(self, sample_lookup_file):
        success, msg = lm.delete_entry('999999999', sample_lookup_file)
        assert success is False
        assert '未找到' in msg

    def test_delete_with_normalization(self, sample_lookup_file):
        success, msg = lm.delete_entry(1090312345678901, sample_lookup_file)
        assert success is True

        entries = lm.read_lookup_entries(sample_lookup_file)
        assert len(entries) == 2


class TestSearchEntries:
    def test_search_by_subject(self, sample_lookup_file):
        results = lm.search_entries('北京', sample_lookup_file)
        assert len(results) == 1
        assert results[0].subject == '北京XX科技有限公司'

    def test_search_by_account(self, sample_lookup_file):
        results = lm.search_entries('38812', sample_lookup_file)
        assert len(results) == 1
        assert results[0].account == '38812345678'

    def test_search_empty_keyword(self, sample_lookup_file):
        results = lm.search_entries('', sample_lookup_file)
        assert len(results) == 3

    def test_search_no_results(self, sample_lookup_file):
        results = lm.search_entries('不存在的关键词', sample_lookup_file)
        assert len(results) == 0

    def test_search_case_insensitive(self, sample_lookup_file):
        results = lm.search_entries('beijing', sample_lookup_file)
        assert len(results) == 0

        results = lm.search_entries('北京', sample_lookup_file)
        assert len(results) == 1


class TestDuplicateEntries:
    def test_no_duplicates(self, sample_lookup_file):
        duplicates = lm.get_duplicate_entries(sample_lookup_file)
        assert len(duplicates) == 0

    def test_with_duplicates(self, tmp_dir):
        path = os.path.join(tmp_dir, 'dup.xlsx')
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.cell(row=1, column=1, value='主体名称')
        ws.cell(row=1, column=2, value='银行账号')
        ws.cell(row=2, column=1, value='公司A')
        ws.cell(row=2, column=2, value='12345')
        ws.cell(row=3, column=1, value='公司B')
        ws.cell(row=3, column=2, value='12345')
        ws.cell(row=4, column=1, value='公司C')
        ws.cell(row=4, column=2, value='67890')
        wb.save(path)
        wb.close()

        duplicates = lm.get_duplicate_entries(path)
        assert len(duplicates) == 1
        assert duplicates[0]['account'] == '12345'
        assert duplicates[0]['count'] == 2
        assert '公司A' in duplicates[0]['subjects']
        assert '公司B' in duplicates[0]['subjects']


class TestImportExport:
    def test_export_excel(self, sample_lookup_file, tmp_dir):
        export_path = os.path.join(tmp_dir, 'export.xlsx')
        success, msg = lm.export_to_excel(export_path, sample_lookup_file)
        assert success is True
        assert os.path.exists(export_path)

        exported_entries = lm.read_lookup_entries(export_path)
        original_entries = lm.read_lookup_entries(sample_lookup_file)
        assert len(exported_entries) == len(original_entries)

    def test_import_merge(self, sample_lookup_file, tmp_dir):
        import_path = os.path.join(tmp_dir, 'import.xlsx')
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.cell(row=1, column=1, value='主体名称')
        ws.cell(row=1, column=2, value='银行账号')
        ws.cell(row=2, column=1, value='新公司')
        ws.cell(row=2, column=2, value='99999')
        ws.cell(row=3, column=1, value='更新后的北京公司')
        ws.cell(row=3, column=2, value='01090312345678901')
        wb.save(import_path)
        wb.close()

        success, msg, stats = lm.import_from_excel(
            import_path, overwrite=False, lookup_file=sample_lookup_file
        )
        assert success is True
        assert stats['imported'] == 1
        assert stats['updated'] == 1
        assert stats['skipped'] == 0

        entries = lm.read_lookup_entries(sample_lookup_file)
        assert len(entries) == 4

        entry = lm.get_entry_by_account('01090312345678901', sample_lookup_file)
        assert entry.subject == '更新后的北京公司'

    def test_import_overwrite(self, sample_lookup_file, tmp_dir):
        import_path = os.path.join(tmp_dir, 'import.xlsx')
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.cell(row=1, column=1, value='主体名称')
        ws.cell(row=1, column=2, value='银行账号')
        ws.cell(row=2, column=1, value='全新公司')
        ws.cell(row=2, column=2, value='11111')
        wb.save(import_path)
        wb.close()

        success, msg, stats = lm.import_from_excel(
            import_path, overwrite=True, lookup_file=sample_lookup_file
        )
        assert success is True
        assert stats['imported'] == 1

        entries = lm.read_lookup_entries(sample_lookup_file)
        assert len(entries) == 1
        assert entries[0].subject == '全新公司'

    def test_import_nonexistent_file(self):
        success, msg, stats = lm.import_from_excel('/nonexistent/file.xlsx')
        assert success is False
        assert '不存在' in msg


class TestAccountNormalization:
    def test_normalize_none(self):
        assert lm._normalize_account_str(None) == ''

    def test_normalize_int(self):
        assert lm._normalize_account_str(12345) == '12345'

    def test_normalize_float_whole(self):
        assert lm._normalize_account_str(12345.0) == '12345'

    def test_normalize_float_with_decimals(self):
        assert lm._normalize_account_str(123.45) == '123.45'

    def test_normalize_string_with_spaces(self):
        assert lm._normalize_account_str('  12345  ') == '12345'

    def test_normalize_string_dot_zero(self):
        assert lm._normalize_account_str('38812345678.0') == '38812345678'

    def test_account_key_leading_zero(self):
        assert lm._account_key('01090312345678901') == '1090312345678901'

    def test_account_key_int(self):
        assert lm._account_key(1090312345678901) == '1090312345678901'

    def test_account_key_equality(self):
        assert lm._account_key('01090312345678901') == lm._account_key(1090312345678901)
