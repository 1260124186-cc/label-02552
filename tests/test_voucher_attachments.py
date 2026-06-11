# -*- coding: utf-8 -*-
"""
凭证附件关联模块单元测试
测试：database.py 的 voucher_attachments CRUD 及 bankcheck.py 的业务逻辑
"""

import os
import sys
import tempfile
import shutil
import pytest

backend_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'backend'))
sys.path.insert(0, backend_path)

import bankcheck
from database import (
    VoucherAttachment, SQLiteBackend, VoucherAttachmentQueryResult,
    add_voucher_attachment, update_voucher_attachment,
    remove_voucher_attachment, get_voucher_attachment_by_id,
    query_voucher_attachments,
)


@pytest.fixture(autouse=True)
def init_logging():
    bankcheck.setup_logging()


@pytest.fixture
def tmp_db_dir():
    d = tempfile.mkdtemp(prefix='voucher_test_')
    yield d
    shutil.rmtree(d, ignore_errors=True)


@pytest.fixture
def db(tmp_db_dir):
    db_path = os.path.join(tmp_db_dir, 'test.db')
    backend = SQLiteBackend(db_path=db_path)
    backend.connect()
    yield backend
    backend.disconnect()


@pytest.fixture
def sample_attachment_records():
    return [
        {
            '交易流水号': 'BJ20240105001',
            '附件路径': '/data/invoices/2024/01/BJ20240105001.pdf',
            '附件类型': '发票',
            '备注': '采购增值税专用发票',
        },
        {
            '交易流水号': 'BJ20240105001',
            '附件路径': '/data/receipts/2024/01/BJ20240105001.jpg',
            '附件类型': '回单',
            '备注': '银行回单扫描件',
        },
        {
            '交易流水号': 'ICBC20240109002',
            '附件路径': '/data/invoices/2024/01/ICBC20240109002.pdf',
            '附件类型': '发票',
            '备注': '销售发票',
        },
        {
            '交易流水号': 'CMB20240106001',
            '附件路径': '/data/misc/contract_2024_001.pdf',
            '附件类型': '其他',
            '备注': '合同附件',
        },
    ]


# ──────────────────────────────────────────────
# 数据模型测试
# ──────────────────────────────────────────────

class TestVoucherAttachmentModel:
    def test_from_dict_full(self):
        d = {
            'id': 1,
            '交易流水号': 'T001',
            '附件路径': '/path/to/file.pdf',
            '附件类型': '发票',
            '备注': '测试',
            '创建时间': '2024-01-01 00:00:00',
            '更新时间': '2024-01-02 00:00:00',
        }
        att = VoucherAttachment.from_dict(d)
        assert att.id == 1
        assert att.交易流水号 == 'T001'
        assert att.附件路径 == '/path/to/file.pdf'
        assert att.附件类型 == '发票'
        assert att.备注 == '测试'

    def test_from_dict_empty(self):
        att = VoucherAttachment.from_dict({})
        assert att.id is None
        assert att.交易流水号 == ''
        assert att.附件路径 == ''
        assert att.附件类型 == ''
        assert att.备注 is None

    def test_from_dict_none_values(self):
        att = VoucherAttachment.from_dict({'交易流水号': None, '附件路径': None, '附件类型': None})
        assert att.交易流水号 == ''
        assert att.附件路径 == ''
        assert att.附件类型 == ''

    def test_to_dict(self):
        att = VoucherAttachment(
            id=5,
            交易流水号='T005',
            附件路径='/tmp/x.jpg',
            附件类型='回单',
            备注='hello',
        )
        d = att.to_dict()
        assert d['id'] == 5
        assert d['交易流水号'] == 'T005'
        assert d['附件路径'] == '/tmp/x.jpg'
        assert d['附件类型'] == '回单'
        assert d['备注'] == 'hello'


# ──────────────────────────────────────────────
# 数据库 CRUD 测试 (SQLiteBackend)
# ──────────────────────────────────────────────

class TestSQLiteVoucherCRUD:
    def test_insert_attachment(self, db, sample_attachment_records):
        rec = sample_attachment_records[0]
        att = VoucherAttachment.from_dict(rec)
        new_id = db.insert_voucher_attachment(att)
        assert isinstance(new_id, int)
        assert new_id > 0

    def test_insert_multiple_same_transaction(self, db, sample_attachment_records):
        ids = []
        for rec in sample_attachment_records:
            att = VoucherAttachment.from_dict(rec)
            ids.append(db.insert_voucher_attachment(att))
        assert len(ids) == 4
        assert len(set(ids)) == 4

    def test_get_attachment_by_id(self, db, sample_attachment_records):
        rec = sample_attachment_records[0]
        att = VoucherAttachment.from_dict(rec)
        new_id = db.insert_voucher_attachment(att)
        fetched = db.get_voucher_attachment(new_id)
        assert fetched is not None
        assert fetched.id == new_id
        assert fetched.交易流水号 == rec['交易流水号']
        assert fetched.附件路径 == rec['附件路径']
        assert fetched.附件类型 == rec['附件类型']
        assert fetched.备注 == rec['备注']

    def test_get_nonexistent_attachment(self, db):
        assert db.get_voucher_attachment(999999) is None

    def test_update_attachment_path(self, db, sample_attachment_records):
        rec = sample_attachment_records[0]
        att = VoucherAttachment.from_dict(rec)
        new_id = db.insert_voucher_attachment(att)
        ok = db.update_voucher_attachment(new_id, attachment_path='/new/path.pdf')
        assert ok is True
        fetched = db.get_voucher_attachment(new_id)
        assert fetched.附件路径 == '/new/path.pdf'
        assert fetched.交易流水号 == rec['交易流水号']

    def test_update_attachment_type(self, db, sample_attachment_records):
        rec = sample_attachment_records[0]
        att = VoucherAttachment.from_dict(rec)
        new_id = db.insert_voucher_attachment(att)
        ok = db.update_voucher_attachment(new_id, attachment_type='回单')
        assert ok is True
        fetched = db.get_voucher_attachment(new_id)
        assert fetched.附件类型 == '回单'

    def test_update_attachment_remark(self, db, sample_attachment_records):
        rec = sample_attachment_records[0]
        att = VoucherAttachment.from_dict(rec)
        new_id = db.insert_voucher_attachment(att)
        ok = db.update_voucher_attachment(new_id, remark='新备注内容')
        assert ok is True
        fetched = db.get_voucher_attachment(new_id)
        assert fetched.备注 == '新备注内容'

    def test_update_attachment_multiple_fields(self, db, sample_attachment_records):
        rec = sample_attachment_records[0]
        att = VoucherAttachment.from_dict(rec)
        new_id = db.insert_voucher_attachment(att)
        ok = db.update_voucher_attachment(
            new_id,
            attachment_path='/updated.pdf',
            attachment_type='其他',
            remark='已更新',
        )
        assert ok is True
        fetched = db.get_voucher_attachment(new_id)
        assert fetched.附件路径 == '/updated.pdf'
        assert fetched.附件类型 == '其他'
        assert fetched.备注 == '已更新'

    def test_update_nonexistent_attachment(self, db):
        ok = db.update_voucher_attachment(999999, attachment_path='/x.pdf')
        assert ok is False

    def test_update_no_fields(self, db, sample_attachment_records):
        rec = sample_attachment_records[0]
        att = VoucherAttachment.from_dict(rec)
        new_id = db.insert_voucher_attachment(att)
        ok = db.update_voucher_attachment(new_id)
        assert ok is False

    def test_delete_attachment(self, db, sample_attachment_records):
        rec = sample_attachment_records[0]
        att = VoucherAttachment.from_dict(rec)
        new_id = db.insert_voucher_attachment(att)
        ok = db.delete_voucher_attachment(new_id)
        assert ok is True
        assert db.get_voucher_attachment(new_id) is None

    def test_delete_nonexistent_attachment(self, db):
        ok = db.delete_voucher_attachment(999999)
        assert ok is False

    def test_query_all(self, db, sample_attachment_records):
        for rec in sample_attachment_records:
            db.insert_voucher_attachment(VoucherAttachment.from_dict(rec))
        result = db.query_voucher_attachments()
        assert isinstance(result, VoucherAttachmentQueryResult)
        assert result.total_count == 4
        assert len(result.records) == 4

    def test_query_by_transaction_id(self, db, sample_attachment_records):
        for rec in sample_attachment_records:
            db.insert_voucher_attachment(VoucherAttachment.from_dict(rec))
        result = db.query_voucher_attachments(transaction_id='BJ20240105001')
        assert result.total_count == 2
        assert all(r.交易流水号 == 'BJ20240105001' for r in result.records)

    def test_query_by_attachment_type_invoice(self, db, sample_attachment_records):
        for rec in sample_attachment_records:
            db.insert_voucher_attachment(VoucherAttachment.from_dict(rec))
        result = db.query_voucher_attachments(attachment_type='发票')
        assert result.total_count == 2
        assert all(r.附件类型 == '发票' for r in result.records)

    def test_query_by_attachment_type_receipt(self, db, sample_attachment_records):
        for rec in sample_attachment_records:
            db.insert_voucher_attachment(VoucherAttachment.from_dict(rec))
        result = db.query_voucher_attachments(attachment_type='回单')
        assert result.total_count == 1
        assert result.records[0].附件类型 == '回单'

    def test_query_by_keyword_remark(self, db, sample_attachment_records):
        for rec in sample_attachment_records:
            db.insert_voucher_attachment(VoucherAttachment.from_dict(rec))
        result = db.query_voucher_attachments(keyword='扫描')
        assert result.total_count == 1
        assert '扫描' in result.records[0].备注

    def test_query_by_keyword_path(self, db, sample_attachment_records):
        for rec in sample_attachment_records:
            db.insert_voucher_attachment(VoucherAttachment.from_dict(rec))
        result = db.query_voucher_attachments(keyword='contract')
        assert result.total_count == 1
        assert 'contract' in result.records[0].附件路径

    def test_query_pagination(self, db, sample_attachment_records):
        for rec in sample_attachment_records:
            db.insert_voucher_attachment(VoucherAttachment.from_dict(rec))
        result = db.query_voucher_attachments(limit=2, offset=0)
        assert len(result.records) == 2
        assert result.total_count == 4

        result2 = db.query_voucher_attachments(limit=2, offset=2)
        assert len(result2.records) == 2
        assert result2.total_count == 4

        ids_page1 = {r.id for r in result.records}
        ids_page2 = {r.id for r in result2.records}
        assert ids_page1.isdisjoint(ids_page2)

    def test_query_order_by_update_time_desc(self, db, sample_attachment_records):
        ids = []
        for rec in sample_attachment_records:
            ids.append(db.insert_voucher_attachment(VoucherAttachment.from_dict(rec)))
        result = db.query_voucher_attachments()
        order_ids = [r.id for r in result.records]
        assert order_ids == sorted(order_ids, reverse=True)

    def test_query_empty_result(self, db):
        result = db.query_voucher_attachments()
        assert result.total_count == 0
        assert result.records == []


# ──────────────────────────────────────────────
# 便捷函数测试 (database.py level)
# ──────────────────────────────────────────────

class TestDatabaseConvenienceFunctions:
    def _config_for(self, tmp_db_dir):
        db_path = os.path.join(tmp_db_dir, 'conv.db')
        return {'backend': 'sqlite', 'sqlite': {'db_path': db_path}}

    def test_add_voucher_attachment(self, tmp_db_dir):
        cfg = self._config_for(tmp_db_dir)
        new_id = add_voucher_attachment(
            transaction_id='TID001',
            attachment_path='/a/b.pdf',
            attachment_type='发票',
            remark='便利函数测试',
            config=cfg,
        )
        assert isinstance(new_id, int)
        assert new_id > 0

    def test_get_voucher_attachment_by_id(self, tmp_db_dir):
        cfg = self._config_for(tmp_db_dir)
        new_id = add_voucher_attachment(
            transaction_id='TID002',
            attachment_path='/c/d.jpg',
            attachment_type='回单',
            config=cfg,
        )
        att = get_voucher_attachment_by_id(new_id, config=cfg)
        assert att is not None
        assert att.交易流水号 == 'TID002'
        assert att.附件路径 == '/c/d.jpg'

    def test_update_voucher_attachment(self, tmp_db_dir):
        cfg = self._config_for(tmp_db_dir)
        new_id = add_voucher_attachment('TID003', '/old.pdf', config=cfg)
        ok = update_voucher_attachment(new_id, attachment_path='/new.pdf', config=cfg)
        assert ok is True
        att = get_voucher_attachment_by_id(new_id, config=cfg)
        assert att.附件路径 == '/new.pdf'

    def test_remove_voucher_attachment(self, tmp_db_dir):
        cfg = self._config_for(tmp_db_dir)
        new_id = add_voucher_attachment('TID004', '/x.pdf', config=cfg)
        ok = remove_voucher_attachment(new_id, config=cfg)
        assert ok is True
        assert get_voucher_attachment_by_id(new_id, config=cfg) is None

    def test_query_voucher_attachments(self, tmp_db_dir):
        cfg = self._config_for(tmp_db_dir)
        for i in range(5):
            add_voucher_attachment(f'TID{i:03d}', f'/f{i}.pdf', config=cfg)
        result = query_voucher_attachments(limit=3, config=cfg)
        assert result.total_count == 5
        assert len(result.records) == 3


# ──────────────────────────────────────────────
# bankcheck.py 业务逻辑测试
# ──────────────────────────────────────────────

class TestBankcheckVoucherAttachments:
    def _tmp_config(self, tmp_db_dir):
        db_path = os.path.join(tmp_db_dir, 'bc.db')
        return {'backend': 'sqlite', 'sqlite': {'db_path': db_path}}

    def _get_db_config(self, tmp_db_dir):
        import database as db_module
        cfg_path = os.path.join(tmp_db_dir, 'database_config.json')
        cfg = self._tmp_config(tmp_db_dir)
        db_module.save_database_config(cfg, script_dir=tmp_db_dir)
        return tmp_db_dir

    def test_add_voucher_attachment_ok(self, tmp_db_dir):
        sdir = self._get_db_config(tmp_db_dir)
        new_id = bankcheck.add_voucher_attachment(
            transaction_id='BC001',
            attachment_path='/vouchers/BC001.pdf',
            attachment_type='发票',
            remark='采购发票',
            script_dir=sdir,
        )
        assert isinstance(new_id, int)
        assert new_id > 0

    def test_add_voucher_attachment_empty_transaction_id_raises(self, tmp_db_dir):
        sdir = self._get_db_config(tmp_db_dir)
        with pytest.raises(ValueError, match='交易流水号不能为空'):
            bankcheck.add_voucher_attachment(
                transaction_id='',
                attachment_path='/x.pdf',
                script_dir=sdir,
            )

    def test_add_voucher_attachment_whitespace_transaction_id_raises(self, tmp_db_dir):
        sdir = self._get_db_config(tmp_db_dir)
        with pytest.raises(ValueError, match='交易流水号不能为空'):
            bankcheck.add_voucher_attachment(
                transaction_id='   ',
                attachment_path='/x.pdf',
                script_dir=sdir,
            )

    def test_add_voucher_attachment_empty_path_raises(self, tmp_db_dir):
        sdir = self._get_db_config(tmp_db_dir)
        with pytest.raises(ValueError, match='附件路径不能为空'):
            bankcheck.add_voucher_attachment(
                transaction_id='TID',
                attachment_path='',
                script_dir=sdir,
            )

    def test_update_voucher_attachment(self, tmp_db_dir):
        sdir = self._get_db_config(tmp_db_dir)
        new_id = bankcheck.add_voucher_attachment('TID', '/old.pdf', script_dir=sdir)
        ok = bankcheck.update_voucher_attachment(
            new_id,
            attachment_path='/new.pdf',
            attachment_type='回单',
            remark='changed',
            script_dir=sdir,
        )
        assert ok is True

    def test_delete_voucher_attachment(self, tmp_db_dir):
        sdir = self._get_db_config(tmp_db_dir)
        new_id = bankcheck.add_voucher_attachment('TID', '/x.pdf', script_dir=sdir)
        ok = bankcheck.delete_voucher_attachment(new_id, script_dir=sdir)
        assert ok is True
        assert bankcheck.get_voucher_attachment(new_id, script_dir=sdir) is None

    def test_get_voucher_attachment(self, tmp_db_dir):
        sdir = self._get_db_config(tmp_db_dir)
        new_id = bankcheck.add_voucher_attachment(
            'TID001', '/a.pdf', '发票', '备注', script_dir=sdir
        )
        att = bankcheck.get_voucher_attachment(new_id, script_dir=sdir)
        assert att is not None
        assert att.交易流水号 == 'TID001'
        assert att.附件路径 == '/a.pdf'
        assert att.附件类型 == '发票'
        assert att.备注 == '备注'

    def test_list_voucher_attachments_filtered(self, tmp_db_dir):
        sdir = self._get_db_config(tmp_db_dir)
        bankcheck.add_voucher_attachment('T_A', '/a.pdf', '发票', script_dir=sdir)
        bankcheck.add_voucher_attachment('T_A', '/b.jpg', '回单', script_dir=sdir)
        bankcheck.add_voucher_attachment('T_B', '/c.pdf', '其他', script_dir=sdir)

        result = bankcheck.list_voucher_attachments(transaction_id='T_A', script_dir=sdir)
        assert result.total_count == 2

    def test_get_voucher_attachments_for_transaction(self, tmp_db_dir):
        sdir = self._get_db_config(tmp_db_dir)
        bankcheck.add_voucher_attachment('TX_MULTI', '/inv.pdf', '发票', script_dir=sdir)
        bankcheck.add_voucher_attachment('TX_MULTI', '/rec.jpg', '回单', script_dir=sdir)
        bankcheck.add_voucher_attachment('TX_OTHER', '/x.pdf', '发票', script_dir=sdir)

        atts = bankcheck.get_voucher_attachments_for_transaction('TX_MULTI', script_dir=sdir)
        assert len(atts) == 2
        assert all(a.交易流水号 == 'TX_MULTI' for a in atts)

    def test_get_voucher_attachments_for_transaction_empty(self, tmp_db_dir):
        sdir = self._get_db_config(tmp_db_dir)
        atts = bankcheck.get_voucher_attachments_for_transaction('NOEXIST', script_dir=sdir)
        assert atts == []

    def test_open_voucher_attachment_missing_id_and_path(self, tmp_db_dir):
        ok, msg = bankcheck.open_voucher_attachment()
        assert ok is False
        assert 'attachment_id' in msg or 'attachment_path' in msg

    def test_open_voucher_attachment_empty_path(self, tmp_db_dir):
        ok, msg = bankcheck.open_voucher_attachment(attachment_path='')
        assert ok is False
        assert '路径为空' in msg

    def test_open_voucher_attachment_nonexistent_file(self, tmp_db_dir):
        sdir = self._get_db_config(tmp_db_dir)
        bad_path = os.path.join(tmp_db_dir, 'definitely_not_exists_xyz123.pdf')
        ok, msg = bankcheck.open_voucher_attachment(attachment_path=bad_path)
        assert ok is False
        assert '不存在' in msg

    def test_open_voucher_attachment_nonexistent_record(self, tmp_db_dir):
        sdir = self._get_db_config(tmp_db_dir)
        ok, msg = bankcheck.open_voucher_attachment(attachment_id=9999999, script_dir=sdir)
        assert ok is False
        assert '未找到' in msg

    def test_open_voucher_attachment_real_file(self, tmp_db_dir):
        sdir = self._get_db_config(tmp_db_dir)
        real_file = os.path.join(tmp_db_dir, 'real.txt')
        with open(real_file, 'w') as f:
            f.write('test content')

        new_id = bankcheck.add_voucher_attachment(
            transaction_id='T_OPEN',
            attachment_path=real_file,
            attachment_type='其他',
            script_dir=sdir,
        )
        ok, msg = bankcheck.open_voucher_attachment(attachment_id=new_id, script_dir=sdir)
        assert ok is True, f'打开失败: {msg}'
        assert '已打开' in msg

    def test_open_voucher_attachments_for_transaction(self, tmp_db_dir):
        sdir = self._get_db_config(tmp_db_dir)
        f1 = os.path.join(tmp_db_dir, 'a.txt')
        f2 = os.path.join(tmp_db_dir, 'b.txt')
        for f in (f1, f2):
            with open(f, 'w') as fp:
                fp.write('x')
        bankcheck.add_voucher_attachment('OPEN_MULTI', f1, '发票', script_dir=sdir)
        bankcheck.add_voucher_attachment('OPEN_MULTI', f2, '回单', script_dir=sdir)
        bankcheck.add_voucher_attachment('OPEN_MULTI', '/nonexistent_xyz.pdf', '其他', script_dir=sdir)

        success_count, errors = bankcheck.open_voucher_attachments_for_transaction(
            'OPEN_MULTI', script_dir=sdir
        )
        assert success_count == 2
        assert len(errors) == 1

    def test_open_voucher_attachments_for_transaction_none(self, tmp_db_dir):
        sdir = self._get_db_config(tmp_db_dir)
        success_count, errors = bankcheck.open_voucher_attachments_for_transaction(
            'NOPE', script_dir=sdir
        )
        assert success_count == 0
        assert errors == []

    def test_voucher_attachment_types_constant(self):
        assert bankcheck.VOUCHER_ATTACHMENT_TYPES == ('发票', '回单', '其他')


# ──────────────────────────────────────────────
# VoucherAttachmentQueryResult 导出功能测试
# ──────────────────────────────────────────────

class TestVoucherAttachmentQueryResult:
    def test_to_dataframe(self, db, sample_attachment_records):
        for rec in sample_attachment_records:
            db.insert_voucher_attachment(VoucherAttachment.from_dict(rec))
        result = db.query_voucher_attachments()
        df = result.to_dataframe()
        assert len(df) == 4
        assert list(df.columns) == [
            'id', '交易流水号', '附件路径', '附件类型', '备注',
            '创建时间', '更新时间',
        ]

    def test_to_excel(self, tmp_db_dir, db, sample_attachment_records):
        for rec in sample_attachment_records:
            db.insert_voucher_attachment(VoucherAttachment.from_dict(rec))
        result = db.query_voucher_attachments()
        out = os.path.join(tmp_db_dir, 'voucher_export.xlsx')
        path = result.to_excel(out)
        assert os.path.exists(path)
        assert os.path.getsize(path) > 0
