# -*- coding: utf-8 -*-
"""
数据库持久化模块单元测试
"""

import os
import sys
import tempfile
import shutil
import pytest
from datetime import datetime

backend_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'backend'))
sys.path.insert(0, backend_path)

import bankcheck
from database import (
    TransactionRecord, SQLiteBackend, QueryResult,
    persist_transactions, query_transactions, get_db_statistics,
    create_database_backend, load_database_config,
)


@pytest.fixture
def setup_logging():
    bankcheck.setup_logging()


@pytest.fixture
def tmp_db_dir():
    d = tempfile.mkdtemp(prefix='bankcheck_db_test_')
    yield d
    shutil.rmtree(d, ignore_errors=True)


@pytest.fixture
def sample_records():
    return [
        {
            '唯一id': '20240101000000000001' + 'a' * 32,
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
        {
            '唯一id': '20240101000000000002' + 'b' * 32,
            '银行': '北京银行',
            '银行账号': '01090312345678901',
            '主体': '北京XX科技有限公司',
            '交易日期': '2024-01-10',
            '付款': None,
            '收款': 80000.0,
            '摘要': '销售收款',
            '对方户名': '客户B公司',
            '余额': 1580000.0,
            '交易流水号': 'BJ20240110002',
        },
        {
            '唯一id': '20240101000000000003' + 'c' * 32,
            '银行': '东亚银行',
            '银行账号': '38812345678',
            '主体': '上海YY贸易有限公司',
            '交易日期': '2024-01-03',
            '付款': -20000.0,
            '收款': None,
            '摘要': '向 张三 付款-货款',
            '对方户名': '张三',
            '余额': 480000.0,
            '交易流水号': 'EA20240103001',
        },
        {
            '唯一id': '20240101000000000004' + 'd' * 32,
            '银行': '东亚银行',
            '银行账号': '38812345678',
            '主体': '上海YY贸易有限公司',
            '交易日期': '2024-01-08',
            '付款': None,
            '收款': 35000.0,
            '摘要': '收到 李四 汇款-服务费',
            '对方户名': '李四',
            '余额': 515000.0,
            '交易流水号': 'EA20240108002',
        },
    ]


class TestTransactionRecord:
    """测试 TransactionRecord 数据类"""

    def test_from_dict(self, setup_logging):
        row_dict = {
            '唯一id': 'test123',
            '银行': '北京银行',
            '银行账号': '123456',
            '主体': '测试公司',
            '交易日期': '2024-01-01',
            '付款': '-1000',
            '收款': None,
            '摘要': '测试',
            '对方户名': '测试对方',
            '余额': '99999',
            '交易流水号': 'TXN001',
        }
        record = TransactionRecord.from_dict(row_dict)
        assert record.唯一id == 'test123'
        assert record.银行 == '北京银行'
        assert record.银行账号 == '123456'
        assert record.主体 == '测试公司'
        assert record.交易日期 == '2024-01-01'
        assert record.付款 == -1000.0
        assert record.收款 is None
        assert record.摘要 == '测试'
        assert record.对方户名 == '测试对方'
        assert record.余额 == 99999.0
        assert record.交易流水号 == 'TXN001'

    def test_to_dict(self, setup_logging):
        record = TransactionRecord(
            唯一id='test123',
            银行='北京银行',
            银行账号='123456',
            主体='测试公司',
        )
        d = record.to_dict()
        assert d['唯一id'] == 'test123'
        assert d['银行'] == '北京银行'
        assert d['银行账号'] == '123456'
        assert d['主体'] == '测试公司'

    def test_compute_match_key(self, setup_logging):
        record = TransactionRecord(
            唯一id='test123',
            银行='北京银行',
            银行账号='01090312345678901',
            交易流水号='BJ20240105001',
            交易日期='2024-01-05',
            付款=-50000.0,
        )
        key = record.compute_match_key()
        assert key is not None
        assert 'BJ20240105001' in key


class TestSQLiteBackend:
    """测试 SQLite 数据库后端"""

    def test_init_schema(self, setup_logging, tmp_db_dir):
        db_path = os.path.join(tmp_db_dir, 'test.db')
        with SQLiteBackend(db_path=db_path) as db:
            assert os.path.exists(db_path)

    def test_insert_and_query(self, setup_logging, tmp_db_dir, sample_records):
        db_path = os.path.join(tmp_db_dir, 'test.db')
        with SQLiteBackend(db_path=db_path) as db:
            records = [TransactionRecord.from_dict(r) for r in sample_records]
            inserted, duplicates = db.insert_records(records, deduplicate=False)

            assert inserted == 4
            assert duplicates == 0

            result = db.query_records()
            assert result.total_count == 4
            assert len(result.records) == 4

    def test_deduplication(self, setup_logging, tmp_db_dir, sample_records):
        db_path = os.path.join(tmp_db_dir, 'test.db')
        with SQLiteBackend(db_path=db_path) as db:
            records = [TransactionRecord.from_dict(r) for r in sample_records]

            inserted1, dup1 = db.insert_records(records, deduplicate=True)
            assert inserted1 == 4
            assert dup1 == 0

            inserted2, dup2 = db.insert_records(records, deduplicate=True)
            assert inserted2 == 0
            assert dup2 == 4

            result = db.query_records()
            assert result.total_count == 4

    def test_query_by_subject(self, setup_logging, tmp_db_dir, sample_records):
        db_path = os.path.join(tmp_db_dir, 'test.db')
        with SQLiteBackend(db_path=db_path) as db:
            records = [TransactionRecord.from_dict(r) for r in sample_records]
            db.insert_records(records, deduplicate=False)

            result = db.query_records(subject='北京XX科技')
            assert result.total_count == 2
            for r in result.records:
                assert '北京XX科技' in (r.主体 or '')

    def test_query_by_account(self, setup_logging, tmp_db_dir, sample_records):
        db_path = os.path.join(tmp_db_dir, 'test.db')
        with SQLiteBackend(db_path=db_path) as db:
            records = [TransactionRecord.from_dict(r) for r in sample_records]
            db.insert_records(records, deduplicate=False)

            result = db.query_records(account='38812345678')
            assert result.total_count == 2
            for r in result.records:
                assert r.银行账号 == '38812345678'

    def test_query_by_bank(self, setup_logging, tmp_db_dir, sample_records):
        db_path = os.path.join(tmp_db_dir, 'test.db')
        with SQLiteBackend(db_path=db_path) as db:
            records = [TransactionRecord.from_dict(r) for r in sample_records]
            db.insert_records(records, deduplicate=False)

            result = db.query_records(bank='北京银行')
            assert result.total_count == 2
            for r in result.records:
                assert r.银行 == '北京银行'

    def test_query_by_date_range(self, setup_logging, tmp_db_dir, sample_records):
        db_path = os.path.join(tmp_db_dir, 'test.db')
        with SQLiteBackend(db_path=db_path) as db:
            records = [TransactionRecord.from_dict(r) for r in sample_records]
            db.insert_records(records, deduplicate=False)

            result = db.query_records(start_date='2024-01-05', end_date='2024-01-08')
            assert result.total_count == 2
            dates = {r.交易日期 for r in result.records}
            assert '2024-01-05' in dates
            assert '2024-01-08' in dates

            result2 = db.query_records(start_date='2024-01-01', end_date='2024-01-05')
            assert result2.total_count == 2
            dates2 = {r.交易日期 for r in result2.records}
            assert '2024-01-03' in dates2
            assert '2024-01-05' in dates2

            result3 = db.query_records(start_date='2024-01-10')
            assert result3.total_count == 1
            assert result3.records[0].交易日期 == '2024-01-10'

    def test_query_by_amount(self, setup_logging, tmp_db_dir, sample_records):
        db_path = os.path.join(tmp_db_dir, 'test.db')
        with SQLiteBackend(db_path=db_path) as db:
            records = [TransactionRecord.from_dict(r) for r in sample_records]
            db.insert_records(records, deduplicate=False)

            result = db.query_records(min_amount=50000)
            assert result.total_count == 2
            for r in result.records:
                abs_amount = max(abs(r.付款 or 0), r.收款 or 0)
                assert abs_amount >= 50000

            result2 = db.query_records(max_amount=35000)
            assert result2.total_count == 2
            for r in result2.records:
                abs_amount = max(abs(r.付款 or 0), r.收款 or 0)
                assert abs_amount <= 35000

    def test_query_by_counterpart(self, setup_logging, tmp_db_dir, sample_records):
        db_path = os.path.join(tmp_db_dir, 'test.db')
        with SQLiteBackend(db_path=db_path) as db:
            records = [TransactionRecord.from_dict(r) for r in sample_records]
            db.insert_records(records, deduplicate=False)

            result = db.query_records(counterpart='供应商')
            assert result.total_count == 1
            assert '供应商' in (result.records[0].对方户名 or '')

    def test_query_by_summary(self, setup_logging, tmp_db_dir, sample_records):
        db_path = os.path.join(tmp_db_dir, 'test.db')
        with SQLiteBackend(db_path=db_path) as db:
            records = [TransactionRecord.from_dict(r) for r in sample_records]
            db.insert_records(records, deduplicate=False)

            result = db.query_records(summary_keyword='采购')
            assert result.total_count == 1
            assert '采购' in (result.records[0].摘要 or '')

    def test_query_combined_conditions(self, setup_logging, tmp_db_dir, sample_records):
        db_path = os.path.join(tmp_db_dir, 'test.db')
        with SQLiteBackend(db_path=db_path) as db:
            records = [TransactionRecord.from_dict(r) for r in sample_records]
            db.insert_records(records, deduplicate=False)

            result = db.query_records(
                subject='北京XX科技',
                bank='北京银行',
                start_date='2024-01-01',
                end_date='2024-01-31',
            )
            assert result.total_count == 2

    def test_query_pagination(self, setup_logging, tmp_db_dir, sample_records):
        db_path = os.path.join(tmp_db_dir, 'test.db')
        with SQLiteBackend(db_path=db_path) as db:
            records = [TransactionRecord.from_dict(r) for r in sample_records]
            db.insert_records(records, deduplicate=False)

            result1 = db.query_records(limit=2, offset=0)
            assert len(result1.records) == 2

            result2 = db.query_records(limit=2, offset=2)
            assert len(result2.records) == 2

            ids1 = {r.唯一id for r in result1.records}
            ids2 = {r.唯一id for r in result2.records}
            assert ids1.isdisjoint(ids2)

    def test_query_ordering(self, setup_logging, tmp_db_dir, sample_records):
        db_path = os.path.join(tmp_db_dir, 'test.db')
        with SQLiteBackend(db_path=db_path) as db:
            records = [TransactionRecord.from_dict(r) for r in sample_records]
            db.insert_records(records, deduplicate=False)

            result_desc = db.query_records(order_by='交易日期', ascending=False)
            dates_desc = [r.交易日期 for r in result_desc.records]
            assert dates_desc == sorted(dates_desc, reverse=True)

            result_asc = db.query_records(order_by='交易日期', ascending=True)
            dates_asc = [r.交易日期 for r in result_asc.records]
            assert dates_asc == sorted(dates_asc)

    def test_query_summary(self, setup_logging, tmp_db_dir, sample_records):
        db_path = os.path.join(tmp_db_dir, 'test.db')
        with SQLiteBackend(db_path=db_path) as db:
            records = [TransactionRecord.from_dict(r) for r in sample_records]
            db.insert_records(records, deduplicate=False)

            result = db.query_records()
            summary = result.summary

            assert summary['记录总数'] == 4
            assert summary['付款笔数'] == 2
            assert summary['收款笔数'] == 2
            assert summary['付款总额'] == -70000.0
            assert summary['收款总额'] == 115000.0
            assert summary['净额'] == 45000.0

    def test_get_statistics(self, setup_logging, tmp_db_dir, sample_records):
        db_path = os.path.join(tmp_db_dir, 'test.db')
        with SQLiteBackend(db_path=db_path) as db:
            records = [TransactionRecord.from_dict(r) for r in sample_records]
            db.insert_records(records, deduplicate=False)

            stats = db.get_statistics()
            assert stats['总记录数'] == 4
            assert stats['导入批次数量'] == 1

            by_bank = stats['按银行统计']
            assert len(by_bank) == 2

            by_subject = stats['按主体统计']
            assert len(by_subject) == 2

            by_account = stats['按账号统计']
            assert len(by_account) == 2

    def test_delete_records(self, setup_logging, tmp_db_dir, sample_records):
        db_path = os.path.join(tmp_db_dir, 'test.db')
        with SQLiteBackend(db_path=db_path) as db:
            records = [TransactionRecord.from_dict(r) for r in sample_records]
            db.insert_records(records, deduplicate=False)

            deleted = db.delete_records(subject='北京XX科技有限公司')
            assert deleted == 2

            result = db.query_records()
            assert result.total_count == 2
            for r in result.records:
                assert r.主体 != '北京XX科技有限公司'

    def test_get_existing_match_keys(self, setup_logging, tmp_db_dir, sample_records):
        db_path = os.path.join(tmp_db_dir, 'test.db')
        with SQLiteBackend(db_path=db_path) as db:
            records = [TransactionRecord.from_dict(r) for r in sample_records]
            db.insert_records(records, deduplicate=False)

            keys = db.get_existing_match_keys()
            assert len(keys) == 4

    def test_batch_tracking(self, setup_logging, tmp_db_dir, sample_records):
        db_path = os.path.join(tmp_db_dir, 'test.db')
        with SQLiteBackend(db_path=db_path) as db:
            records = [TransactionRecord.from_dict(r) for r in sample_records]
            inserted, _ = db.insert_records(records, batch_id='TEST_BATCH_001', deduplicate=False)
            assert inserted == 4

            stats = db.get_statistics()
            assert stats['导入批次数量'] >= 1


class TestQueryResult:
    """测试 QueryResult 类"""

    def test_to_dataframe(self, setup_logging, sample_records):
        records = [TransactionRecord.from_dict(r) for r in sample_records]
        result = QueryResult(records=records, total_count=4)
        df = result.to_dataframe()
        assert len(df) == 4
        assert '唯一id' in df.columns
        assert '银行' in df.columns
        assert '主体' in df.columns

    def test_to_excel(self, setup_logging, tmp_db_dir, sample_records):
        records = [TransactionRecord.from_dict(r) for r in sample_records]
        result = QueryResult(records=records, total_count=4)
        output_path = os.path.join(tmp_db_dir, 'query_result.xlsx')
        result.to_excel(output_path)
        assert os.path.exists(output_path)


class TestConvenienceFunctions:
    """测试便捷函数"""

    def test_persist_transactions(self, setup_logging, tmp_db_dir, sample_records):
        config = {
            'backend': 'sqlite',
            'sqlite': {'db_path': os.path.join(tmp_db_dir, 'test.db')},
            'auto_persist': True,
        }

        inserted, duplicates = persist_transactions(
            sample_records,
            batch_id='BATCH001',
            deduplicate=True,
            script_dir=tmp_db_dir,
            config=config,
        )
        assert inserted == 4
        assert duplicates == 0

        inserted2, duplicates2 = persist_transactions(
            sample_records,
            batch_id='BATCH002',
            deduplicate=True,
            script_dir=tmp_db_dir,
            config=config,
        )
        assert inserted2 == 0
        assert duplicates2 == 4

    def test_query_transactions(self, setup_logging, tmp_db_dir, sample_records):
        config = {
            'backend': 'sqlite',
            'sqlite': {'db_path': os.path.join(tmp_db_dir, 'test.db')},
            'auto_persist': True,
        }

        persist_transactions(
            sample_records,
            script_dir=tmp_db_dir,
            config=config,
        )

        result = query_transactions(
            subject='北京XX科技',
            script_dir=tmp_db_dir,
            config=config,
        )
        assert result.total_count == 2

    def test_get_db_statistics(self, setup_logging, tmp_db_dir, sample_records):
        config = {
            'backend': 'sqlite',
            'sqlite': {'db_path': os.path.join(tmp_db_dir, 'test.db')},
            'auto_persist': True,
        }

        persist_transactions(
            sample_records,
            script_dir=tmp_db_dir,
            config=config,
        )

        stats = get_db_statistics(script_dir=tmp_db_dir, config=config)
        assert stats['总记录数'] == 4


class TestDatabaseFactory:
    """测试数据库工厂"""

    def test_create_sqlite_backend(self, setup_logging, tmp_db_dir):
        config = {
            'backend': 'sqlite',
            'sqlite': {'db_path': os.path.join(tmp_db_dir, 'test.db')},
        }
        db = create_database_backend(config, script_dir=tmp_db_dir)
        assert isinstance(db, SQLiteBackend)

    def test_load_database_config_default(self, setup_logging, tmp_db_dir):
        config = load_database_config(script_dir=tmp_db_dir)
        assert config['backend'] == 'sqlite'
        assert 'auto_persist' in config

    def test_load_database_config_with_file(self, setup_logging, tmp_db_dir):
        import json
        config_data = {
            'backend': 'sqlite',
            'auto_persist': False,
        }
        config_path = os.path.join(tmp_db_dir, 'database_config.json')
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(config_data, f)

        config = load_database_config(script_dir=tmp_db_dir)
        assert config['backend'] == 'sqlite'
        assert config['auto_persist'] is False


class TestIntegrationWithBankcheck:
    """测试与 bankcheck 主流程的集成"""

    def test_processing_result_has_db_fields(self, setup_logging):
        from bankcheck import ProcessingResult
        result = ProcessingResult()
        assert hasattr(result, 'db_inserted_count')
        assert hasattr(result, 'db_duplicate_count')
        assert result.db_inserted_count == 0
        assert result.db_duplicate_count == 0

    def test_database_module_import(self, setup_logging):
        assert bankcheck.HAS_DATABASE is True
        assert bankcheck.db_module is not None
        assert hasattr(bankcheck.db_module, 'SQLiteBackend')
        assert hasattr(bankcheck.db_module, 'persist_transactions')
        assert hasattr(bankcheck.db_module, 'query_transactions')
        assert hasattr(bankcheck.db_module, 'get_db_statistics')
