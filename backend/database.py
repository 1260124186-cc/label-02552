# -*- coding: utf-8 -*-
"""
银行流水数据库持久化模块
功能：
  1. 将提取后的流水记录持久化到 SQLite 或 PostgreSQL
  2. 提供按主体、账号、时间范围的 SQL 查询能力
  3. 支持增量写入，自动去重
  4. 突破单次 Excel 总表的查询与分析局限
"""

import os
import sys
import logging
import json
import sqlite3
from dataclasses import dataclass, field, asdict
from datetime import datetime
from typing import List, Optional, Dict, Any, Union, Tuple
from abc import ABC, abstractmethod


def get_script_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def get_logger():
    return logging.getLogger('bankcheck')


def _normalize_account_str(value):
    if value is None:
        return ''
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        if value != value:
            return ''
        if value == int(value):
            return str(int(value))
        return str(value)
    s = str(value).strip()
    if '.' in s:
        try:
            f = float(s)
            if f == int(f):
                return str(int(f))
        except (ValueError, TypeError, OverflowError):
            pass
    return s


def _account_key(value):
    normalized = _normalize_account_str(value)
    return normalized.lstrip('0') or '0'


def to_float(value):
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).strip())
    except (ValueError, TypeError):
        return None


def _make_match_key(row):
    transaction_id = row.get('交易流水号')
    bank_account = _account_key(row.get('银行账号'))
    if transaction_id is not None and str(transaction_id).strip():
        tid = str(transaction_id).strip()
        return f"{bank_account}::{tid}"
    payment = row.get('付款')
    receipt = row.get('收款')
    trade_date = row.get('交易日期')
    p = '' if payment is None else str(payment)
    r = '' if receipt is None else str(receipt)
    d = '' if trade_date is None else str(trade_date)
    return f"{bank_account}::{d}::{p}::{r}"


# ──────────────────────────────────────────────
# 数据模型
# ──────────────────────────────────────────────

TRANSACTION_DB_FILENAME = 'transactions.db'

TRANSACTION_COLUMNS = [
    '唯一id', '银行', '银行账号', '主体', '交易日期',
    '付款', '收款', '摘要', '对方户名', '余额', '交易流水号',
    '导入批次号', '导入时间', '匹配键',
    '黑白名单标签', '命中规则名称', '命中关键词',
]


@dataclass
class TransactionRecord:
    """交易记录数据类，与 Excel 总表字段保持一致"""
    唯一id: str
    银行: str
    银行账号: Optional[str] = None
    主体: Optional[str] = None
    交易日期: Optional[Union[str, datetime]] = None
    付款: Optional[float] = None
    收款: Optional[float] = None
    摘要: Optional[str] = None
    对方户名: Optional[str] = None
    余额: Optional[float] = None
    交易流水号: Optional[str] = None
    导入批次号: Optional[str] = None
    导入时间: Optional[str] = None
    匹配键: Optional[str] = None
    黑白名单标签: Optional[str] = None
    命中规则名称: Optional[str] = None
    命中关键词: Optional[str] = None

    @classmethod
    def from_dict(cls, row_dict: Dict[str, Any]) -> 'TransactionRecord':
        """从字典创建记录，自动处理类型转换"""
        return cls(
            唯一id=str(row_dict.get('唯一id', '')),
            银行=str(row_dict.get('银行', '')),
            银行账号=row_dict.get('银行账号'),
            主体=row_dict.get('主体'),
            交易日期=row_dict.get('交易日期'),
            付款=to_float(row_dict.get('付款')) if row_dict.get('付款') is not None else None,
            收款=to_float(row_dict.get('收款')) if row_dict.get('收款') is not None else None,
            摘要=str(row_dict.get('摘要')) if row_dict.get('摘要') is not None else None,
            对方户名=str(row_dict.get('对方户名')) if row_dict.get('对方户名') is not None else None,
            余额=to_float(row_dict.get('余额')) if row_dict.get('余额') is not None else None,
            交易流水号=str(row_dict.get('交易流水号')) if row_dict.get('交易流水号') is not None else None,
            导入批次号=str(row_dict.get('导入批次号')) if row_dict.get('导入批次号') is not None else None,
            导入时间=str(row_dict.get('导入时间')) if row_dict.get('导入时间') is not None else None,
            匹配键=str(row_dict.get('匹配键')) if row_dict.get('匹配键') is not None else None,
            黑白名单标签=str(row_dict.get('黑白名单标签')) if row_dict.get('黑白名单标签') is not None else None,
            命中规则名称=str(row_dict.get('命中规则名称')) if row_dict.get('命中规则名称') is not None else None,
            命中关键词=str(row_dict.get('命中关键词')) if row_dict.get('命中关键词') is not None else None,
        )

    def to_dict(self) -> Dict[str, Any]:
        """转换为字典"""
        return asdict(self)

    def compute_match_key(self) -> str:
        """计算匹配键，用于去重检测"""
        if self.匹配键:
            return self.匹配键
        row = {
            '银行账号': self.银行账号,
            '交易流水号': self.交易流水号,
            '交易日期': self.交易日期,
            '付款': self.付款,
            '收款': self.收款,
        }
        return _make_match_key(row)


@dataclass
class QueryResult:
    """查询结果封装"""
    records: List[TransactionRecord] = field(default_factory=list)
    total_count: int = 0
    summary: Dict[str, Any] = field(default_factory=dict)

    def to_dataframe(self):
        """转换为 pandas DataFrame"""
        import pandas as pd
        data = [r.to_dict() for r in self.records]
        return pd.DataFrame(data, columns=TRANSACTION_COLUMNS)

    def to_excel(self, output_path: str) -> str:
        """导出到 Excel"""
        df = self.to_dataframe()
        df.to_excel(output_path, index=False, engine='openpyxl')
        return output_path


# ──────────────────────────────────────────────
# 抽象数据库接口
# ──────────────────────────────────────────────

class DatabaseBackend(ABC):
    """数据库后端抽象基类"""

    @abstractmethod
    def connect(self):
        """建立数据库连接"""
        pass

    @abstractmethod
    def disconnect(self):
        """关闭数据库连接"""
        pass

    @abstractmethod
    def init_schema(self):
        """初始化数据库表结构"""
        pass

    @abstractmethod
    def insert_records(self, records: List[TransactionRecord],
                       batch_id: Optional[str] = None,
                       deduplicate: bool = True) -> Tuple[int, int]:
        """
        批量插入记录
        Returns: (实际插入数量, 重复数量)
        """
        pass

    @abstractmethod
    def query_records(self,
                      subject: Optional[str] = None,
                      account: Optional[str] = None,
                      bank: Optional[str] = None,
                      start_date: Optional[str] = None,
                      end_date: Optional[str] = None,
                      min_amount: Optional[float] = None,
                      max_amount: Optional[float] = None,
                      counterpart: Optional[str] = None,
                      summary_keyword: Optional[str] = None,
                      limit: Optional[int] = None,
                      offset: int = 0,
                      order_by: str = '交易日期',
                      ascending: bool = False) -> QueryResult:
        """
        多条件组合查询
        支持按主体、账号、时间范围、金额范围、对方户名、摘要关键词等查询
        """
        pass

    @abstractmethod
    def get_existing_match_keys(self) -> set:
        """获取所有已存在记录的匹配键集合，用于增量去重"""
        pass

    @abstractmethod
    def get_statistics(self) -> Dict[str, Any]:
        """获取统计信息：总记录数、按银行/主体/账号分组统计"""
        pass

    @abstractmethod
    def delete_records(self,
                       subject: Optional[str] = None,
                       account: Optional[str] = None,
                       start_date: Optional[str] = None,
                       end_date: Optional[str] = None) -> int:
        """删除指定条件的记录，返回删除数量"""
        pass

    def __enter__(self):
        self.connect()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.disconnect()
        return False


# ──────────────────────────────────────────────
# SQLite 实现
# ──────────────────────────────────────────────

class SQLiteBackend(DatabaseBackend):
    """SQLite 数据库后端实现"""

    def __init__(self, db_path: Optional[str] = None):
        if db_path is None:
            db_path = os.path.join(get_script_dir(), TRANSACTION_DB_FILENAME)
        self.db_path = db_path
        self.conn: Optional[sqlite3.Connection] = None
        self.logger = get_logger()

    def connect(self):
        if self.conn is None:
            self.conn = sqlite3.connect(self.db_path)
            self.conn.row_factory = sqlite3.Row
            self.conn.execute("PRAGMA journal_mode = WAL")
            self.conn.execute("PRAGMA synchronous = NORMAL")
            self.conn.execute("PRAGMA foreign_keys = ON")
            self.init_schema()
            self.logger.info('SQLite 数据库已连接: %s', self.db_path)

    def disconnect(self):
        if self.conn is not None:
            self.conn.close()
            self.conn = None
            self.logger.info('SQLite 数据库已断开')

    def init_schema(self):
        if self.conn is None:
            self.connect()

        cursor = self.conn.cursor()

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS transactions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                唯一id TEXT NOT NULL UNIQUE,
                银行 TEXT NOT NULL,
                银行账号 TEXT,
                主体 TEXT,
                交易日期 TEXT,
                付款 REAL,
                收款 REAL,
                摘要 TEXT,
                对方户名 TEXT,
                余额 REAL,
                交易流水号 TEXT,
                导入批次号 TEXT,
                导入时间 TEXT,
                匹配键 TEXT UNIQUE,
                黑白名单标签 TEXT,
                命中规则名称 TEXT,
                命中关键词 TEXT,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
        ''')

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS import_batches (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                batch_id TEXT NOT NULL UNIQUE,
                source_directory TEXT,
                total_records INTEGER DEFAULT 0,
                inserted_records INTEGER DEFAULT 0,
                duplicate_records INTEGER DEFAULT 0,
                started_at TEXT NOT NULL,
                completed_at TEXT,
                status TEXT NOT NULL
            )
        ''')

        cursor.execute('CREATE INDEX IF NOT EXISTS idx_transactions_subject ON transactions(主体)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_transactions_account ON transactions(银行账号)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_transactions_bank ON transactions(银行)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_transactions_date ON transactions(交易日期)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_transactions_counterpart ON transactions(对方户名)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_transactions_batch ON transactions(导入批次号)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_transactions_match_key ON transactions(匹配键)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_transactions_counterparty_tag ON transactions(黑白名单标签)')

        cursor.execute('CREATE INDEX IF NOT EXISTS idx_batches_date ON import_batches(started_at)')

        self.conn.commit()
        self.logger.debug('数据库表结构初始化完成')

    def _record_to_tuple(self, record: TransactionRecord) -> tuple:
        """将记录转换为插入元组"""
        match_key = record.compute_match_key() if not record.匹配键 else record.匹配键
        import_time = record.导入时间 or datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')
        trade_date = record.交易日期
        if isinstance(trade_date, datetime):
            trade_date = trade_date.strftime('%Y-%m-%d %H:%M:%S')
        elif trade_date is not None:
            trade_date = str(trade_date)

        return (
            record.唯一id,
            record.银行,
            record.银行账号,
            record.主体,
            trade_date,
            record.付款,
            record.收款,
            record.摘要,
            record.对方户名,
            record.余额,
            record.交易流水号,
            record.导入批次号,
            import_time,
            match_key,
            record.黑白名单标签,
            record.命中规则名称,
            record.命中关键词,
        )

    def insert_records(self, records: List[TransactionRecord],
                       batch_id: Optional[str] = None,
                       deduplicate: bool = True) -> Tuple[int, int]:
        if self.conn is None:
            self.connect()

        if not records:
            self.logger.warning('无记录可插入')
            return 0, 0

        if batch_id is None:
            batch_id = f"BATCH{datetime.now().strftime('%Y%m%d%H%M%S%f')}"

        cursor = self.conn.cursor()
        started_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')

        cursor.execute('''
            INSERT INTO import_batches (
                batch_id, total_records, started_at, status
            ) VALUES (?, ?, ?, ?)
        ''', (batch_id, len(records), started_at, 'processing'))

        existing_keys = set()
        if deduplicate:
            existing_keys = self.get_existing_match_keys()
            self.logger.info('已加载 %d 个历史匹配键用于去重', len(existing_keys))

        insert_count = 0
        duplicate_count = 0
        insert_tuples = []

        for record in records:
            match_key = record.compute_match_key()
            if deduplicate and match_key in existing_keys:
                duplicate_count += 1
                continue
            record.匹配键 = match_key
            record.导入批次号 = batch_id
            insert_tuples.append(self._record_to_tuple(record))
            existing_keys.add(match_key)
            insert_count += 1

        if insert_tuples:
            cursor.executemany('''
                INSERT OR IGNORE INTO transactions (
                    唯一id, 银行, 银行账号, 主体, 交易日期,
                    付款, 收款, 摘要, 对方户名, 余额,
                    交易流水号, 导入批次号, 导入时间, 匹配键,
                    黑白名单标签, 命中规则名称, 命中关键词
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', insert_tuples)

        completed_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')
        cursor.execute('''
            UPDATE import_batches SET
                inserted_records = ?,
                duplicate_records = ?,
                completed_at = ?,
                status = ?
            WHERE batch_id = ?
        ''', (insert_count, duplicate_count, completed_at, 'completed', batch_id))

        self.conn.commit()

        self.logger.info(
            '批量导入完成: 批次 %s, 共 %d 条, 新增 %d 条, 重复 %d 条',
            batch_id, len(records), insert_count, duplicate_count
        )

        return insert_count, duplicate_count

    def get_existing_match_keys(self) -> set:
        if self.conn is None:
            self.connect()

        cursor = self.conn.cursor()
        cursor.execute('SELECT 匹配键 FROM transactions WHERE 匹配键 IS NOT NULL')
        rows = cursor.fetchall()
        return {row['匹配键'] for row in rows if row['匹配键']}

    def query_records(self,
                      subject: Optional[str] = None,
                      account: Optional[str] = None,
                      bank: Optional[str] = None,
                      start_date: Optional[str] = None,
                      end_date: Optional[str] = None,
                      min_amount: Optional[float] = None,
                      max_amount: Optional[float] = None,
                      counterpart: Optional[str] = None,
                      summary_keyword: Optional[str] = None,
                      limit: Optional[int] = None,
                      offset: int = 0,
                      order_by: str = '交易日期',
                      ascending: bool = False) -> QueryResult:
        if self.conn is None:
            self.connect()

        cursor = self.conn.cursor()

        conditions = []
        params = []

        if subject:
            conditions.append('主体 LIKE ?')
            params.append(f'%{subject}%')
        if account:
            account_key = _account_key(account)
            conditions.append('(银行账号 LIKE ? OR 银行账号 LIKE ?)')
            params.append(f'%{account}%')
            params.append(f'%{account_key}%')
        if bank:
            conditions.append('银行 = ?')
            params.append(bank)
        if start_date:
            conditions.append('date(交易日期) >= date(?)')
            params.append(start_date)
        if end_date:
            conditions.append('date(交易日期) <= date(?)')
            params.append(end_date)
        if min_amount is not None:
            conditions.append('(ABS(付款) >= ? OR 收款 >= ?)')
            params.append(min_amount)
            params.append(min_amount)
        if max_amount is not None:
            conditions.append('(ABS(付款) <= ? OR 收款 <= ?)')
            params.append(abs(max_amount))
            params.append(max_amount)
        if counterpart:
            conditions.append('对方户名 LIKE ?')
            params.append(f'%{counterpart}%')
        if summary_keyword:
            conditions.append('摘要 LIKE ?')
            params.append(f'%{summary_keyword}%')

        where_clause = ' WHERE ' + ' AND '.join(conditions) if conditions else ''

        count_query = f'SELECT COUNT(*) as cnt FROM transactions{where_clause}'
        cursor.execute(count_query, params)
        total_count = cursor.fetchone()['cnt']

        valid_order_columns = {'交易日期', '付款', '收款', '余额', '银行', '主体', '银行账号', '导入时间'}
        order_col = order_by if order_by in valid_order_columns else '交易日期'
        order_dir = 'ASC' if ascending else 'DESC'

        query = f'''
            SELECT * FROM transactions{where_clause}
            ORDER BY {order_col} {order_dir}
        '''

        if limit is not None:
            query += ' LIMIT ? OFFSET ?'
            params.extend([limit, offset])

        cursor.execute(query, params)
        rows = cursor.fetchall()

        records = []
        for row in rows:
            record_dict = {
                '唯一id': row['唯一id'],
                '银行': row['银行'],
                '银行账号': row['银行账号'],
                '主体': row['主体'],
                '交易日期': row['交易日期'],
                '付款': row['付款'],
                '收款': row['收款'],
                '摘要': row['摘要'],
                '对方户名': row['对方户名'],
                '余额': row['余额'],
                '交易流水号': row['交易流水号'],
                '导入批次号': row['导入批次号'],
                '导入时间': row['导入时间'],
                '匹配键': row['匹配键'],
                '黑白名单标签': row['黑白名单标签'],
                '命中规则名称': row['命中规则名称'],
                '命中关键词': row['命中关键词'],
            }
            records.append(TransactionRecord.from_dict(record_dict))

        summary = self._compute_query_summary(records)

        return QueryResult(
            records=records,
            total_count=total_count,
            summary=summary,
        )

    def _compute_query_summary(self, records: List[TransactionRecord]) -> Dict[str, Any]:
        """计算查询结果的汇总统计"""
        if not records:
            return {}

        total_payment = sum(r.付款 or 0 for r in records if r.付款 is not None and r.付款 < 0)
        total_receipt = sum(r.收款 or 0 for r in records if r.收款 is not None and r.收款 > 0)
        net_amount = total_payment + total_receipt

        payment_count = sum(1 for r in records if r.付款 is not None and r.付款 < 0)
        receipt_count = sum(1 for r in records if r.收款 is not None and r.收款 > 0)

        banks = {}
        subjects = {}
        accounts = {}

        for r in records:
            bank = r.银行 or '未知'
            subject = r.主体 or '未知'
            account = r.银行账号 or '未知'

            banks[bank] = banks.get(bank, 0) + 1
            subjects[subject] = subjects.get(subject, 0) + 1
            accounts[account] = accounts.get(account, 0) + 1

        return {
            '记录总数': len(records),
            '付款笔数': payment_count,
            '收款笔数': receipt_count,
            '付款总额': round(total_payment, 2),
            '收款总额': round(total_receipt, 2),
            '净额': round(net_amount, 2),
            '银行分布': banks,
            '主体分布': subjects,
            '账号分布': accounts,
        }

    def get_statistics(self) -> Dict[str, Any]:
        if self.conn is None:
            self.connect()

        cursor = self.conn.cursor()

        cursor.execute('SELECT COUNT(*) as cnt FROM transactions')
        total_count = cursor.fetchone()['cnt']

        cursor.execute('''
            SELECT 银行, COUNT(*) as cnt,
                   SUM(CASE WHEN 付款 < 0 THEN 付款 ELSE 0 END) as total_payment,
                   SUM(CASE WHEN 收款 > 0 THEN 收款 ELSE 0 END) as total_receipt
            FROM transactions
            GROUP BY 银行
            ORDER BY cnt DESC
        ''')
        by_bank = [dict(row) for row in cursor.fetchall()]

        cursor.execute('''
            SELECT 主体, COUNT(*) as cnt,
                   SUM(CASE WHEN 付款 < 0 THEN 付款 ELSE 0 END) as total_payment,
                   SUM(CASE WHEN 收款 > 0 THEN 收款 ELSE 0 END) as total_receipt
            FROM transactions
            GROUP BY 主体
            ORDER BY cnt DESC
        ''')
        by_subject = [dict(row) for row in cursor.fetchall()]

        cursor.execute('''
            SELECT 银行账号, 主体, 银行, COUNT(*) as cnt
            FROM transactions
            GROUP BY 银行账号, 主体, 银行
            ORDER BY cnt DESC
        ''')
        by_account = [dict(row) for row in cursor.fetchall()]

        cursor.execute('''
            SELECT date(交易日期) as dt, COUNT(*) as cnt,
                   SUM(CASE WHEN 付款 < 0 THEN 付款 ELSE 0 END) as total_payment,
                   SUM(CASE WHEN 收款 > 0 THEN 收款 ELSE 0 END) as total_receipt
            FROM transactions
            WHERE 交易日期 IS NOT NULL
            GROUP BY date(交易日期)
            ORDER BY dt DESC
            LIMIT 30
        ''')
        by_date = [dict(row) for row in cursor.fetchall()]

        cursor.execute('SELECT MIN(交易日期) as min_d, MAX(交易日期) as max_d FROM transactions')
        date_range = cursor.fetchone()

        cursor.execute('''
            SELECT COUNT(*) as cnt FROM import_batches
        ''')
        batch_count = cursor.fetchone()['cnt']

        return {
            '总记录数': total_count,
            '导入批次数量': batch_count,
            '日期范围': {
                '最早交易日期': date_range['min_d'],
                '最晚交易日期': date_range['max_d'],
            },
            '按银行统计': by_bank,
            '按主体统计': by_subject,
            '按账号统计': by_account,
            '近30天交易趋势': by_date,
        }

    def delete_records(self,
                       subject: Optional[str] = None,
                       account: Optional[str] = None,
                       start_date: Optional[str] = None,
                       end_date: Optional[str] = None) -> int:
        if self.conn is None:
            self.connect()

        cursor = self.conn.cursor()

        conditions = []
        params = []

        if subject:
            conditions.append('主体 = ?')
            params.append(subject)
        if account:
            conditions.append('银行账号 = ?')
            params.append(account)
        if start_date:
            conditions.append('date(交易日期) >= date(?)')
            params.append(start_date)
        if end_date:
            conditions.append('date(交易日期) <= date(?)')
            params.append(end_date)

        if not conditions:
            self.logger.warning('删除操作未指定条件，将删除所有记录')
            confirm = input('确定要删除所有记录吗？(yes/N): ')
            if confirm.lower() != 'yes':
                self.logger.info('已取消删除操作')
                return 0

        where_clause = ' WHERE ' + ' AND '.join(conditions) if conditions else ''

        cursor.execute(f'SELECT COUNT(*) as cnt FROM transactions{where_clause}', params)
        count = cursor.fetchone()['cnt']

        cursor.execute(f'DELETE FROM transactions{where_clause}', params)
        self.conn.commit()

        self.logger.info('已删除 %d 条记录', count)
        return count


# ──────────────────────────────────────────────
# PostgreSQL 实现（可选）
# ──────────────────────────────────────────────

class PostgreSQLBackend(DatabaseBackend):
    """PostgreSQL 数据库后端实现（可选）"""

    def __init__(self, host: str = 'localhost', port: int = 5432,
                 database: str = 'bank_transactions',
                 user: Optional[str] = None, password: Optional[str] = None):
        try:
            import psycopg2
            self.psycopg2 = psycopg2
        except ImportError:
            raise ImportError(
                '使用 PostgreSQL 后端需要安装 psycopg2 库。'
                '请运行: pip install psycopg2-binary'
            )

        self.host = host
        self.port = port
        self.database = database
        self.user = user or os.environ.get('PGUSER')
        self.password = password or os.environ.get('PGPASSWORD')
        self.conn = None
        self.logger = get_logger()

    def connect(self):
        if self.conn is None:
            self.conn = self.psycopg2.connect(
                host=self.host,
                port=self.port,
                database=self.database,
                user=self.user,
                password=self.password,
            )
            self.init_schema()
            self.logger.info('PostgreSQL 数据库已连接: %s:%s/%s',
                             self.host, self.port, self.database)

    def disconnect(self):
        if self.conn is not None:
            self.conn.close()
            self.conn = None
            self.logger.info('PostgreSQL 数据库已断开')

    def init_schema(self):
        if self.conn is None:
            self.connect()

        cursor = self.conn.cursor()

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS transactions (
                id SERIAL PRIMARY KEY,
                唯一id VARCHAR(255) NOT NULL UNIQUE,
                银行 VARCHAR(100) NOT NULL,
                银行账号 VARCHAR(255),
                主体 VARCHAR(255),
                交易日期 TIMESTAMP,
                付款 NUMERIC(20, 2),
                收款 NUMERIC(20, 2),
                摘要 TEXT,
                对方户名 VARCHAR(255),
                余额 NUMERIC(20, 2),
                交易流水号 VARCHAR(255),
                导入批次号 VARCHAR(255),
                导入时间 TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                匹配键 VARCHAR(255) UNIQUE,
                黑白名单标签 VARCHAR(255),
                命中规则名称 VARCHAR(255),
                命中关键词 TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS import_batches (
                id SERIAL PRIMARY KEY,
                batch_id VARCHAR(255) NOT NULL UNIQUE,
                source_directory TEXT,
                total_records INTEGER DEFAULT 0,
                inserted_records INTEGER DEFAULT 0,
                duplicate_records INTEGER DEFAULT 0,
                started_at TIMESTAMP NOT NULL,
                completed_at TIMESTAMP,
                status VARCHAR(50) NOT NULL
            )
        ''')

        cursor.execute('CREATE INDEX IF NOT EXISTS idx_transactions_subject ON transactions(主体)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_transactions_account ON transactions(银行账号)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_transactions_bank ON transactions(银行)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_transactions_date ON transactions(交易日期)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_transactions_counterpart ON transactions(对方户名)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_transactions_batch ON transactions(导入批次号)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_transactions_counterparty_tag ON transactions(黑白名单标签)')

        self.conn.commit()
        self.logger.debug('PostgreSQL 表结构初始化完成')

    def insert_records(self, records: List[TransactionRecord],
                       batch_id: Optional[str] = None,
                       deduplicate: bool = True) -> Tuple[int, int]:
        if self.conn is None:
            self.connect()

        if not records:
            return 0, 0

        if batch_id is None:
            batch_id = f"BATCH{datetime.now().strftime('%Y%m%d%H%M%S%f')}"

        cursor = self.conn.cursor()
        started_at = datetime.now()

        cursor.execute('''
            INSERT INTO import_batches (batch_id, total_records, started_at, status)
            VALUES (%s, %s, %s, %s)
        ''', (batch_id, len(records), started_at, 'processing'))

        existing_keys = set()
        if deduplicate:
            existing_keys = self.get_existing_match_keys()

        insert_count = 0
        duplicate_count = 0

        for record in records:
            match_key = record.compute_match_key()
            if deduplicate and match_key in existing_keys:
                duplicate_count += 1
                continue
            record.匹配键 = match_key
            record.导入批次号 = batch_id

            trade_date = record.交易日期
            if isinstance(trade_date, str):
                try:
                    trade_date = datetime.fromisoformat(trade_date)
                except ValueError:
                    pass

            try:
                cursor.execute('''
                    INSERT INTO transactions (
                        唯一id, 银行, 银行账号, 主体, 交易日期,
                        付款, 收款, 摘要, 对方户名, 余额,
                        交易流水号, 导入批次号, 导入时间, 匹配键,
                        黑白名单标签, 命中规则名称, 命中关键词
                    ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                    ON CONFLICT (匹配键) DO NOTHING
                    RETURNING id
                ''', (
                    record.唯一id, record.银行, record.银行账号, record.主体, trade_date,
                    record.付款, record.收款, record.摘要, record.对方户名, record.余额,
                    record.交易流水号, record.导入批次号, datetime.now(), record.匹配键,
                    record.黑白名单标签, record.命中规则名称, record.命中关键词,
                ))
                if cursor.fetchone():
                    insert_count += 1
                    existing_keys.add(match_key)
                else:
                    duplicate_count += 1
            except self.psycopg2.IntegrityError:
                duplicate_count += 1
                self.conn.rollback()

        cursor.execute('''
            UPDATE import_batches SET
                inserted_records = %s,
                duplicate_records = %s,
                completed_at = %s,
                status = %s
            WHERE batch_id = %s
        ''', (insert_count, duplicate_count, datetime.now(), 'completed', batch_id))

        self.conn.commit()
        return insert_count, duplicate_count

    def get_existing_match_keys(self) -> set:
        if self.conn is None:
            self.connect()

        cursor = self.conn.cursor()
        cursor.execute('SELECT 匹配键 FROM transactions WHERE 匹配键 IS NOT NULL')
        rows = cursor.fetchall()
        return {row[0] for row in rows if row[0]}

    def query_records(self,
                      subject: Optional[str] = None,
                      account: Optional[str] = None,
                      bank: Optional[str] = None,
                      start_date: Optional[str] = None,
                      end_date: Optional[str] = None,
                      min_amount: Optional[float] = None,
                      max_amount: Optional[float] = None,
                      counterpart: Optional[str] = None,
                      summary_keyword: Optional[str] = None,
                      limit: Optional[int] = None,
                      offset: int = 0,
                      order_by: str = '交易日期',
                      ascending: bool = False) -> QueryResult:
        if self.conn is None:
            self.connect()

        cursor = self.conn.cursor()

        conditions = []
        params = []

        if subject:
            conditions.append('主体 ILIKE %s')
            params.append(f'%{subject}%')
        if account:
            conditions.append('银行账号 ILIKE %s')
            params.append(f'%{account}%')
        if bank:
            conditions.append('银行 = %s')
            params.append(bank)
        if start_date:
            conditions.append('交易日期 >= %s::date')
            params.append(start_date)
        if end_date:
            conditions.append('交易日期 <= %s::date')
            params.append(end_date)
        if min_amount is not None:
            conditions.append('(ABS(付款) >= %s OR 收款 >= %s)')
            params.extend([min_amount, min_amount])
        if max_amount is not None:
            conditions.append('(ABS(付款) <= %s OR 收款 <= %s)')
            params.extend([abs(max_amount), max_amount])
        if counterpart:
            conditions.append('对方户名 ILIKE %s')
            params.append(f'%{counterpart}%')
        if summary_keyword:
            conditions.append('摘要 ILIKE %s')
            params.append(f'%{summary_keyword}%')

        where_clause = ' WHERE ' + ' AND '.join(conditions) if conditions else ''

        cursor.execute(f'SELECT COUNT(*) FROM transactions{where_clause}', params)
        total_count = cursor.fetchone()[0]

        valid_order_columns = {'交易日期', '付款', '收款', '余额', '银行', '主体', '银行账号', '导入时间'}
        order_col = order_by if order_by in valid_order_columns else '交易日期'
        order_dir = 'ASC' if ascending else 'DESC'

        query = f'SELECT * FROM transactions{where_clause} ORDER BY {order_col} {order_dir}'

        if limit is not None:
            query += ' LIMIT %s OFFSET %s'
            params.extend([limit, offset])

        cursor.execute(query, params)
        rows = cursor.fetchall()

        colnames = [desc[0] for desc in cursor.description]

        records = []
        for row in rows:
            row_dict = dict(zip(colnames, row))
            records.append(TransactionRecord.from_dict(row_dict))

        summary = SQLiteBackend._compute_query_summary(None, records)

        return QueryResult(
            records=records,
            total_count=total_count,
            summary=summary,
        )

    def get_statistics(self) -> Dict[str, Any]:
        if self.conn is None:
            self.connect()

        cursor = self.conn.cursor()

        cursor.execute('SELECT COUNT(*) FROM transactions')
        total_count = cursor.fetchone()[0]

        return {
            '总记录数': total_count,
            '数据库类型': 'PostgreSQL',
            '主机': self.host,
            '数据库': self.database,
        }

    def delete_records(self,
                       subject: Optional[str] = None,
                       account: Optional[str] = None,
                       start_date: Optional[str] = None,
                       end_date: Optional[str] = None) -> int:
        if self.conn is None:
            self.connect()

        cursor = self.conn.cursor()

        conditions = []
        params = []

        if subject:
            conditions.append('主体 = %s')
            params.append(subject)
        if account:
            conditions.append('银行账号 = %s')
            params.append(account)
        if start_date:
            conditions.append('交易日期 >= %s::date')
            params.append(start_date)
        if end_date:
            conditions.append('交易日期 <= %s::date')
            params.append(end_date)

        where_clause = ' WHERE ' + ' AND '.join(conditions) if conditions else ''

        cursor.execute(f'SELECT COUNT(*) FROM transactions{where_clause}', params)
        count = cursor.fetchone()[0]

        cursor.execute(f'DELETE FROM transactions{where_clause}', params)
        self.conn.commit()

        return count


# ──────────────────────────────────────────────
# 数据库工厂
# ──────────────────────────────────────────────

DATABASE_CONFIG_FILE = 'database_config.json'


def load_database_config(script_dir: Optional[str] = None) -> Dict[str, Any]:
    """加载数据库配置"""
    if script_dir is None:
        script_dir = get_script_dir()

    config_path = os.path.join(script_dir, DATABASE_CONFIG_FILE)

    default_config = {
        'backend': 'sqlite',
        'sqlite': {
            'db_path': os.path.join(script_dir, TRANSACTION_DB_FILENAME),
        },
        'postgresql': {
            'host': 'localhost',
            'port': 5432,
            'database': 'bank_transactions',
            'user': None,
            'password': None,
        },
        'auto_persist': True,
    }

    if os.path.exists(config_path):
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                user_config = json.load(f)
                default_config.update(user_config)
        except Exception as e:
            logger = get_logger()
            logger.warning('读取数据库配置失败，使用默认配置: %s', e)

    return default_config


def save_database_config(config: Dict[str, Any],
                         script_dir: Optional[str] = None) -> str:
    """保存数据库配置"""
    if script_dir is None:
        script_dir = get_script_dir()

    config_path = os.path.join(script_dir, DATABASE_CONFIG_FILE)

    with open(config_path, 'w', encoding='utf-8') as f:
        json.dump(config, f, ensure_ascii=False, indent=2)

    return config_path


def create_database_backend(config: Optional[Dict[str, Any]] = None,
                            script_dir: Optional[str] = None) -> DatabaseBackend:
    """
    根据配置创建数据库后端实例
    支持通过环境变量 BANKCHECK_DB_BACKEND 覆盖配置
    """
    if config is None:
        config = load_database_config(script_dir)

    backend = os.environ.get('BANKCHECK_DB_BACKEND', config.get('backend', 'sqlite')).lower()

    if backend == 'sqlite':
        sqlite_config = config.get('sqlite', {})
        db_path = sqlite_config.get('db_path')
        return SQLiteBackend(db_path=db_path)
    elif backend in ('postgresql', 'postgres', 'pg'):
        pg_config = config.get('postgresql', {})
        return PostgreSQLBackend(
            host=pg_config.get('host', 'localhost'),
            port=pg_config.get('port', 5432),
            database=pg_config.get('database', 'bank_transactions'),
            user=pg_config.get('user'),
            password=pg_config.get('password'),
        )
    else:
        raise ValueError(f'不支持的数据库后端: {backend}')


# ──────────────────────────────────────────────
# 便捷函数
# ──────────────────────────────────────────────

def persist_transactions(records: List[Dict[str, Any]],
                         batch_id: Optional[str] = None,
                         deduplicate: bool = True,
                         script_dir: Optional[str] = None,
                         config: Optional[Dict[str, Any]] = None) -> Tuple[int, int]:
    """
    将交易记录持久化到数据库的便捷函数

    Args:
        records: 交易记录字典列表（与 Excel 总表格式一致）
        batch_id: 导入批次号
        deduplicate: 是否启用去重
        script_dir: 脚本目录
        config: 数据库配置

    Returns:
        (实际插入数量, 重复数量)
    """
    logger = get_logger()

    if config is None:
        config = load_database_config(script_dir)

    if not config.get('auto_persist', True):
        logger.info('数据库自动持久化已禁用，跳过写入')
        return 0, 0

    transaction_records = [TransactionRecord.from_dict(r) for r in records]

    try:
        with create_database_backend(config, script_dir) as db:
            insert_count, duplicate_count = db.insert_records(
                transaction_records,
                batch_id=batch_id,
                deduplicate=deduplicate,
            )
            return insert_count, duplicate_count
    except ImportError as e:
        logger.warning('数据库后端导入失败，跳过持久化: %s', e)
        return 0, 0
    except Exception as e:
        logger.error('数据库持久化失败: %s', e, exc_info=True)
        return 0, 0


def query_transactions(subject: Optional[str] = None,
                       account: Optional[str] = None,
                       bank: Optional[str] = None,
                       start_date: Optional[str] = None,
                       end_date: Optional[str] = None,
                       min_amount: Optional[float] = None,
                       max_amount: Optional[float] = None,
                       counterpart: Optional[str] = None,
                       summary_keyword: Optional[str] = None,
                       limit: Optional[int] = None,
                       offset: int = 0,
                       script_dir: Optional[str] = None,
                       config: Optional[Dict[str, Any]] = None) -> QueryResult:
    """
    查询交易记录的便捷函数
    """
    if config is None:
        config = load_database_config(script_dir)

    with create_database_backend(config, script_dir) as db:
        return db.query_records(
            subject=subject,
            account=account,
            bank=bank,
            start_date=start_date,
            end_date=end_date,
            min_amount=min_amount,
            max_amount=max_amount,
            counterpart=counterpart,
            summary_keyword=summary_keyword,
            limit=limit,
            offset=offset,
        )


def get_db_statistics(script_dir: Optional[str] = None,
                      config: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """
    获取数据库统计信息的便捷函数
    """
    if config is None:
        config = load_database_config(script_dir)

    with create_database_backend(config, script_dir) as db:
        return db.get_statistics()
