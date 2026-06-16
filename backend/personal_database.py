# -*- coding: utf-8 -*-
"""
个人理财专版 - 数据库模块
独立于对公版数据库，仅用于个人理财数据存储
支持 SQLite 存储，提供增删改查与统计分析能力
"""

import os
import sys
import sqlite3
import json
import logging
from dataclasses import dataclass, field, asdict
from datetime import datetime
from typing import List, Optional, Dict, Any, Tuple
from contextlib import contextmanager


PERSONAL_DB_FILENAME = 'personal_finance.db'


def get_logger():
    return logging.getLogger('bankcheck')


def get_script_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def get_personal_db_path() -> str:
    """获取个人理财数据库路径"""
    return os.path.join(get_script_dir(), PERSONAL_DB_FILENAME)


@contextmanager
def get_db_connection(db_path: str = None):
    """获取数据库连接上下文管理器"""
    if db_path is None:
        db_path = get_personal_db_path()

    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    conn.execute('PRAGMA journal_mode=WAL')
    conn.execute('PRAGMA foreign_keys=ON')

    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def init_personal_db(db_path: str = None) -> bool:
    """
    初始化个人理财数据库
    创建表结构，返回是否成功
    """
    logger = get_logger()

    if db_path is None:
        db_path = get_personal_db_path()

    try:
        with get_db_connection(db_path) as conn:
            cursor = conn.cursor()

            cursor.execute('''
                CREATE TABLE IF NOT EXISTS transactions (
                    id TEXT PRIMARY KEY,
                    bank_name TEXT,
                    card_type TEXT,
                    account_number TEXT,
                    trade_date TEXT,
                    transaction_time TEXT,
                    post_date TEXT,
                    amount REAL,
                    direction TEXT,
                    balance REAL,
                    summary TEXT,
                    merchant TEXT,
                    transaction_type TEXT,
                    transaction_id TEXT,
                    category TEXT,
                    category_icon TEXT,
                    category_color TEXT,
                    installments INTEGER DEFAULT 0,
                    status TEXT,
                    source_file TEXT,
                    created_at TEXT,
                    updated_at TEXT
                )
            ''')

            cursor.execute('''
                CREATE TABLE IF NOT EXISTS accounts (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    bank_name TEXT,
                    card_type TEXT,
                    account_number TEXT,
                    account_name TEXT,
                    credit_limit REAL DEFAULT 0,
                    current_balance REAL DEFAULT 0,
                    created_at TEXT,
                    updated_at TEXT,
                    UNIQUE(bank_name, account_number)
                )
            ''')

            cursor.execute('''
                CREATE TABLE IF NOT EXISTS categories (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT UNIQUE,
                    icon TEXT,
                    color TEXT,
                    direction TEXT,
                    type TEXT DEFAULT 'system',
                    sort_order INTEGER DEFAULT 0
                )
            ''')

            cursor.execute('''
                CREATE TABLE IF NOT EXISTS import_batches (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    batch_name TEXT,
                    source_path TEXT,
                    total_files INTEGER DEFAULT 0,
                    parsed_files INTEGER DEFAULT 0,
                    failed_files INTEGER DEFAULT 0,
                    total_records INTEGER DEFAULT 0,
                    status TEXT DEFAULT 'pending',
                    created_at TEXT,
                    finished_at TEXT
                )
            ''')

            cursor.execute('''
                CREATE INDEX IF NOT EXISTS idx_transactions_trade_date
                ON transactions(trade_date)
            ''')

            cursor.execute('''
                CREATE INDEX IF NOT EXISTS idx_transactions_category
                ON transactions(category)
            ''')

            cursor.execute('''
                CREATE INDEX IF NOT EXISTS idx_transactions_direction
                ON transactions(direction)
            ''')

            cursor.execute('''
                CREATE INDEX IF NOT EXISTS idx_transactions_merchant
                ON transactions(merchant)
            ''')

            cursor.execute('''
                CREATE INDEX IF NOT EXISTS idx_transactions_bank
                ON transactions(bank_name)
            ''')

            logger.info('个人理财数据库初始化成功: %s', db_path)
            return True

    except Exception as e:
        logger.error('个人理财数据库初始化失败: %s', e)
        return False


def ensure_db_initialized(db_path: str = None) -> bool:
    """确保数据库已初始化"""
    if db_path is None:
        db_path = get_personal_db_path()

    if not os.path.exists(db_path):
        return init_personal_db(db_path)
    return True


def insert_transactions(transactions: list, batch_id: int = None,
                        db_path: str = None) -> Tuple[int, int]:
    """
    批量插入交易记录
    注意：重复记录（ID已存在）会被跳过，不会更新
    Returns: (新增数量, 重复数量)
    """
    logger = get_logger()

    if not transactions:
        return 0, 0

    ensure_db_initialized(db_path)

    tx_dicts = []
    for tx in transactions:
        if isinstance(tx, dict):
            tx_dicts.append(tx)
        else:
            tx_dicts.append(tx.to_dict())

    tx_ids = [tx.get('id', '') for tx in tx_dicts]

    inserted = 0
    duplicates = 0
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    try:
        with get_db_connection(db_path) as conn:
            cursor = conn.cursor()

            if tx_ids:
                placeholders = ','.join('?' * len(tx_ids))
                cursor.execute(
                    f'SELECT id FROM transactions WHERE id IN ({placeholders})',
                    tx_ids
                )
                existing_ids = set(row['id'] for row in cursor.fetchall())
            else:
                existing_ids = set()

            for tx_dict in tx_dicts:
                tx_id = tx_dict.get('id', '')
                if not tx_id:
                    logger.warning('交易记录缺少 ID，跳过')
                    continue

                if tx_id in existing_ids:
                    duplicates += 1
                    continue

                try:
                    cursor.execute('''
                        INSERT INTO transactions (
                            id, bank_name, card_type, account_number,
                            trade_date, transaction_time, post_date,
                            amount, direction, balance,
                            summary, merchant, transaction_type, transaction_id,
                            category, category_icon, category_color,
                            installments, status, source_file,
                            created_at, updated_at
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (
                        tx_id,
                        tx_dict.get('bank_name', ''),
                        tx_dict.get('card_type', ''),
                        tx_dict.get('account_number', ''),
                        tx_dict.get('trade_date', ''),
                        tx_dict.get('transaction_time', ''),
                        tx_dict.get('post_date', ''),
                        float(tx_dict.get('amount', 0)),
                        tx_dict.get('direction', ''),
                        float(tx_dict.get('balance', 0)),
                        tx_dict.get('summary', ''),
                        tx_dict.get('merchant', ''),
                        tx_dict.get('transaction_type', ''),
                        tx_dict.get('transaction_id', ''),
                        tx_dict.get('category', ''),
                        tx_dict.get('category_icon', ''),
                        tx_dict.get('category_color', ''),
                        int(tx_dict.get('installments', 0)),
                        tx_dict.get('status', ''),
                        tx_dict.get('source_file', ''),
                        now,
                        now,
                    ))

                    inserted += 1
                    existing_ids.add(tx_id)

                except sqlite3.IntegrityError:
                    duplicates += 1
                    existing_ids.add(tx_id)
                except Exception as e:
                    logger.warning('插入交易记录失败: %s - %s', tx_id, e)

        logger.info('交易记录插入完成: 新增 %d 条，重复 %d 条', inserted, duplicates)
        return inserted, duplicates

    except Exception as e:
        logger.error('批量插入交易记录失败: %s', e)
        return 0, 0


def query_transactions(
    start_date: str = None,
    end_date: str = None,
    category: str = None,
    direction: str = None,
    merchant: str = None,
    bank_name: str = None,
    min_amount: float = None,
    max_amount: float = None,
    limit: int = 1000,
    offset: int = 0,
    order_by: str = 'trade_date',
    order_dir: str = 'DESC',
    db_path: str = None
) -> Tuple[List[Dict], int]:
    """
    查询交易记录
    Returns: (记录列表, 总数)
    """
    ensure_db_initialized(db_path)

    query = 'SELECT * FROM transactions WHERE 1=1'
    count_query = 'SELECT COUNT(*) as total FROM transactions WHERE 1=1'
    params = []

    if start_date:
        query += ' AND trade_date >= ?'
        count_query += ' AND trade_date >= ?'
        params.append(start_date)

    if end_date:
        query += ' AND trade_date <= ?'
        count_query += ' AND trade_date <= ?'
        params.append(end_date)

    if category:
        query += ' AND category = ?'
        count_query += ' AND category = ?'
        params.append(category)

    if direction:
        query += ' AND direction = ?'
        count_query += ' AND direction = ?'
        params.append(direction)

    if merchant:
        query += ' AND merchant LIKE ?'
        count_query += ' AND merchant LIKE ?'
        params.append(f'%{merchant}%')

    if bank_name:
        query += ' AND bank_name = ?'
        count_query += ' AND bank_name = ?'
        params.append(bank_name)

    if min_amount is not None:
        query += ' AND amount >= ?'
        count_query += ' AND amount >= ?'
        params.append(min_amount)

    if max_amount is not None:
        query += ' AND amount <= ?'
        count_query += ' AND amount <= ?'
        params.append(max_amount)

    valid_order_cols = ['trade_date', 'amount', 'category', 'merchant', 'created_at']
    if order_by not in valid_order_cols:
        order_by = 'trade_date'
    if order_dir.upper() not in ['ASC', 'DESC']:
        order_dir = 'DESC'

    query += f' ORDER BY {order_by} {order_dir.upper()}'
    query += ' LIMIT ? OFFSET ?'
    query_params = params + [limit, offset]

    try:
        with get_db_connection(db_path) as conn:
            cursor = conn.cursor()

            cursor.execute(count_query, params)
            total = cursor.fetchone()['total']

            cursor.execute(query, query_params)
            rows = cursor.fetchall()

            results = []
            for row in rows:
                results.append(dict(row))

            return results, total

    except Exception as e:
        logger = get_logger()
        logger.error('查询交易记录失败: %s', e)
        return [], 0


def get_summary(
    start_date: str = None,
    end_date: str = None,
    db_path: str = None
) -> Dict[str, Any]:
    """获取汇总统计"""
    ensure_db_initialized(db_path)

    params = []
    where_clause = "WHERE direction != 'transfer'"

    if start_date:
        where_clause += ' AND trade_date >= ?'
        params.append(start_date)
    if end_date:
        where_clause += ' AND trade_date <= ?'
        params.append(end_date)

    try:
        with get_db_connection(db_path) as conn:
            cursor = conn.cursor()

            cursor.execute(f'''
                SELECT
                    COUNT(*) as total_records,
                    SUM(CASE WHEN direction = 'income' THEN amount ELSE 0 END) as total_income,
                    SUM(CASE WHEN direction = 'expense' THEN amount ELSE 0 END) as total_expense,
                    SUM(CASE WHEN direction = 'income' THEN amount
                             WHEN direction = 'expense' THEN -amount
                             ELSE 0 END) as net_amount,
                    MIN(trade_date) as min_date,
                    MAX(trade_date) as max_date
                FROM transactions
                {where_clause}
            ''', params)

            row = cursor.fetchone()

            cursor.execute('SELECT COUNT(DISTINCT merchant) as merchant_count FROM transactions')
            merchant_count = cursor.fetchone()['merchant_count']

            cursor.execute('SELECT COUNT(DISTINCT category) as category_count FROM transactions')
            category_count = cursor.fetchone()['category_count']

            cursor.execute('SELECT COUNT(DISTINCT bank_name) as bank_count FROM transactions')
            bank_count = cursor.fetchone()['bank_count']

            return {
                'total_records': row['total_records'] or 0,
                'total_income': row['total_income'] or 0.0,
                'total_expense': row['total_expense'] or 0.0,
                'net_amount': row['net_amount'] or 0.0,
                'merchant_count': merchant_count or 0,
                'category_count': category_count or 0,
                'bank_count': bank_count or 0,
                'date_range': {
                    'start': row['min_date'] or '',
                    'end': row['max_date'] or ''
                }
            }

    except Exception as e:
        logger = get_logger()
        logger.error('获取汇总统计失败: %s', e)
        return {
            'total_records': 0,
            'total_income': 0.0,
            'total_expense': 0.0,
            'net_amount': 0.0,
            'merchant_count': 0,
            'category_count': 0,
            'bank_count': 0,
            'date_range': {'start': '', 'end': ''}
        }


def get_category_breakdown(
    direction: str = 'expense',
    start_date: str = None,
    end_date: str = None,
    db_path: str = None
) -> List[Dict[str, Any]]:
    """获取分类统计"""
    ensure_db_initialized(db_path)

    params = [direction]
    where_clause = 'WHERE direction = ?'

    if start_date:
        where_clause += ' AND trade_date >= ?'
        params.append(start_date)
    if end_date:
        where_clause += ' AND trade_date <= ?'
        params.append(end_date)

    try:
        with get_db_connection(db_path) as conn:
            cursor = conn.cursor()

            cursor.execute(f'''
                SELECT
                    category as name,
                    category_icon as icon,
                    category_color as color,
                    SUM(amount) as amount,
                    COUNT(*) as count
                FROM transactions
                {where_clause}
                GROUP BY category
                ORDER BY amount DESC
            ''', params)

            rows = cursor.fetchall()
            return [dict(row) for row in rows]

    except Exception as e:
        logger = get_logger()
        logger.error('获取分类统计失败: %s', e)
        return []


def get_monthly_trend(
    start_date: str = None,
    end_date: str = None,
    db_path: str = None
) -> List[Dict[str, Any]]:
    """获取月度趋势"""
    ensure_db_initialized(db_path)

    params = []
    where_clause = "WHERE direction != 'transfer'"

    if start_date:
        where_clause += ' AND trade_date >= ?'
        params.append(start_date)
    if end_date:
        where_clause += ' AND trade_date <= ?'
        params.append(end_date)

    try:
        with get_db_connection(db_path) as conn:
            cursor = conn.cursor()

            cursor.execute(f'''
                SELECT
                    substr(trade_date, 1, 7) as month,
                    SUM(CASE WHEN direction = 'income' THEN amount ELSE 0 END) as income,
                    SUM(CASE WHEN direction = 'expense' THEN amount ELSE 0 END) as expense,
                    SUM(CASE WHEN direction = 'income' THEN amount
                             WHEN direction = 'expense' THEN -amount
                             ELSE 0 END) as net
                FROM transactions
                {where_clause}
                GROUP BY substr(trade_date, 1, 7)
                ORDER BY month
            ''', params)

            rows = cursor.fetchall()
            return [dict(row) for row in rows]

    except Exception as e:
        logger = get_logger()
        logger.error('获取月度趋势失败: %s', e)
        return []


def get_top_merchants(
    limit: int = 10,
    direction: str = 'expense',
    start_date: str = None,
    end_date: str = None,
    db_path: str = None
) -> List[Dict[str, Any]]:
    """获取 Top 商户"""
    ensure_db_initialized(db_path)

    params = [direction]
    where_clause = 'WHERE direction = ?'

    if start_date:
        where_clause += ' AND trade_date >= ?'
        params.append(start_date)
    if end_date:
        where_clause += ' AND trade_date <= ?'
        params.append(end_date)

    params.append(limit)

    try:
        with get_db_connection(db_path) as conn:
            cursor = conn.cursor()

            cursor.execute(f'''
                SELECT
                    merchant as name,
                    SUM(amount) as amount,
                    COUNT(*) as count
                FROM transactions
                {where_clause}
                AND merchant != ''
                GROUP BY merchant
                ORDER BY amount DESC
                LIMIT ?
            ''', params)

            rows = cursor.fetchall()
            return [dict(row) for row in rows]

    except Exception as e:
        logger = get_logger()
        logger.error('获取 Top 商户失败: %s', e)
        return []


def delete_transactions(ids: List[str], db_path: str = None) -> int:
    """删除交易记录"""
    if not ids:
        return 0

    ensure_db_initialized(db_path)

    try:
        with get_db_connection(db_path) as conn:
            cursor = conn.cursor()
            placeholders = ','.join('?' * len(ids))
            cursor.execute(f'DELETE FROM transactions WHERE id IN ({placeholders})', ids)
            return cursor.rowcount

    except Exception as e:
        logger = get_logger()
        logger.error('删除交易记录失败: %s', e)
        return 0


def update_transaction_category(tx_id: str, category: str,
                                category_icon: str = '',
                                category_color: str = '',
                                db_path: str = None) -> bool:
    """更新交易分类"""
    ensure_db_initialized(db_path)

    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    try:
        with get_db_connection(db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('''
                UPDATE transactions
                SET category = ?, category_icon = ?, category_color = ?, updated_at = ?
                WHERE id = ?
            ''', (category, category_icon, category_color, now, tx_id))
            return cursor.rowcount > 0

    except Exception as e:
        logger = get_logger()
        logger.error('更新交易分类失败: %s', e)
        return False


def create_import_batch(batch_name: str, source_path: str = '',
                        db_path: str = None) -> int:
    """创建导入批次"""
    ensure_db_initialized(db_path)

    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    try:
        with get_db_connection(db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('''
                INSERT INTO import_batches (batch_name, source_path, status, created_at)
                VALUES (?, ?, 'processing', ?)
            ''', (batch_name, source_path, now))
            return cursor.lastrowid

    except Exception as e:
        logger = get_logger()
        logger.error('创建导入批次失败: %s', e)
        return 0


def update_import_batch(batch_id: int, status: str,
                        total_files: int = None,
                        parsed_files: int = None,
                        failed_files: int = None,
                        total_records: int = None,
                        db_path: str = None) -> bool:
    """更新导入批次状态"""
    ensure_db_initialized(db_path)

    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    try:
        with get_db_connection(db_path) as conn:
            cursor = conn.cursor()

            updates = []
            params = []

            if status is not None:
                updates.append('status = ?')
                params.append(status)
            if total_files is not None:
                updates.append('total_files = ?')
                params.append(total_files)
            if parsed_files is not None:
                updates.append('parsed_files = ?')
                params.append(parsed_files)
            if failed_files is not None:
                updates.append('failed_files = ?')
                params.append(failed_files)
            if total_records is not None:
                updates.append('total_records = ?')
                params.append(total_records)

            if status in ('completed', 'failed'):
                updates.append('finished_at = ?')
                params.append(now)

            params.append(batch_id)

            cursor.execute(f'''
                UPDATE import_batches
                SET {', '.join(updates)}
                WHERE id = ?
            ''', params)

            return cursor.rowcount > 0

    except Exception as e:
        logger = get_logger()
        logger.error('更新导入批次失败: %s', e)
        return False


def get_recent_batches(limit: int = 10, db_path: str = None) -> List[Dict]:
    """获取最近的导入批次"""
    ensure_db_initialized(db_path)

    try:
        with get_db_connection(db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('''
                SELECT * FROM import_batches
                ORDER BY id DESC
                LIMIT ?
            ''', (limit,))
            rows = cursor.fetchall()
            return [dict(row) for row in rows]

    except Exception as e:
        logger = get_logger()
        logger.error('获取导入批次失败: %s', e)
        return []


def clear_all_data(db_path: str = None) -> bool:
    """清空所有个人理财数据（谨慎使用）"""
    ensure_db_initialized(db_path)

    try:
        with get_db_connection(db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('DELETE FROM transactions')
            cursor.execute('DELETE FROM accounts')
            cursor.execute('DELETE FROM import_batches')
            return True

    except Exception as e:
        logger = get_logger()
        logger.error('清空数据失败: %s', e)
        return False
