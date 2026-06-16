# -*- coding: utf-8 -*-
"""
个人理财专版 - 核心解析模块
功能：
  1. 解析个人借记卡、信用卡账单
  2. 自动消费分类
  3. 轻量资金视图汇总
  4. 与对公版配置完全隔离
"""

import os
import sys
import re
import yaml
import hashlib
import logging
from dataclasses import dataclass, field
from datetime import datetime
from typing import List, Dict, Any, Optional, Tuple
from collections import defaultdict

try:
    from bankcheck import (
        open_workbook_compat,
        read_csv_as_workbook,
        detect_file_format,
        iter_sheet_records,
        get_script_dir as _get_script_dir,
    )
    HAS_BANKCHECK = True
except ImportError:
    HAS_BANKCHECK = False


PERSONAL_RULES_FILENAME = 'personal_bank_rules.yaml'
CATEGORY_RULES_FILENAME = 'personal_category_rules.yaml'


def get_logger():
    return logging.getLogger('bankcheck')


def get_personal_rules_path():
    """获取个人银行规则配置文件路径"""
    return os.path.join(_get_script_dir() if HAS_BANKCHECK else os.path.dirname(os.path.abspath(__file__)),
                        PERSONAL_RULES_FILENAME)


def get_category_rules_path():
    """获取消费分类规则配置文件路径"""
    return os.path.join(_get_script_dir() if HAS_BANKCHECK else os.path.dirname(os.path.abspath(__file__)),
                        CATEGORY_RULES_FILENAME)


def load_personal_rules():
    """加载个人银行规则配置"""
    path = get_personal_rules_path()
    if not os.path.exists(path):
        return None
    try:
        with open(path, 'r', encoding='utf-8') as f:
            return yaml.safe_load(f)
    except Exception as e:
        logger = get_logger()
        logger.error('加载个人银行规则配置失败: %s', e)
        return None


def load_category_rules():
    """加载消费分类规则配置"""
    path = get_category_rules_path()
    if not os.path.exists(path):
        return None
    try:
        with open(path, 'r', encoding='utf-8') as f:
            return yaml.safe_load(f)
    except Exception as e:
        logger = get_logger()
        logger.error('加载消费分类规则配置失败: %s', e)
        return None


@dataclass
class PersonalTransaction:
    """个人交易记录"""
    id: str = ''
    bank_name: str = ''
    card_type: str = ''  # debit / credit
    account_number: str = ''
    trade_date: str = ''
    transaction_time: str = ''
    post_date: str = ''
    amount: float = 0.0
    direction: str = ''  # income / expense / transfer
    balance: float = 0.0
    summary: str = ''
    merchant: str = ''
    transaction_type: str = ''
    transaction_id: str = ''
    category: str = ''
    category_icon: str = ''
    category_color: str = ''
    installments: int = 0
    status: str = ''
    raw_data: dict = field(default_factory=dict)

    def __post_init__(self):
        if not self.id:
            id_src = f"{self.bank_name}|{self.trade_date}|{self.amount}|{self.merchant}|{self.summary}"
            self.id = hashlib.md5(id_src.encode('utf-8')).hexdigest()

    def to_dict(self):
        return {
            'id': self.id,
            'bank_name': self.bank_name,
            'card_type': self.card_type,
            'account_number': self.account_number,
            'trade_date': self.trade_date,
            'transaction_time': self.transaction_time,
            'post_date': self.post_date,
            'amount': self.amount,
            'direction': self.direction,
            'balance': self.balance,
            'summary': self.summary,
            'merchant': self.merchant,
            'transaction_type': self.transaction_type,
            'transaction_id': self.transaction_id,
            'category': self.category,
            'category_icon': self.category_icon,
            'category_color': self.category_color,
            'installments': self.installments,
            'status': self.status,
        }


class CategoryClassifier:
    """消费分类器"""

    def __init__(self, category_rules=None):
        self.rules = category_rules or load_category_rules()
        self.expense_categories = []
        self.income_categories = []
        self.special_rules = []
        self._load_categories()

    def _load_categories(self):
        if not self.rules:
            return
        self.expense_categories = sorted(
            self.rules.get('expense_categories', []),
            key=lambda x: x.get('priority', 0),
            reverse=True
        )
        self.income_categories = sorted(
            self.rules.get('income_categories', []),
            key=lambda x: x.get('priority', 0),
            reverse=True
        )
        self.special_rules = self.rules.get('special_rules', [])

    def classify(self, transaction: PersonalTransaction) -> Tuple[str, str, str]:
        """
        对交易进行分类
        Returns: (category_name, category_icon, category_color)
        """
        text = f"{transaction.merchant} {transaction.summary} {transaction.transaction_type}"
        text = text.lower()

        for rule in self.special_rules:
            if self._match_keywords(text, rule.get('keywords', [])):
                if rule.get('type') == 'transfer':
                    transaction.direction = 'transfer'
                return (rule.get('name', '转账'), '🔄', '#6B7280')

        if transaction.direction == 'income':
            categories = self.income_categories
            default_cat = self.rules.get('general', {}).get('default_income_category', '其他收入') if self.rules else '其他收入'
        else:
            categories = self.expense_categories
            default_cat = self.rules.get('general', {}).get('default_category', '其他支出') if self.rules else '其他支出'

        for cat in categories:
            if self._match_keywords(text, cat.get('keywords', [])):
                return (cat.get('name', default_cat),
                        cat.get('icon', '📝'),
                        cat.get('color', '#6B7280'))

        if transaction.direction == 'income':
            default_info = self._get_default_income_category()
        else:
            default_info = self._get_default_expense_category()
        return default_info

    def _match_keywords(self, text: str, keywords: list) -> bool:
        """检查文本是否包含任一关键词"""
        if not keywords:
            return False
        for kw in keywords:
            kw_str = str(kw).lower()
            if kw_str in text:
                return True
        return False

    def _get_default_expense_category(self) -> Tuple[str, str, str]:
        for cat in self.expense_categories:
            if cat.get('name') == '其他支出':
                return (cat['name'], cat.get('icon', '📝'), cat.get('color', '#6B7280'))
        return ('其他支出', '📝', '#6B7280')

    def _get_default_income_category(self) -> Tuple[str, str, str]:
        for cat in self.income_categories:
            if cat.get('name') == '其他收入':
                return (cat['name'], cat.get('icon', '🎉'), cat.get('color', '#6B7280'))
        return ('其他收入', '🎉', '#6B7280')

    def get_all_categories(self, direction: str = 'expense') -> list:
        """获取所有分类"""
        if direction == 'income':
            return self.income_categories
        return self.expense_categories


def detect_bank_by_filename(filename: str, rules: dict) -> Optional[dict]:
    """
    根据文件名判断银行类型
    Returns: 银行规则配置字典或 None
    """
    if not rules:
        return None

    filename_lower = filename.lower()

    all_cards = []
    all_cards.extend(rules.get('debit_cards', []))
    all_cards.extend(rules.get('credit_cards', []))

    best_match = None
    best_match_score = 0

    for card in all_cards:
        if not card.get('enabled', True):
            continue
        keywords = card.get('filename_keywords', [])
        score = 0
        for kw in keywords:
            if kw.lower() in filename_lower:
                score += 1
        if score > best_match_score:
            best_match_score = score
            best_match = card

    return best_match if best_match_score > 0 else None


def detect_bank_by_headers(headers: list, rules: dict) -> Optional[dict]:
    """
    根据表头判断银行类型
    Returns: 银行规则配置字典或 None
    """
    if not rules or not headers:
        return None

    header_set = set(str(h).strip() for h in headers if h)

    all_cards = []
    all_cards.extend(rules.get('debit_cards', []))
    all_cards.extend(rules.get('credit_cards', []))

    best_match = None
    best_match_score = 0

    for card in all_cards:
        if not card.get('enabled', True):
            continue
        column_mapping = card.get('column_mapping', {})
        expected_headers = set(column_mapping.values())
        match_count = len(header_set & expected_headers)
        if match_count > best_match_score:
            best_match_score = match_count
            best_match = card

    return best_match if best_match_score >= 2 else None


def _parse_amount(value) -> float:
    """解析金额为浮点数"""
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip()
    s = s.replace(',', '').replace('，', '').replace('¥', '').replace('￥', '')
    s = s.replace('元', '').replace(' ', '')
    try:
        return float(s)
    except (ValueError, TypeError):
        return 0.0


def _generate_transaction_id(*args) -> str:
    """生成交易唯一ID"""
    raw = '|'.join(str(a) for a in args)
    return hashlib.md5(raw.encode('utf-8')).hexdigest()


def parse_personal_bill(filepath: str,
                        bank_rule: dict = None,
                        classifier: CategoryClassifier = None) -> List[PersonalTransaction]:
    """
    解析个人账单文件

    Args:
        filepath: 账单文件路径
        bank_rule: 银行规则，None 则自动检测
        classifier: 分类器，None 则使用默认

    Returns:
        list[PersonalTransaction]: 交易记录列表
    """
    logger = get_logger()
    logger.info('开始解析个人账单: %s', os.path.basename(filepath))

    if not os.path.exists(filepath):
        logger.error('文件不存在: %s', filepath)
        return []

    rules = load_personal_rules()
    if bank_rule is None:
        bank_rule = detect_bank_by_filename(os.path.basename(filepath), rules)

    if classifier is None:
        classifier = CategoryClassifier()

    try:
        wb, tmp_path = open_workbook_compat(filepath, read_only=False)
    except Exception as e:
        logger.error('打开文件失败: %s - %s', filepath, e)
        return []

    try:
        ws = wb.active if hasattr(wb, 'active') else wb.worksheets[0]
        if ws is None:
            logger.error('无法获取工作表')
            return []

        if bank_rule is None:
            header_row = 1
            headers = []
            for col in range(1, ws.max_column + 1):
                val = ws.cell(row=header_row, column=col).value
                headers.append(val)
            bank_rule = detect_bank_by_headers(headers, rules)

        if bank_rule is None:
            logger.warning('无法识别账单类型: %s', os.path.basename(filepath))
            return []

        logger.info('识别到账单类型: %s', bank_rule.get('bank_name', '未知'))

        transactions = _parse_sheet_with_rule(ws, bank_rule, classifier, filepath)

        logger.info('解析完成，共 %d 条记录', len(transactions))
        return transactions

    finally:
        if hasattr(wb, 'close'):
            wb.close()
        if tmp_path and os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except Exception:
                pass


def _parse_sheet_with_rule(ws, bank_rule: dict,
                           classifier: CategoryClassifier,
                           filepath: str) -> List[PersonalTransaction]:
    """根据规则解析工作表"""
    logger = get_logger()

    header_row = bank_rule.get('header_row', 1)
    start_row = bank_rule.get('start_row', header_row + 1)
    column_mapping = bank_rule.get('column_mapping', {})
    amount_sign_cfg = bank_rule.get('amount_sign', {})
    date_format = bank_rule.get('date_format', '%Y-%m-%d')
    bank_name = bank_rule.get('bank_name', '')
    card_type = bank_rule.get('card_type', 'debit')

    headers = {}
    for col in range(1, ws.max_column + 1):
        val = ws.cell(row=header_row, column=col).value
        if val:
            headers[str(val).strip()] = col

    col_index_map = {}
    for field_name, col_name in column_mapping.items():
        if isinstance(col_name, list):
            for cn in col_name:
                if cn in headers:
                    col_index_map[field_name] = headers[cn]
                    break
        else:
            if col_name in headers:
                col_index_map[field_name] = headers[col_name]

    if not col_index_map:
        logger.warning('未匹配到任何列')
        return []

    transactions = []
    file_basename = os.path.basename(filepath)

    for row_idx in range(start_row, ws.max_row + 1):
        record = {}
        has_data = False
        for field_name, col_idx in col_index_map.items():
            val = ws.cell(row=row_idx, column=col_idx).value
            if val is not None and str(val).strip() != '':
                has_data = True
            record[field_name] = val

        if not has_data:
            continue

        tx = _build_transaction(record, bank_rule, amount_sign_cfg,
                                bank_name, card_type, file_basename, row_idx)

        if tx.amount != 0 or tx.merchant or tx.summary:
            cat_name, cat_icon, cat_color = classifier.classify(tx)
            tx.category = cat_name
            tx.category_icon = cat_icon
            tx.category_color = cat_color
            tx.raw_data = record
            transactions.append(tx)

    return transactions


def _build_transaction(record: dict, bank_rule: dict,
                       amount_sign_cfg: dict,
                       bank_name: str, card_type: str,
                       file_basename: str, row_idx: int) -> PersonalTransaction:
    """构建交易记录对象"""
    tx = PersonalTransaction()
    tx.bank_name = bank_name
    tx.card_type = card_type

    tx.trade_date = _format_date(record.get('trade_date'),
                                  bank_rule.get('date_format', '%Y-%m-%d'))
    tx.transaction_time = str(record.get('transaction_time', '')) if record.get('transaction_time') else ''
    tx.post_date = _format_date(record.get('post_date'),
                                 bank_rule.get('date_format', '%Y-%m-%d'))

    amount = _parse_amount(record.get('amount', 0))
    tx.amount = abs(amount)

    direction = _determine_direction(amount, amount_sign_cfg, record)
    tx.direction = direction

    tx.balance = _parse_amount(record.get('balance', 0))
    tx.summary = str(record.get('summary', '') or '').strip()
    tx.merchant = str(record.get('merchant', '') or '').strip()
    tx.transaction_type = str(record.get('transaction_type', '') or '').strip()
    tx.transaction_id = str(record.get('transaction_id', '') or '').strip()
    tx.status = str(record.get('status', '') or '').strip()
    tx.installments = int(record.get('installments', 0) or 0)

    tx_id_src = f"{file_basename}|{row_idx}|{tx.trade_date}|{tx.amount}|{tx.merchant}"
    tx.id = _generate_transaction_id(tx_id_src)

    if not tx.transaction_id:
        tx.transaction_id = tx.id

    return tx


def _format_date(value, input_format: str = None) -> str:
    """格式化日期为 YYYY-MM-DD"""
    if value is None or value == '':
        return ''

    if isinstance(value, datetime):
        return value.strftime('%Y-%m-%d')

    s = str(value).strip()

    try:
        if input_format:
            dt = datetime.strptime(s, input_format)
            return dt.strftime('%Y-%m-%d')
    except (ValueError, TypeError):
        pass

    date_patterns = [
        '%Y-%m-%d',
        '%Y/%m/%d',
        '%Y年%m月%d日',
        '%Y%m%d',
        '%m/%d/%Y',
        '%d/%m/%Y',
        '%Y-%m-%d %H:%M:%S',
        '%Y/%m/%d %H:%M:%S',
    ]
    for fmt in date_patterns:
        try:
            dt = datetime.strptime(s, fmt)
            return dt.strftime('%Y-%m-%d')
        except (ValueError, TypeError):
            continue

    return s


def _determine_direction(amount: float, cfg: dict, record: dict) -> str:
    """判断收支方向"""
    use_type = cfg.get('use_type_column', False)

    if use_type:
        tx_type = str(record.get('transaction_type', '') or '').strip()
        income_types = cfg.get('income_types', [])
        expense_types = cfg.get('expense_types', [])

        for it in income_types:
            if it in tx_type:
                return 'income'
        for et in expense_types:
            if et in tx_type:
                return 'expense'

    direction = cfg.get('direction', 'negative')
    if direction == 'positive':
        if amount > 0:
            return 'income'
        elif amount < 0:
            return 'expense'
    else:
        if amount > 0:
            return 'expense'
        elif amount < 0:
            return 'income'

    return 'expense'


@dataclass
class PersonalSummary:
    """个人理财汇总"""
    total_income: float = 0.0
    total_expense: float = 0.0
    net_amount: float = 0.0
    total_records: int = 0
    income_count: int = 0
    expense_count: int = 0
    category_breakdown: List[Dict[str, Any]] = field(default_factory=list)
    monthly_trend: List[Dict[str, Any]] = field(default_factory=list)
    top_merchants: List[Dict[str, Any]] = field(default_factory=list)
    date_range: Dict[str, str] = field(default_factory=dict)


def summarize_transactions(transactions: List[PersonalTransaction]) -> PersonalSummary:
    """汇总交易数据"""
    summary = PersonalSummary()

    if not transactions:
        return summary

    valid_txs = [t for t in transactions if t.direction != 'transfer']

    summary.total_records = len(valid_txs)

    category_map = defaultdict(lambda: {
        'name': '',
        'icon': '',
        'color': '',
        'amount': 0.0,
        'count': 0,
        'direction': 'expense'
    })

    monthly_map = defaultdict(lambda: {
        'month': '',
        'income': 0.0,
        'expense': 0.0,
        'net': 0.0
    })

    merchant_map = defaultdict(lambda: {
        'name': '',
        'amount': 0.0,
        'count': 0,
        'direction': 'expense'
    })

    dates = []

    for tx in valid_txs:
        if tx.direction == 'income':
            summary.total_income += tx.amount
            summary.income_count += 1
        elif tx.direction == 'expense':
            summary.total_expense += tx.amount
            summary.expense_count += 1

        if tx.category:
            cat_key = tx.category
            cat_info = category_map[cat_key]
            cat_info['name'] = tx.category
            cat_info['icon'] = tx.category_icon
            cat_info['color'] = tx.category_color
            cat_info['amount'] += tx.amount
            cat_info['count'] += 1
            cat_info['direction'] = tx.direction

        if tx.trade_date:
            dates.append(tx.trade_date)
            month = tx.trade_date[:7]
            month_info = monthly_map[month]
            month_info['month'] = month
            if tx.direction == 'income':
                month_info['income'] += tx.amount
            elif tx.direction == 'expense':
                month_info['expense'] += tx.amount
            month_info['net'] = month_info['income'] - month_info['expense']

        if tx.merchant and tx.direction == 'expense':
            merch_key = tx.merchant
            merch_info = merchant_map[merch_key]
            merch_info['name'] = tx.merchant
            merch_info['amount'] += tx.amount
            merch_info['count'] += 1

    summary.net_amount = summary.total_income - summary.total_expense

    summary.category_breakdown = sorted(
        list(category_map.values()),
        key=lambda x: x['amount'],
        reverse=True
    )

    summary.monthly_trend = sorted(
        list(monthly_map.values()),
        key=lambda x: x['month']
    )

    summary.top_merchants = sorted(
        list(merchant_map.values()),
        key=lambda x: x['amount'],
        reverse=True
    )[:20]

    if dates:
        summary.date_range = {
            'start': min(dates),
            'end': max(dates)
        }

    return summary


def parse_personal_folder(folder_path: str) -> Tuple[List[PersonalTransaction], Dict[str, Any]]:
    """
    解析整个文件夹中的个人账单

    Returns:
        (transactions, stats)
    """
    logger = get_logger()
    logger.info('开始解析个人账单文件夹: %s', folder_path)

    all_transactions = []
    stats = {
        'total_files': 0,
        'parsed_files': 0,
        'failed_files': 0,
        'total_records': 0,
        'files': []
    }

    if not os.path.isdir(folder_path):
        logger.error('文件夹不存在: %s', folder_path)
        return [], stats

    rules = load_personal_rules()
    classifier = CategoryClassifier()

    supported_ext = ('.xlsx', '.xls', '.csv')

    for root, dirs, files in os.walk(folder_path):
        for filename in files:
            if not filename.lower().endswith(supported_ext):
                continue
            if filename.startswith('~$'):
                continue

            filepath = os.path.join(root, filename)
            stats['total_files'] += 1

            try:
                bank_rule = detect_bank_by_filename(filename, rules)
                txs = parse_personal_bill(filepath, bank_rule, classifier)

                if txs:
                    all_transactions.extend(txs)
                    stats['parsed_files'] += 1
                    stats['total_records'] += len(txs)
                    stats['files'].append({
                        'filename': filename,
                        'records': len(txs),
                        'bank_name': txs[0].bank_name if txs else '',
                        'status': 'success'
                    })
                else:
                    stats['failed_files'] += 1
                    stats['files'].append({
                        'filename': filename,
                        'records': 0,
                        'bank_name': '',
                        'status': 'no_data'
                    })

            except Exception as e:
                logger.error('解析文件失败 %s: %s', filename, e)
                stats['failed_files'] += 1
                stats['files'].append({
                    'filename': filename,
                    'records': 0,
                    'bank_name': '',
                    'status': 'error',
                    'error': str(e)
                })

    logger.info('解析完成: 共 %d 个文件，成功 %d 个，失败 %d 个，%d 条记录',
                stats['total_files'], stats['parsed_files'],
                stats['failed_files'], stats['total_records'])

    return all_transactions, stats
