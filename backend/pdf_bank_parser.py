# -*- coding: utf-8 -*-
"""
PDF 银行流水解析模块
功能：
  1. 通过表格提取（pdfplumber）解析网银导出的 PDF 流水
  2. 通过 OCR（pytesseract + Pillow）解析扫描件 PDF 流水
  3. 将 PDF 内容转换为与 Excel 解析兼容的统一中间格式
  4. 支持银行类型识别和多账号区块检测

统一中间格式（与 Excel 解析保持一致）：
    {
        '银行': str,
        '银行账号': str,
        '主体': str,
        '交易日期': str/datetime,
        '付款': float/None,   # 负数表示支出
        '收款': float/None,   # 正数表示收入
        '摘要': str,
        '对方户名': str,
        '余额': float/None,
        '交易流水号': str,
    }
"""

import os
import re
import logging
from dataclasses import dataclass, field
from typing import List, Optional, Dict, Any, Tuple, Callable

import pandas as pd

try:
    import pdfplumber
    HAS_PDFPLUMBER = True
except ImportError:
    HAS_PDFPLUMBER = False

try:
    from PIL import Image
    import pytesseract
    HAS_OCR = True
except ImportError:
    HAS_OCR = False
    Image = None
    pytesseract = None

from bankcheck import (
    get_logger, is_numeric, to_float, _normalize_width, _strip_separators,
    _SEP_PATTERN, get_subject_info, generate_unique_id,
)


PDF_EXTS = ('.pdf',)

STANDARD_COLUMNS = [
    '唯一id', '银行', '银行账号', '主体', '交易日期',
    '付款', '收款', '摘要', '对方户名', '余额', '交易流水号',
]


@dataclass
class PdfParseResult:
    rows: List[dict] = field(default_factory=list)
    bank_name: Optional[str] = None
    account: Optional[str] = None
    page_count: int = 0
    method: str = 'table'
    raw_tables: List[List[List[str]]] = field(default_factory=list)
    error: Optional[str] = None


def is_pdf_file(filepath: str) -> bool:
    return isinstance(filepath, str) and filepath.lower().endswith(PDF_EXTS)


# ──────────────────────────────────────────────
# 银行关键词配置（用于 PDF 内容识别）
# ──────────────────────────────────────────────

BANK_KEYWORDS: Dict[str, List[str]] = {
    '北京银行': ['北京银行', 'BANK OF BEIJING', 'BOB'],
    '东亚银行': ['东亚银行', 'BANK OF EAST ASIA', 'BEA'],
    '工商银行': ['中国工商银行', '工商银行', 'ICBC', 'INDUSTRIAL AND COMMERCIAL BANK'],
    '建设银行': ['中国建设银行', '建设银行', 'CCB', 'CHINA CONSTRUCTION BANK'],
    '招商银行': ['招商银行', 'CMB', 'CHINA MERCHANTS BANK'],
    '农业银行': ['中国农业银行', '农业银行', 'ABC', 'AGRICULTURAL BANK OF CHINA'],
    '中国银行': ['中国银行', 'BOC', 'BANK OF CHINA'],
    '交通银行': ['交通银行', 'BOCOM', 'BANK OF COMMUNICATIONS'],
    '浦发银行': ['浦东发展银行', '浦发银行', 'SPDB', 'SHANGHAI PUDONG DEVELOPMENT BANK'],
    '民生银行': ['民生银行', 'CMBC', 'CHINA MINSHENG BANK'],
    '兴业银行': ['兴业银行', 'CIB', 'INDUSTRIAL BANK'],
    '平安银行': ['平安银行', 'PING AN BANK', 'SPABANK'],
    '中信银行': ['中信银行', 'CITIC', 'CHINA CITIC BANK'],
    '光大银行': ['光大银行', 'CEB', 'CHINA EVERBRIGHT BANK'],
    '华夏银行': ['华夏银行', 'HXB', 'HUAXIA BANK'],
    '广发银行': ['广发银行', 'CGB', 'CHINA GUANGFA BANK'],
    '邮储银行': ['邮政储蓄银行', '邮储银行', 'PSBC', 'POSTAL SAVINGS BANK'],
}


ACCOUNT_PATTERNS = [
    re.compile(r'(?:账号|账户|账\s*号|账\s*户)[：:\s]*([0-9]{6,25}[-\s]?[0-9]*)'),
    re.compile(r'(?:Account|A/C|ACCT)[：:\s#]*([0-9]{6,25}[-\s]?[0-9]*)', re.IGNORECASE),
    re.compile(r'(\d{16,22})'),
]


# ──────────────────────────────────────────────
# PDF 文本提取（表格 + 纯文本）
# ──────────────────────────────────────────────

def extract_tables_with_pdfplumber(pdf_path: str) -> Tuple[List[List[List[str]]], List[str]]:
    """
    使用 pdfplumber 提取 PDF 中的所有表格及每页纯文本。

    Returns:
        (tables, page_texts):
            tables: List[page_table -> List[row -> List[cell]]]
            page_texts: List[page_text_str]
    """
    logger = get_logger()
    if not HAS_PDFPLUMBER:
        logger.warning('pdfplumber 未安装，无法提取 PDF 表格')
        return [], []

    tables: List[List[List[str]]] = []
    page_texts: List[str] = []

    try:
        with pdfplumber.open(pdf_path) as pdf:
            logger.info('PDF 共 %d 页，开始提取表格...', len(pdf.pages))
            for i, page in enumerate(pdf.pages):
                text = page.extract_text() or ''
                page_texts.append(text)

                page_tables = page.extract_tables({
                    'vertical_strategy': 'text',
                    'horizontal_strategy': 'text',
                    'intersection_tolerance': 5,
                    'snap_tolerance': 3,
                    'join_tolerance': 3,
                    'edge_min_length': 3,
                }) or []

                cleaned_tables = _clean_extracted_tables(page_tables)
                tables.extend(cleaned_tables)

                logger.debug('第 %d 页：提取到 %d 个表格', i + 1, len(cleaned_tables))
    except Exception as e:
        logger.error('pdfplumber 提取失败: %s', e)
        return [], []

    return tables, page_texts


def _clean_extracted_tables(page_tables: List) -> List[List[List[str]]]:
    """清洗 pdfplumber 提取的原始表格，去除空行空列"""
    result = []
    for table in page_tables:
        if not table:
            continue
        cleaned_rows = []
        for row in table:
            if row is None:
                continue
            cleaned_cells = [
                str(c).strip() if c is not None else ''
                for c in row
            ]
            if any(cell for cell in cleaned_cells):
                cleaned_rows.append(cleaned_cells)
        if cleaned_rows:
            result.append(cleaned_rows)
    return result


def extract_text_with_ocr(pdf_path: str) -> str:
    """
    使用 OCR 提取扫描件 PDF 文本（fallback 方案）。
    需要系统安装 Tesseract OCR 引擎。
    """
    logger = get_logger()
    if not HAS_OCR:
        logger.warning('OCR 依赖未安装（Pillow / pytesseract），跳过 OCR')
        return ''
    if not HAS_PDFPLUMBER:
        logger.warning('pdfplumber 未安装，无法将 PDF 渲染为图片供 OCR 使用')
        return ''

    full_text_parts = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            logger.info('开始 OCR 识别，共 %d 页...', len(pdf.pages))
            for i, page in enumerate(pdf.pages):
                img = page.to_image(resolution=300)
                pil_img = img.original
                text = pytesseract.image_to_string(pil_img, lang='chi_sim+eng')
                full_text_parts.append(text)
                logger.debug('第 %d 页 OCR 完成: %d 字符', i + 1, len(text))
    except pytesseract.TesseractNotFoundError:
        logger.error('未检测到 Tesseract OCR 引擎，请先安装 tesseract-ocr')
        return ''
    except Exception as e:
        logger.error('OCR 提取失败: %s', e)
        return ''

    return '\n'.join(full_text_parts)


# ──────────────────────────────────────────────
# 银行类型识别
# ──────────────────────────────────────────────

def identify_bank_from_pdf(pdf_path: str) -> Optional[str]:
    """
    从 PDF 文件中识别银行类型。

    识别策略：
    1. 文件名匹配（复用现有逻辑）
    2. PDF 文本内容关键词匹配
    """
    from bankcheck import identify_bank as _id_bank_excel

    direct = _id_bank_excel(pdf_path)
    if direct:
        return direct

    logger = get_logger()
    basename = os.path.basename(pdf_path)

    _, page_texts = extract_tables_with_pdfplumber(pdf_path)
    full_text = '\n'.join(page_texts)

    if not full_text.strip():
        full_text = extract_text_with_ocr(pdf_path)

    if not full_text.strip():
        logger.warning('PDF「%s」无法提取任何文本内容', basename)
        return None

    normalized_text = _normalize_width(full_text)
    stripped_text = _strip_separators(normalized_text)

    scores: Dict[str, int] = {}
    for bank_name, keywords in BANK_KEYWORDS.items():
        score = 0
        for kw in keywords:
            norm_kw = _normalize_width(kw)
            stripped_kw = _strip_separators(norm_kw)
            if stripped_kw and stripped_kw in stripped_text:
                score += len(stripped_kw)
        if score > 0:
            scores[bank_name] = score

    if scores:
        best = max(scores.items(), key=lambda x: x[1])
        logger.info('PDF「%s」通过内容关键词识别为: %s（得分: %d）', basename, best[0], best[1])
        return best[0]

    logger.warning('PDF「%s」文本中未识别出银行关键词', basename)
    return None


def extract_account_from_pdf(pdf_path: str) -> Optional[str]:
    """从 PDF 文本中提取银行账号"""
    logger = get_logger()
    _, page_texts = extract_tables_with_pdfplumber(pdf_path)
    full_text = '\n'.join(page_texts)

    if not full_text.strip():
        full_text = extract_text_with_ocr(pdf_path)

    if not full_text:
        return None

    for pattern in ACCOUNT_PATTERNS:
        match = pattern.search(full_text)
        if match:
            account = re.sub(r'[-\s]', '', match.group(1))
            if len(account) >= 6:
                logger.info('PDF 中提取到账号: %s', account)
                return account

    return None


# ──────────────────────────────────────────────
# 表头智能检测
# ──────────────────────────────────────────────

COLUMN_KEYWORDS: Dict[str, List[str]] = {
    'trade_date': ['交易日期', '记账日期', '日期', 'Date', 'Trade Date'],
    'payment': ['支出', '付出', '付款', '借方', '支取', 'Debit', 'Withdrawal', '支出金额', '借方发生额'],
    'receipt': ['收入', '存入', '收款', '贷方', 'Credit', 'Deposit', '收入金额', '贷方发生额'],
    'amount': ['金额', '交易金额', 'Amount'],
    'direction': ['借贷标志', '借贷', '收付', '方向'],
    'summary': ['摘要', '用途', '备注', '附言', '说明', 'Description', 'Remark', 'Summary'],
    'counterpart': ['对方户名', '对方单位', '收款人', '付款人', '对方名称', '交易对手', 'Counterparty', 'Beneficiary'],
    'counterpart_account': ['对方账号', '对方账户'],
    'counterpart_bank': ['对方开户行', '对方银行'],
    'balance': ['余额', '结余', 'Balance'],
    'transaction_id': ['交易流水号', '流水号', '交易编号', '凭证号', 'Reference No', 'Transaction ID'],
    'currency': ['币种', '货币', 'Currency'],
}


def detect_header_row(table: List[List[str]]) -> Optional[int]:
    """
    在表格中智能检测表头行索引。
    通过计算每行匹配到的字段关键词数量来确定。
    """
    if not table:
        return None

    best_row = None
    best_score = 0

    max_scan_rows = min(len(table), 10)
    for row_idx in range(max_scan_rows):
        row = table[row_idx]
        score = 0
        matched_fields = set()
        for cell in row:
            if not cell:
                continue
            cell_text = _normalize_width(str(cell)).strip()
            if not cell_text:
                continue
            for field_key, keywords in COLUMN_KEYWORDS.items():
                for kw in keywords:
                    if kw in cell_text and field_key not in matched_fields:
                        score += 1
                        matched_fields.add(field_key)
                        break
        if score > best_score and score >= 3:
            best_score = score
            best_row = row_idx

    return best_row


def map_columns_by_header(header_row: List[str]) -> Dict[str, int]:
    """
    根据表头行，建立 {字段key: 列索引} 的映射。
    """
    mapping: Dict[str, int] = {}
    for col_idx, cell in enumerate(header_row):
        if not cell:
            continue
        cell_text = _normalize_width(str(cell)).strip()
        if not cell_text:
            continue
        for field_key, keywords in COLUMN_KEYWORDS.items():
            if field_key in mapping:
                continue
            for kw in keywords:
                if kw in cell_text:
                    mapping[field_key] = col_idx
                    break
    return mapping


# ──────────────────────────────────────────────
# 表格 -> 统一格式转换
# ──────────────────────────────────────────────

def _parse_amount(value: str) -> Optional[float]:
    """解析金额字符串，支持千分位、货币符号、全角数字"""
    if not value:
        return None
    s = str(value).strip()
    if not s:
        return None
    s = _normalize_width(s)
    s = s.replace(',', '').replace('，', '')
    s = s.replace('¥', '').replace('￥', '')
    s = s.replace('CNY', '').replace('RMB', '').replace('¥', '')
    s = s.strip()
    if s in ('', '-', '--', '—', '无'):
        return None
    sign = 1.0
    if s.startswith('(') and s.endswith(')'):
        sign = -1.0
        s = s[1:-1]
    if s.startswith('-'):
        sign = -1.0
        s = s[1:]
    try:
        return float(s) * sign
    except (ValueError, TypeError):
        return None


def _parse_date(value: str) -> Optional[str]:
    """解析日期字符串，返回标准格式 YYYY-MM-DD"""
    if not value:
        return None
    s = _normalize_width(str(value)).strip()
    if not s:
        return None
    patterns = [
        (r'(\d{4})[-/年.](\d{1,2})[-/月.](\d{1,2})', lambda m: f"{m[1]}-{int(m[2]):02d}-{int(m[3]):02d}"),
        (r'(\d{4})(\d{2})(\d{2})', lambda m: f"{m[1]}-{m[2]}-{m[3]}"),
        (r'(\d{1,2})[-/](\d{1,2})[-/](\d{4})', lambda m: f"{m[3]}-{int(m[1]):02d}-{int(m[2]):02d}"),
    ]
    for pat, fmt in patterns:
        m = re.search(pat, s)
        if m:
            try:
                return fmt(m)
            except (ValueError, IndexError):
                continue
    return s


def convert_table_to_rows(
    table: List[List[str]],
    bank_name: str,
    account: str,
    lookup_source,
    header_row_idx: Optional[int] = None,
    column_mapping: Optional[Dict[str, int]] = None,
) -> List[dict]:
    """
    将提取的表格转换为标准银行流水记录列表。

    Args:
        table: 原始二维表格数据
        bank_name: 银行名称
        account: 银行账号
        lookup_source: 主体查找表
        header_row_idx: 表头行索引（None 表示自动检测）
        column_mapping: 列映射（None 表示根据表头自动推断）

    Returns:
        标准记录列表（与 Excel 解析格式一致）
    """
    logger = get_logger()
    if not table or len(table) < 2:
        return []

    if header_row_idx is None:
        header_row_idx = detect_header_row(table)
    if header_row_idx is None:
        logger.warning('无法检测表头行，尝试从第0行开始')
        header_row_idx = 0

    if header_row_idx >= len(table) - 1:
        return []

    header = table[header_row_idx]
    if column_mapping is None:
        column_mapping = map_columns_by_header(header)

    if not column_mapping or 'trade_date' not in column_mapping:
        logger.warning('列映射中缺少必要字段 trade_date，跳过此表格')
        return []

    subject_info = get_subject_info(account, lookup_source)
    subject = subject_info.get('subject', '')

    rows: List[dict] = []
    data_start = header_row_idx + 1

    has_separate_pr = ('payment' in column_mapping) and ('receipt' in column_mapping)
    has_amount = 'amount' in column_mapping
    has_direction = 'direction' in column_mapping

    for row_idx in range(data_start, len(table)):
        row = table[row_idx]
        if not row or all(not str(c).strip() for c in row):
            continue

        def _get(col_key: str) -> str:
            idx = column_mapping.get(col_key)
            if idx is None or idx >= len(row):
                return ''
            return str(row[idx]).strip() if row[idx] is not None else ''

        trade_date = _parse_date(_get('trade_date'))
        if not trade_date:
            continue

        payment = None
        receipt = None

        if has_separate_pr:
            payment = _parse_amount(_get('payment'))
            receipt = _parse_amount(_get('receipt'))
            if payment is not None:
                payment = -abs(payment)
            if receipt is not None:
                receipt = abs(receipt)
        elif has_amount:
            amount = _parse_amount(_get('amount'))
            if amount is not None:
                direction = _get('direction')
                if has_direction and direction:
                    d = _normalize_width(direction)
                    if any(kw in d for kw in ['借', '支出', '付款', '付', 'Debit', 'Dr']):
                        payment = -abs(amount)
                    else:
                        receipt = abs(amount)
                else:
                    if amount < 0:
                        payment = amount
                    else:
                        receipt = amount

        balance = _parse_amount(_get('balance'))
        summary = _get('summary')
        counterpart = _get('counterpart')
        transaction_id = _get('transaction_id')

        record = {
            '银行': bank_name,
            '银行账号': account,
            '主体': subject,
            '交易日期': trade_date,
            '付款': payment,
            '收款': receipt,
            '摘要': summary,
            '对方户名': counterpart,
            '余额': balance,
            '交易流水号': transaction_id,
        }
        rows.append(record)

    logger.info('表格转换完成，共 %d 条有效记录', len(rows))
    return rows


# ──────────────────────────────────────────────
# 多账号表格分段解析
# ──────────────────────────────────────────────

def detect_account_blocks_in_tables(
    tables: List[List[List[str]]],
) -> List[Dict[str, Any]]:
    """
    扫描所有表格，检测其中的账号切换，将表格按账号分段。

    Returns:
        [{ 'table_idx', 'start_row', 'end_row', 'account' }, ...]
    """
    blocks = []
    account_re = re.compile(r'(?:账号|账户)[：:\s]*([0-9]{6,25}[-\s]?[0-9]*)')
    fallback_account_re = re.compile(r'(\d{12,22})')

    for t_idx, table in enumerate(tables):
        if not table:
            continue
        current_account = None
        block_start = 0

        for r_idx, row in enumerate(table):
            row_text = ' '.join(str(c) for c in row if c)
            if not row_text:
                continue

            matched_account = None
            m = account_re.search(row_text)
            if m:
                matched_account = re.sub(r'[-\s]', '', m.group(1))
            else:
                fm = fallback_account_re.search(row_text)
                if fm:
                    candidate = fm.group(1)
                    if candidate and not candidate.startswith('20') and not candidate.startswith('202'):
                        matched_account = candidate

            if matched_account and matched_account != current_account:
                if current_account is not None:
                    blocks.append({
                        'table_idx': t_idx,
                        'start_row': block_start,
                        'end_row': r_idx - 1,
                        'account': current_account,
                    })
                current_account = matched_account
                block_start = r_idx + 1

        if current_account is not None:
            blocks.append({
                'table_idx': t_idx,
                'start_row': block_start,
                'end_row': len(table) - 1,
                'account': current_account,
            })

    return blocks


# ──────────────────────────────────────────────
# 主解析入口
# ──────────────────────────────────────────────

def parse_pdf_bank_statement(
    pdf_path: str,
    lookup_source,
    bank_name_hint: Optional[str] = None,
    enable_ocr_fallback: bool = True,
) -> PdfParseResult:
    """
    解析 PDF 银行流水主入口。

    Args:
        pdf_path: PDF 文件路径
        lookup_source: 主体查找表（路径或预加载 dict）
        bank_name_hint: 银行名称提示（如已知可加速）
        enable_ocr_fallback: 表格提取失败时是否启用 OCR

    Returns:
        PdfParseResult 对象，包含解析结果
    """
    logger = get_logger()
    result = PdfParseResult()

    if not os.path.isfile(pdf_path):
        result.error = f'文件不存在: {pdf_path}'
        logger.error(result.error)
        return result

    if not is_pdf_file(pdf_path):
        result.error = f'不是 PDF 文件: {pdf_path}'
        logger.error(result.error)
        return result

    # 1. 识别银行
    bank_name = bank_name_hint or identify_bank_from_pdf(pdf_path)
    if not bank_name:
        logger.warning('无法识别 PDF「%s」的银行类型，将尝试通用解析', os.path.basename(pdf_path))
        bank_name = '通用银行'
    result.bank_name = bank_name

    # 2. 提取账号
    account = extract_account_from_pdf(pdf_path) or ''
    result.account = account

    # 3. 提取表格
    tables, page_texts = extract_tables_with_pdfplumber(pdf_path)
    result.page_count = len(page_texts)
    result.raw_tables = tables

    all_rows: List[dict] = []

    if tables:
        result.method = 'table'
        blocks = detect_account_blocks_in_tables(tables)

        if blocks:
            logger.info('检测到 %d 个账号分段', len(blocks))
            for block in blocks:
                t = tables[block['table_idx']]
                sub_table = t[block['start_row']:block['end_row'] + 1]
                if len(sub_table) >= 2:
                    rows = convert_table_to_rows(
                        sub_table, bank_name, block['account'], lookup_source,
                    )
                    all_rows.extend(rows)
        else:
            for table in tables:
                rows = convert_table_to_rows(table, bank_name, account, lookup_source)
                all_rows.extend(rows)

    # 4. 表格提取失败 -> 尝试 OCR 纯文本解析
    if not all_rows and enable_ocr_fallback:
        logger.info('表格提取无有效记录，尝试 OCR 纯文本解析...')
        raw_text = '\n'.join(page_texts)
        if not raw_text.strip():
            raw_text = extract_text_with_ocr(pdf_path)

        if raw_text.strip():
            result.method = 'ocr' if not page_texts or not any(page_texts) else 'text'
            text_rows = parse_text_to_rows(raw_text, bank_name, account, lookup_source)
            all_rows.extend(text_rows)

    result.rows = all_rows
    if not all_rows:
        result.error = '未从 PDF 中解析到任何有效流水记录'
        logger.warning(result.error)
    else:
        logger.info('PDF 解析完成，共 %d 条记录（方法: %s）', len(all_rows), result.method)

    return result


# ──────────────────────────────────────────────
# 纯文本 -> 表格（用于 OCR 回退或纯文本 PDF）
# ──────────────────────────────────────────────

def parse_text_to_rows(
    raw_text: str,
    bank_name: str,
    account: str,
    lookup_source,
) -> List[dict]:
    """
    将 OCR 或纯文本 PDF 的文本解析为记录列表。

    策略：按行切分，正则匹配类似流水记录的行（含日期 + 金额）。
    """
    logger = get_logger()
    rows: List[dict] = []

    lines = raw_text.splitlines()
    subject_info = get_subject_info(account, lookup_source)
    subject = subject_info.get('subject', '')

    date_pat = r'(?:\d{4}[-/年.]\d{1,2}[-/月.]\d{1,2}|\d{8})'
    amount_pat = r'[-+]?\d{1,3}(?:,\d{3})*(?:\.\d{2})?|[-+]?\d+(?:\.\d{2})?'
    record_pat = re.compile(
        r'(' + date_pat + r')'
        r'.*?'
        r'(' + amount_pat + r')'
    )

    for line in lines:
        line = _normalize_width(line).strip()
        if len(line) < 10:
            continue

        m = record_pat.search(line)
        if not m:
            continue

        trade_date = _parse_date(m.group(1))
        if not trade_date:
            continue

        amounts = re.findall(amount_pat, line)
        payment = None
        receipt = None
        balance = None

        numeric_amounts = []
        for a in amounts:
            v = _parse_amount(a)
            if v is not None:
                numeric_amounts.append(v)

        if len(numeric_amounts) >= 2:
            v1 = numeric_amounts[0]
            if v1 < 0:
                payment = v1
            else:
                receipt = abs(v1)
            balance = numeric_amounts[-1]
        elif len(numeric_amounts) == 1:
            v = numeric_amounts[0]
            if v < 0:
                payment = v
            else:
                receipt = abs(v)

        rest = line[m.end():].strip()
        tokens = [t for t in re.split(r'\s{2,}|\t', rest) if t.strip()]

        summary = tokens[0] if tokens else ''
        counterpart = tokens[1] if len(tokens) > 1 else ''

        rows.append({
            '银行': bank_name,
            '银行账号': account,
            '主体': subject,
            '交易日期': trade_date,
            '付款': payment,
            '收款': receipt,
            '摘要': summary,
            '对方户名': counterpart,
            '余额': balance,
            '交易流水号': '',
        })

    if rows:
        logger.info('纯文本解析得到 %d 条记录', len(rows))
    return rows


# ──────────────────────────────────────────────
# 与 bankcheck 主流程集成的接口
# ──────────────────────────────────────────────

def process_pdf_file(
    filepath: str,
    lookup_source,
    bank_name_hint: Optional[str] = None,
) -> List[dict]:
    """
    与 BANK_PROCESSORS 接口兼容的单文件处理函数。

    Args:
        filepath: PDF 文件路径
        lookup_source: 主体查找表
        bank_name_hint: 银行名称提示

    Returns:
        标准记录列表（每行一个 dict，包含 唯一id 等字段）
    """
    logger = get_logger()
    parse_result = parse_pdf_bank_statement(filepath, lookup_source, bank_name_hint)
    rows = parse_result.rows

    for r in rows:
        if '唯一id' not in r:
            r['唯一id'] = generate_unique_id()
        extra = get_subject_info(r.get('银行账号', ''), lookup_source)
        if extra and extra.get('subject') and not r.get('主体'):
            r['主体'] = extra['subject']
        for col in STANDARD_COLUMNS:
            if col not in r:
                r[col] = '' if col not in ('付款', '收款', '余额') else None

    return rows


def scan_pdf_files(folder: str) -> List[str]:
    """递归扫描文件夹中的 PDF 文件（与 scan_excel_files 对应）"""
    logger = get_logger()
    pdf_files = []
    for root, _dirs, files in os.walk(folder):
        for f in files:
            if f.startswith('~$'):
                continue
            if f.lower().endswith(PDF_EXTS):
                full_path = os.path.join(root, f)
                pdf_files.append(full_path)
                logger.debug('发现 PDF 文件: %s', full_path)
    if pdf_files:
        logger.info('共扫描到 %d 个 PDF 文件', len(pdf_files))
    return pdf_files


def scan_bank_files(folder: str) -> List[str]:
    """
    通用文件扫描：同时扫描 Excel 和 PDF 文件。
    这是对 scan_excel_files 的扩展版本。
    """
    from bankcheck import scan_excel_files
    excel_files = scan_excel_files(folder)
    pdf_files = scan_pdf_files(folder)
    return excel_files + pdf_files
