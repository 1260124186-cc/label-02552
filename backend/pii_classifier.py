# -*- coding: utf-8 -*-
"""
PII（个人可识别信息）分级与脱敏模块

设计原则：
  1. INFO_SAFE（可写 INFO 日志）：非敏感聚合信息，不包含可识别个人/主体的数据
  2. DEBUG_ONLY（仅 DEBUG 日志允许，且需脱敏）：部分敏感字段，仅在开发调试时使用
  3. FORBIDDEN（禁止落盘，任何日志均不允许）：核心敏感数据，严禁出现在日志文件中

字段分类（基于银行流水业务场景）：
  - INFO_SAFE：银行名称、记录条数、处理状态、月份级日期、标签类字段
  - DEBUG_ONLY（需脱敏）：唯一ID（截断）、主体名称（掩码）、摘要（掩码）、对方户名（掩码）
  - FORBIDDEN：银行账号、付款金额、收款金额、余额、交易流水号、匹配键、日级精确日期

使用方式：
  from pii_classifier import (
      PIILevel, classify_field, mask_value,
      sanitize_for_log, PIILogFilter, setup_pii_aware_logging
  )
"""

import re
import logging
from enum import Enum
from typing import Any, Dict, List, Optional, Tuple, Union


class PIILevel(str, Enum):
    """PII 安全级别枚举"""

    INFO_SAFE = 'info_safe'
    DEBUG_ONLY = 'debug_only'
    FORBIDDEN = 'forbidden'


_FIELD_CLASSIFICATION: Dict[str, PIILevel] = {
    '银行': PIILevel.INFO_SAFE,
    '银行名称': PIILevel.INFO_SAFE,
    '记录数': PIILevel.INFO_SAFE,
    '条数': PIILevel.INFO_SAFE,
    '总数': PIILevel.INFO_SAFE,
    '处理状态': PIILevel.INFO_SAFE,
    '状态': PIILevel.INFO_SAFE,
    '模式': PIILevel.INFO_SAFE,
    '黑白名单标签': PIILevel.INFO_SAFE,
    '命中规则名称': PIILevel.INFO_SAFE,
    '命中关键词': PIILevel.INFO_SAFE,
    '工作表': PIILevel.INFO_SAFE,
    'sheet_name': PIILevel.INFO_SAFE,
    '模板': PIILevel.INFO_SAFE,
    '类型': PIILevel.INFO_SAFE,
    '币种': PIILevel.INFO_SAFE,
    '序号': PIILevel.INFO_SAFE,

    '唯一id': PIILevel.DEBUG_ONLY,
    '唯一ID': PIILevel.DEBUG_ONLY,
    'transaction_id': PIILevel.DEBUG_ONLY,
    '主体': PIILevel.DEBUG_ONLY,
    '主体名称': PIILevel.DEBUG_ONLY,
    'subject': PIILevel.DEBUG_ONLY,
    '摘要': PIILevel.DEBUG_ONLY,
    'summary': PIILevel.DEBUG_ONLY,
    '对方户名': PIILevel.DEBUG_ONLY,
    'counterparty': PIILevel.DEBUG_ONLY,
    '交易描述': PIILevel.DEBUG_ONLY,
    '备注': PIILevel.DEBUG_ONLY,
    'remark': PIILevel.DEBUG_ONLY,
    '附件类型': PIILevel.DEBUG_ONLY,

    '银行账号': PIILevel.FORBIDDEN,
    'account': PIILevel.FORBIDDEN,
    '对方账号': PIILevel.FORBIDDEN,
    '付款': PIILevel.FORBIDDEN,
    'payment': PIILevel.FORBIDDEN,
    '支出金额': PIILevel.FORBIDDEN,
    '收款': PIILevel.FORBIDDEN,
    'receipt': PIILevel.FORBIDDEN,
    '收入金额': PIILevel.FORBIDDEN,
    '余额': PIILevel.FORBIDDEN,
    'balance': PIILevel.FORBIDDEN,
    '交易流水号': PIILevel.FORBIDDEN,
    '匹配键': PIILevel.FORBIDDEN,
    'match_key': PIILevel.FORBIDDEN,
    '交易日期': PIILevel.FORBIDDEN,
    '交易时间': PIILevel.FORBIDDEN,
    '导入批次号': PIILevel.DEBUG_ONLY,
    '导入时间': PIILevel.DEBUG_ONLY,
    '附件路径': PIILevel.FORBIDDEN,
    '对方行名': PIILevel.FORBIDDEN,
    '凭证种类': PIILevel.FORBIDDEN,
    '凭证号码': PIILevel.FORBIDDEN,
    '手续费': PIILevel.FORBIDDEN,
    '利息': PIILevel.FORBIDDEN,
    '税费': PIILevel.FORBIDDEN,
    'id': PIILevel.DEBUG_ONLY,
    'workflow_id': PIILevel.DEBUG_ONLY,
    'amount': PIILevel.FORBIDDEN,
    'description': PIILevel.DEBUG_ONLY,
    'handled_by': PIILevel.DEBUG_ONLY,
    'created_at': PIILevel.DEBUG_ONLY,
    'updated_at': PIILevel.DEBUG_ONLY,
}


_KEYWORD_HINTS_FORBIDDEN = [
    '账号', '账户', '金额', '余额', '流水号', '流水', '卡号',
    'account', 'balance', 'payment', 'receipt', 'amount',
]

_KEYWORD_HINTS_DEBUG = [
    '主体', '户名', '摘要', '备注', '名称', '描述',
    'subject', 'counterparty', 'summary', 'remark', 'description',
]


def classify_field(field_name: str) -> PIILevel:
    """
    根据字段名判断其 PII 安全级别。

    先精确匹配分类表，未命中时根据关键词模糊推断。

    Args:
        field_name: 字段名称（中英文均可）

    Returns:
        PIILevel: 该字段对应的安全级别
    """
    if not field_name:
        return PIILevel.INFO_SAFE

    normalized = str(field_name).strip()

    if normalized in _FIELD_CLASSIFICATION:
        return _FIELD_CLASSIFICATION[normalized]

    lower = normalized.lower()
    for kw in _KEYWORD_HINTS_FORBIDDEN:
        if kw.lower() in lower:
            return PIILevel.FORBIDDEN

    for kw in _KEYWORD_HINTS_DEBUG:
        if kw.lower() in lower:
            return PIILevel.DEBUG_ONLY

    return PIILevel.INFO_SAFE


def _mask_bank_account(value: str) -> str:
    """银行账号脱敏：保留前3后3，中间用 * 替代"""
    if not value:
        return value
    s = value.strip()
    if len(s) <= 6:
        return '*' * len(s)
    return s[:3] + '*' * (len(s) - 6) + s[-3:]


def _mask_subject_name(value: str) -> str:
    """主体名称/对方户名脱敏：保留首尾各1字符"""
    if not value:
        return value
    s = value.strip()
    if len(s) <= 2:
        return s[0] + '*' if len(s) > 1 else '*'
    return s[0] + '*' * (len(s) - 2) + s[-1]


def _mask_text(value: str, keep_prefix: int = 2) -> str:
    """文本类字段（摘要、备注）脱敏：保留前 N 字，其余用 *** 替代"""
    if not value:
        return value
    s = value.strip()
    if len(s) <= keep_prefix:
        return s[:1] + '***' if s else '***'
    return s[:keep_prefix] + '***'


def _mask_transaction_id(value: str) -> str:
    """交易流水号脱敏：仅保留前4位"""
    if not value:
        return value
    s = value.strip()
    if len(s) <= 4:
        return '*' * len(s)
    return s[:4] + '*' * (len(s) - 4)


def _mask_unique_id(value: str) -> str:
    """唯一ID脱敏：保留前8位"""
    if not value:
        return value
    s = str(value).strip()
    if len(s) <= 8:
        return s[:4] + '****' if len(s) > 4 else '*' * len(s)
    return s[:8] + '****'


def _mask_amount(value: Any) -> str:
    """金额/余额脱敏：显示区间而非精确值"""
    try:
        if value is None or (isinstance(value, float) and value != value):
            return '[金额已隐藏]'
        v = abs(float(value))
        if v == 0:
            return '[0]'
        if v < 1000:
            return '[<1千]'
        if v < 10000:
            return '[1千-1万]'
        if v < 100000:
            return '[1万-10万]'
        if v < 1000000:
            return '[10万-100万]'
        if v < 10000000:
            return '[100万-1千万]'
        return '[>=1千万]'
    except (TypeError, ValueError):
        return '[金额已隐藏]'


def _mask_date(value: str) -> str:
    """日期脱敏：仅保留年月"""
    if not value:
        return value
    s = str(value).strip()
    m = re.match(r'(\d{4})[-/.年]?(\d{1,2})', s)
    if m:
        return f"{m.group(1)}年{m.group(2)}月"
    return '[日期已隐藏]'


def _mask_path(value: str) -> str:
    """文件路径脱敏：仅保留文件名"""
    if not value:
        return value
    import os
    try:
        return os.path.basename(str(value))
    except Exception:
        return '[路径已隐藏]'


_MASK_FORBIDDEN_MAP = {
    '银行账号': _mask_bank_account,
    '对方账号': _mask_bank_account,
    '付款': _mask_amount,
    '支出金额': _mask_amount,
    '收款': _mask_amount,
    '收入金额': _mask_amount,
    '余额': _mask_amount,
    '交易流水号': _mask_transaction_id,
    '匹配键': lambda v: '[匹配键已隐藏]',
    '交易日期': _mask_date,
    '交易时间': lambda v: '[时间已隐藏]',
    '附件路径': _mask_path,
    '对方行名': lambda v: '[对方行已隐藏]',
    '凭证种类': lambda v: '[凭证种类已隐藏]',
    '凭证号码': lambda v: '[凭证号已隐藏]',
    '手续费': _mask_amount,
    '利息': _mask_amount,
    '税费': _mask_amount,
    'amount': _mask_amount,
    'account': _mask_bank_account,
    'balance': _mask_amount,
    'payment': _mask_amount,
    'receipt': _mask_amount,
    'match_key': lambda v: '[match_key hidden]',
}

_MASK_DEBUG_MAP = {
    '主体': _mask_subject_name,
    '主体名称': _mask_subject_name,
    '对方户名': _mask_subject_name,
    '摘要': _mask_text,
    '交易描述': _mask_text,
    '备注': lambda v: _mask_text(v, keep_prefix=2),
    '唯一id': _mask_unique_id,
    '唯一ID': _mask_unique_id,
    'transaction_id': _mask_unique_id,
    'workflow_id': _mask_unique_id,
    'subject': _mask_subject_name,
    'counterparty': _mask_subject_name,
    'summary': _mask_text,
    'remark': lambda v: _mask_text(v, keep_prefix=2),
    'description': _mask_text,
    'handled_by': _mask_subject_name,
    'id': _mask_unique_id,
}

_FORBIDDEN_PLACEHOLDER = '[已脱敏-禁止落盘]'
_DEBUG_PLACEHOLDER = '[已脱敏]'


def mask_value(
    field_name: str,
    value: Any,
    target_level: PIILevel = PIILevel.INFO_SAFE,
) -> Any:
    """
    根据目标日志级别，对字段值进行脱敏处理。

    Args:
        field_name: 字段名称
        value: 字段原始值
        target_level: 目标日志级别（INFO_SAFE 最严格）

    Returns:
        脱敏后的值；如完全禁止则返回占位符
    """
    if value is None:
        return value

    field_level = classify_field(field_name)

    if field_level == PIILevel.INFO_SAFE:
        return value

    if field_level == PIILevel.DEBUG_ONLY:
        if target_level == PIILevel.DEBUG_ONLY:
            mask_fn = _MASK_DEBUG_MAP.get(field_name)
            if mask_fn is not None:
                try:
                    return mask_fn(str(value))
                except Exception:
                    return _DEBUG_PLACEHOLDER
            return _mask_text(str(value))
        else:
            mask_fn = _MASK_DEBUG_MAP.get(field_name)
            if mask_fn is not None:
                try:
                    return mask_fn(str(value))
                except Exception:
                    return _DEBUG_PLACEHOLDER
            return _DEBUG_PLACEHOLDER

    if field_level == PIILevel.FORBIDDEN:
        if target_level == PIILevel.DEBUG_ONLY:
            mask_fn = _MASK_FORBIDDEN_MAP.get(field_name)
            if mask_fn is not None:
                try:
                    return mask_fn(str(value) if value is not None else '')
                except Exception:
                    return _FORBIDDEN_PLACEHOLDER
            return _FORBIDDEN_PLACEHOLDER
        else:
            mask_fn = _MASK_FORBIDDEN_MAP.get(field_name)
            if mask_fn is not None:
                try:
                    return mask_fn(str(value) if value is not None else '')
                except Exception:
                    return _FORBIDDEN_PLACEHOLDER
            return _FORBIDDEN_PLACEHOLDER

    return value


def sanitize_dict_for_log(
    data: Dict[str, Any],
    target_level: PIILevel = PIILevel.INFO_SAFE,
    allowed_fields: Optional[List[str]] = None,
) -> Dict[str, Any]:
    """
    对字典结构的日志数据进行 PII 脱敏。

    Args:
        data: 原始数据字典（字段名 -> 字段值）
        target_level: 目标日志安全级别
        allowed_fields: 显式允许的字段名列表（不受分级限制）

    Returns:
        脱敏后的数据字典
    """
    if not isinstance(data, dict):
        return data

    sanitized: Dict[str, Any] = {}
    allowed = set(allowed_fields or [])

    for key, value in data.items():
        if key in allowed:
            sanitized[key] = value
            continue

        if isinstance(value, dict):
            sanitized[key] = sanitize_dict_for_log(value, target_level, allowed_fields)
        elif isinstance(value, list):
            sanitized[key] = [
                sanitize_dict_for_log(item, target_level, allowed_fields)
                if isinstance(item, dict) else item
                for item in value
            ]
        else:
            sanitized[key] = mask_value(key, value, target_level)

    return sanitized


def _looks_like_bank_account(value: str) -> bool:
    """启发式判断是否像银行账号：纯数字或字母数字组合，长度>=8"""
    if not value:
        return False
    s = str(value).strip().replace('-', '').replace(' ', '')
    if len(s) < 8:
        return False
    return bool(re.fullmatch(r'[A-Za-z0-9]+', s))


def _looks_like_amount(value: Any) -> bool:
    """
    启发式判断是否像金额。
    严格条件：避免将普通计数（记录数、字段数、行数）误判为金额。
    判断规则（满足任一）：
      1. 浮点型且不是整数（带小数）
      2. 数值 >= 1000 且不是 2/4/8/16... 等典型 2 的幂次方计数
      3. 字符串含千分号逗号 或 含小数点后两位（如 1,234.56）
      4. 负数（支出类金额）
    """
    try:
        if isinstance(value, bool):
            return False

        raw_str = str(value).strip() if not isinstance(value, (int, float)) else ''

        if isinstance(value, (int, float)):
            v = float(value)
            abs_v = abs(v)
            if v < 0:
                return True
            if abs_v != int(abs_v):
                return True
            if abs_v >= 10000:
                return True
            if 1000 <= abs_v < 10000 and not (abs_v in (1024, 2048, 4096, 8192)):
                return True
            return False

        s = raw_str.replace(',', '')
        if not s:
            return False

        if re.search(r'\.\d{2}$', raw_str):
            return True
        if ',' in raw_str and re.search(r'\d{1,3}(,\d{3})+', raw_str):
            return True

        v = abs(float(s))
        if v < 0:
            return True
        if v != int(v):
            return True
        if v >= 10000:
            return True
        if 1000 <= v < 10000 and not (v in (1024, 2048, 4096, 8192)):
            return True
        return False

    except (TypeError, ValueError):
        return False


def _looks_like_transaction_id(value: str) -> bool:
    """启发式判断是否像交易流水号：字母数字混合长度>=8"""
    if not value:
        return False
    s = str(value).strip()
    if len(s) < 8:
        return False
    return bool(re.search(r'[A-Za-z]', s)) and bool(re.search(r'\d', s))


def _looks_like_date(value: str) -> bool:
    """启发式判断是否像日期"""
    if not value:
        return False
    s = str(value).strip()
    patterns = [
        r'^\d{4}[-/.年]\d{1,2}[-/.月]\d{1,2}',
        r'^\d{4}\d{2}\d{2}',
        r'^\d{1,2}[-/.]\d{1,2}[-/.]\d{4}',
    ]
    return any(re.match(p, s) for p in patterns)


def _looks_like_unique_id(value: str) -> bool:
    """启发式判断是否像唯一ID：含短横线或长度>=12字母数字"""
    if not value:
        return False
    s = str(value).strip()
    if '-' in s and len(s) >= 10:
        return True
    if len(s) >= 12 and bool(re.fullmatch(r'[A-Za-z0-9]+', s)):
        return True
    return False


def _heuristic_mask_value(value: Any, target_level: PIILevel) -> Any:
    """
    当字段名缺失时，基于值的特征做启发式脱敏。
    宁可误脱敏，不可漏脱敏。
    """
    if value is None:
        return value

    if isinstance(value, (int, float)):
        if _looks_like_amount(value):
            return _mask_amount(value)
        return value

    if isinstance(value, str):
        s = value.strip()
        if not s:
            return value

        if _looks_like_date(s):
            return _mask_date(s)
        if _looks_like_transaction_id(s):
            return _mask_transaction_id(s)
        if _looks_like_bank_account(s):
            return _mask_bank_account(s)
        if _looks_like_unique_id(s):
            return _mask_unique_id(s)

        try:
            v = abs(float(s.replace(',', '')))
            if v > 10:
                return _mask_amount(s)
        except (TypeError, ValueError):
            pass

        return value

    return value


def _extract_field_keywords_from_msg(msg: str) -> List[str]:
    """
    从日志消息文本中提取与格式化占位符（%s, %d 等）对应的字段关键词。

    策略：
      1. 先找到所有格式化占位符及其位置
      2. 对每个占位符，在其前面 20 字符内寻找最近的 PII 关键词
      3. 关键词到占位符之间不能有其他占位符或数字（如"%d 个 ... %s"中，"个"之前的数字不相关）

    返回与占位符位置对应的字段名列表（长度=占位符数量）。
    占位符未匹配到关键词则为 None。
    """
    if not msg or not isinstance(msg, str):
        return []

    placeholder_pattern = re.compile(r'%(?:\([^)]+\))?[.\-0-9]*[sdfro]')
    placeholders = [(m.start(), m.end()) for m in placeholder_pattern.finditer(msg)]
    if not placeholders:
        return []

    keyword_patterns = [
        (r'(银行账号|对方账号|account)\s*[：:]?\s*$', '银行账号'),
        (r'(主体|主体名称|subject)\s*[：:]?\s*$', '主体'),
        (r'(对方户名|counterparty)\s*[：:]?\s*$', '对方户名'),
        (r'(付款|收款|支出金额|收入金额|金额|余额|payment|receipt|balance|amount)\s*[：:]?\s*$', '余额'),
        (r'(交易流水号|流水号|transaction_id)\s*[：:]?\s*$', '交易流水号'),
        (r'(匹配键|match_key)\s*[：:]?\s*$', '匹配键'),
        (r'(交易日期|交易时间|日期|时间)\s*[：:]?\s*$', '交易日期'),
        (r'(摘要|summary|description)\s*[：:]?\s*$', '摘要'),
        (r'(唯一id|唯一ID|^id)\s*[：:]?\s*$', '唯一id'),
        (r'(附件路径|路径|file|path)\s*[：:]?\s*$', '附件路径'),
        (r'(备注|remark)\s*[：:]?\s*$', '备注'),
    ]

    result: List[Optional[str]] = [None] * len(placeholders)
    for idx, (ph_start, _ph_end) in enumerate(placeholders):
        prev_end = placeholders[idx - 1][1] if idx > 0 else 0
        search_start = max(prev_end, ph_start - 40)
        context = msg[search_start:ph_start]

        for pat_re, field_name in keyword_patterns:
            if re.search(pat_re, context, re.IGNORECASE):
                result[idx] = field_name
                break

    return result


def sanitize_for_log(
    message: str,
    extra: Optional[Dict[str, Any]] = None,
    target_level: PIILevel = PIILevel.INFO_SAFE,
) -> Tuple[str, Dict[str, Any]]:
    """
    统一入口：对日志消息与 extra 字典进行脱敏。

    Args:
        message: 日志消息字符串
        extra: logging 的 extra 参数字典
        target_level: 目标安全级别

    Returns:
        (脱敏后的消息, 脱敏后的 extra 字典)
    """
    if extra is None:
        extra = {}

    sanitized_extra = sanitize_dict_for_log(extra, target_level)

    if not message:
        return message, sanitized_extra

    sanitized_msg = message

    has_format_placeholders = bool(re.search(r'%(?:\([^)]+\))?[.\-0-9]*[sdfro]', sanitized_msg))

    if not has_format_placeholders:
        inline_patterns = [
            (r'银行账号\s*[：:]\s*([\w\-]+)', lambda m: '银行账号: ' + _mask_bank_account(m.group(1))),
            (r'对方账号\s*[：:]\s*([\w\-]+)', lambda m: '对方账号: ' + _mask_bank_account(m.group(1))),
            (r'account\s*[=:]\s*([\w\-]+)', lambda m: 'account=' + _mask_bank_account(m.group(1)), re.IGNORECASE),
            (r'余额\s*[：:]\s*([\-]?\d[\d.,]*)', lambda m: '余额: ' + str(_mask_amount(m.group(1)))),
            (r'balance\s*[=:]\s*([\-]?\d[\d.,]*)', lambda m: 'balance=' + str(_mask_amount(m.group(1))), re.IGNORECASE),
            (r'付款\s*[：:]\s*([\-]?\d[\d.,]*)', lambda m: '付款: ' + str(_mask_amount(m.group(1)))),
            (r'收款\s*[：:]\s*([\-]?\d[\d.,]*)', lambda m: '收款: ' + str(_mask_amount(m.group(1)))),
            (r'金额\s*[：:]\s*([\-]?\d[\d.,]*)', lambda m: '金额: ' + str(_mask_amount(m.group(1)))),
            (r'transaction_id\s*[=:]\s*([A-Za-z0-9\-]+)', lambda m: 'transaction_id=' + _mask_transaction_id(m.group(1)), re.IGNORECASE),
            (r'交易流水号\s*[：:]\s*([A-Za-z0-9\-]+)', lambda m: '交易流水号: ' + _mask_transaction_id(m.group(1))),
            (r'交易日期\s*[：:]\s*(\d{4}[-/.年]?\d{1,2}[-/.月]?\d{0,2}\d*)', lambda m: '交易日期: ' + _mask_date(m.group(1))),
            (r'主体\s*[：:]\s*([^\s,，;；]+)', lambda m: '主体: ' + _mask_subject_name(m.group(1))),
            (r'对方户名\s*[：:]\s*([^\s,，;；]+)', lambda m: '对方户名: ' + _mask_subject_name(m.group(1))),
        ]

        for pat in inline_patterns:
            try:
                if len(pat) == 3:
                    p, r, flags = pat
                    sanitized_msg = re.sub(p, r, sanitized_msg, flags=flags)
                else:
                    p, r = pat
                    sanitized_msg = re.sub(p, r, sanitized_msg)
            except Exception:
                pass

    return sanitized_msg, sanitized_extra


class PIILogFilter(logging.Filter):
    """
    logging.Filter 实现：按日志级别自动对消息与 extra 进行 PII 脱敏。

    - WARNING 及以上（WARNING / ERROR / CRITICAL）：使用 INFO_SAFE 严格脱敏
    - INFO：使用 INFO_SAFE 严格脱敏
    - DEBUG：使用 DEBUG_ONLY 部分脱敏

    注意：FORBIDDEN 级别字段在 DEBUG 日志中也仅显示区间掩码，
    不输出精确值，确保即使 DEBUG 日志被泄露也无法还原核心数据。
    """

    def filter(self, record: logging.LogRecord) -> bool:
        try:
            target_level = (
                PIILevel.DEBUG_ONLY
                if record.levelno <= logging.DEBUG
                else PIILevel.INFO_SAFE
            )

            field_keywords = []
            if isinstance(record.msg, str):
                field_keywords = _extract_field_keywords_from_msg(record.msg)
                record.msg, _ = sanitize_for_log(record.msg, None, target_level)

            if isinstance(record.args, dict):
                sanitized_args = {}
                for k, v in record.args.items():
                    sanitized_args[k] = mask_value(str(k), v, target_level)
                record.args = tuple(sanitized_args.values())

            elif isinstance(record.args, (tuple, list)):
                new_args = []
                for idx, arg in enumerate(record.args):
                    if isinstance(arg, dict):
                        new_args.append(sanitize_dict_for_log(arg, target_level))
                    elif isinstance(arg, (list, tuple)):
                        new_args.append(type(arg)(
                            sanitize_dict_for_log(item, target_level)
                            if isinstance(item, dict)
                            else self._sanitize_primitive(
                                item,
                                field_keywords[idx] if idx < len(field_keywords) else None,
                                target_level
                            )
                            for item in arg
                        ))
                    else:
                        inferred_field = field_keywords[idx] if idx < len(field_keywords) else None
                        new_args.append(self._sanitize_primitive(arg, inferred_field, target_level))
                record.args = tuple(new_args)

        except Exception:
            pass

        return True

    @staticmethod
    def _sanitize_primitive(arg: Any, inferred_field: Optional[str], target_level: PIILevel) -> Any:
        if inferred_field:
            try:
                return mask_value(inferred_field, arg, target_level)
            except Exception:
                pass
        return _heuristic_mask_value(arg, target_level)


def build_safe_log_context(
    record_count: Optional[int] = None,
    bank_name: Optional[str] = None,
    sheet_name: Optional[str] = None,
    status: Optional[str] = None,
    month: Optional[str] = None,
    **extras: Any,
) -> Dict[str, Any]:
    """
    构建可安全写入 INFO 日志的上下文字典。

    推荐在调用 logger.info(..., extra=build_safe_log_context(...)) 时使用，
    避免手动拼接字段导致误写敏感数据。

    Args:
        record_count: 记录条数
        bank_name: 银行名称
        sheet_name: 工作表名
        status: 处理状态
        month: 月份（格式：YYYY-MM）
        **extras: 其他字段；会自动按 PII 分级进行脱敏

    Returns:
        可直接作为 logging extra 的安全字典
    """
    ctx: Dict[str, Any] = {}
    if record_count is not None:
        ctx['记录数'] = record_count
    if bank_name is not None:
        ctx['银行'] = bank_name
    if sheet_name is not None:
        ctx['工作表'] = sheet_name
    if status is not None:
        ctx['状态'] = status
    if month is not None:
        ctx['月份'] = month

    if extras:
        for k, v in extras.items():
            ctx[k] = mask_value(k, v, PIILevel.INFO_SAFE)

    return ctx


def setup_pii_aware_logging(
    logger_name: str = 'bankcheck',
    log_file: Optional[str] = None,
    console_level: int = logging.INFO,
    file_level: int = logging.DEBUG,
) -> logging.Logger:
    """
    初始化带 PII 脱敏的日志系统。

    对所有 handler 附加 PIILogFilter，确保无论级别如何，敏感字段都会被正确脱敏。

    Args:
        logger_name: logger 名称
        log_file: 日志文件路径（如为 None 则不添加 FileHandler）
        console_level: 控制台日志级别
        file_level: 文件日志级别

    Returns:
        配置好的 logger 实例
    """
    import os
    import sys

    logger = logging.getLogger(logger_name)
    logger.setLevel(logging.DEBUG)

    pii_filter = PIILogFilter()

    if logger.handlers:
        for h in logger.handlers:
            h.addFilter(pii_filter)
        return logger

    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(console_level)
    console_fmt = logging.Formatter(
        '[%(asctime)s] %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S',
    )
    console_handler.setFormatter(console_fmt)
    console_handler.addFilter(pii_filter)
    logger.addHandler(console_handler)

    if log_file:
        log_dir = os.path.dirname(log_file)
        if log_dir and not os.path.exists(log_dir):
            os.makedirs(log_dir, exist_ok=True)
        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setLevel(file_level)
        file_fmt = logging.Formatter(
            '[%(asctime)s] %(levelname)s [%(funcName)s] %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S',
        )
        file_handler.setFormatter(file_fmt)
        file_handler.addFilter(pii_filter)
        logger.addHandler(file_handler)

    return logger


def is_field_forbidden(field_name: str) -> bool:
    """判断字段是否属于 FORBIDDEN 级别（禁止落盘）"""
    return classify_field(field_name) == PIILevel.FORBIDDEN


def is_field_debug_only(field_name: str) -> bool:
    """判断字段是否属于 DEBUG_ONLY 级别"""
    return classify_field(field_name) == PIILevel.DEBUG_ONLY


def is_field_info_safe(field_name: str) -> bool:
    """判断字段是否属于 INFO_SAFE 级别"""
    return classify_field(field_name) == PIILevel.INFO_SAFE


def get_export_field_whitelist(level: PIILevel = PIILevel.DEBUG_ONLY) -> List[str]:
    """
    获取指定安全级别允许导出的字段白名单。

    用于导出场景：确保导出报告不包含超出授权级别的字段。

    Args:
        level: 允许的最高敏感级别（默认 DEBUG_ONLY，即不包含 FORBIDDEN）

    Returns:
        允许导出的字段名列表
    """
    level_order = {
        PIILevel.INFO_SAFE: 0,
        PIILevel.DEBUG_ONLY: 1,
        PIILevel.FORBIDDEN: 2,
    }
    max_level = level_order[level]
    return [
        field
        for field, field_level in _FIELD_CLASSIFICATION.items()
        if level_order[field_level] <= max_level
    ]
