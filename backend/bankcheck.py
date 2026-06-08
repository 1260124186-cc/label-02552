# -*- coding: utf-8 -*-
"""
银行流水检验工具 - bankcheck.py
功能：
  1. 用户选择一个文件夹
  2. 在同路径下复制一份，命名为"原名＋检验版"
  3. 对复制后的文件夹递归扫描所有 Excel 文件（.xlsx 和 .xls）
  4. 根据文件名前缀判断银行类型，提取指定列
  5. 生成唯一ID，合并输出总表
  6. 删除已处理的 Excel 文件，保留无法识别银行的文件

当前支持的银行：北京银行、东亚银行
"""

import os
import sys
import shutil
import uuid
import logging
import tempfile
import sqlite3
import json
import getpass
import hashlib
from dataclasses import dataclass, field
from datetime import datetime
from typing import List, Optional, Tuple, Dict, Any

import openpyxl
import pandas as pd

# ──────────────────────────────────────────────
# tkinter 兼容：尝试导入并安全测试，失败则回退命令行模式
# ──────────────────────────────────────────────

def _verify_tkinter():
    import subprocess
    import sys
    try:
        # 在子进程中测试 tk.Tk()。这样如果底层动态库版本不匹配发生 Abort trap: 6 崩溃，
        # 只会使子进程报错退出（返回非 0 状态码），不会导致主程序崩溃。
        code = "import tkinter; tkinter.Tk().destroy()"
        result = subprocess.run([sys.executable, '-c', code], capture_output=True, text=True, timeout=5)
        return result.returncode == 0
    except Exception:
        return False

HAS_TKINTER = False
try:
    import tkinter as tk
    from tkinter import filedialog, messagebox
    if _verify_tkinter():
        HAS_TKINTER = True
    else:
        tk = None
except (ImportError, ModuleNotFoundError):
    # 在无桌面或缺少 tk 的环境中，回退到命令行模式
    tk = None


def cli_askdirectory(title='请选择文件夹'):
    """命令行模式下让用户输入文件夹路径"""
    print(f'\n{title}')
    path = input('请输入文件夹路径: ').strip().strip('"').strip("'")
    if path and os.path.isdir(path):
        return path
    return ''


def cli_showinfo(title, message):
    """命令行模式下打印信息"""
    print(f'\n[{title}] {message}')


def cli_showwarning(title, message):
    """命令行模式下打印警告"""
    print(f'\n[警告 - {title}] {message}')


def cli_askfile(title='请选择文件'):
    """命令行模式下让用户输入文件路径"""
    print(f'\n{title}')
    path = input('请输入文件路径: ').strip().strip('"').strip("'")
    if path and os.path.isfile(path):
        return path
    return ''


def cli_askmode():
    """命令行模式下让用户选择运行模式"""
    print('\n请选择运行模式：')
    print('  1) 主流程：处理银行流水文件夹，输出总表')
    print('  2) 变更对比：对比两次总表的差异（新增/删除/变更）')
    choice = input('请输入选项（1 或 2，直接回车默认为 1）: ').strip()
    if choice == '2':
        return 'diff'
    return 'pipeline'


def gui_askdirectory(title='请选择银行流水文件夹'):
    """GUI 模式选择文件夹"""
    root = tk.Tk()
    root.withdraw()
    folder = filedialog.askdirectory(title=title)
    if not folder:
        messagebox.showinfo('提示', '未选择文件夹，程序退出。')
    root.destroy()
    return folder


def gui_showinfo(title, message):
    """GUI 模式弹出信息"""
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo(title, message)
    root.destroy()


def gui_showwarning(title, message):
    """GUI 模式弹出警告"""
    root = tk.Tk()
    root.withdraw()
    messagebox.showwarning(title, message)
    root.destroy()


def gui_askfile(title='请选择总表文件'):
    """GUI 模式选择文件"""
    root = tk.Tk()
    root.withdraw()
    filepath = filedialog.askopenfilename(
        title=title,
        filetypes=[('Excel 文件', '*.xlsx *.xls'), ('所有文件', '*.*')],
    )
    if not filepath:
        messagebox.showinfo('提示', '未选择文件，程序退出。')
    root.destroy()
    return filepath


def gui_askmode():
    """GUI 模式下让用户选择运行模式"""
    root = tk.Tk()
    root.withdraw()
    choice = messagebox.askyesnocancel(
        '选择运行模式',
        '是 = 主流程：处理流水文件夹，输出总表\n\n否 = 变更对比：对比两次总表的差异',
    )
    root.destroy()
    if choice is None:
        return None
    return 'pipeline' if choice else 'diff'


# 根据 tkinter 是否可用，选择交互方式
if HAS_TKINTER:
    ask_directory = gui_askdirectory
    show_info = gui_showinfo
    show_warning = gui_showwarning
    ask_file = gui_askfile
    ask_mode = gui_askmode
else:
    ask_directory = cli_askdirectory
    show_info = cli_showinfo
    show_warning = cli_showwarning
    ask_file = cli_askfile
    ask_mode = cli_askmode


# ──────────────────────────────────────────────
# 日志配置
# ──────────────────────────────────────────────

def get_script_dir():
    """获取脚本（或打包后的 exe）所在目录"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def setup_logging():
    """
    初始化日志系统。
    - 控制台输出 INFO 级别及以上日志
    - 日志文件（bankcheck.log）记录 DEBUG 级别及以上日志，
      文件保存在脚本/exe 所在目录下
    """
    log_dir = get_script_dir()
    log_file = os.path.join(log_dir, 'bankcheck.log')

    logger = logging.getLogger('bankcheck')
    logger.setLevel(logging.DEBUG)

    # 避免重复添加 handler
    if logger.handlers:
        return logger

    # 控制台 handler
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)
    console_fmt = logging.Formatter(
        '[%(asctime)s] %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S',
    )
    console_handler.setFormatter(console_fmt)

    # 文件 handler
    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    file_handler.setLevel(logging.DEBUG)
    file_fmt = logging.Formatter(
        '[%(asctime)s] %(levelname)s [%(funcName)s] %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S',
    )
    file_handler.setFormatter(file_fmt)

    logger.addHandler(console_handler)
    logger.addHandler(file_handler)

    logger.info('日志系统初始化完成，日志文件: %s', log_file)
    if not HAS_TKINTER:
        logger.info('未检测到 tkinter，将使用命令行交互模式')
    return logger


def get_logger():
    """获取名为 'bankcheck' 的全局 logger"""
    return logging.getLogger('bankcheck')


# ──────────────────────────────────────────────
# .xls 兼容：将 .xls 转换为 .xlsx
# ──────────────────────────────────────────────

def convert_xls_to_xlsx(xls_path):
    """
    将 .xls 文件转换为 .xlsx 文件（临时文件），返回临时 .xlsx 路径。
    使用 xlrd 读取 .xls，再用 openpyxl 写入 .xlsx。
    """
    logger = get_logger()
    try:
        import xlrd
    except ImportError:
        logger.error('处理 .xls 文件需要 xlrd 库，请运行: pip install xlrd')
        raise ImportError('缺少 xlrd 库，无法处理 .xls 文件。请运行: pip install xlrd')

    logger.info('正在将 .xls 文件转换为 .xlsx: %s', xls_path)

    xls_book = xlrd.open_workbook(xls_path)
    wb = openpyxl.Workbook()
    # 删除默认 sheet
    wb.remove(wb.active)

    for sheet_name in xls_book.sheet_names():
        xls_sheet = xls_book.sheet_by_name(sheet_name)
        ws = wb.create_sheet(title=sheet_name)
        for row_idx in range(xls_sheet.nrows):
            for col_idx in range(xls_sheet.ncols):
                cell_value = xls_sheet.cell_value(row_idx, col_idx)
                cell_type = xls_sheet.cell_type(row_idx, col_idx)
                # xlrd 日期类型转换
                if cell_type == xlrd.XL_CELL_DATE:
                    try:
                        date_tuple = xlrd.xldate_as_tuple(cell_value, xls_book.datemode)
                        cell_value = datetime(*date_tuple)
                    except Exception:
                        pass
                ws.cell(row=row_idx + 1, column=col_idx + 1, value=cell_value)

    # 保存为临时 .xlsx 文件
    tmp_dir = tempfile.gettempdir()
    base_name = os.path.splitext(os.path.basename(xls_path))[0]
    tmp_path = os.path.join(tmp_dir, f'{base_name}_converted.xlsx')
    wb.save(tmp_path)
    wb.close()
    xls_book.release_resources()

    logger.info('.xls 转换完成: %s -> %s', xls_path, tmp_path)
    return tmp_path


def open_workbook_compat(filepath):
    """
    兼容打开 .xlsx 和 .xls 文件，统一返回 (openpyxl.Workbook, 临时文件路径或None)。
    如果是 .xls 文件，先转换为 .xlsx 再打开。
    调用方负责在使用完毕后清理临时文件。
    """
    tmp_path = None
    if filepath.lower().endswith('.xls') and not filepath.lower().endswith('.xlsx'):
        tmp_path = convert_xls_to_xlsx(filepath)
        wb = openpyxl.load_workbook(tmp_path, data_only=True)
    else:
        wb = openpyxl.load_workbook(filepath, data_only=True)
    return wb, tmp_path


def cleanup_temp_file(tmp_path):
    """清理临时转换的 .xlsx 文件"""
    if tmp_path and os.path.exists(tmp_path):
        try:
            os.remove(tmp_path)
        except OSError:
            pass


# ──────────────────────────────────────────────
# 文件扫描
# ──────────────────────────────────────────────

def scan_excel_files(folder):
    """递归扫描文件夹中的所有 Excel 文件（.xlsx 和 .xls），排除临时文件"""
    logger = get_logger()
    excel_files = []
    for root, _dirs, files in os.walk(folder):
        for f in files:
            if f.startswith('~$'):
                continue
            if f.lower().endswith(('.xlsx', '.xls')):
                full_path = os.path.join(root, f)
                excel_files.append(full_path)
                logger.debug('发现 Excel 文件: %s', full_path)
    logger.info('共扫描到 %d 个 Excel 文件', len(excel_files))
    return excel_files


# ──────────────────────────────────────────────
# 银行识别
# ──────────────────────────────────────────────

BANK_PREFIXES = ['北京银行', '东亚银行']


def identify_bank(filepath):
    """根据文件名开头判断银行，返回银行名称或 None"""
    logger = get_logger()
    basename = os.path.basename(filepath)
    for prefix in BANK_PREFIXES:
        if basename.startswith(prefix):
            logger.info('文件「%s」识别为: %s', basename, prefix)
            return prefix
    logger.warning('文件「%s」无法识别银行类型', basename)
    return None


# ──────────────────────────────────────────────
# 主体查找
# ──────────────────────────────────────────────

# 主体查找表的推荐文件名（优先精确匹配）
LOOKUP_FILE_NAMES = ['主体查找表.xlsx', '主体查找表.xls']


def find_lookup_file(script_dir):
    """
    在脚本所在目录下查找主体查找表 Excel 文件。

    查找策略（按优先级）：
    1. 优先按文件名精确匹配 "主体查找表.xlsx" 或 "主体查找表.xls"
    2. 若未精确匹配到，回退到查找目录下唯一的 Excel 文件（排除输出总表和临时文件）
    """
    logger = get_logger()

    # ── 策略 1：按文件名精确匹配 ──
    for name in LOOKUP_FILE_NAMES:
        candidate = os.path.join(script_dir, name)
        if os.path.isfile(candidate):
            logger.info('精确匹配到主体查找表: %s', candidate)
            return candidate

    # ── 策略 2：回退到唯一 Excel 文件 ──
    excel_exts = ('.xlsx', '.xls')
    exclude_names = {'银行流水总表.xlsx', '银行流水总表.xls'}
    excel_files = []
    for f in os.listdir(script_dir):
        if f.startswith('~$'):
            continue
        if f in exclude_names:
            continue
        if f.lower().endswith(excel_exts):
            excel_files.append(os.path.join(script_dir, f))

    if len(excel_files) == 1:
        logger.info('找到主体查找表（唯一 Excel 文件）: %s', excel_files[0])
        return excel_files[0]
    elif len(excel_files) == 0:
        logger.warning('程序目录下未找到任何 Excel 文件作为主体查找表')
    else:
        logger.warning(
            '程序目录下存在 %d 个 Excel 文件，无法确定唯一查找表: %s。'
            '建议将查找表文件命名为 "主体查找表.xlsx"',
            len(excel_files),
            [os.path.basename(f) for f in excel_files],
        )
    return None


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


def get_subject(bank_account, lookup_file):
    """
    根据银行账号在查找表中找到对应的主体。
    查找表中 B 列为银行账号，取同一行 A 列的值作为主体。
    """
    logger = get_logger()

    if not lookup_file or not os.path.exists(lookup_file):
        logger.warning('主体查找表不存在或未指定，银行账号「%s」的主体将为空', bank_account)
        return ''
    if bank_account is None:
        logger.warning('银行账号为空，无法查找主体')
        return ''

    target_key = _account_key(bank_account)
    tmp_path = None
    try:
        wb, tmp_path = open_workbook_compat(lookup_file)
        ws = wb.active
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=2, max_col=2):
            cell = row[0]
            if cell.value is not None and _account_key(cell.value) == target_key:
                subject = ws.cell(row=cell.row, column=1).value
                wb.close()
                cleanup_temp_file(tmp_path)
                logger.debug('银行账号「%s」匹配到主体: %s', bank_account, subject)
                return subject if subject else ''
        wb.close()
        logger.warning('银行账号「%s」在查找表中未找到对应主体', bank_account)
    except Exception as e:
        logger.error('读取主体查找表「%s」时发生错误: %s', lookup_file, e, exc_info=True)
    finally:
        cleanup_temp_file(tmp_path)
    return ''


# ──────────────────────────────────────────────
# 唯一 ID
# ──────────────────────────────────────────────

def generate_unique_id():
    """生成唯一 ID：当前时间戳 + UUID"""
    timestamp = datetime.now().strftime('%Y%m%d%H%M%S%f')
    uid = uuid.uuid4().hex
    return f"{timestamp}{uid}"


# ──────────────────────────────────────────────
# 数值工具
# ──────────────────────────────────────────────

def is_numeric(value):
    """判断值是否为数字（int / float / 可转换的字符串）"""
    if value is None:
        return False
    if isinstance(value, (int, float)):
        return True
    try:
        float(str(value).strip())
        return True
    except (ValueError, TypeError):
        return False


def to_float(value):
    """安全地将值转为 float"""
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).strip())
    except (ValueError, TypeError):
        return None


# ──────────────────────────────────────────────
# 银行处理器
# ──────────────────────────────────────────────

def process_beijing_bank(filepath, lookup_file):
    """
    处理北京银行流水 Excel 文件
    ─────────────────────────────────
    银行账号：B2
    交易日期：B 列，第 4 行起
    付款：D 列第 4 行起，有数字则取负
    收款：E 列第 4 行起，有数字则直接填入
    摘要：L 列第 4 行起
    对方户名：G 列第 4 行起
    余额：F 列第 4 行起
    交易流水号：P 列第 4 行起
    """
    logger = get_logger()
    logger.info('开始处理北京银行文件: %s', filepath)

    wb, tmp_path = open_workbook_compat(filepath)
    try:
        ws = wb.active

        bank_account = ws['B2'].value
        if bank_account is None:
            logger.warning('文件「%s」B2 单元格为空，银行账号缺失', filepath)

        subject = get_subject(bank_account, lookup_file)

        rows = []
        start_row = 4
        for row_idx in range(start_row, ws.max_row + 1):
            trade_date = ws.cell(row=row_idx, column=2).value  # B 列
            if trade_date is None:
                continue

            payment_val = ws.cell(row=row_idx, column=4).value
            payment = -abs(to_float(payment_val)) if is_numeric(payment_val) else None

            receipt_val = ws.cell(row=row_idx, column=5).value
            receipt = to_float(receipt_val) if is_numeric(receipt_val) else None

            summary = ws.cell(row=row_idx, column=12).value       # L 列
            counterpart = ws.cell(row=row_idx, column=7).value     # G 列
            balance = ws.cell(row=row_idx, column=6).value         # F 列
            transaction_id = ws.cell(row=row_idx, column=16).value # P 列

            rows.append({
                '唯一id': generate_unique_id(),
                '银行': '北京银行',
                '银行账号': bank_account,
                '主体': subject,
                '交易日期': trade_date,
                '付款': payment,
                '收款': receipt,
                '摘要': summary,
                '对方户名': counterpart,
                '余额': balance,
                '交易流水号': transaction_id,
            })

        wb.close()
        logger.info('北京银行文件处理完成，提取 %d 条记录', len(rows))
        return rows
    finally:
        cleanup_temp_file(tmp_path)


def process_east_asia_bank(filepath, lookup_file):
    """
    处理东亚银行流水 Excel 文件
    ─────────────────────────────────
    银行账号：B1
    交易日期：A 列，第 5 行起
    付款：D 列第 5 行起，有数字则取负
    收款：E 列第 5 行起，有数字则直接填入
    摘要：L 列第 5 行起
    对方户名：L 列第 5 行起
    余额：I 列第 5 行起
    交易流水号：K 列第 5 行起
    """
    logger = get_logger()
    logger.info('开始处理东亚银行文件: %s', filepath)

    wb, tmp_path = open_workbook_compat(filepath)
    try:
        ws = wb.active

        bank_account = ws['B1'].value
        if bank_account is None:
            logger.warning('文件「%s」B1 单元格为空，银行账号缺失', filepath)

        subject = get_subject(bank_account, lookup_file)

        rows = []
        start_row = 5
        for row_idx in range(start_row, ws.max_row + 1):
            trade_date = ws.cell(row=row_idx, column=1).value  # A 列
            if trade_date is None:
                continue

            payment_val = ws.cell(row=row_idx, column=4).value
            payment = -abs(to_float(payment_val)) if is_numeric(payment_val) else None

            receipt_val = ws.cell(row=row_idx, column=5).value
            receipt = to_float(receipt_val) if is_numeric(receipt_val) else None

            summary = ws.cell(row=row_idx, column=12).value       # L 列
            counterpart = ws.cell(row=row_idx, column=12).value    # L 列（同摘要）
            balance = ws.cell(row=row_idx, column=9).value         # I 列
            transaction_id = ws.cell(row=row_idx, column=11).value # K 列

            rows.append({
                '唯一id': generate_unique_id(),
                '银行': '东亚银行',
                '银行账号': bank_account,
                '主体': subject,
                '交易日期': trade_date,
                '付款': payment,
                '收款': receipt,
                '摘要': summary,
                '对方户名': counterpart,
                '余额': balance,
                '交易流水号': transaction_id,
            })

        wb.close()
        logger.info('东亚银行文件处理完成，提取 %d 条记录', len(rows))
        return rows
    finally:
        cleanup_temp_file(tmp_path)


# 银行处理器注册表（方便后续扩展新银行）
BANK_PROCESSORS = {
    '北京银行': process_beijing_bank,
    '东亚银行': process_east_asia_bank,
}


# ──────────────────────────────────────────────
# 主流程
# ──────────────────────────────────────────────

@dataclass
class ProcessingResult:
    all_rows: List[dict] = field(default_factory=list)
    processed_files: List[str] = field(default_factory=list)
    unprocessed_files: List[str] = field(default_factory=list)
    error_files: List[Tuple[str, str]] = field(default_factory=list)
    output_path: Optional[str] = None
    lookup_missing: bool = False
    folder_empty: bool = False


def run_pipeline(folder, script_dir):
    logger = get_logger()

    lookup_file = find_lookup_file(script_dir)
    lookup_missing = lookup_file is None
    if lookup_missing:
        logger.warning('未找到主体查找表，"主体"列将为空')

    folder_name = os.path.basename(folder.rstrip('/\\'))
    parent_dir = os.path.dirname(folder.rstrip('/\\'))
    new_folder = os.path.join(parent_dir, f"{folder_name}＋检验版")

    if os.path.exists(new_folder):
        logger.info('＋检验版文件夹已存在，先删除: %s', new_folder)
        shutil.rmtree(new_folder)
    shutil.copytree(folder, new_folder)
    logger.info('已复制文件夹为＋检验版: %s', new_folder)

    excel_files = scan_excel_files(new_folder)
    if not excel_files:
        logger.warning('检验版文件夹中未发现任何 Excel 文件')
        return ProcessingResult(lookup_missing=lookup_missing, folder_empty=True)

    all_rows = []
    processed_files = []
    unprocessed_files = []
    error_files = []

    for filepath in excel_files:
        bank = identify_bank(filepath)
        if bank and bank in BANK_PROCESSORS:
            try:
                processor = BANK_PROCESSORS[bank]
                rows = processor(filepath, lookup_file)
                all_rows.extend(rows)
                processed_files.append(filepath)
                logger.info('成功处理文件: %s（%d 条记录）', filepath, len(rows))
            except Exception as e:
                error_files.append((filepath, str(e)))
                logger.error('处理文件「%s」时发生错误: %s', filepath, e, exc_info=True)
        else:
            unprocessed_files.append(filepath)

    error_file_paths = {f for f, _ in error_files}
    delete_processed_files(excel_files, set(unprocessed_files) | error_file_paths)

    output_path = None
    if all_rows:
        columns = [
            '唯一id', '银行', '银行账号', '主体', '交易日期',
            '付款', '收款', '摘要', '对方户名', '余额', '交易流水号',
        ]
        df = pd.DataFrame(all_rows, columns=columns)
        output_path = os.path.join(script_dir, '银行流水总表.xlsx')
        df.to_excel(output_path, index=False, engine='openpyxl')
        logger.info('总表输出完成: %s（共 %d 条记录）', output_path, len(all_rows))
    else:
        logger.warning('未提取到任何银行流水记录')

    return ProcessingResult(
        all_rows=all_rows,
        processed_files=processed_files,
        unprocessed_files=unprocessed_files,
        error_files=error_files,
        output_path=output_path,
        lookup_missing=lookup_missing,
    )


def format_result_message(result):
    if result.folder_empty:
        return '文件夹中未发现任何 Excel 文件。'

    if result.all_rows:
        msg = (
            f'处理完成！\n\n'
            f'已处理文件数：{len(result.processed_files)}\n'
            f'提取记录数：{len(result.all_rows)}\n'
            f'总表路径：{result.output_path}'
        )
    else:
        msg = '未提取到任何银行流水记录。'

    if result.unprocessed_files:
        names = '\n  '.join(os.path.basename(f) for f in result.unprocessed_files)
        msg += f'\n\n无法识别的文件（{len(result.unprocessed_files)} 个，已保留）：\n  {names}'
    if result.error_files:
        err_info = '\n  '.join(f'{os.path.basename(f)}: {e}' for f, e in result.error_files)
        msg += f'\n\n处理出错的文件（{len(result.error_files)} 个，已保留）：\n  {err_info}'

    return msg


def delete_processed_files(excel_files, keep_set):
    logger = get_logger()
    for filepath in excel_files:
        if filepath not in keep_set:
            try:
                os.remove(filepath)
                logger.debug('已删除文件: %s', filepath)
            except OSError as e:
                logger.error('删除文件「%s」失败: %s', filepath, e)


# ──────────────────────────────────────────────
# 流水文件变更对比
# ──────────────────────────────────────────────

DIFF_COLUMNS = [
    '银行', '银行账号', '主体', '交易日期',
    '付款', '收款', '摘要', '对方户名', '余额', '交易流水号',
]

AMOUNT_FIELDS = ['付款', '收款', '余额']


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


def _is_nan_or_none(value):
    if value is None:
        return True
    if isinstance(value, float) and value != value:
        return True
    return False


def _values_equal(val_a, val_b, field_name):
    if field_name in AMOUNT_FIELDS:
        fa = to_float(val_a)
        fb = to_float(val_b)
        if _is_nan_or_none(fa) and _is_nan_or_none(fb):
            return True
        if _is_nan_or_none(fa) or _is_nan_or_none(fb):
            return False
        return abs(fa - fb) < 0.005
    sa = '' if _is_nan_or_none(val_a) else str(val_a).strip()
    sb = '' if _is_nan_or_none(val_b) else str(val_b).strip()
    return sa == sb


@dataclass
class DiffRecord:
    change_type: str
    match_key: str
    old_row: Optional[dict] = None
    new_row: Optional[dict] = None
    changed_fields: Optional[List[str]] = None

    def to_flat_dict(self):
        source = self.new_row if self.new_row else self.old_row
        result = {'变更类型': self.change_type}
        for col in DIFF_COLUMNS:
            result[col] = source.get(col) if source else None
        if self.change_type == '变更' and self.changed_fields:
            parts = []
            for f in self.changed_fields:
                old_v = self.old_row.get(f) if self.old_row else None
                new_v = self.new_row.get(f) if self.new_row else None
                parts.append(f"{f}: {old_v} → {new_v}")
            result['变更明细'] = '; '.join(parts)
        else:
            result['变更明细'] = ''
        if self.change_type == '变更':
            for f in (self.changed_fields or []):
                if f in AMOUNT_FIELDS:
                    old_v = self.old_row.get(f) if self.old_row else None
                    new_v = self.new_row.get(f) if self.new_row else None
                    result[f'{f}(旧)'] = old_v
                    result[f'{f}(新)'] = new_v
        return result


@dataclass
class DiffResult:
    records: List[DiffRecord] = field(default_factory=list)
    added_count: int = 0
    deleted_count: int = 0
    changed_count: int = 0
    unchanged_count: int = 0
    output_path: Optional[str] = None


def diff_transactions(old_df, new_df):
    logger = get_logger()

    old_keyed = {}
    for _, row in old_df.iterrows():
        key = _make_match_key(row)
        old_keyed[key] = row.to_dict()

    new_keyed = {}
    for _, row in new_df.iterrows():
        key = _make_match_key(row)
        new_keyed[key] = row.to_dict()

    all_keys = list(dict.fromkeys(list(old_keyed.keys()) + list(new_keyed.keys())))

    records = []
    added = deleted = changed = unchanged = 0

    for key in all_keys:
        in_old = key in old_keyed
        in_new = key in new_keyed

        if in_old and not in_new:
            records.append(DiffRecord(
                change_type='删除', match_key=key, old_row=old_keyed[key],
            ))
            deleted += 1
        elif in_new and not in_old:
            records.append(DiffRecord(
                change_type='新增', match_key=key, new_row=new_keyed[key],
            ))
            added += 1
        else:
            old_r = old_keyed[key]
            new_r = new_keyed[key]
            changed_fields = []
            for col in DIFF_COLUMNS:
                if not _values_equal(old_r.get(col), new_r.get(col), col):
                    changed_fields.append(col)
            if changed_fields:
                records.append(DiffRecord(
                    change_type='变更', match_key=key,
                    old_row=old_r, new_row=new_r,
                    changed_fields=changed_fields,
                ))
                changed += 1
            else:
                records.append(DiffRecord(
                    change_type='未变更', match_key=key,
                    old_row=old_r, new_row=new_r,
                ))
                unchanged += 1

    logger.info(
        '变更对比完成: 新增 %d, 删除 %d, 变更 %d, 未变更 %d',
        added, deleted, changed, unchanged,
    )
    return DiffResult(
        records=records,
        added_count=added,
        deleted_count=deleted,
        changed_count=changed,
        unchanged_count=unchanged,
    )


DIFF_HIGHLIGHT_COLORS = {
    '新增': 'C6EFCE',
    '删除': 'FFC7CE',
    '变更': 'FFEB9C',
}


def export_diff_result(diff_result, output_path):
    logger = get_logger()

    flat_rows = [r.to_flat_dict() for r in diff_result.records]
    if not flat_rows:
        logger.warning('无可对比结果，不生成输出文件')
        return None

    has_changes = any(r['变更明细'] for r in flat_rows)
    extra_cols = []
    if has_changes:
        extra_cols.append('变更明细')
    amount_change_cols = []
    for f in AMOUNT_FIELDS:
        col_old = f'{f}(旧)'
        col_new = f'{f}(新)'
        if any(col_old in r for r in flat_rows):
            amount_change_cols.extend([col_old, col_new])
    extra_cols.extend(amount_change_cols)

    columns = ['变更类型'] + DIFF_COLUMNS + extra_cols

    df = pd.DataFrame(flat_rows)
    for col in columns:
        if col not in df.columns:
            df[col] = None
    df = df[columns]

    df.to_excel(output_path, index=False, engine='openpyxl')

    wb = openpyxl.load_workbook(output_path)
    ws = wb.active

    type_col_idx = 1
    for row_idx in range(2, ws.max_row + 1):
        cell_type = ws.cell(row=row_idx, column=type_col_idx)
        change_type = str(cell_type.value).strip() if cell_type.value else ''
        fill_color = DIFF_HIGHLIGHT_COLORS.get(change_type)
        if fill_color:
            fill = openpyxl.styles.PatternFill(
                start_color=fill_color, end_color=fill_color, fill_type='solid',
            )
            for col_idx in range(1, ws.max_column + 1):
                ws.cell(row=row_idx, column=col_idx).fill = fill

    wb.save(output_path)
    wb.close()

    logger.info('变更对比结果已输出: %s', output_path)
    return output_path


def run_diff(old_path, new_path, output_dir=None):
    logger = get_logger()
    logger.info('========== 流水变更对比开始 ==========')
    logger.info('旧批次文件: %s', old_path)
    logger.info('新批次文件: %s', new_path)

    if not os.path.exists(old_path):
        logger.error('旧批次总表文件不存在: %s', old_path)
        raise FileNotFoundError(f'旧批次总表文件不存在: {old_path}')
    if not os.path.exists(new_path):
        logger.error('新批次总表文件不存在: %s', new_path)
        raise FileNotFoundError(f'新批次总表文件不存在: {new_path}')

    old_df = pd.read_excel(old_path, engine='openpyxl')
    new_df = pd.read_excel(new_path, engine='openpyxl')

    required_cols = ['银行账号', '交易流水号']
    for col in required_cols:
        if col not in old_df.columns:
            logger.warning('旧批次总表缺少列: %s', col)
        if col not in new_df.columns:
            logger.warning('新批次总表缺少列: %s', col)

    diff_result = diff_transactions(old_df, new_df)

    if output_dir is None:
        output_dir = get_script_dir()
    os.makedirs(output_dir, exist_ok=True)

    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_filename = f'流水变更对比_{timestamp}.xlsx'
    output_path = os.path.join(output_dir, output_filename)

    export_diff_result(diff_result, output_path)
    diff_result.output_path = output_path

    logger.info('========== 流水变更对比结束 ==========')
    return diff_result


def format_diff_message(diff_result):
    if not diff_result.records:
        return '两份总表无数据可对比。'

    total = len(diff_result.records)
    msg = (
        f'变更对比完成！\n\n'
        f'总记录数：{total}\n'
        f'新增交易：{diff_result.added_count}\n'
        f'删除交易：{diff_result.deleted_count}\n'
        f'金额/内容变更：{diff_result.changed_count}\n'
        f'未变更：{diff_result.unchanged_count}'
    )
    if diff_result.output_path:
        msg += f'\n\n对比结果文件：{diff_result.output_path}'
    return msg


def run_pipeline_flow(script_dir):
    """主流程：处理银行流水文件夹，输出总表"""
    logger = get_logger()

    folder = ask_directory('请选择银行流水文件夹')
    if not folder:
        show_info('提示', '未选择文件夹，程序退出。')
        logger.info('用户未选择文件夹，程序退出')
        return

    logger.info('用户选择文件夹: %s', folder)

    result = run_pipeline(folder, script_dir)

    if result.lookup_missing:
        show_warning(
            '警告',
            '在程序所在目录下未找到主体查找表文件，\n"主体"列将为空。\n'
            '建议将查找表文件命名为"主体查找表.xlsx"并放在程序所在目录下。'
        )

    msg = format_result_message(result)
    show_info('完成' if result.all_rows else '提示', msg)


def run_diff_flow(script_dir):
    """变更对比流程：选择两次总表，输出对比结果"""
    logger = get_logger()

    old_path = ask_file('请选择【旧批次】银行流水总表')
    if not old_path:
        show_info('提示', '未选择旧批次文件，程序退出。')
        logger.info('用户未选择旧批次文件，程序退出')
        return
    logger.info('用户选择旧批次文件: %s', old_path)

    new_path = ask_file('请选择【新批次】银行流水总表')
    if not new_path:
        show_info('提示', '未选择新批次文件，程序退出。')
        logger.info('用户未选择新批次文件，程序退出')
        return
    logger.info('用户选择新批次文件: %s', new_path)

    try:
        diff_result = run_diff(old_path, new_path, script_dir)
        msg = format_diff_message(diff_result)

        has_changes = (
            diff_result.added_count > 0
            or diff_result.deleted_count > 0
            or diff_result.changed_count > 0
        )
        title = '对比完成（发现差异' if has_changes else '对比完成（无差异）'
        show_info(title, msg)
    except FileNotFoundError as e:
        show_warning('错误', str(e))
        logger.error('对比失败: %s', e)


# ──────────────────────────────────────────────
# 多用户与操作审计模块
# ──────────────────────────────────────────────

AUDIT_DB_FILENAME = 'audit_log.db'


def get_audit_db_path(script_dir=None):
    """获取审计数据库文件路径"""
    if script_dir is None:
        script_dir = get_script_dir()
    return os.path.join(script_dir, AUDIT_DB_FILENAME)


def init_audit_db(db_path=None):
    """
    初始化审计数据库，创建所需表结构。
    包含三张核心表：
    - users: 用户信息表
    - audit_logs: 操作审计主表
    - config_changes: 配置变更历史表
    """
    if db_path is None:
        db_path = get_audit_db_path()

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL UNIQUE,
            display_name TEXT,
            role TEXT DEFAULT 'operator',
            created_at TEXT NOT NULL,
            last_login TEXT,
            is_active INTEGER DEFAULT 1
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS audit_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            audit_id TEXT NOT NULL UNIQUE,
            user_id INTEGER,
            username TEXT NOT NULL,
            operation_type TEXT NOT NULL,
            input_directory TEXT,
            output_path TEXT,
            processed_files INTEGER DEFAULT 0,
            extracted_records INTEGER DEFAULT 0,
            unprocessed_files INTEGER DEFAULT 0,
            error_files INTEGER DEFAULT 0,
            lookup_file TEXT,
            lookup_missing INTEGER DEFAULT 0,
            status TEXT NOT NULL,
            error_message TEXT,
            config_snapshot TEXT,
            input_files_hash TEXT,
            output_hash TEXT,
            started_at TEXT NOT NULL,
            completed_at TEXT,
            duration_ms INTEGER,
            client_ip TEXT,
            client_hostname TEXT,
            FOREIGN KEY (user_id) REFERENCES users (id)
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS config_changes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            change_id TEXT NOT NULL UNIQUE,
            user_id INTEGER,
            username TEXT NOT NULL,
            config_type TEXT NOT NULL,
            config_name TEXT NOT NULL,
            old_value TEXT,
            new_value TEXT,
            old_hash TEXT,
            new_hash TEXT,
            change_reason TEXT,
            changed_at TEXT NOT NULL,
            FOREIGN KEY (user_id) REFERENCES users (id)
        )
    ''')

    cursor.execute('CREATE INDEX IF NOT EXISTS idx_audit_logs_username ON audit_logs(username)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_audit_logs_operation ON audit_logs(operation_type)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_audit_logs_started_at ON audit_logs(started_at)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_config_changes_config ON config_changes(config_type, config_name)')

    conn.commit()
    conn.close()

    logger = get_logger()
    logger.info('审计数据库初始化完成: %s', db_path)


def get_current_user():
    """
    获取当前操作用户。
    优先级：环境变量 BANKCHECK_USER > 系统登录用户 > 'unknown'
    """
    env_user = os.environ.get('BANKCHECK_USER', '').strip()
    if env_user:
        return env_user
    try:
        return getpass.getuser()
    except Exception:
        return 'unknown'


def get_client_info():
    """获取客户端信息（主机名、IP等）"""
    import socket
    hostname = ''
    ip = ''
    try:
        hostname = socket.gethostname()
        ip = socket.gethostbyname(hostname)
    except Exception:
        pass
    return {'hostname': hostname, 'ip': ip}


def _ensure_user(username, db_path=None):
    """确保用户存在于数据库中，不存在则创建"""
    if db_path is None:
        db_path = get_audit_db_path()

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    cursor.execute('SELECT id FROM users WHERE username = ?', (username,))
    row = cursor.fetchone()

    if row:
        user_id = row[0]
        cursor.execute(
            'UPDATE users SET last_login = ? WHERE id = ?',
            (datetime.now().strftime('%Y-%m-%d %H:%M:%S'), user_id)
        )
    else:
        now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        cursor.execute(
            'INSERT INTO users (username, display_name, created_at, last_login) VALUES (?, ?, ?, ?)',
            (username, username, now, now)
        )
        user_id = cursor.lastrowid

    conn.commit()
    conn.close()
    return user_id


def compute_file_hash(filepath):
    """计算文件的 SHA256 哈希值，用于完整性校验"""
    if not filepath or not os.path.exists(filepath):
        return None
    sha256 = hashlib.sha256()
    try:
        with open(filepath, 'rb') as f:
            for chunk in iter(lambda: f.read(8192), b''):
                sha256.update(chunk)
        return sha256.hexdigest()
    except Exception:
        return None


def compute_config_snapshot(script_dir):
    """
    生成当前配置快照，包含：
    - 支持的银行列表
    - 银行处理器配置
    - 查找表文件信息
    """
    lookup_file = find_lookup_file(script_dir)
    snapshot = {
        'supported_banks': BANK_PREFIXES,
        'bank_processors': list(BANK_PROCESSORS.keys()),
        'lookup_file': lookup_file,
        'lookup_file_hash': compute_file_hash(lookup_file) if lookup_file else None,
        'diff_columns': DIFF_COLUMNS,
        'amount_fields': AMOUNT_FIELDS,
        'timestamp': datetime.now().isoformat(),
    }
    return json.dumps(snapshot, ensure_ascii=False)


@dataclass
class AuditRecord:
    """审计记录数据类"""
    audit_id: str
    username: str
    operation_type: str
    input_directory: Optional[str] = None
    output_path: Optional[str] = None
    processed_files: int = 0
    extracted_records: int = 0
    unprocessed_files: int = 0
    error_files: int = 0
    lookup_file: Optional[str] = None
    lookup_missing: bool = False
    status: str = 'running'
    error_message: Optional[str] = None
    config_snapshot: Optional[str] = None
    input_files_hash: Optional[str] = None
    output_hash: Optional[str] = None
    started_at: Optional[str] = None
    completed_at: Optional[str] = None
    duration_ms: Optional[int] = None
    client_ip: Optional[str] = None
    client_hostname: Optional[str] = None


class AuditLogger:
    """
    审计日志核心类，提供上下文管理器支持。
    使用方式：
        with AuditLogger('pipeline', script_dir) as audit:
            audit.record_input(folder)
            ... 执行操作 ...
            audit.record_result(result)
    """

    def __init__(self, operation_type, script_dir=None, username=None):
        self.script_dir = script_dir or get_script_dir()
        self.db_path = get_audit_db_path(self.script_dir)
        init_audit_db(self.db_path)

        self.username = username or get_current_user()
        self.user_id = _ensure_user(self.username, self.db_path)

        self.audit_id = f"AUD{datetime.now().strftime('%Y%m%d%H%M%S')}{uuid.uuid4().hex[:8].upper()}"
        self.operation_type = operation_type

        client_info = get_client_info()

        self.record = AuditRecord(
            audit_id=self.audit_id,
            username=self.username,
            operation_type=operation_type,
            status='running',
            started_at=datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f'),
            client_ip=client_info.get('ip'),
            client_hostname=client_info.get('hostname'),
            config_snapshot=compute_config_snapshot(self.script_dir),
        )

        self.logger = get_logger()
        self._start_time = datetime.now()
        self._save_record()
        self.logger.info('审计记录已创建 [%s] 操作: %s, 用户: %s',
                         self.audit_id, operation_type, self.username)

    def _save_record(self):
        """保存审计记录到数据库"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        cursor.execute('SELECT id FROM audit_logs WHERE audit_id = ?', (self.audit_id,))
        exists = cursor.fetchone()

        if exists:
            cursor.execute('''
                UPDATE audit_logs SET
                    input_directory = ?, output_path = ?,
                    processed_files = ?, extracted_records = ?,
                    unprocessed_files = ?, error_files = ?,
                    lookup_file = ?, lookup_missing = ?,
                    status = ?, error_message = ?,
                    config_snapshot = ?, input_files_hash = ?,
                    output_hash = ?, started_at = ?,
                    completed_at = ?, duration_ms = ?,
                    client_ip = ?, client_hostname = ?
                WHERE audit_id = ?
            ''', (
                self.record.input_directory, self.record.output_path,
                self.record.processed_files, self.record.extracted_records,
                self.record.unprocessed_files, self.record.error_files,
                self.record.lookup_file, 1 if self.record.lookup_missing else 0,
                self.record.status, self.record.error_message,
                self.record.config_snapshot, self.record.input_files_hash,
                self.record.output_hash, self.record.started_at,
                self.record.completed_at, self.record.duration_ms,
                self.record.client_ip, self.record.client_hostname,
                self.audit_id,
            ))
        else:
            cursor.execute('''
                INSERT INTO audit_logs (
                    audit_id, user_id, username, operation_type,
                    input_directory, output_path,
                    processed_files, extracted_records,
                    unprocessed_files, error_files,
                    lookup_file, lookup_missing,
                    status, error_message,
                    config_snapshot, input_files_hash,
                    output_hash, started_at,
                    completed_at, duration_ms,
                    client_ip, client_hostname
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                self.audit_id, self.user_id, self.username, self.operation_type,
                self.record.input_directory, self.record.output_path,
                self.record.processed_files, self.record.extracted_records,
                self.record.unprocessed_files, self.record.error_files,
                self.record.lookup_file, 1 if self.record.lookup_missing else 0,
                self.record.status, self.record.error_message,
                self.record.config_snapshot, self.record.input_files_hash,
                self.record.output_hash, self.record.started_at,
                self.record.completed_at, self.record.duration_ms,
                self.record.client_ip, self.record.client_hostname,
            ))

        conn.commit()
        conn.close()

    def record_input(self, input_directory, lookup_file=None):
        """记录输入信息"""
        self.record.input_directory = input_directory
        if lookup_file:
            self.record.lookup_file = lookup_file
        elif self.record.input_directory:
            self.record.lookup_file = find_lookup_file(self.script_dir)
        self._save_record()
        self.logger.debug('审计记录 [%s] 已记录输入目录: %s', self.audit_id, input_directory)

    def record_result(self, result):
        """
        根据 ProcessingResult 或 DiffResult 记录处理结果
        """
        if isinstance(result, ProcessingResult):
            self.record.processed_files = len(result.processed_files)
            self.record.extracted_records = len(result.all_rows)
            self.record.unprocessed_files = len(result.unprocessed_files)
            self.record.error_files = len(result.error_files)
            self.record.output_path = result.output_path
            self.record.lookup_missing = result.lookup_missing
            if result.output_path:
                self.record.output_hash = compute_file_hash(result.output_path)
        elif isinstance(result, DiffResult):
            self.record.extracted_records = result.added_count + result.deleted_count + result.changed_count
            self.record.output_path = result.output_path
            if result.output_path:
                self.record.output_hash = compute_file_hash(result.output_path)

        self._save_record()

    def record_success(self):
        """标记操作成功完成"""
        self.record.status = 'success'
        self.record.completed_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')
        duration = (datetime.now() - self._start_time).total_seconds() * 1000
        self.record.duration_ms = int(duration)
        self._save_record()
        self.logger.info('审计记录 [%s] 操作成功完成，耗时 %d ms', self.audit_id, self.record.duration_ms)

    def record_failure(self, error_message):
        """标记操作失败"""
        self.record.status = 'failed'
        self.record.error_message = str(error_message)
        self.record.completed_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')
        duration = (datetime.now() - self._start_time).total_seconds() * 1000
        self.record.duration_ms = int(duration)
        self._save_record()
        self.logger.error('审计记录 [%s] 操作失败: %s', self.audit_id, error_message)

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if exc_type is not None:
            self.record_failure(f'{exc_type.__name__}: {exc_val}')
        elif self.record.status == 'running':
            self.record_success()
        return False


def record_config_change(config_type, config_name, old_value, new_value,
                         change_reason='', script_dir=None, username=None):
    """
    记录配置变更历史。
    适用于查找表更新、银行配置调整等场景。
    """
    if script_dir is None:
        script_dir = get_script_dir()
    db_path = get_audit_db_path(script_dir)
    init_audit_db(db_path)

    username = username or get_current_user()
    user_id = _ensure_user(username, db_path)

    change_id = f"CFG{datetime.now().strftime('%Y%m%d%H%M%S')}{uuid.uuid4().hex[:6].upper()}"
    changed_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')

    old_hash = hashlib.sha256(str(old_value or '').encode('utf-8')).hexdigest()
    new_hash = hashlib.sha256(str(new_value or '').encode('utf-8')).hexdigest()

    old_str = json.dumps(old_value, ensure_ascii=False) if isinstance(old_value, (dict, list)) else str(old_value)
    new_str = json.dumps(new_value, ensure_ascii=False) if isinstance(new_value, (dict, list)) else str(new_value)

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO config_changes (
            change_id, user_id, username, config_type, config_name,
            old_value, new_value, old_hash, new_hash, change_reason, changed_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (
        change_id, user_id, username, config_type, config_name,
        old_str, new_str, old_hash, new_hash, change_reason, changed_at
    ))
    conn.commit()
    conn.close()

    logger = get_logger()
    logger.info('配置变更已记录 [%s] %s.%s 由 %s 修改',
                change_id, config_type, config_name, username)

    return change_id


def query_audit_logs(script_dir=None, username=None, operation_type=None,
                     start_date=None, end_date=None, status=None, limit=100):
    """
    查询审计日志记录。

    Args:
        script_dir: 脚本目录
        username: 按用户名过滤
        operation_type: 按操作类型过滤
        start_date: 开始日期 (YYYY-MM-DD)
        end_date: 结束日期 (YYYY-MM-DD)
        status: 按状态过滤
        limit: 返回记录数限制

    Returns:
        List[Dict] 审计记录列表
    """
    if script_dir is None:
        script_dir = get_script_dir()
    db_path = get_audit_db_path(script_dir)

    if not os.path.exists(db_path):
        return []

    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    query = 'SELECT * FROM audit_logs WHERE 1=1'
    params = []

    if username:
        query += ' AND username = ?'
        params.append(username)
    if operation_type:
        query += ' AND operation_type = ?'
        params.append(operation_type)
    if start_date:
        query += ' AND date(started_at) >= date(?)'
        params.append(start_date)
    if end_date:
        query += ' AND date(started_at) <= date(?)'
        params.append(end_date)
    if status:
        query += ' AND status = ?'
        params.append(status)

    query += ' ORDER BY started_at DESC LIMIT ?'
    params.append(limit)

    cursor.execute(query, params)
    rows = cursor.fetchall()
    conn.close()

    return [dict(row) for row in rows]


def query_config_changes(script_dir=None, config_type=None, config_name=None,
                         username=None, limit=100):
    """
    查询配置变更历史。
    """
    if script_dir is None:
        script_dir = get_script_dir()
    db_path = get_audit_db_path(script_dir)

    if not os.path.exists(db_path):
        return []

    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    query = 'SELECT * FROM config_changes WHERE 1=1'
    params = []

    if config_type:
        query += ' AND config_type = ?'
        params.append(config_type)
    if config_name:
        query += ' AND config_name = ?'
        params.append(config_name)
    if username:
        query += ' AND username = ?'
        params.append(username)

    query += ' ORDER BY changed_at DESC LIMIT ?'
    params.append(limit)

    cursor.execute(query, params)
    rows = cursor.fetchall()
    conn.close()

    return [dict(row) for row in rows]


def export_audit_logs(output_path, script_dir=None, **query_kwargs):
    """
    导出审计日志到 Excel 文件。
    """
    records = query_audit_logs(script_dir=script_dir, **query_kwargs)
    if not records:
        return None

    df = pd.DataFrame(records)
    columns_order = [
        'audit_id', 'username', 'operation_type', 'status',
        'input_directory', 'output_path',
        'processed_files', 'extracted_records',
        'unprocessed_files', 'error_files',
        'lookup_file', 'lookup_missing',
        'started_at', 'completed_at', 'duration_ms',
        'client_ip', 'client_hostname',
        'error_message', 'config_snapshot',
        'input_files_hash', 'output_hash', 'id'
    ]
    for col in columns_order:
        if col not in df.columns:
            df[col] = None
    df = df[columns_order]

    df.to_excel(output_path, index=False, engine='openpyxl')
    logger = get_logger()
    logger.info('审计日志已导出到: %s (共 %d 条记录)', output_path, len(records))
    return output_path


def export_config_changes(output_path, script_dir=None, **query_kwargs):
    """
    导出配置变更历史到 Excel 文件。
    """
    records = query_config_changes(script_dir=script_dir, **query_kwargs)
    if not records:
        return None

    df = pd.DataFrame(records)
    columns_order = [
        'change_id', 'username', 'config_type', 'config_name',
        'old_value', 'new_value', 'change_reason',
        'old_hash', 'new_hash', 'changed_at', 'user_id', 'id'
    ]
    for col in columns_order:
        if col not in df.columns:
            df[col] = None
    df = df[columns_order]

    df.to_excel(output_path, index=False, engine='openpyxl')
    logger = get_logger()
    logger.info('配置变更历史已导出到: %s (共 %d 条记录)', output_path, len(records))
    return output_path


def run_pipeline_flow(script_dir):
    """主流程：处理银行流水文件夹，输出总表"""
    logger = get_logger()

    folder = ask_directory('请选择银行流水文件夹')
    if not folder:
        show_info('提示', '未选择文件夹，程序退出。')
        logger.info('用户未选择文件夹，程序退出')
        return

    logger.info('用户选择文件夹: %s', folder)

    with AuditLogger('pipeline', script_dir) as audit:
        audit.record_input(folder)

        result = run_pipeline(folder, script_dir)
        audit.record_result(result)

        if result.lookup_missing:
            show_warning(
                '警告',
                '在程序所在目录下未找到主体查找表文件，\n"主体"列将为空。\n'
                '建议将查找表文件命名为"主体查找表.xlsx"并放在程序所在目录下。'
            )

        msg = format_result_message(result)
        msg += f'\n\n审计编号: {audit.audit_id}'
        show_info('完成' if result.all_rows else '提示', msg)


def run_diff_flow(script_dir):
    """变更对比流程：选择两次总表，输出对比结果"""
    logger = get_logger()

    old_path = ask_file('请选择【旧批次】银行流水总表')
    if not old_path:
        show_info('提示', '未选择旧批次文件，程序退出。')
        logger.info('用户未选择旧批次文件，程序退出')
        return
    logger.info('用户选择旧批次文件: %s', old_path)

    new_path = ask_file('请选择【新批次】银行流水总表')
    if not new_path:
        show_info('提示', '未选择新批次文件，程序退出。')
        logger.info('用户未选择新批次文件，程序退出')
        return
    logger.info('用户选择新批次文件: %s', new_path)

    with AuditLogger('diff', script_dir) as audit:
        audit.record_input(f'旧:{old_path} | 新:{new_path}')

        try:
            diff_result = run_diff(old_path, new_path, script_dir)
            audit.record_result(diff_result)

            msg = format_diff_message(diff_result)

            has_changes = (
                diff_result.added_count > 0
                or diff_result.deleted_count > 0
                or diff_result.changed_count > 0
            )
            title = '对比完成（发现差异' if has_changes else '对比完成（无差异）'
            msg += f'\n\n审计编号: {audit.audit_id}'
            show_info(title, msg)
        except FileNotFoundError as e:
            show_warning('错误', str(e))
            logger.error('对比失败: %s', e)


def main():
    setup_logging()
    logger = get_logger()
    logger.info('========== 银行流水检验工具启动 ==========')

    script_dir = get_script_dir()

    init_audit_db(get_audit_db_path(script_dir))
    logger.info('当前操作用户: %s', get_current_user())

    mode = ask_mode()
    if mode is None:
        show_info('提示', '未选择模式，程序退出。')
        logger.info('用户未选择模式，程序退出')
        return
    logger.info('用户选择模式: %s', mode)

    if mode == 'pipeline':
        run_pipeline_flow(script_dir)
    elif mode == 'diff':
        run_diff_flow(script_dir)

    logger.info('========== 银行流水检验工具运行结束 ==========')


if __name__ == '__main__':
    main()
