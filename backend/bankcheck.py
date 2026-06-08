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

try:
    import database as db_module
    HAS_DATABASE = True
except ImportError:
    HAS_DATABASE = False
    db_module = None

try:
    import batch_manager as batch_module
    HAS_BATCH_MANAGER = True
except ImportError:
    HAS_BATCH_MANAGER = False
    batch_module = None

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


def cli_askmode():
    """命令行模式下让用户选择运行模式"""
    print('\n请选择运行模式：')
    print('  1) 主流程：处理银行流水文件夹，输出总表')
    print('  2) 变更对比：对比两次总表的差异（新增/删除/变更）')
    print('  3) 监控面板：运行监控与告警管理')
    print('  4) 定时调度：定时批处理调度管理')
    print('  5) 财务导出：按用友/金蝶等模板导出凭证或日记账')
    print('  6) 数据库查询：按主体/账号/时间范围查询流水记录')
    print('  7) 数据库统计：查看数据汇总统计信息')
    print('  8) 批次管理：查看历史批次与版本回溯')
    print('  9) 预设管理：管理任务配置预设，一键加载常用方案')
    choice = input('请输入选项（1-9，直接回车默认为 1）: ').strip()
    if choice == '2':
        return 'diff'
    elif choice == '3':
        return 'monitor'
    elif choice == '4':
        return 'scheduler'
    elif choice == '5':
        return 'export'
    elif choice == '6':
        return 'db_query'
    elif choice == '7':
        return 'db_stats'
    elif choice == '8':
        return 'batch_history'
    elif choice == '9':
        return 'preset'
    return 'pipeline'


def gui_askmode():
    """GUI 模式下让用户选择运行模式"""
    if tk is None:
        return cli_askmode()

    try:
        result = _gui_askmode_full()
        if result is not None:
            return result
    except Exception:
        pass

    try:
        root = tk.Tk()
        root.withdraw()
        choice = messagebox.askyesnocancel(
            '选择运行模式',
            '是 = 主流程：处理流水文件夹，输出总表\n\n否 = 变更对比：对比两次总表的差异\n\n取消 = 财务导出：按用友/金蝶等模板导出\n\n提示：\n- 使用命令行参数 --scheduler-menu 可进入定时调度管理\n- 使用命令行参数 --export 可直接进入财务导出\n- 使用命令行参数 --monitor 可进入监控面板\n- 使用命令行参数 --preset-menu 可进入预设管理',
        )
        root.destroy()
        if choice is None:
            return 'export'
        return 'pipeline' if choice else 'diff'
    except Exception:
        return cli_askmode()


def _gui_askmode_full():
    """完整的GUI模式选择界面"""
    if not hasattr(tk, 'Tk') or not callable(tk.Tk):
        raise RuntimeError('Tk not available')

    root = tk.Tk()
    if not hasattr(root, 'mainloop') or not callable(root.mainloop):
        root.destroy()
        raise RuntimeError('Mock Tk detected')

    root.title('银行流水检验工具 - 选择功能')
    root.geometry('480x520')
    root.resizable(False, False)

    result = {'mode': None}

    def select_mode(mode):
        result['mode'] = mode
        root.destroy()

    tk.Label(root, text='请选择运行模式', font=('Arial', 16, 'bold')).pack(pady=20)

    button_frame = tk.Frame(root)
    button_frame.pack(pady=10)

    modes = [
        ('主流程', 'pipeline', '处理银行流水文件夹，输出总表', '#4CAF50'),
        ('变更对比', 'diff', '对比两次总表的差异', '#2196F3'),
        ('监控面板', 'monitor', '运行监控与告警管理', '#FF9800'),
        ('定时调度', 'scheduler', '定时批处理调度管理', '#9C27B0'),
        ('财务导出', 'export', '按用友/金蝶等模板导出', '#607D8B'),
        ('数据库查询', 'db_query', '按条件查询流水记录', '#00BCD4'),
        ('数据库统计', 'db_stats', '查看数据汇总统计', '#795548'),
        ('批次管理', 'batch_history', '历史批次与版本回溯', '#E91E63'),
        ('预设管理', 'preset', '管理任务配置预设', '#FF5722'),
    ]

    for i, (name, mode, desc, color) in enumerate(modes):
        row = i // 3
        col = i % 3
        btn = tk.Button(
            button_frame,
            text=name,
            width=14,
            height=3,
            bg=color,
            fg='white',
            font=('Arial', 11, 'bold'),
            command=lambda m=mode: select_mode(m),
        )
        btn.grid(row=row, column=col, padx=8, pady=8)
        tk.Label(button_frame, text=desc, font=('Arial', 8), fg='#666').grid(
            row=row * 2 + 1, column=col, padx=8, pady=(0, 8)
        )

    tk.Button(
        root,
        text='退出',
        width=12,
        command=lambda: select_mode(None),
        bg='#f44336',
        fg='white',
        font=('Arial', 10, 'bold'),
    ).pack(pady=20)

    root.mainloop()
    return result['mode']


def _ask_monitor_or_scheduler():
    """二级菜单：选择监控或调度"""
    if tk is not None:
        try:
            root = tk.Tk()
            root.withdraw()
            choice = messagebox.askyesnocancel(
                '选择功能',
                '是 = 监控面板：运行监控与告警管理\n\n否 = 定时调度：定时批处理调度管理\n\n取消 = 返回主菜单',
            )
            root.destroy()
            if choice is None:
                return 'monitor'
            return 'monitor' if choice else 'scheduler'
        except Exception:
            pass

    print('\n请选择功能：')
    print('  1) 监控面板：运行监控与告警管理')
    print('  2) 定时调度：定时批处理调度管理')
    print('  3) 返回主菜单')
    choice = input('请输入选项: ').strip()
    if choice == '1':
        return 'monitor'
    elif choice == '2':
        return 'scheduler'
    return 'monitor'


def cli_ask_monitor_export_menu():
    """二级菜单：选择监控/调度/财务导出"""
    print('\n请选择功能：')
    print('  1) 监控面板：运行监控与告警管理')
    print('  2) 定时调度：定时批处理调度管理')
    print('  3) 财务导出：按用友/金蝶等模板导出凭证或日记账')
    print('  4) 返回主菜单')
    choice = input('请输入选项: ').strip()
    if choice == '1':
        return 'monitor'
    elif choice == '2':
        return 'scheduler'
    elif choice == '3':
        return 'export'
    return 'monitor'


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
    incremental_mode: bool = False
    existing_record_count: int = 0
    new_record_count: int = 0
    duplicate_record_count: int = 0
    db_inserted_count: int = 0
    db_duplicate_count: int = 0


# ──────────────────────────────────────────────
# 增量合并模块
# ──────────────────────────────────────────────

SUMMARY_TABLE_FILENAME = '银行流水总表.xlsx'


def get_summary_table_path(script_dir):
    """获取历史总表文件路径"""
    return os.path.join(script_dir, SUMMARY_TABLE_FILENAME)


def load_existing_keys(summary_path):
    """
    读取历史总表，提取所有记录的匹配键集合。
    用于快速检测新记录是否已存在。

    Args:
        summary_path: 历史总表文件路径

    Returns:
        tuple: (existing_keys_set, existing_records_list)
            - existing_keys_set: 匹配键集合，用于 O(1) 复杂度的存在性检测
            - existing_records_list: 历史记录列表，用于合并输出
    """
    logger = get_logger()

    if not summary_path or not os.path.exists(summary_path):
        logger.info('未找到历史总表，将以全量模式运行')
        return set(), []

    try:
        df = pd.read_excel(summary_path, engine='openpyxl')
        if df.empty:
            logger.info('历史总表为空，将以全量模式运行')
            return set(), []

        required_cols = ['银行账号', '交易流水号', '交易日期', '付款', '收款']
        for col in required_cols:
            if col not in df.columns:
                logger.warning('历史总表缺少必要列「%s」，无法进行增量检测，将以全量模式运行', col)
                return set(), []

        existing_keys = set()
        existing_records = []

        for _, row in df.iterrows():
            row_dict = row.to_dict()
            key = _make_match_key(row_dict)
            existing_keys.add(key)
            existing_records.append(row_dict)

        logger.info('已加载历史总表，共 %d 条记录，%d 个唯一匹配键',
                    len(existing_records), len(existing_keys))
        return existing_keys, existing_records

    except Exception as e:
        logger.error('读取历史总表失败: %s，将以全量模式运行', e, exc_info=True)
        return set(), []


def filter_incremental_records(new_rows, existing_keys):
    """
    过滤新提取的记录，只返回增量（未在历史总表中出现过的）记录。

    Args:
        new_rows: 新提取的记录列表
        existing_keys: 历史总表的匹配键集合

    Returns:
        tuple: (incremental_rows, duplicate_count)
            - incremental_rows: 增量记录列表
            - duplicate_count: 检测到的重复记录数量
    """
    logger = get_logger()

    if not existing_keys:
        logger.info('无历史记录，所有 %d 条记录均为新增', len(new_rows))
        return new_rows, 0

    incremental_rows = []
    duplicate_count = 0

    for row in new_rows:
        key = _make_match_key(row)
        if key in existing_keys:
            duplicate_count += 1
            logger.debug('检测到重复记录，匹配键: %s', key)
        else:
            incremental_rows.append(row)
            existing_keys.add(key)

    logger.info('增量过滤完成: 新记录 %d 条，重复 %d 条，实际新增 %d 条',
                len(new_rows), duplicate_count, len(incremental_rows))
    return incremental_rows, duplicate_count


def merge_and_export_summary(existing_records, incremental_rows, script_dir):
    """
    合并历史记录与增量记录，并输出到总表。

    Args:
        existing_records: 历史记录列表
        incremental_rows: 新增记录列表
        script_dir: 脚本目录

    Returns:
        str: 输出文件路径
    """
    logger = get_logger()

    columns = [
        '唯一id', '银行', '银行账号', '主体', '交易日期',
        '付款', '收款', '摘要', '对方户名', '余额', '交易流水号',
    ]

    merged_records = existing_records + incremental_rows

    if not merged_records:
        logger.warning('无任何记录可输出')
        return None

    df = pd.DataFrame(merged_records, columns=columns)
    output_path = get_summary_table_path(script_dir)
    df.to_excel(output_path, index=False, engine='openpyxl')

    logger.info('总表输出完成: %s（历史 %d 条 + 新增 %d 条 = 共 %d 条）',
                output_path, len(existing_records), len(incremental_rows), len(merged_records))
    return output_path


def cli_ask_incremental_mode():
    """命令行模式下询问用户是否启用增量合并"""
    print('\n请选择运行模式：')
    print('  1) 增量合并（推荐）：检测历史记录，仅追加新增数据')
    print('  2) 全量覆盖：重新生成总表，覆盖历史数据')
    choice = input('请输入选项（直接回车默认为 1 增量模式）: ').strip()
    return choice != '2'


def gui_ask_incremental_mode():
    """GUI 模式下询问用户是否启用增量合并"""
    root = tk.Tk()
    root.withdraw()
    choice = messagebox.askyesnocancel(
        '选择运行模式',
        '是 = 增量合并（推荐）：检测历史记录，仅追加新增数据\n\n'
        '否 = 全量覆盖：重新生成总表，覆盖历史数据\n\n'
        '取消 = 返回主菜单',
    )
    root.destroy()
    if choice is None:
        return None
    return choice


if HAS_TKINTER:
    ask_incremental_mode = gui_ask_incremental_mode
else:
    ask_incremental_mode = cli_ask_incremental_mode


def run_pipeline(folder, script_dir, incremental=True, batch_id=None):
    logger = get_logger()

    lookup_file = find_lookup_file(script_dir)
    lookup_missing = lookup_file is None
    if lookup_missing:
        logger.warning('未找到主体查找表，"主体"列将为空')

    existing_keys = set()
    existing_records = []
    actual_incremental = False
    duplicate_count = 0
    new_record_count = 0

    if incremental:
        summary_path = get_summary_table_path(script_dir)
        existing_keys, existing_records = load_existing_keys(summary_path)
        actual_incremental = len(existing_records) > 0
        if actual_incremental:
            logger.info('===== 增量合并模式已启用 =====')
        else:
            logger.info('无历史数据，将以全量模式运行')

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
        return ProcessingResult(
            lookup_missing=lookup_missing,
            folder_empty=True,
            incremental_mode=actual_incremental,
            existing_record_count=len(existing_records),
        )

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
    final_rows = []

    if all_rows:
        if actual_incremental:
            incremental_rows, duplicate_count = filter_incremental_records(all_rows, existing_keys)
            new_record_count = len(incremental_rows)
            output_path = merge_and_export_summary(existing_records, incremental_rows, script_dir)
            final_rows = existing_records + incremental_rows
        else:
            columns = [
                '唯一id', '银行', '银行账号', '主体', '交易日期',
                '付款', '收款', '摘要', '对方户名', '余额', '交易流水号',
            ]
            df = pd.DataFrame(all_rows, columns=columns)
            output_path = get_summary_table_path(script_dir)
            df.to_excel(output_path, index=False, engine='openpyxl')
            logger.info('总表输出完成: %s（共 %d 条记录）', output_path, len(all_rows))
            final_rows = all_rows
            new_record_count = len(all_rows)
    else:
        logger.warning('未提取到任何银行流水记录')
        if existing_records:
            output_path = merge_and_export_summary(existing_records, [], script_dir)
            final_rows = existing_records

    db_inserted = 0
    db_duplicates = 0
    if HAS_DATABASE and final_rows:
        try:
            if batch_id is None:
                batch_id = f"BATCH{datetime.now().strftime('%Y%m%d%H%M%S')}"
            db_inserted, db_duplicates = db_module.persist_transactions(
                final_rows,
                batch_id=batch_id,
                deduplicate=True,
                script_dir=script_dir,
            )
            logger.info(
                '数据库持久化完成: 批次 %s, 插入 %d 条, 去重跳过 %d 条',
                batch_id, db_inserted, db_duplicates,
            )
        except Exception as e:
            logger.error('数据库持久化失败: %s', e, exc_info=True)

    return ProcessingResult(
        all_rows=final_rows,
        processed_files=processed_files,
        unprocessed_files=unprocessed_files,
        error_files=error_files,
        output_path=output_path,
        lookup_missing=lookup_missing,
        incremental_mode=actual_incremental,
        existing_record_count=len(existing_records),
        new_record_count=new_record_count,
        duplicate_record_count=duplicate_count,
        db_inserted_count=db_inserted,
        db_duplicate_count=db_duplicates,
    )


def format_result_message(result):
    if result.folder_empty:
        return '文件夹中未发现任何 Excel 文件。'

    if result.all_rows:
        if result.incremental_mode:
            msg = (
                f'增量合并处理完成！\n\n'
                f'运行模式：增量合并\n'
                f'已处理文件数：{len(result.processed_files)}\n'
                f'历史总记录数：{result.existing_record_count}\n'
                f'本次新提取记录数：{result.new_record_count + result.duplicate_record_count}\n'
                f'├─ 重复记录（已跳过）：{result.duplicate_record_count}\n'
                f'└─ 新增记录（已追加）：{result.new_record_count}\n'
                f'总表当前总记录数：{len(result.all_rows)}\n'
                f'总表路径：{result.output_path}'
            )
        else:
            msg = (
                f'处理完成！\n\n'
                f'运行模式：全量覆盖\n'
                f'已处理文件数：{len(result.processed_files)}\n'
                f'提取记录数：{len(result.all_rows)}\n'
                f'总表路径：{result.output_path}'
            )

        if HAS_DATABASE and (result.db_inserted_count > 0 or result.db_duplicate_count > 0):
            msg += (
                f'\n\n数据库持久化：\n'
                f'├─ 新增入库：{result.db_inserted_count} 条\n'
                f'└─ 重复跳过：{result.db_duplicate_count} 条'
            )
    else:
        if result.incremental_mode and result.existing_record_count > 0:
            msg = (
                f'本次未提取到任何新增银行流水记录。\n\n'
                f'运行模式：增量合并\n'
                f'历史记录保留：{result.existing_record_count} 条\n'
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

    incremental = ask_incremental_mode()
    if incremental is None:
        logger.info('用户取消增量模式选择，返回主菜单')
        return
    logger.info('用户选择运行模式: %s', '增量合并' if incremental else '全量覆盖')

    result = run_pipeline(folder, script_dir, incremental=incremental)

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
    包含四张核心表：
    - users: 用户信息表
    - audit_logs: 操作审计主表
    - config_changes: 配置变更历史表
    - lookup_snapshots: 查找表快照表（用于自动检测变更）
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
            change_details TEXT,
            FOREIGN KEY (user_id) REFERENCES users (id)
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS lookup_snapshots (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            snapshot_id TEXT NOT NULL UNIQUE,
            lookup_file TEXT NOT NULL,
            file_hash TEXT NOT NULL,
            content_json TEXT NOT NULL,
            created_at TEXT NOT NULL,
            created_by TEXT NOT NULL
        )
    ''')

    cursor.execute('CREATE INDEX IF NOT EXISTS idx_audit_logs_username ON audit_logs(username)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_audit_logs_operation ON audit_logs(operation_type)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_audit_logs_started_at ON audit_logs(started_at)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_config_changes_config ON config_changes(config_type, config_name)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_lookup_snapshots_file ON lookup_snapshots(lookup_file)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_lookup_snapshots_created ON lookup_snapshots(created_at)')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS alert_rules (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            rule_id TEXT NOT NULL UNIQUE,
            rule_name TEXT NOT NULL,
            rule_type TEXT NOT NULL,
            metric TEXT NOT NULL,
            operator TEXT NOT NULL,
            threshold REAL NOT NULL,
            window_minutes INTEGER DEFAULT 60,
            severity TEXT DEFAULT 'warning',
            enabled INTEGER DEFAULT 1,
            description TEXT,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            created_by TEXT NOT NULL
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS alerts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            alert_id TEXT NOT NULL UNIQUE,
            rule_id TEXT NOT NULL,
            rule_name TEXT NOT NULL,
            severity TEXT NOT NULL,
            alert_type TEXT NOT NULL,
            metric TEXT NOT NULL,
            current_value REAL,
            threshold REAL,
            operator TEXT,
            window_minutes INTEGER,
            description TEXT,
            triggered_at TEXT NOT NULL,
            resolved_at TEXT,
            status TEXT DEFAULT 'active',
            acknowledged INTEGER DEFAULT 0,
            acknowledged_by TEXT,
            acknowledged_at TEXT,
            audit_log_id TEXT,
            details TEXT,
            FOREIGN KEY (rule_id) REFERENCES alert_rules (rule_id)
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS file_processing_details (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            detail_id TEXT NOT NULL UNIQUE,
            audit_id TEXT NOT NULL,
            file_path TEXT NOT NULL,
            file_name TEXT NOT NULL,
            bank TEXT,
            file_size INTEGER,
            record_count INTEGER DEFAULT 0,
            status TEXT NOT NULL,
            error_message TEXT,
            processing_duration_ms INTEGER,
            started_at TEXT NOT NULL,
            completed_at TEXT,
            FOREIGN KEY (audit_id) REFERENCES audit_logs (audit_id)
        )
    ''')

    cursor.execute('CREATE INDEX IF NOT EXISTS idx_alert_rules_type ON alert_rules(rule_type)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_alert_rules_enabled ON alert_rules(enabled)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_alerts_status ON alerts(status)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_alerts_severity ON alerts(severity)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_alerts_triggered ON alerts(triggered_at)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_alerts_rule ON alerts(rule_id)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_file_details_audit ON file_processing_details(audit_id)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_file_details_status ON file_processing_details(status)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_file_details_bank ON file_processing_details(bank)')

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

    def set_extra_info(self, extra_info):
        """记录额外信息到config_snapshot中"""
        try:
            if self.record.config_snapshot:
                snapshot = json.loads(self.record.config_snapshot)
            else:
                snapshot = {}
            snapshot.update(extra_info)
            self.record.config_snapshot = json.dumps(snapshot, ensure_ascii=False)
            self._save_record()
            self.logger.debug('审计记录 [%s] 已添加额外信息: %s', self.audit_id, extra_info)
        except Exception as e:
            self.logger.warning('添加额外信息失败: %s', e)

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
                         change_reason='', script_dir=None, username=None,
                         change_details=None):
    """
    记录配置变更历史。
    适用于查找表更新、银行配置调整等场景。

    Args:
        config_type: 配置类型（如 'lookup_table', 'bank_config'）
        config_name: 配置名称（如 '主体查找表.xlsx'）
        old_value: 变更前的值
        new_value: 变更后的值
        change_reason: 变更原因说明
        script_dir: 脚本目录
        username: 操作用户（默认自动获取）
        change_details: 详细变更内容（字典或列表，描述具体变更项）

    Returns:
        change_id: 变更记录ID
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

    details_str = None
    if change_details is not None:
        details_str = json.dumps(change_details, ensure_ascii=False)

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO config_changes (
            change_id, user_id, username, config_type, config_name,
            old_value, new_value, old_hash, new_hash, change_reason, changed_at,
            change_details
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (
        change_id, user_id, username, config_type, config_name,
        old_str, new_str, old_hash, new_hash, change_reason, changed_at,
        details_str
    ))
    conn.commit()
    conn.close()

    logger = get_logger()
    if change_details:
        change_count = len(change_details) if isinstance(change_details, list) else 1
        logger.info('配置变更已记录 [%s] %s.%s 由 %s 修改，%d 项变更',
                    change_id, config_type, config_name, username, change_count)
    else:
        logger.info('配置变更已记录 [%s] %s.%s 由 %s 修改',
                    change_id, config_type, config_name, username)

    return change_id


def read_lookup_table_content(lookup_file):
    """
    读取查找表的完整内容（单元格级），返回结构化数据用于比对。

    返回格式：
    {
        'file_path': '...',
        'file_hash': '...',
        'sheet_name': '...',
        'headers': ['主体名称', '银行账号'],
        'rows': [
            {'row_num': 2, '主体名称': 'XX公司', '银行账号': '12345'},
            {'row_num': 3, '主体名称': 'YY公司', '银行账号': '67890'}
        ],
        'account_map': {'12345': {'row_num': 2, '主体名称': 'XX公司'}}
    }
    """
    logger = get_logger()
    if not lookup_file or not os.path.exists(lookup_file):
        return None

    file_hash = compute_file_hash(lookup_file)
    tmp_path = None
    try:
        wb, tmp_path = open_workbook_compat(lookup_file)
        ws = wb.active

        headers = []
        for col in range(1, ws.max_column + 1):
            val = ws.cell(row=1, column=col).value
            headers.append(str(val) if val is not None else f'列{col}')

        rows = []
        account_map = {}

        for row_num in range(2, ws.max_row + 1):
            row_data = {'row_num': row_num}
            for col_idx, header in enumerate(headers):
                val = ws.cell(row=row_num, column=col_idx + 1).value
                row_data[header] = val

            account = row_data.get('银行账号') or row_data.get('账号')
            if account:
                account_key = _account_key(account)
                account_map[account_key] = row_data

            has_content = any(v is not None and str(v).strip() for v in row_data.values() if v != row_data['row_num'])
            if has_content:
                rows.append(row_data)

        wb.close()
        cleanup_temp_file(tmp_path)

        result = {
            'file_path': lookup_file,
            'file_hash': file_hash,
            'sheet_name': ws.title,
            'headers': headers,
            'rows': rows,
            'account_map': account_map,
        }

        logger.debug('读取查找表内容完成: %s (%d 条记录)', lookup_file, len(rows))
        return result

    except Exception as e:
        logger.error('读取查找表内容失败: %s', e)
        cleanup_temp_file(tmp_path)
        return None


def _save_lookup_snapshot(lookup_content, script_dir=None, username=None):
    """保存查找表快照到数据库"""
    if script_dir is None:
        script_dir = get_script_dir()
    db_path = get_audit_db_path(script_dir)
    init_audit_db(db_path)

    username = username or get_current_user()
    _ensure_user(username, db_path)

    snapshot_id = f"SNP{datetime.now().strftime('%Y%m%d%H%M%S')}{uuid.uuid4().hex[:6].upper()}"
    created_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')

    content_json = json.dumps(lookup_content, ensure_ascii=False)

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO lookup_snapshots (
            snapshot_id, lookup_file, file_hash, content_json,
            created_at, created_by
        ) VALUES (?, ?, ?, ?, ?, ?)
    ''', (
        snapshot_id,
        lookup_content.get('file_path', ''),
        lookup_content.get('file_hash', ''),
        content_json,
        created_at,
        username
    ))
    conn.commit()
    conn.close()

    logger = get_logger()
    logger.debug('查找表快照已保存: %s', snapshot_id)
    return snapshot_id


def _get_last_lookup_snapshot(lookup_file, script_dir=None):
    """获取指定查找表的最新快照"""
    if script_dir is None:
        script_dir = get_script_dir()
    db_path = get_audit_db_path(script_dir)

    if not os.path.exists(db_path):
        return None

    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute('''
        SELECT * FROM lookup_snapshots
        WHERE lookup_file = ?
        ORDER BY created_at DESC
        LIMIT 1
    ''', (lookup_file,))
    row = cursor.fetchone()
    conn.close()

    if row:
        return {
            'snapshot_id': row['snapshot_id'],
            'lookup_file': row['lookup_file'],
            'file_hash': row['file_hash'],
            'content': json.loads(row['content_json']),
            'created_at': row['created_at'],
            'created_by': row['created_by'],
        }
    return None


def _diff_lookup_snapshots(old_content, new_content):
    """
    比对两个查找表快照，生成详细变更列表。

    返回格式：
    [
        {
            'change_type': 'add' | 'remove' | 'modify',
            'account': '银行账号',
            'row_num': 行号,
            'field': '变更字段',
            'old_value': '旧值',
            'new_value': '新值',
            'description': '可读描述'
        },
        ...
    ]
    """
    changes = []

    if old_content is None and new_content is None:
        return changes

    if old_content is None:
        for row in (new_content or {}).get('rows', []):
            account = row.get('银行账号') or row.get('账号') or ''
            changes.append({
                'change_type': 'add',
                'account': str(account),
                'row_num': row.get('row_num'),
                'field': None,
                'old_value': None,
                'new_value': {k: v for k, v in row.items() if k != 'row_num'},
                'description': f'新增账号「{account}」: {row.get("主体名称") or row.get("主体") or "未命名主体"}'
            })
        return changes

    if new_content is None:
        for row in old_content.get('rows', []):
            account = row.get('银行账号') or row.get('账号') or ''
            changes.append({
                'change_type': 'remove',
                'account': str(account),
                'row_num': row.get('row_num'),
                'field': None,
                'old_value': {k: v for k, v in row.items() if k != 'row_num'},
                'new_value': None,
                'description': f'删除账号「{account}」: {row.get("主体名称") or row.get("主体") or "未命名主体"}'
            })
        return changes

    old_map = old_content.get('account_map', {})
    new_map = new_content.get('account_map', {})

    all_keys = set(list(old_map.keys()) + list(new_map.keys()))

    for key in all_keys:
        old_row = old_map.get(key)
        new_row = new_map.get(key)

        if old_row is None and new_row is not None:
            account = new_row.get('银行账号') or new_row.get('账号') or key
            changes.append({
                'change_type': 'add',
                'account': str(account),
                'row_num': new_row.get('row_num'),
                'field': None,
                'old_value': None,
                'new_value': {k: v for k, v in new_row.items() if k != 'row_num'},
                'description': f'新增账号「{account}」: {new_row.get("主体名称") or new_row.get("主体") or "未命名主体"}'
            })
        elif old_row is not None and new_row is None:
            account = old_row.get('银行账号') or old_row.get('账号') or key
            changes.append({
                'change_type': 'remove',
                'account': str(account),
                'row_num': old_row.get('row_num'),
                'field': None,
                'old_value': {k: v for k, v in old_row.items() if k != 'row_num'},
                'new_value': None,
                'description': f'删除账号「{account}」: {old_row.get("主体名称") or old_row.get("主体") or "未命名主体"}'
            })
        else:
            for field in set(list(old_row.keys()) + list(new_row.keys())):
                if field == 'row_num':
                    continue
                old_val = old_row.get(field)
                new_val = new_row.get(field)

                old_str = '' if old_val is None else str(old_val).strip()
                new_str = '' if new_val is None else str(new_val).strip()

                if old_str != new_str:
                    account = old_row.get('银行账号') or old_row.get('账号') or key
                    changes.append({
                        'change_type': 'modify',
                        'account': str(account),
                        'row_num': new_row.get('row_num'),
                        'field': field,
                        'old_value': old_val,
                        'new_value': new_val,
                        'description': f'账号「{account}」的「{field}」变更: {old_val} → {new_val}'
                    })

    return changes


@dataclass
class ChangeDetectionResult:
    """变更检测结果"""
    has_changes: bool = False
    change_id: Optional[str] = None
    change_details: List[Dict[str, Any]] = field(default_factory=list)
    old_snapshot: Optional[Dict] = None
    new_snapshot: Optional[Dict] = None
    old_content: Optional[Dict] = None
    new_content: Optional[Dict] = None


def detect_and_record_lookup_change(script_dir=None, username=None,
                                    change_reason='自动检测到查找表变更'):
    """
    自动检测查找表是否发生变更，如果有变更则记录到配置变更历史。

    工作流程：
    1. 读取当前查找表内容
    2. 获取数据库中保存的最新快照
    3. 比对两者差异（文件哈希 + 内容比对）
    4. 如果有变更，记录详细变更并保存新快照
    5. 返回变更检测结果

    Args:
        script_dir: 脚本目录
        username: 操作用户（默认自动获取）
        change_reason: 变更原因说明

    Returns:
        ChangeDetectionResult: 变更检测结果
    """
    if script_dir is None:
        script_dir = get_script_dir()
    logger = get_logger()

    lookup_file = find_lookup_file(script_dir)
    if lookup_file is None:
        logger.debug('未找到查找表，跳过变更检测')
        return ChangeDetectionResult(has_changes=False)

    username = username or get_current_user()

    current_content = read_lookup_table_content(lookup_file)
    if current_content is None:
        logger.warning('无法读取当前查找表内容，跳过变更检测')
        return ChangeDetectionResult(has_changes=False)

    last_snapshot = _get_last_lookup_snapshot(lookup_file, script_dir)
    last_content = last_snapshot.get('content') if last_snapshot else None

    current_hash = current_content.get('file_hash')
    last_hash = last_content.get('file_hash') if last_content else None

    if last_content is None:
        logger.info('首次运行，保存查找表初始快照')
        _save_lookup_snapshot(current_content, script_dir, username)
        return ChangeDetectionResult(
            has_changes=False,
            new_snapshot={'snapshot_id': 'initial'},
            new_content=current_content
        )

    if current_hash == last_hash:
        logger.debug('查找表文件哈希未变化，无变更')
        return ChangeDetectionResult(
            has_changes=False,
            old_snapshot=last_snapshot,
            old_content=last_content,
            new_content=current_content
        )

    changes = _diff_lookup_snapshots(last_content, current_content)

    if not changes:
        logger.info('查找表哈希变化但内容无差异，可能是格式或元数据变更')
        _save_lookup_snapshot(current_content, script_dir, username)
        return ChangeDetectionResult(
            has_changes=False,
            old_snapshot=last_snapshot,
            old_content=last_content,
            new_content=current_content
        )

    change_id = record_config_change(
        config_type='lookup_table',
        config_name=os.path.basename(lookup_file),
        old_value=last_content,
        new_value=current_content,
        change_reason=change_reason,
        script_dir=script_dir,
        username=username,
        change_details=changes
    )

    _save_lookup_snapshot(current_content, script_dir, username)

    add_count = sum(1 for c in changes if c['change_type'] == 'add')
    remove_count = sum(1 for c in changes if c['change_type'] == 'remove')
    modify_count = sum(1 for c in changes if c['change_type'] == 'modify')

    logger.warning(
        '检测到查找表变更 [%s]: 新增 %d, 删除 %d, 修改 %d',
        change_id, add_count, remove_count, modify_count
    )

    return ChangeDetectionResult(
        has_changes=True,
        change_id=change_id,
        change_details=changes,
        old_snapshot=last_snapshot,
        new_content=current_content,
        old_content=last_content
    )


def manual_record_lookup_change(script_dir=None, username=None, change_reason=''):
    """
    手动触发查找表变更记录。
    用于用户明确修改了查找表后需要立即记录的场景。

    Args:
        script_dir: 脚本目录
        username: 操作用户
        change_reason: 变更原因

    Returns:
        ChangeDetectionResult: 变更检测结果
    """
    return detect_and_record_lookup_change(
        script_dir=script_dir,
        username=username,
        change_reason=change_reason or '手动记录查找表变更'
    )


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
                         username=None, limit=100, parse_details=True):
    """
    查询配置变更历史。

    Args:
        script_dir: 脚本目录
        config_type: 按配置类型过滤
        config_name: 按配置名称过滤
        username: 按用户名过滤
        limit: 返回记录数限制
        parse_details: 是否解析 change_details JSON 字段

    Returns:
        List[Dict] 配置变更记录列表，包含 change_details 解析结果
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

    result = []
    for row in rows:
        record = dict(row)
        if parse_details and record.get('change_details'):
            try:
                record['change_details'] = json.loads(record['change_details'])
            except (json.JSONDecodeError, TypeError):
                pass
        result.append(record)

    return result


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


def export_config_changes(output_path, script_dir=None, expand_details=False, **query_kwargs):
    """
    导出配置变更历史到 Excel 文件。

    Args:
        output_path: 输出文件路径
        script_dir: 脚本目录
        expand_details: 是否展开详细变更记录为多行
        **query_kwargs: 查询参数

    Returns:
        输出文件路径
    """
    records = query_config_changes(script_dir=script_dir, **query_kwargs)
    if not records:
        return None

    if expand_details:
        expanded_rows = []
        for record in records:
            details = record.get('change_details')
            if isinstance(details, list) and details:
                for idx, detail in enumerate(details, 1):
                    row = record.copy()
                    row['detail_index'] = idx
                    row['change_type'] = detail.get('change_type', '')
                    row['account'] = detail.get('account', '')
                    row['row_num'] = detail.get('row_num', '')
                    row['field'] = detail.get('field', '')
                    row['old_value_detail'] = json.dumps(detail.get('old_value'), ensure_ascii=False) if isinstance(detail.get('old_value'), (dict, list)) else str(detail.get('old_value', ''))
                    row['new_value_detail'] = json.dumps(detail.get('new_value'), ensure_ascii=False) if isinstance(detail.get('new_value'), (dict, list)) else str(detail.get('new_value', ''))
                    row['description'] = detail.get('description', '')
                    expanded_rows.append(row)
            else:
                row = record.copy()
                row['detail_index'] = None
                row['change_type'] = ''
                row['account'] = ''
                row['row_num'] = ''
                row['field'] = ''
                row['old_value_detail'] = ''
                row['new_value_detail'] = ''
                row['description'] = ''
                expanded_rows.append(row)

        df = pd.DataFrame(expanded_rows)
        columns_order = [
            'change_id', 'username', 'config_type', 'config_name',
            'detail_index', 'change_type', 'account', 'row_num', 'field',
            'old_value_detail', 'new_value_detail', 'description',
            'change_reason', 'old_hash', 'new_hash', 'changed_at',
            'old_value', 'new_value', 'change_details', 'user_id', 'id'
        ]
    else:
        df = pd.DataFrame(records)
        if 'change_details' in df.columns:
            df['change_details'] = df['change_details'].apply(
                lambda x: json.dumps(x, ensure_ascii=False) if isinstance(x, (dict, list)) else str(x) if x is not None else ''
            )
        columns_order = [
            'change_id', 'username', 'config_type', 'config_name',
            'old_value', 'new_value', 'change_reason',
            'old_hash', 'new_hash', 'changed_at', 'change_details', 'user_id', 'id'
        ]

    for col in columns_order:
        if col not in df.columns:
            df[col] = None
    df = df[columns_order]

    df.to_excel(output_path, index=False, engine='openpyxl')
    logger = get_logger()
    logger.info('配置变更历史已导出到: %s (共 %d 条记录)', output_path, len(records))
    return output_path


# ──────────────────────────────────────────────
# 财务软件对接导出模块
# ──────────────────────────────────────────────

@dataclass
class StandardTransaction:
    """标准化交易记录，作为总表到各财务软件模板的中间格式"""
    transaction_date: Optional[datetime] = None
    voucher_date: Optional[datetime] = None
    voucher_number: str = ''
    summary: str = ''
    subject_code: str = ''
    subject_name: str = ''
    debit_amount: float = 0.0
    credit_amount: float = 0.0
    bank_account: str = ''
    bank_name: str = ''
    entity: str = ''
    counterparty: str = ''
    transaction_id: str = ''
    balance: float = 0.0
    direction: str = ''
    department: str = ''
    personnel: str = ''
    customer: str = ''
    supplier: str = ''
    project: str = ''
    attachment_count: int = 1
    prepared_by: str = ''
    reviewed_by: str = ''
    posted_by: str = ''
    remark: str = ''


@dataclass
class StandardVoucher:
    """标准化凭证，包含多条借贷分录"""
    voucher_date: Optional[datetime] = None
    voucher_number: str = ''
    voucher_type: str = '记'
    attachment_count: int = 1
    prepared_by: str = ''
    reviewed_by: str = ''
    posted_by: str = ''
    entries: List[StandardTransaction] = field(default_factory=list)
    source_transaction: Optional[dict] = None


FINANCIAL_EXPORT_TEMPLATES = {
    'yonyou_voucher': {
        'name': '用友凭证导入模板',
        'description': '用友U8/U9/NC系列财务软件凭证导入格式',
        'file_suffix': '_用友凭证导入',
    },
    'kingdee_voucher': {
        'name': '金蝶凭证导入模板',
        'description': '金蝶K3/KIS/EAS系列财务软件凭证导入格式',
        'file_suffix': '_金蝶凭证导入',
    },
    'bank_journal': {
        'name': '银行日记账模板',
        'description': '标准银行日记账格式，可直接导入财务软件',
        'file_suffix': '_银行日记账',
    },
}

DEFAULT_ACCOUNT_MAPPING = {
    'cash': {'code': '1001', 'name': '库存现金'},
    'bank_deposit': {'code': '1002', 'name': '银行存款'},
    'accounts_receivable': {'code': '1122', 'name': '应收账款'},
    'accounts_payable': {'code': '2202', 'name': '应付账款'},
    'operating_revenue': {'code': '6001', 'name': '主营业务收入'},
    'operating_cost': {'code': '6401', 'name': '主营业务成本'},
    'management_expense': {'code': '6602', 'name': '管理费用'},
    'sales_expense': {'code': '6601', 'name': '销售费用'},
    'financial_expense': {'code': '6603', 'name': '财务费用'},
}


def _normalize_date(value):
    """规范化日期为 datetime 对象"""
    if value is None:
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, str):
        for fmt in ['%Y-%m-%d', '%Y/%m/%d', '%Y.%m.%d', '%Y%m%d', '%Y-%m-%d %H:%M:%S', '%Y/%m/%d %H:%M:%S']:
            try:
                return datetime.strptime(value.strip(), fmt)
            except ValueError:
                continue
        try:
            return pd.to_datetime(value).to_pydatetime()
        except Exception:
            pass
    return None


def _format_date_for_export(dt, format_str='%Y-%m-%d'):
    """格式化日期用于导出"""
    if dt is None:
        return ''
    return dt.strftime(format_str)


def _format_amount(amount):
    """格式化金额，保留2位小数"""
    if amount is None:
        return 0.0
    try:
        return round(float(amount), 2)
    except (ValueError, TypeError):
        return 0.0


def total_to_standard_transactions(total_records, account_mapping=None, operator=''):
    """
    将总表数据转换为标准化交易记录列表。

    标准化层是财务导出的核心，它将银行流水的每一条记录转换为：
    1. 借方分录（收款时）或贷方分录（付款时）- 银行存款科目
    2. 对应的对方科目分录（需要根据摘要/对方户名智能判断）

    Args:
        total_records: 总表记录列表，每条包含银行流水字段
        account_mapping: 科目映射字典，可选
        operator: 制单人名称

    Returns:
        List[StandardVoucher]: 标准化凭证列表
    """
    logger = get_logger()

    if account_mapping is None:
        account_mapping = DEFAULT_ACCOUNT_MAPPING

    bank_account_info = account_mapping['bank_deposit']
    vouchers = []

    for idx, record in enumerate(total_records, 1):
        payment = to_float(record.get('付款'))
        receipt = to_float(record.get('收款'))
        trade_date = _normalize_date(record.get('交易日期'))
        voucher_num = f"记-{datetime.now().strftime('%Y%m')}-{idx:04d}"

        bank_entry = StandardTransaction(
            transaction_date=trade_date,
            voucher_date=trade_date,
            voucher_number=voucher_num,
            summary=str(record.get('摘要', '')) or '银行流水',
            subject_code=bank_account_info['code'],
            subject_name=bank_account_info['name'],
            bank_account=str(record.get('银行账号', '')),
            bank_name=str(record.get('银行', '')),
            entity=str(record.get('主体', '')),
            counterparty=str(record.get('对方户名', '')),
            transaction_id=str(record.get('交易流水号', '')),
            balance=_format_amount(record.get('余额')),
            prepared_by=operator,
            attachment_count=1,
        )

        counter_entry = StandardTransaction(
            transaction_date=trade_date,
            voucher_date=trade_date,
            voucher_number=voucher_num,
            summary=str(record.get('摘要', '')) or '银行流水',
            bank_account=str(record.get('银行账号', '')),
            bank_name=str(record.get('银行', '')),
            entity=str(record.get('主体', '')),
            counterparty=str(record.get('对方户名', '')),
            transaction_id=str(record.get('交易流水号', '')),
            prepared_by=operator,
            attachment_count=1,
        )

        if receipt and receipt > 0:
            bank_entry.debit_amount = _format_amount(receipt)
            bank_entry.direction = '借'

            counter_entry.credit_amount = _format_amount(receipt)
            summary_lower = str(record.get('摘要', '')).lower()
            counterparty_lower = str(record.get('对方户名', '')).lower()

            if any(k in summary_lower for k in ['收入', '销售', '货款', '营收', '主营业务']):
                counter_entry.subject_code = account_mapping['operating_revenue']['code']
                counter_entry.subject_name = account_mapping['operating_revenue']['name']
            elif any(k in summary_lower or k in counterparty_lower for k in ['客户', '应收']):
                counter_entry.subject_code = account_mapping['accounts_receivable']['code']
                counter_entry.subject_name = account_mapping['accounts_receivable']['name']
            else:
                counter_entry.subject_code = account_mapping['accounts_receivable']['code']
                counter_entry.subject_name = account_mapping['accounts_receivable']['name']

        elif payment and payment < 0:
            payment_abs = abs(payment)
            bank_entry.credit_amount = _format_amount(payment_abs)
            bank_entry.direction = '贷'

            counter_entry.debit_amount = _format_amount(payment_abs)
            summary_lower = str(record.get('摘要', '')).lower()
            counterparty_lower = str(record.get('对方户名', '')).lower()

            if any(k in summary_lower for k in ['成本', '主营成本']):
                counter_entry.subject_code = account_mapping['operating_cost']['code']
                counter_entry.subject_name = account_mapping['operating_cost']['name']
            elif any(k in summary_lower for k in ['管理费用', '办公费', '差旅费', '招待费', '工资']):
                counter_entry.subject_code = account_mapping['management_expense']['code']
                counter_entry.subject_name = account_mapping['management_expense']['name']
            elif any(k in summary_lower for k in ['销售费用', '广告费', '推广费']):
                counter_entry.subject_code = account_mapping['sales_expense']['code']
                counter_entry.subject_name = account_mapping['sales_expense']['name']
            elif any(k in summary_lower for k in ['财务费用', '手续费', '利息']):
                counter_entry.subject_code = account_mapping['financial_expense']['code']
                counter_entry.subject_name = account_mapping['financial_expense']['name']
            elif any(k in summary_lower or k in counterparty_lower for k in ['供应商', '应付', '采购', '货款']):
                counter_entry.subject_code = account_mapping['accounts_payable']['code']
                counter_entry.subject_name = account_mapping['accounts_payable']['name']
            else:
                counter_entry.subject_code = account_mapping['accounts_payable']['code']
                counter_entry.subject_name = account_mapping['accounts_payable']['name']
        else:
            continue

        voucher = StandardVoucher(
            voucher_date=trade_date,
            voucher_number=voucher_num,
            voucher_type='记',
            attachment_count=1,
            prepared_by=operator,
            entries=[bank_entry, counter_entry],
            source_transaction=record,
        )
        vouchers.append(voucher)

    logger.info('总表 %d 条记录转换为 %d 张标准化凭证', len(total_records), len(vouchers))
    return vouchers


def load_total_table(total_path):
    """加载总表数据"""
    logger = get_logger()

    if not total_path or not os.path.exists(total_path):
        logger.error('总表文件不存在: %s', total_path)
        return []

    try:
        df = pd.read_excel(total_path, engine='openpyxl')
        records = df.to_dict('records')
        logger.info('加载总表成功，共 %d 条记录', len(records))
        return records
    except Exception as e:
        logger.error('加载总表失败: %s', e, exc_info=True)
        return []


def export_yonyou_voucher(vouchers, output_path):
    """
    导出用友凭证导入格式。

    用友U8标准凭证导入格式包含以下字段：
    凭证类别字、凭证编号、凭证日期、附单据数、制单人、审核人、记账人、
    摘要、科目编码、科目名称、借方金额、贷方金额、
    部门编码、部门名称、个人编码、个人名称、客户编码、客户名称、
    供应商编码、供应商名称、项目编码、项目名称、银行账号、票据号
    """
    logger = get_logger()

    columns = [
        '凭证类别字', '凭证编号', '凭证日期', '附单据数',
        '制单人', '审核人', '记账人',
        '摘要', '科目编码', '科目名称',
        '借方金额', '贷方金额',
        '部门编码', '部门名称',
        '个人编码', '个人名称',
        '客户编码', '客户名称',
        '供应商编码', '供应商名称',
        '项目编码', '项目名称',
        '银行账号', '票据号',
    ]

    rows = []
    for voucher in vouchers:
        for entry in voucher.entries:
            rows.append({
                '凭证类别字': voucher.voucher_type,
                '凭证编号': voucher.voucher_number,
                '凭证日期': _format_date_for_export(voucher.voucher_date),
                '附单据数': voucher.attachment_count,
                '制单人': voucher.prepared_by,
                '审核人': voucher.reviewed_by,
                '记账人': voucher.posted_by,
                '摘要': entry.summary,
                '科目编码': str(entry.subject_code) if entry.subject_code is not None else '',
                '科目名称': entry.subject_name,
                '借方金额': entry.debit_amount,
                '贷方金额': entry.credit_amount,
                '部门编码': '',
                '部门名称': entry.department,
                '个人编码': '',
                '个人名称': entry.personnel,
                '客户编码': '',
                '客户名称': entry.customer,
                '供应商编码': '',
                '供应商名称': entry.supplier,
                '项目编码': '',
                '项目名称': entry.project,
                '银行账号': entry.bank_account,
                '票据号': entry.transaction_id,
            })

    if not rows:
        logger.warning('无凭证数据可导出')
        return None

    df = pd.DataFrame(rows, columns=columns)

    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='凭证导入', index=False)

            ws = writer.sheets['凭证导入']
            for col in ws.columns:
                max_length = max(len(str(cell.value)) if cell.value is not None else 0 for cell in col)
                ws.column_dimensions[col[0].column_letter].width = min(max_length + 2, 50)

        logger.info('用友凭证导出成功: %s (共 %d 张凭证，%d 条分录)',
                    output_path, len(vouchers), len(rows))
        return output_path
    except Exception as e:
        logger.error('导出用友凭证失败: %s', e, exc_info=True)
        return None


def export_kingdee_voucher(vouchers, output_path):
    """
    导出金蝶凭证导入格式。

    金蝶K3标准凭证导入格式包含以下字段：
    凭证字号、凭证日期、附件数、制单人、审核人、过账人、
    摘要、科目代码、科目名称、借方金额、贷方金额、
    核算项目类别、核算项目代码、核算项目名称、
    币别、汇率、原币金额、结算方式、结算号、业务日期
    """
    logger = get_logger()

    columns = [
        '凭证字号', '凭证日期', '附件数',
        '制单人', '审核人', '过账人',
        '摘要', '科目代码', '科目名称',
        '借方金额', '贷方金额',
        '核算项目类别', '核算项目代码', '核算项目名称',
        '币别', '汇率', '原币金额',
        '结算方式', '结算号', '业务日期',
    ]

    rows = []
    for voucher in vouchers:
        for entry in voucher.entries:
            rows.append({
                '凭证字号': voucher.voucher_number,
                '凭证日期': _format_date_for_export(voucher.voucher_date),
                '附件数': voucher.attachment_count,
                '制单人': voucher.prepared_by,
                '审核人': voucher.reviewed_by,
                '过账人': voucher.posted_by,
                '摘要': entry.summary,
                '科目代码': str(entry.subject_code) if entry.subject_code is not None else '',
                '科目名称': entry.subject_name,
                '借方金额': entry.debit_amount,
                '贷方金额': entry.credit_amount,
                '核算项目类别': '',
                '核算项目代码': '',
                '核算项目名称': '',
                '币别': '人民币',
                '汇率': 1.0,
                '原币金额': entry.debit_amount if entry.debit_amount else entry.credit_amount,
                '结算方式': '银行转账',
                '结算号': entry.transaction_id,
                '业务日期': _format_date_for_export(entry.transaction_date),
            })

    if not rows:
        logger.warning('无凭证数据可导出')
        return None

    df = pd.DataFrame(rows, columns=columns)

    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='凭证导入', index=False)

            ws = writer.sheets['凭证导入']
            for col in ws.columns:
                max_length = max(len(str(cell.value)) if cell.value is not None else 0 for cell in col)
                ws.column_dimensions[col[0].column_letter].width = min(max_length + 2, 50)

        logger.info('金蝶凭证导出成功: %s (共 %d 张凭证，%d 条分录)',
                    output_path, len(vouchers), len(rows))
        return output_path
    except Exception as e:
        logger.error('导出金蝶凭证失败: %s', e, exc_info=True)
        return None


def export_bank_journal(vouchers, output_path):
    """
    导出银行日记账格式。

    标准银行日记账格式包含以下字段：
    日期、凭证号、摘要、对方科目、借方金额、贷方金额、方向、余额、
    银行账号、开户银行、核算主体、对方户名、交易流水号、备注
    """
    logger = get_logger()

    columns = [
        '日期', '凭证号', '摘要', '对方科目',
        '借方金额', '贷方金额', '方向', '余额',
        '银行账号', '开户银行', '核算主体',
        '对方户名', '交易流水号', '备注',
    ]

    rows = []
    for voucher in vouchers:
        bank_entry = None
        other_entry = None

        for entry in voucher.entries:
            if '银行' in entry.subject_name or entry.subject_code.startswith('1002'):
                bank_entry = entry
            else:
                other_entry = entry

        if bank_entry is None:
            continue

        rows.append({
            '日期': _format_date_for_export(bank_entry.transaction_date),
            '凭证号': voucher.voucher_number,
            '摘要': bank_entry.summary,
            '对方科目': other_entry.subject_name if other_entry else '',
            '借方金额': bank_entry.debit_amount,
            '贷方金额': bank_entry.credit_amount,
            '方向': bank_entry.direction,
            '余额': bank_entry.balance,
            '银行账号': bank_entry.bank_account,
            '开户银行': bank_entry.bank_name,
            '核算主体': bank_entry.entity,
            '对方户名': bank_entry.counterparty,
            '交易流水号': bank_entry.transaction_id,
            '备注': bank_entry.remark,
        })

    if not rows:
        logger.warning('无日记账数据可导出')
        return None

    df = pd.DataFrame(rows, columns=columns)

    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='银行日记账', index=False)

            ws = writer.sheets['银行日记账']
            for col in ws.columns:
                max_length = max(len(str(cell.value)) if cell.value is not None else 0 for cell in col)
                ws.column_dimensions[col[0].column_letter].width = min(max_length + 2, 50)

            for row in ws.iter_rows(min_row=2):
                for cell in row:
                    if cell.column_letter in ['E', 'F', 'H']:
                        cell.number_format = '#,##0.00'

        logger.info('银行日记账导出成功: %s (共 %d 条记录)', output_path, len(rows))
        return output_path
    except Exception as e:
        logger.error('导出银行日记账失败: %s', e, exc_info=True)
        return None


def export_financial_template(total_path, template_type, output_dir=None,
                              account_mapping=None, operator=''):
    """
    标准化导出入口：从总表到财务软件模板的统一入口。

    这是在总表之上增加的一步标准化导出流程：
    总表数据 → 标准化转换 → 模板格式化 → 导出文件

    Args:
        total_path: 总表文件路径
        template_type: 模板类型 ('yonyou_voucher', 'kingdee_voucher', 'bank_journal')
        output_dir: 输出目录，默认为总表所在目录
        account_mapping: 科目映射字典，可选
        operator: 制单人名称

    Returns:
        dict: 导出结果，包含 output_path, voucher_count, entry_count 等信息
    """
    logger = get_logger()

    if template_type not in FINANCIAL_EXPORT_TEMPLATES:
        raise ValueError(f'不支持的导出模板类型: {template_type}')

    template_info = FINANCIAL_EXPORT_TEMPLATES[template_type]

    if output_dir is None:
        output_dir = os.path.dirname(total_path) or get_script_dir()

    os.makedirs(output_dir, exist_ok=True)

    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    base_name = os.path.splitext(os.path.basename(total_path))[0]
    output_filename = f'{base_name}{template_info["file_suffix"]}_{timestamp}.xlsx'
    output_path = os.path.join(output_dir, output_filename)

    logger.info('开始财务导出: 模板=%s, 总表=%s, 输出=%s',
                template_info['name'], total_path, output_path)

    total_records = load_total_table(total_path)
    if not total_records:
        logger.error('总表无数据可导出')
        return {'success': False, 'error': '总表无数据', 'output_path': None}

    vouchers = total_to_standard_transactions(total_records, account_mapping, operator)
    if not vouchers:
        logger.error('无有效凭证可导出')
        return {'success': False, 'error': '无有效凭证', 'output_path': None}

    export_func = {
        'yonyou_voucher': export_yonyou_voucher,
        'kingdee_voucher': export_kingdee_voucher,
        'bank_journal': export_bank_journal,
    }[template_type]

    result = export_func(vouchers, output_path)

    if result:
        entry_count = sum(len(v.entries) for v in vouchers)
        return {
            'success': True,
            'output_path': result,
            'voucher_count': len(vouchers),
            'entry_count': entry_count,
            'template_type': template_type,
            'template_name': template_info['name'],
        }
    else:
        return {'success': False, 'error': '导出失败', 'output_path': None}


def cli_ask_export_template():
    """命令行模式下询问导出模板类型"""
    print('\n请选择导出模板：')
    for idx, (key, info) in enumerate(FINANCIAL_EXPORT_TEMPLATES.items(), 1):
        print(f'  {idx}) {info["name"]}')
        print(f'     {info["description"]}')

    choice = input('\n请输入选项（直接回车默认为 1 用友凭证）: ').strip()

    if not choice:
        return 'yonyou_voucher'

    if choice.isdigit():
        keys = list(FINANCIAL_EXPORT_TEMPLATES.keys())
        idx = int(choice) - 1
        if 0 <= idx < len(keys):
            return keys[idx]

    if choice in FINANCIAL_EXPORT_TEMPLATES:
        return choice

    return 'yonyou_voucher'


def run_export_flow(script_dir):
    """财务导出主流程"""
    logger = get_logger()
    logger.info('========== 财务软件对接导出开始 ==========')

    total_path = ask_file('请选择【银行流水总表】文件')
    if not total_path:
        show_info('提示', '未选择总表文件，程序退出。')
        logger.info('用户未选择总表文件，退出导出流程')
        return

    logger.info('用户选择总表文件: %s', total_path)

    template_type = cli_ask_export_template()
    template_info = FINANCIAL_EXPORT_TEMPLATES[template_type]
    logger.info('用户选择导出模板: %s', template_info['name'])

    operator = input('请输入制单人名称（可选，直接回车为空）: ').strip()

    output_dir = input(f'请输入输出目录（直接回车默认为总表所在目录）: ').strip()
    if not output_dir:
        output_dir = os.path.dirname(total_path) or script_dir

    result = export_financial_template(
        total_path=total_path,
        template_type=template_type,
        output_dir=output_dir,
        operator=operator,
    )

    if result['success']:
        msg = (
            f'导出完成！\n\n'
            f'导出模板：{result["template_name"]}\n'
            f'凭证张数：{result["voucher_count"]}\n'
            f'分录条数：{result["entry_count"]}\n'
            f'输出文件：{result["output_path"]}'
        )
        show_info('导出成功', msg)
        logger.info('财务导出完成: %s', result["output_path"])
    else:
        msg = f'导出失败：{result.get("error", "未知错误")}'
        show_warning('导出失败', msg)
        logger.error('财务导出失败: %s', result.get("error"))

    logger.info('========== 财务软件对接导出结束 ==========')


# ──────────────────────────────────────────────
# 文件处理详情记录
# ──────────────────────────────────────────────

def record_file_processing_detail(audit_id, file_path, file_name, bank=None,
                                  status='processing', error_message=None,
                                  record_count=0, processing_duration_ms=None,
                                  started_at=None, completed_at=None, script_dir=None):
    """记录单个文件的处理详情"""
    if script_dir is None:
        script_dir = get_script_dir()
    db_path = get_audit_db_path(script_dir)
    init_audit_db(db_path)

    detail_id = f"FPD{datetime.now().strftime('%Y%m%d%H%M%S')}{uuid.uuid4().hex[:6].upper()}"
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')

    file_size = None
    try:
        if os.path.exists(file_path):
            file_size = os.path.getsize(file_path)
    except Exception:
        pass

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO file_processing_details (
            detail_id, audit_id, file_path, file_name, bank, file_size,
            record_count, status, error_message, processing_duration_ms,
            started_at, completed_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (
        detail_id, audit_id, file_path, file_name, bank, file_size,
        record_count, status, error_message, processing_duration_ms,
        started_at or now, completed_at
    ))
    conn.commit()
    conn.close()
    return detail_id


def update_file_processing_detail(detail_id, status=None, error_message=None,
                                  record_count=None, processing_duration_ms=None,
                                  completed_at=None, script_dir=None):
    """更新文件处理详情"""
    if script_dir is None:
        script_dir = get_script_dir()
    db_path = get_audit_db_path(script_dir)

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    updates = []
    params = []
    if status is not None:
        updates.append('status = ?')
        params.append(status)
    if error_message is not None:
        updates.append('error_message = ?')
        params.append(error_message)
    if record_count is not None:
        updates.append('record_count = ?')
        params.append(record_count)
    if processing_duration_ms is not None:
        updates.append('processing_duration_ms = ?')
        params.append(processing_duration_ms)
    if completed_at is not None:
        updates.append('completed_at = ?')
        params.append(completed_at)

    if not updates:
        conn.close()
        return

    params.append(detail_id)
    query = f"UPDATE file_processing_details SET {', '.join(updates)} WHERE detail_id = ?"
    cursor.execute(query, params)
    conn.commit()
    conn.close()


# ──────────────────────────────────────────────
# 告警规则管理
# ──────────────────────────────────────────────

@dataclass
class AlertRule:
    rule_id: str
    rule_name: str
    rule_type: str
    metric: str
    operator: str
    threshold: float
    window_minutes: int = 60
    severity: str = 'warning'
    enabled: bool = True
    description: Optional[str] = None


DEFAULT_ALERT_RULES = [
    {
        'rule_name': '处理成功率过低',
        'rule_type': 'pipeline',
        'metric': 'success_rate',
        'operator': '<',
        'threshold': 80.0,
        'window_minutes': 60,
        'severity': 'critical',
        'description': '最近1小时处理成功率低于80%',
    },
    {
        'rule_name': '单次运行耗时过长',
        'rule_type': 'pipeline',
        'metric': 'duration_ms',
        'operator': '>',
        'threshold': 300000,
        'window_minutes': 1,
        'severity': 'warning',
        'description': '单次运行耗时超过5分钟',
    },
    {
        'rule_name': '错误文件过多',
        'rule_type': 'pipeline',
        'metric': 'error_files',
        'operator': '>',
        'threshold': 5,
        'window_minutes': 60,
        'severity': 'warning',
        'description': '最近1小时错误文件数超过5个',
    },
    {
        'rule_name': '无法识别文件过多',
        'rule_type': 'pipeline',
        'metric': 'unprocessed_files',
        'operator': '>',
        'threshold': 10,
        'window_minutes': 60,
        'severity': 'warning',
        'description': '最近1小时无法识别文件数超过10个',
    },
    {
        'rule_name': '连续运行失败',
        'rule_type': 'pipeline',
        'metric': 'consecutive_failures',
        'operator': '>=',
        'threshold': 3,
        'window_minutes': 120,
        'severity': 'critical',
        'description': '最近2小时内连续3次运行失败',
    },
    {
        'rule_name': '处理量突增',
        'rule_type': 'pipeline',
        'metric': 'processed_files',
        'operator': '>',
        'threshold': 100,
        'window_minutes': 60,
        'severity': 'info',
        'description': '最近1小时处理文件数超过100个',
    },
]


def init_default_alert_rules(script_dir=None, username=None):
    """初始化默认告警规则"""
    if script_dir is None:
        script_dir = get_script_dir()
    db_path = get_audit_db_path(script_dir)
    init_audit_db(db_path)

    username = username or get_current_user()
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')
    created_count = 0

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    for rule_def in DEFAULT_ALERT_RULES:
        cursor.execute('SELECT rule_id FROM alert_rules WHERE rule_name = ?',
                       (rule_def['rule_name'],))
        if cursor.fetchone():
            continue

        rule_id = f"RUL{datetime.now().strftime('%Y%m%d%H%M%S')}{uuid.uuid4().hex[:6].upper()}"
        cursor.execute('''
            INSERT INTO alert_rules (
                rule_id, rule_name, rule_type, metric, operator, threshold,
                window_minutes, severity, enabled, description, created_at, updated_at, created_by
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            rule_id, rule_def['rule_name'], rule_def['rule_type'], rule_def['metric'],
            rule_def['operator'], rule_def['threshold'], rule_def['window_minutes'],
            rule_def['severity'], 1, rule_def.get('description'),
            now, now, username
        ))
        created_count += 1

    conn.commit()
    conn.close()

    if created_count > 0:
        logger = get_logger()
        logger.info('已初始化 %d 条默认告警规则', created_count)
    return created_count


def query_alert_rules(script_dir=None, rule_type=None, enabled=None, limit=100):
    """查询告警规则"""
    if script_dir is None:
        script_dir = get_script_dir()
    db_path = get_audit_db_path(script_dir)

    if not os.path.exists(db_path):
        return []

    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    query = 'SELECT * FROM alert_rules WHERE 1=1'
    params = []

    if rule_type:
        query += ' AND rule_type = ?'
        params.append(rule_type)
    if enabled is not None:
        query += ' AND enabled = ?'
        params.append(1 if enabled else 0)

    query += ' ORDER BY created_at DESC LIMIT ?'
    params.append(limit)

    cursor.execute(query, params)
    rows = cursor.fetchall()
    conn.close()

    return [dict(row) for row in rows]


def create_alert_rule(rule_name, rule_type, metric, operator, threshold,
                      window_minutes=60, severity='warning', description=None,
                      script_dir=None, username=None):
    """创建新的告警规则"""
    if script_dir is None:
        script_dir = get_script_dir()
    db_path = get_audit_db_path(script_dir)
    init_audit_db(db_path)

    username = username or get_current_user()
    rule_id = f"RUL{datetime.now().strftime('%Y%m%d%H%M%S')}{uuid.uuid4().hex[:6].upper()}"
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO alert_rules (
            rule_id, rule_name, rule_type, metric, operator, threshold,
            window_minutes, severity, enabled, description, created_at, updated_at, created_by
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (
        rule_id, rule_name, rule_type, metric, operator, threshold,
        window_minutes, severity, 1, description, now, now, username
    ))
    conn.commit()
    conn.close()

    logger = get_logger()
    logger.info('已创建告警规则 [%s] %s', rule_id, rule_name)
    return rule_id


def toggle_alert_rule(rule_id, enabled, script_dir=None):
    """启用/禁用告警规则"""
    if script_dir is None:
        script_dir = get_script_dir()
    db_path = get_audit_db_path(script_dir)

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute('''
        UPDATE alert_rules SET enabled = ?, updated_at = ? WHERE rule_id = ?
    ''', (1 if enabled else 0, datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f'), rule_id))
    conn.commit()
    conn.close()

    logger = get_logger()
    logger.info('告警规则 [%s] 已%s', rule_id, '启用' if enabled else '禁用')


def delete_alert_rule(rule_id, script_dir=None):
    """删除告警规则"""
    if script_dir is None:
        script_dir = get_script_dir()
    db_path = get_audit_db_path(script_dir)

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute('DELETE FROM alert_rules WHERE rule_id = ?', (rule_id,))
    conn.commit()
    conn.close()

    logger = get_logger()
    logger.info('告警规则 [%s] 已删除', rule_id)


# ──────────────────────────────────────────────
# 告警检测引擎
# ──────────────────────────────────────────────

def _evaluate_condition(current_value, operator, threshold):
    """评估告警条件"""
    if current_value is None:
        return False

    ops = {
        '>': lambda a, b: a > b,
        '>=': lambda a, b: a >= b,
        '<': lambda a, b: a < b,
        '<=': lambda a, b: a <= b,
        '==': lambda a, b: a == b,
        '!=': lambda a, b: a != b,
    }

    func = ops.get(operator)
    if func is None:
        return False
    return func(current_value, threshold)


def _get_metric_value(metric, records):
    """从审计记录中计算指标值"""
    if not records:
        return None

    if metric == 'success_rate':
        total = len(records)
        success = sum(1 for r in records if r['status'] == 'success')
        return (success / total * 100) if total > 0 else 100.0

    if metric == 'duration_ms':
        latest = max(records, key=lambda r: r['started_at'] or '')
        return latest.get('duration_ms')

    if metric == 'error_files':
        return sum(r.get('error_files', 0) for r in records)

    if metric == 'unprocessed_files':
        return sum(r.get('unprocessed_files', 0) for r in records)

    if metric == 'processed_files':
        return sum(r.get('processed_files', 0) for r in records)

    if metric == 'consecutive_failures':
        count = 0
        for r in sorted(records, key=lambda x: x['started_at'] or '', reverse=True):
            if r['status'] == 'failed':
                count += 1
            else:
                break
        return count

    return None


def create_alert(rule, current_value, audit_log_id=None, details=None, script_dir=None):
    """创建告警记录"""
    if script_dir is None:
        script_dir = get_script_dir()
    db_path = get_audit_db_path(script_dir)
    init_audit_db(db_path)

    alert_id = f"ALT{datetime.now().strftime('%Y%m%d%H%M%S')}{uuid.uuid4().hex[:8].upper()}"
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')

    description = rule.get('description', '')
    if current_value is not None:
        description += f' 当前值: {current_value}, 阈值: {rule["operator"]}{rule["threshold"]}'

    details_json = json.dumps(details, ensure_ascii=False) if details else None

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    cursor.execute('''
        SELECT alert_id FROM alerts
        WHERE rule_id = ? AND status = 'active'
        ORDER BY triggered_at DESC LIMIT 1
    ''', (rule['rule_id'],))
    existing = cursor.fetchone()

    if existing:
        conn.close()
        return None

    cursor.execute('''
        INSERT INTO alerts (
            alert_id, rule_id, rule_name, severity, alert_type, metric,
            current_value, threshold, operator, window_minutes, description,
            triggered_at, status, audit_log_id, details
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (
        alert_id, rule['rule_id'], rule['rule_name'], rule['severity'],
        rule['rule_type'], rule['metric'], current_value, rule['threshold'],
        rule['operator'], rule['window_minutes'], description,
        now, 'active', audit_log_id, details_json
    ))
    conn.commit()
    conn.close()

    logger = get_logger()
    logger.warning('告警触发 [%s] %s: %s', alert_id, rule['severity'].upper(),
                   rule['rule_name'])
    return alert_id


def run_alert_detection(script_dir=None, username=None):
    """运行告警检测，检查所有启用的规则"""
    if script_dir is None:
        script_dir = get_script_dir()
    logger = get_logger()

    init_default_alert_rules(script_dir, username)

    rules = query_alert_rules(script_dir=script_dir, enabled=True)
    if not rules:
        return []

    triggered_alerts = []

    for rule in rules:
        window_minutes = rule['window_minutes']
        from datetime import timedelta
        start_time = (datetime.now() - timedelta(minutes=window_minutes)).strftime('%Y-%m-%d %H:%M:%S')

        records = query_audit_logs(
            script_dir=script_dir,
            start_date=start_time.split(' ')[0],
            limit=1000
        )

        records_in_window = [
            r for r in records
            if r.get('started_at') and r['started_at'] >= start_time
        ]

        if not records_in_window:
            continue

        current_value = _get_metric_value(rule['metric'], records_in_window)
        if current_value is None:
            continue

        if _evaluate_condition(current_value, rule['operator'], rule['threshold']):
            latest_record = max(records_in_window, key=lambda r: r['started_at'] or '')
            alert_id = create_alert(
                rule, current_value,
                audit_log_id=latest_record.get('audit_id'),
                details={'records_analyzed': len(records_in_window)},
                script_dir=script_dir
            )
            if alert_id:
                triggered_alerts.append({
                    'alert_id': alert_id,
                    'rule': rule,
                    'current_value': current_value
                })

    if triggered_alerts:
        logger.warning('告警检测完成，触发 %d 条告警', len(triggered_alerts))
    else:
        logger.info('告警检测完成，无异常')

    return triggered_alerts


def query_alerts(script_dir=None, status=None, severity=None, limit=100):
    """查询告警记录"""
    if script_dir is None:
        script_dir = get_script_dir()
    db_path = get_audit_db_path(script_dir)

    if not os.path.exists(db_path):
        return []

    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    query = 'SELECT * FROM alerts WHERE 1=1'
    params = []

    if status:
        query += ' AND status = ?'
        params.append(status)
    if severity:
        query += ' AND severity = ?'
        params.append(severity)

    query += ' ORDER BY triggered_at DESC LIMIT ?'
    params.append(limit)

    cursor.execute(query, params)
    rows = cursor.fetchall()
    conn.close()

    result = []
    for row in rows:
        record = dict(row)
        if record.get('details'):
            try:
                record['details'] = json.loads(record['details'])
            except (json.JSONDecodeError, TypeError):
                pass
        result.append(record)
    return result


def acknowledge_alert(alert_id, script_dir=None, username=None):
    """确认告警"""
    if script_dir is None:
        script_dir = get_script_dir()
    db_path = get_audit_db_path(script_dir)

    username = username or get_current_user()
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute('''
        UPDATE alerts SET acknowledged = 1, acknowledged_by = ?, acknowledged_at = ?
        WHERE alert_id = ?
    ''', (username, now, alert_id))
    conn.commit()
    conn.close()

    logger = get_logger()
    logger.info('告警 [%s] 已由 %s 确认', alert_id, username)


def resolve_alert(alert_id, script_dir=None, username=None):
    """解决告警"""
    if script_dir is None:
        script_dir = get_script_dir()
    db_path = get_audit_db_path(script_dir)

    username = username or get_current_user()
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute('''
        UPDATE alerts SET status = 'resolved', resolved_at = ?,
                   acknowledged = 1, acknowledged_by = ?, acknowledged_at = ?
        WHERE alert_id = ?
    ''', (now, username, now, alert_id))
    conn.commit()
    conn.close()

    logger = get_logger()
    logger.info('告警 [%s] 已由 %s 标记为已解决', alert_id, username)


# ──────────────────────────────────────────────
# 监控统计分析
# ──────────────────────────────────────────────

@dataclass
class MonitorStats:
    total_runs: int = 0
    success_count: int = 0
    failed_count: int = 0
    success_rate: float = 0.0
    avg_duration_ms: float = 0.0
    max_duration_ms: int = 0
    min_duration_ms: int = 0
    total_processed_files: int = 0
    total_extracted_records: int = 0
    total_error_files: int = 0
    total_unprocessed_files: int = 0


def get_monitor_stats(script_dir=None, days=7):
    """获取指定天数内的监控统计数据"""
    from datetime import timedelta
    if script_dir is None:
        script_dir = get_script_dir()

    start_date = (datetime.now() - timedelta(days=days)).strftime('%Y-%m-%d')
    records = query_audit_logs(
        script_dir=script_dir,
        start_date=start_date,
        limit=10000
    )

    if not records:
        return MonitorStats()

    durations = [r['duration_ms'] for r in records if r.get('duration_ms') is not None]

    stats = MonitorStats(
        total_runs=len(records),
        success_count=sum(1 for r in records if r['status'] == 'success'),
        failed_count=sum(1 for r in records if r['status'] == 'failed'),
        total_processed_files=sum(r.get('processed_files', 0) for r in records),
        total_extracted_records=sum(r.get('extracted_records', 0) for r in records),
        total_error_files=sum(r.get('error_files', 0) for r in records),
        total_unprocessed_files=sum(r.get('unprocessed_files', 0) for r in records),
    )

    if stats.total_runs > 0:
        stats.success_rate = (stats.success_count / stats.total_runs) * 100

    if durations:
        stats.avg_duration_ms = sum(durations) / len(durations)
        stats.max_duration_ms = max(durations)
        stats.min_duration_ms = min(durations)

    return stats


def get_daily_trend(script_dir=None, days=7):
    """获取每日运行趋势数据"""
    from datetime import timedelta
    if script_dir is None:
        script_dir = get_script_dir()
    db_path = get_audit_db_path(script_dir)

    if not os.path.exists(db_path):
        return []

    start_date = (datetime.now() - timedelta(days=days - 1)).strftime('%Y-%m-%d')

    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    cursor.execute('''
        SELECT
            date(started_at) as run_date,
            COUNT(*) as total_runs,
            SUM(CASE WHEN status = 'success' THEN 1 ELSE 0 END) as success_count,
            SUM(CASE WHEN status = 'failed' THEN 1 ELSE 0 END) as failed_count,
            AVG(duration_ms) as avg_duration_ms,
            SUM(processed_files) as total_processed_files,
            SUM(extracted_records) as total_extracted_records,
            SUM(error_files) as total_error_files,
            SUM(unprocessed_files) as total_unprocessed_files
        FROM audit_logs
        WHERE date(started_at) >= date(?)
        GROUP BY date(started_at)
        ORDER BY run_date DESC
    ''', (start_date,))

    rows = cursor.fetchall()
    conn.close()

    trend_data = []
    for row in rows:
        data = dict(row)
        total = data['total_runs'] or 0
        success = data['success_count'] or 0
        data['success_rate'] = (success / total * 100) if total > 0 else 0
        trend_data.append(data)

    return trend_data


def get_bank_distribution(script_dir=None, days=7):
    """获取各银行处理分布"""
    from datetime import timedelta
    if script_dir is None:
        script_dir = get_script_dir()
    db_path = get_audit_db_path(script_dir)

    if not os.path.exists(db_path):
        return []

    start_date = (datetime.now() - timedelta(days=days)).strftime('%Y-%m-%d')

    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    cursor.execute('''
        SELECT
            fpd.bank,
            COUNT(*) as file_count,
            SUM(fpd.record_count) as total_records,
            AVG(fpd.processing_duration_ms) as avg_duration_ms,
            SUM(CASE WHEN fpd.status = 'success' THEN 1 ELSE 0 END) as success_count,
            SUM(CASE WHEN fpd.status = 'error' THEN 1 ELSE 0 END) as error_count
        FROM file_processing_details fpd
        JOIN audit_logs al ON fpd.audit_id = al.audit_id
        WHERE date(al.started_at) >= date(?) AND fpd.bank IS NOT NULL
        GROUP BY fpd.bank
        ORDER BY file_count DESC
    ''', (start_date,))

    rows = cursor.fetchall()
    conn.close()

    result = []
    for row in rows:
        data = dict(row)
        total = data['file_count'] or 0
        success = data['success_count'] or 0
        data['success_rate'] = (success / total * 100) if total > 0 else 0
        result.append(data)

    return result


def get_recent_runs(script_dir=None, limit=10):
    """获取最近运行记录列表"""
    records = query_audit_logs(script_dir=script_dir, limit=limit)
    return records


def get_active_alerts_summary(script_dir=None):
    """获取活跃告警摘要"""
    alerts = query_alerts(script_dir=script_dir, status='active', limit=100)
    summary = {
        'total': len(alerts),
        'critical': sum(1 for a in alerts if a['severity'] == 'critical'),
        'warning': sum(1 for a in alerts if a['severity'] == 'warning'),
        'info': sum(1 for a in alerts if a['severity'] == 'info'),
        'alerts': alerts
    }
    return summary


# ──────────────────────────────────────────────
# 运行监控面板
# ──────────────────────────────────────────────

def _format_duration(ms):
    """格式化毫秒为可读时间"""
    if ms is None:
        return 'N/A'
    seconds = ms / 1000
    if seconds < 60:
        return f'{seconds:.1f}秒'
    minutes = seconds / 60
    if minutes < 60:
        return f'{minutes:.1f}分钟'
    hours = minutes / 60
    return f'{hours:.2f}小时'


def _format_rate(rate):
    """格式化百分比"""
    if rate is None:
        return 'N/A'
    return f'{rate:.2f}%'


def _get_severity_color(severity):
    """获取告警严重度显示符号"""
    colors = {
        'critical': '🔴',
        'warning': '🟡',
        'info': '🔵',
    }
    return colors.get(severity, '⚪')


def _get_status_symbol(status):
    """获取状态显示符号"""
    symbols = {
        'success': '✅',
        'failed': '❌',
        'running': '⏳',
        'active': '🔴',
        'resolved': '✅',
        'acknowledged': '🟡',
    }
    return symbols.get(status, '⚪')


def render_monitor_dashboard(script_dir=None):
    """渲染监控面板主界面"""
    if script_dir is None:
        script_dir = get_script_dir()

    init_default_alert_rules(script_dir)

    stats = get_monitor_stats(script_dir, days=7)
    daily_trend = get_daily_trend(script_dir, days=7)
    recent_runs = get_recent_runs(script_dir, limit=5)
    alerts_summary = get_active_alerts_summary(script_dir)
    bank_dist = get_bank_distribution(script_dir, days=7)

    print('\n' + '=' * 80)
    print('📊 运行监控与告警面板'.center(70))
    print('=' * 80)
    print(f'统计周期: 最近7天')
    print()

    print('┌' + '─' * 78 + '┐')
    print('│  📈 运行概览' + ' ' * 65 + '│')
    print('├' + '─' * 78 + '┤')
    print(f'│  总运行次数: {stats.total_runs:<10}  |  成功: {stats.success_count:<10}  |  失败: {stats.failed_count:<10}  |  成功率: {_format_rate(stats.success_rate):>10} │')
    print(f'│  平均耗时: {_format_duration(stats.avg_duration_ms):<15} |  最大: {_format_duration(stats.max_duration_ms):<15} |  最小: {_format_duration(stats.min_duration_ms):<15} │')
    print(f'│  处理文件: {stats.total_processed_files:<10} |  提取记录: {stats.total_extracted_records:<10} |  错误文件: {stats.total_error_files:<10} |  未识别: {stats.total_unprocessed_files:<10} │')
    print('└' + '─' * 78 + '┘')
    print()

    if alerts_summary['total'] > 0:
        print('┌' + '─' * 78 + '┐')
        print(f'│  🚨 活跃告警 ({alerts_summary["total"]}条)' + ' ' * 56 + '│')
        print('├' + '─' * 78 + '┤')
        print(f'│  🔴 严重: {alerts_summary["critical"]:<5} | 🟡 警告: {alerts_summary["warning"]:<5} | 🔵 信息: {alerts_summary["info"]:<5}' + ' ' * 45 + '│')
        for alert in alerts_summary['alerts'][:3]:
            sev = _get_severity_color(alert['severity'])
            short_desc = alert.get('description', '')[:60]
            print(f'│  {sev} [{alert["alert_id"]}] {alert["rule_name"]}: {short_desc} │')
        if alerts_summary['total'] > 3:
            print(f'│  ... 还有 {alerts_summary["total"] - 3} 条告警，请查看详情' + ' ' * 38 + '│')
        print('└' + '─' * 78 + '┘')
        print()

    if daily_trend:
        print('┌' + '─' * 78 + '┐')
        print('│  📅 每日运行趋势' + ' ' * 63 + '│')
        print('├' + '─' * 78 + '┤')
        print(f'│  {"日期":<12} {"次数":<6} {"成功":<6} {"失败":<6} {"成功率":<9} {"处理":<6} {"错误":<6} {"未识别":<8} {"耗时":<10} │')
        print('│' + '─' * 78 + '│')
        for day in daily_trend[:7]:
            err = day.get('total_error_files', 0) or 0
            unproc = day.get('total_unprocessed_files', 0) or 0
            print(f'│  {day["run_date"]:<12} {day["total_runs"]:<6} {day["success_count"]:<6} {day["failed_count"]:<6} {_format_rate(day["success_rate"]):<9} {day["total_processed_files"]:<6} {err:<6} {unproc:<8} {_format_duration(day["avg_duration_ms"]):<10} │')
        print('└' + '─' * 78 + '┘')
        print()

    if bank_dist:
        print('┌' + '─' * 78 + '┐')
        print('│  🏦 银行处理分布' + ' ' * 63 + '│')
        print('├' + '─' * 78 + '┤')
        print(f'│  {"银行":<12} {"文件数":<8} {"记录数":<10} {"成功率":<10} {"平均耗时":<12} {"错误数":<8} │')
        print('│' + '─' * 78 + '│')
        for bank in bank_dist:
            name = bank.get('bank', '未知')[:10]
            print(f'│  {name:<12} {bank["file_count"]:<8} {bank["total_records"]:<10} {_format_rate(bank["success_rate"]):<10} {_format_duration(bank["avg_duration_ms"]):<12} {bank["error_count"]:<8} │')
        print('└' + '─' * 78 + '┘')
        print()

    if recent_runs:
        print('┌' + '─' * 78 + '┐')
        print('│  🕐 最近运行记录' + ' ' * 63 + '│')
        print('├' + '─' * 78 + '┤')
        print(f'│  {"审计ID":<14} {"时间":<20} {"类型":<10} {"状态":<8} {"耗时":<12} {"文件":<6} │')
        print('│' + '─' * 78 + '│')
        for run in recent_runs:
            status_sym = _get_status_symbol(run['status'])
            audit_id = run['audit_id'][:12]
            start_time = run.get('started_at', '')[:19]
            op_type = run.get('operation_type', '')[:8]
            duration = _format_duration(run.get('duration_ms'))
            files = run.get('processed_files', 0)
            print(f'│  {audit_id:<14} {start_time:<20} {op_type:<10} {status_sym} {run["status"]:<6} {duration:<12} {files:<6} │')
        print('└' + '─' * 78 + '┘')
    print()


def show_monitor_menu(script_dir=None):
    """显示监控面板菜单并处理用户选择"""
    if script_dir is None:
        script_dir = get_script_dir()
    logger = get_logger()

    while True:
        run_alert_detection(script_dir)
        render_monitor_dashboard(script_dir)

        print('\n请选择操作：')
        print('  1) 🔄 刷新面板')
        print('  2) 🚨 查看所有告警')
        print('  3) 📋 查看详细运行历史')
        print('  4) ⚙️  告警规则管理')
        print('  5) 📊 导出监控报告')
        print('  6) ✅ 确认/解决告警')
        print('  0) 返回主菜单')

        choice = input('\n请输入选项: ').strip()

        if choice == '0':
            break
        elif choice == '1':
            continue
        elif choice == '2':
            _show_all_alerts(script_dir)
        elif choice == '3':
            _show_run_history(script_dir)
        elif choice == '4':
            _show_alert_rules_menu(script_dir)
        elif choice == '5':
            _export_monitor_report(script_dir)
        elif choice == '6':
            _acknowledge_alert_menu(script_dir)
        else:
            print('无效选项，请重试')
            input('\n按回车继续...')


def _show_all_alerts(script_dir=None):
    """显示所有告警列表"""
    if script_dir is None:
        script_dir = get_script_dir()

    alerts = query_alerts(script_dir=script_dir, limit=50)

    print('\n' + '=' * 80)
    print('📋 告警记录列表'.center(70))
    print('=' * 80)
    print(f'{"告警ID":<16} {"时间":<20} {"严重度":<8} {"规则名称":<20} {"状态":<10}')
    print('-' * 80)

    for alert in alerts:
        sev = _get_severity_color(alert['severity'])
        status_sym = _get_status_symbol(alert['status'])
        alert_id = alert['alert_id'][:14]
        triggered = alert.get('triggered_at', '')[:19]
        rule_name = alert.get('rule_name', '')[:18]
        status = alert.get('status', '')
        print(f'{alert_id:<16} {triggered:<20} {sev} {alert["severity"]:<6} {rule_name:<20} {status_sym} {status:<8}')
        print(f'  → {alert.get("description", "")[:70]}')

    if not alerts:
        print('暂无告警记录')

    input('\n按回车继续...')


def _show_run_history(script_dir=None):
    """显示详细运行历史"""
    if script_dir is None:
        script_dir = get_script_dir()

    days_input = input('查看最近几天的记录 (默认7): ').strip()
    days = int(days_input) if days_input.isdigit() else 7

    records = query_audit_logs(script_dir=script_dir, limit=50)
    daily_trend = get_daily_trend(script_dir, days=days)

    print('\n' + '=' * 80)
    print(f'📊 运行历史记录 (最近{days}天)'.center(70))
    print('=' * 80)
    print(f'{"审计ID":<14} {"开始时间":<20} {"类型":<10} {"状态":<8} {"耗时":<10} {"处理":<6} {"错误":<6} {"未识别":<8} {"记录":<8}')
    print('-' * 80)

    for run in records:
        status_sym = _get_status_symbol(run['status'])
        audit_id = run['audit_id'][:12]
        start_time = run.get('started_at', '')[:19]
        op_type = run.get('operation_type', '')[:8]
        duration = _format_duration(run.get('duration_ms'))
        files = run.get('processed_files', 0)
        err = run.get('error_files', 0)
        unproc = run.get('unprocessed_files', 0)
        recs = run.get('extracted_records', 0)
        print(f'{audit_id:<14} {start_time:<20} {op_type:<10} {status_sym} {run["status"]:<6} {duration:<10} {files:<6} {err:<6} {unproc:<8} {recs:<8}')

    if daily_trend:
        print('\n' + '=' * 80)
        print('📈 处理量趋势图 (ASCII)'.center(70))
        print('=' * 80)
        max_files = max(d.get('total_processed_files', 0) for d in daily_trend) or 1
        for day in reversed(daily_trend):
            date = day['run_date']
            files = day.get('total_processed_files', 0)
            bar_len = int((files / max_files) * 40)
            bar = '█' * bar_len
            print(f'{date} | {bar:<40} | {files:>5} 文件')

        print('\n' + '=' * 80)
        print('🚨 错误文件趋势图 (ASCII)'.center(70))
        print('=' * 80)
        max_err = max((d.get('total_error_files', 0) or 0) for d in daily_trend) or 1
        for day in reversed(daily_trend):
            date = day['run_date']
            err = day.get('total_error_files', 0) or 0
            bar_len = int((err / max_err) * 40) if max_err > 0 else 0
            bar = '▓' * bar_len
            print(f'{date} | {bar:<40} | {err:>5} 错误')

        print('\n' + '=' * 80)
        print('⚠️  未识别文件趋势图 (ASCII)'.center(70))
        print('=' * 80)
        max_unproc = max((d.get('total_unprocessed_files', 0) or 0) for d in daily_trend) or 1
        for day in reversed(daily_trend):
            date = day['run_date']
            unproc = day.get('total_unprocessed_files', 0) or 0
            bar_len = int((unproc / max_unproc) * 40) if max_unproc > 0 else 0
            bar = '▒' * bar_len
            print(f'{date} | {bar:<40} | {unproc:>5} 未识别')

    input('\n按回车继续...')


def _show_alert_rules_menu(script_dir=None):
    """告警规则管理菜单"""
    if script_dir is None:
        script_dir = get_script_dir()

    while True:
        rules = query_alert_rules(script_dir=script_dir)

        print('\n' + '=' * 80)
        print('⚙️  告警规则管理'.center(70))
        print('=' * 80)
        print(f'{"规则ID":<12} {"名称":<20} {"指标":<16} {"条件":<16} {"严重度":<8} {"状态":<8}')
        print('-' * 80)

        for rule in rules:
            rid = rule['rule_id'][:10]
            name = rule['rule_name'][:18]
            metric = rule['metric'][:14]
            condition = f'{rule["operator"]}{rule["threshold"]}'
            enabled = '✅启用' if rule['enabled'] else '❌禁用'
            print(f'{rid:<12} {name:<20} {metric:<16} {condition:<16} {rule["severity"]:<8} {enabled:<8}')

        print('\n请选择操作：')
        print('  1) 启用/禁用规则')
        print('  2) 新增规则')
        print('  3) 删除规则')
        print('  4) 重置为默认规则')
        print('  0) 返回')

        choice = input('\n请输入选项: ').strip()

        if choice == '0':
            break
        elif choice == '1':
            rule_id = input('请输入规则ID: ').strip()
            rule = next((r for r in rules if r['rule_id'] == rule_id), None)
            if rule:
                new_state = not rule['enabled']
                toggle_alert_rule(rule_id, new_state, script_dir)
                print(f'规则已{"启用" if new_state else "禁用"}')
            else:
                print('未找到该规则')
        elif choice == '2':
            _create_new_rule_interactive(script_dir)
        elif choice == '3':
            rule_id = input('请输入要删除的规则ID: ').strip()
            confirm = input(f'确认删除规则 {rule_id}? (y/N): ').strip().lower()
            if confirm == 'y':
                delete_alert_rule(rule_id, script_dir)
                print('规则已删除')
        elif choice == '4':
            confirm = input('确认重置所有规则为默认? (y/N): ').strip().lower()
            if confirm == 'y':
                conn = sqlite3.connect(get_audit_db_path(script_dir))
                cursor = conn.cursor()
                cursor.execute('DELETE FROM alert_rules')
                conn.commit()
                conn.close()
                init_default_alert_rules(script_dir)
                print('已重置为默认规则')
        input('\n按回车继续...')


def _create_new_rule_interactive(script_dir=None):
    """交互式创建新告警规则"""
    print('\n--- 新建告警规则 ---')
    rule_name = input('规则名称: ').strip()
    if not rule_name:
        print('名称不能为空')
        return

    print('\n指标类型:')
    print('  success_rate - 成功率')
    print('  duration_ms - 运行耗时(ms)')
    print('  error_files - 错误文件数')
    print('  unprocessed_files - 未识别文件数')
    print('  processed_files - 处理文件数')
    print('  consecutive_failures - 连续失败次数')
    metric = input('请输入指标: ').strip()

    print('\n操作符: >, >=, <, <=, ==, !=')
    operator = input('请输入操作符: ').strip()

    threshold_input = input('阈值: ').strip()
    try:
        threshold = float(threshold_input)
    except ValueError:
        print('阈值必须是数字')
        return

    window_input = input('时间窗口(分钟，默认60): ').strip()
    window_minutes = int(window_input) if window_input.isdigit() else 60

    print('\n严重度: critical, warning, info')
    severity = input('严重度 (默认warning): ').strip() or 'warning'

    description = input('描述 (可选): ').strip() or None

    rule_id = create_alert_rule(
        rule_name=rule_name,
        rule_type='pipeline',
        metric=metric,
        operator=operator,
        threshold=threshold,
        window_minutes=window_minutes,
        severity=severity,
        description=description,
        script_dir=script_dir
    )
    print(f'规则已创建，ID: {rule_id}')


def _acknowledge_alert_menu(script_dir=None):
    """确认/解决告警菜单"""
    if script_dir is None:
        script_dir = get_script_dir()

    alerts = query_alerts(script_dir=script_dir, status='active', limit=50)

    if not alerts:
        print('\n没有活跃的告警需要处理')
        input('按回车继续...')
        return

    print('\n' + '=' * 80)
    print('✅ 告警处理'.center(70))
    print('=' * 80)
    print(f'{"告警ID":<16} {"时间":<20} {"严重度":<8} {"规则名称":<20}')
    print('-' * 80)

    for alert in alerts:
        sev = _get_severity_color(alert['severity'])
        alert_id = alert['alert_id'][:14]
        triggered = alert.get('triggered_at', '')[:19]
        rule_name = alert.get('rule_name', '')[:18]
        print(f'{alert_id:<16} {triggered:<20} {sev} {alert["severity"]:<6} {rule_name:<20}')
        print(f'  → {alert.get("description", "")[:70]}')

    alert_id = input('\n请输入要处理的告警ID: ').strip()
    if not alert_id:
        return

    alert = next((a for a in alerts if a['alert_id'] == alert_id), None)
    if not alert:
        print('未找到该告警')
        input('按回车继续...')
        return

    print('\n请选择操作：')
    print('  1) 确认告警')
    print('  2) 标记为已解决')
    action = input('请输入选项: ').strip()

    if action == '1':
        acknowledge_alert(alert_id, script_dir)
        print('告警已确认')
    elif action == '2':
        resolve_alert(alert_id, script_dir)
        print('告警已标记为已解决')

    input('按回车继续...')


def _export_monitor_report(script_dir=None):
    """导出监控报告到Excel"""
    if script_dir is None:
        script_dir = get_script_dir()
    logger = get_logger()

    days_input = input('导出最近几天的数据 (默认7): ').strip()
    days = int(days_input) if days_input.isdigit() else 7

    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_path = os.path.join(script_dir, f'监控报告_{timestamp}.xlsx')

    stats = get_monitor_stats(script_dir, days=days)
    daily_trend = get_daily_trend(script_dir, days=days)
    bank_dist = get_bank_distribution(script_dir, days=days)
    alerts = query_alerts(script_dir=script_dir, limit=1000)
    runs = query_audit_logs(script_dir=script_dir, limit=1000)

    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            summary_data = [{
                '统计项目': '总运行次数',
                '数值': stats.total_runs,
            }, {
                '统计项目': '成功次数',
                '数值': stats.success_count,
            }, {
                '统计项目': '失败次数',
                '数值': stats.failed_count,
            }, {
                '统计项目': '成功率(%)',
                '数值': round(stats.success_rate, 2),
            }, {
                '统计项目': '平均耗时(秒)',
                '数值': round(stats.avg_duration_ms / 1000, 2) if stats.avg_duration_ms else None,
            }, {
                '统计项目': '最大耗时(秒)',
                '数值': round(stats.max_duration_ms / 1000, 2) if stats.max_duration_ms else None,
            }, {
                '统计项目': '最小耗时(秒)',
                '数值': round(stats.min_duration_ms / 1000, 2) if stats.min_duration_ms else None,
            }, {
                '统计项目': '总处理文件数',
                '数值': stats.total_processed_files,
            }, {
                '统计项目': '总提取记录数',
                '数值': stats.total_extracted_records,
            }, {
                '统计项目': '总错误文件数',
                '数值': stats.total_error_files,
            }, {
                '统计项目': '总未识别文件数',
                '数值': stats.total_unprocessed_files,
            }, {
                '统计项目': '活跃告警数',
                '数值': sum(1 for a in alerts if a['status'] == 'active'),
            }, {
                '统计项目': '严重告警数',
                '数值': sum(1 for a in alerts if a['status'] == 'active' and a['severity'] == 'critical'),
            }]
            pd.DataFrame(summary_data).to_excel(writer, sheet_name='概览', index=False)

            if daily_trend:
                df_trend = pd.DataFrame(daily_trend)
                df_trend.to_excel(writer, sheet_name='每日趋势', index=False)

            if bank_dist:
                df_bank = pd.DataFrame(bank_dist)
                df_bank.to_excel(writer, sheet_name='银行分布', index=False)

            if runs:
                df_runs = pd.DataFrame(runs)
                df_runs.to_excel(writer, sheet_name='运行历史', index=False)

            if alerts:
                for a in alerts:
                    if isinstance(a.get('details'), (dict, list)):
                        a['details'] = json.dumps(a['details'], ensure_ascii=False)
                df_alerts = pd.DataFrame(alerts)
                df_alerts.to_excel(writer, sheet_name='告警记录', index=False)

        logger.info('监控报告已导出: %s', output_path)
        print(f'\n✅ 监控报告已导出到: {output_path}')
    except Exception as e:
        logger.error('导出监控报告失败: %s', e)
        print(f'\n❌ 导出失败: {e}')

    input('\n按回车继续...')


def run_monitor_flow(script_dir):
    """监控面板流程"""
    logger = get_logger()
    logger.info('========== 运行监控面板启动 ==========')

    init_default_alert_rules(script_dir)
    run_alert_detection(script_dir)
    show_monitor_menu(script_dir)

    logger.info('========== 运行监控面板关闭 ==========')


def run_pipeline_flow(script_dir):
    """主流程：处理银行流水文件夹，输出总表"""
    logger = get_logger()

    folder = ask_directory('请选择银行流水文件夹')
    if not folder:
        show_info('提示', '未选择文件夹，程序退出。')
        logger.info('用户未选择文件夹，程序退出')
        return

    logger.info('用户选择文件夹: %s', folder)

    change_result = detect_and_record_lookup_change(script_dir)
    if change_result.has_changes:
        add_count = sum(1 for c in change_result.change_details if c['change_type'] == 'add')
        remove_count = sum(1 for c in change_result.change_details if c['change_type'] == 'remove')
        modify_count = sum(1 for c in change_result.change_details if c['change_type'] == 'modify')
        warning_msg = (
            f'检测到查找表已变更！\n\n'
            f'变更编号: {change_result.change_id}\n'
            f'新增记录: {add_count} 条\n'
            f'删除记录: {remove_count} 条\n'
            f'修改记录: {modify_count} 条\n\n'
        )
        show_warning('查找表变更提醒', warning_msg)
        logger.warning('查找表变更已记录: %s', change_result.change_id)

    incremental = ask_incremental_mode()
    if incremental is None:
        logger.info('用户取消增量模式选择，返回主菜单')
        return
    logger.info('用户选择运行模式: %s', '增量合并' if incremental else '全量覆盖')

    batch_id = None
    batch_manager = None
    if HAS_BATCH_MANAGER:
        try:
            batch_manager = batch_module.get_batch_manager(script_dir)
            operator = get_current_user()
            batch_info = batch_manager.start_batch(input_folder=folder, operator=operator)
            batch_id = batch_info.batch_id
            logger.info('批次管理已启用，批次号: %s', batch_id)
        except Exception as e:
            logger.error('批次创建失败: %s', e, exc_info=True)
            batch_manager = None

    with AuditLogger('pipeline', script_dir) as audit:
        audit.record_input(folder)

        result = run_pipeline(folder, script_dir, incremental=incremental, batch_id=batch_id)
        audit.record_result(result)

        if result.lookup_missing:
            show_warning(
                '警告',
                '在程序所在目录下未找到主体查找表文件，\n"主体"列将为空。\n'
                '建议将查找表文件命名为"主体查找表.xlsx"并放在程序所在目录下。'
            )

        msg = format_result_message(result)
        msg += f'\n\n审计编号: {audit.audit_id}'
        if change_result.change_id:
            msg += f'\n配置变更编号: {change_result.change_id}'

        if batch_manager and batch_id:
            try:
                log_file = os.path.join(script_dir, 'bankcheck.log')
                result_data = {
                    'total_records': len(result.all_rows),
                    'new_records': result.new_record_count,
                    'duplicate_records': result.duplicate_record_count,
                    'processed_files': result.processed_files,
                    'unprocessed_files': result.unprocessed_files,
                    'error_files': result.error_files,
                    'incremental_mode': result.incremental_mode,
                    'output_folder': folder,
                    'summary_table_path': result.output_path,
                    'log_file_path': log_file,
                    'audit_id': audit.audit_id,
                }
                status = 'success' if result.all_rows or result.existing_record_count > 0 else 'warning'
                if result.error_files:
                    status = 'warning'
                batch_manager.finish_batch(batch_id, result_data, status=status)
                msg += f'\n批次号: {batch_id}'
                msg += f'\n归档目录: {batch_info.batch_dir}'
            except Exception as e:
                logger.error('批次归档失败: %s', e, exc_info=True)
                if batch_manager:
                    try:
                        batch_manager.finish_batch(batch_id, {}, status='failed', error_message=str(e))
                    except Exception:
                        pass

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

    change_result = detect_and_record_lookup_change(script_dir)
    if change_result.has_changes:
        add_count = sum(1 for c in change_result.change_details if c['change_type'] == 'add')
        remove_count = sum(1 for c in change_result.change_details if c['change_type'] == 'remove')
        modify_count = sum(1 for c in change_result.change_details if c['change_type'] == 'modify')
        warning_msg = (
            f'检测到查找表已变更！\n\n'
            f'变更编号: {change_result.change_id}\n'
            f'新增记录: {add_count} 条\n'
            f'删除记录: {remove_count} 条\n'
            f'修改记录: {modify_count} 条\n\n'
        )
        show_warning('查找表变更提醒', warning_msg)
        logger.warning('查找表变更已记录: %s', change_result.change_id)

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
            if change_result.change_id:
                msg += f'\n配置变更编号: {change_result.change_id}'
            show_info(title, msg)
        except FileNotFoundError as e:
            show_warning('错误', str(e))
            logger.error('对比失败: %s', e)


# ──────────────────────────────────────────────
# 定时批处理调度模块
# ──────────────────────────────────────────────

SCHEDULER_CONFIG_FILENAME = 'scheduler_config.json'
PROCESSED_FILES_DB_FILENAME = 'processed_files.db'
SCHEDULE_LOG_FILENAME = 'scheduler.log'


@dataclass
class ScheduleJobConfig:
    job_id: str
    name: str
    watch_directory: str
    cron_expression: str = '0 0 * * *'
    interval_minutes: Optional[int] = None
    schedule_type: str = 'cron'
    incremental: bool = True
    enabled: bool = True
    description: str = ''
    created_at: str = ''
    updated_at: str = ''


@dataclass
class ProcessedFileRecord:
    file_path: str
    file_hash: str
    file_size: int
    last_modified: float
    processed_at: str
    job_id: str
    record_count: int = 0
    status: str = 'success'


def get_scheduler_config_path(script_dir=None):
    if script_dir is None:
        script_dir = get_script_dir()
    return os.path.join(script_dir, SCHEDULER_CONFIG_FILENAME)


def get_processed_files_db_path(script_dir=None):
    if script_dir is None:
        script_dir = get_script_dir()
    return os.path.join(script_dir, PROCESSED_FILES_DB_FILENAME)


def get_scheduler_log_path(script_dir=None):
    if script_dir is None:
        script_dir = get_script_dir()
    return os.path.join(script_dir, SCHEDULE_LOG_FILENAME)


def init_processed_files_db(db_path=None):
    if db_path is None:
        db_path = get_processed_files_db_path()

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS processed_files (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            file_path TEXT NOT NULL,
            file_hash TEXT NOT NULL,
            file_size INTEGER NOT NULL,
            last_modified REAL NOT NULL,
            processed_at TEXT NOT NULL,
            job_id TEXT NOT NULL,
            record_count INTEGER DEFAULT 0,
            status TEXT DEFAULT 'success',
            UNIQUE(file_path, file_hash)
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS scheduler_runs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            run_id TEXT NOT NULL UNIQUE,
            job_id TEXT NOT NULL,
            job_name TEXT NOT NULL,
            started_at TEXT NOT NULL,
            completed_at TEXT,
            status TEXT NOT NULL,
            files_scanned INTEGER DEFAULT 0,
            files_new INTEGER DEFAULT 0,
            files_processed INTEGER DEFAULT 0,
            files_skipped INTEGER DEFAULT 0,
            files_error INTEGER DEFAULT 0,
            records_extracted INTEGER DEFAULT 0,
            error_message TEXT,
            output_path TEXT,
            duration_ms INTEGER
        )
    ''')

    cursor.execute('CREATE INDEX IF NOT EXISTS idx_processed_files_job ON processed_files(job_id)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_processed_files_path ON processed_files(file_path)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_scheduler_runs_job ON scheduler_runs(job_id)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_scheduler_runs_started ON scheduler_runs(started_at)')

    conn.commit()
    conn.close()

    logger = get_logger()
    logger.debug('已处理文件数据库初始化完成: %s', db_path)


def load_scheduler_config(script_dir=None):
    config_path = get_scheduler_config_path(script_dir)
    logger = get_logger()

    if not os.path.exists(config_path):
        logger.info('调度配置文件不存在，将创建默认配置: %s', config_path)
        default_config = {
            'jobs': [],
            'settings': {
                'max_concurrent_jobs': 1,
                'check_interval_seconds': 30,
                'enable_alerts': True,
                'log_level': 'INFO'
            }
        }
        save_scheduler_config(default_config, script_dir)
        return default_config

    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)
        logger.info('已加载调度配置，共 %d 个任务', len(config.get('jobs', [])))
        return config
    except Exception as e:
        logger.error('加载调度配置失败: %s', e)
        return {'jobs': [], 'settings': {}}


def save_scheduler_config(config, script_dir=None):
    config_path = get_scheduler_config_path(script_dir)
    try:
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
        logger = get_logger()
        logger.info('调度配置已保存: %s', config_path)
        return True
    except Exception as e:
        logger = get_logger()
        logger.error('保存调度配置失败: %s', e)
        return False


def add_schedule_job(job_config, script_dir=None):
    config = load_scheduler_config(script_dir)
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    if not job_config.get('job_id'):
        job_config['job_id'] = f"JOB{datetime.now().strftime('%Y%m%d%H%M%S')}{uuid.uuid4().hex[:6].upper()}"

    job_config['created_at'] = now
    job_config['updated_at'] = now
    job_config['enabled'] = job_config.get('enabled', True)

    config['jobs'].append(job_config)
    save_scheduler_config(config, script_dir)

    logger = get_logger()
    logger.info('已添加调度任务 [%s] %s', job_config['job_id'], job_config.get('name', ''))
    return job_config['job_id']


def update_schedule_job(job_id, updates, script_dir=None):
    config = load_scheduler_config(script_dir)
    logger = get_logger()

    for job in config['jobs']:
        if job['job_id'] == job_id:
            job.update(updates)
            job['updated_at'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            save_scheduler_config(config, script_dir)
            logger.info('已更新调度任务 [%s]', job_id)
            return True

    logger.warning('未找到调度任务 [%s]', job_id)
    return False


def remove_schedule_job(job_id, script_dir=None):
    config = load_scheduler_config(script_dir)
    logger = get_logger()

    config['jobs'] = [job for job in config['jobs'] if job['job_id'] != job_id]
    save_scheduler_config(config, script_dir)
    logger.info('已删除调度任务 [%s]', job_id)


def list_schedule_jobs(script_dir=None):
    config = load_scheduler_config(script_dir)
    return config.get('jobs', [])


def is_file_processed(file_path, job_id, db_path=None):
    if db_path is None:
        db_path = get_processed_files_db_path()

    file_hash = compute_file_hash(file_path)
    if not file_hash:
        return False, None

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute('''
        SELECT file_path, processed_at FROM processed_files
        WHERE file_path = ? AND file_hash = ? AND job_id = ?
    ''', (file_path, file_hash, job_id))
    result = cursor.fetchone()
    conn.close()

    return result is not None, file_hash


def mark_file_processed(file_path, file_hash, job_id, record_count=0, status='success', db_path=None):
    if db_path is None:
        db_path = get_processed_files_db_path()
    init_processed_files_db(db_path)

    file_size = os.path.getsize(file_path) if os.path.exists(file_path) else 0
    last_modified = os.path.getmtime(file_path) if os.path.exists(file_path) else 0
    processed_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    try:
        cursor.execute('''
            INSERT OR REPLACE INTO processed_files (
                file_path, file_hash, file_size, last_modified,
                processed_at, job_id, record_count, status
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        ''', (file_path, file_hash, file_size, last_modified,
              processed_at, job_id, record_count, status))
        conn.commit()
    finally:
        conn.close()

    logger = get_logger()
    logger.debug('已标记文件为已处理: %s (job=%s)', file_path, job_id)


def scan_new_files(watch_directory, job_id, script_dir=None):
    logger = get_logger()
    db_path = get_processed_files_db_path(script_dir)
    init_processed_files_db(db_path)

    all_files = scan_excel_files(watch_directory)
    new_files = []

    for file_path in all_files:
        processed, file_hash = is_file_processed(file_path, job_id, db_path)
        if not processed:
            new_files.append((file_path, file_hash))
            logger.debug('发现新文件: %s', file_path)

    logger.info('扫描目录 %s: 共 %d 个文件，新增 %d 个',
                watch_directory, len(all_files), len(new_files))
    return new_files


def record_scheduler_run(run_data, script_dir=None):
    db_path = get_processed_files_db_path(script_dir)
    init_processed_files_db(db_path)

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    try:
        cursor.execute('''
            INSERT INTO scheduler_runs (
                run_id, job_id, job_name, started_at, completed_at,
                status, files_scanned, files_new, files_processed,
                files_skipped, files_error, records_extracted,
                error_message, output_path, duration_ms
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            run_data['run_id'], run_data['job_id'], run_data['job_name'],
            run_data['started_at'], run_data.get('completed_at'),
            run_data['status'], run_data.get('files_scanned', 0),
            run_data.get('files_new', 0), run_data.get('files_processed', 0),
            run_data.get('files_skipped', 0), run_data.get('files_error', 0),
            run_data.get('records_extracted', 0), run_data.get('error_message'),
            run_data.get('output_path'), run_data.get('duration_ms')
        ))
        conn.commit()
    finally:
        conn.close()


def run_scheduled_pipeline(job_config, script_dir=None):
    if script_dir is None:
        script_dir = get_script_dir()

    logger = get_logger()
    job_id = job_config['job_id']
    job_name = job_config.get('name', job_id)
    watch_directory = job_config['watch_directory']
    incremental = job_config.get('incremental', True)

    logger.info('========== 定时任务启动 [%s] %s ==========', job_id, job_name)
    logger.info('监控目录: %s', watch_directory)
    logger.info('运行模式: %s', '增量合并' if incremental else '全量覆盖')

    run_id = f"SCH{datetime.now().strftime('%Y%m%d%H%M%S')}{uuid.uuid4().hex[:8].upper()}"
    start_time = datetime.now()

    run_data = {
        'run_id': run_id,
        'job_id': job_id,
        'job_name': job_name,
        'started_at': start_time.strftime('%Y-%m-%d %H:%M:%S.%f'),
        'status': 'running',
    }

    try:
        change_result = detect_and_record_lookup_change(script_dir, username='scheduler')
        if change_result.has_changes:
            logger.warning('检测到查找表变更: %s', change_result.change_id)

        if not os.path.exists(watch_directory):
            raise FileNotFoundError(f'监控目录不存在: {watch_directory}')

        new_files = scan_new_files(watch_directory, job_id, script_dir)
        run_data['files_scanned'] = len(scan_excel_files(watch_directory))
        run_data['files_new'] = len(new_files)

        if not new_files:
            logger.info('没有新文件需要处理，任务结束')
            run_data['status'] = 'success'
            run_data['completed_at'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')
            run_data['duration_ms'] = int((datetime.now() - start_time).total_seconds() * 1000)
            return run_data

        with AuditLogger('scheduled_pipeline', script_dir, username='scheduler') as audit:
            audit.record_input(watch_directory)

            result = run_pipeline(watch_directory, script_dir, incremental=incremental)
            audit.record_result(result)

            db_path = get_processed_files_db_path(script_dir)
            for file_path, file_hash in new_files:
                record_count = 0
                status = 'success'
                for proc_file in result.processed_files:
                    if os.path.basename(file_path) in proc_file:
                        record_count = len(result.all_rows)
                        break
                for err_file, _ in result.error_files:
                    if os.path.basename(file_path) in err_file:
                        status = 'error'
                        break

                mark_file_processed(file_path, file_hash, job_id, record_count, status, db_path)

            run_data['files_processed'] = len(result.processed_files)
            run_data['files_error'] = len(result.error_files)
            run_data['files_skipped'] = len(result.unprocessed_files)
            run_data['records_extracted'] = result.new_record_count
            run_data['output_path'] = result.output_path
            run_data['status'] = 'success' if not result.error_files else 'partial'

            if result.lookup_missing:
                logger.warning('未找到主体查找表，主体列为空')

            logger.info('定时任务完成: 新增文件 %d 个，处理 %d 个，错误 %d 个，提取记录 %d 条',
                        len(new_files), len(result.processed_files),
                        len(result.error_files), result.new_record_count)

            if result.output_path:
                logger.info('输出总表: %s', result.output_path)

    except Exception as e:
        logger.error('定时任务执行失败: %s', e, exc_info=True)
        run_data['status'] = 'failed'
        run_data['error_message'] = str(e)

    finally:
        run_data['completed_at'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')
        run_data['duration_ms'] = int((datetime.now() - start_time).total_seconds() * 1000)
        record_scheduler_run(run_data, script_dir)

        if run_data['status'] == 'failed' and load_scheduler_config(script_dir).get('settings', {}).get('enable_alerts', True):
            run_alert_detection(script_dir, username='scheduler')

    logger.info('========== 定时任务结束 [%s] 状态: %s 耗时: %dms ==========',
                job_id, run_data['status'], run_data['duration_ms'])

    return run_data


def _try_import_apscheduler():
    try:
        from apscheduler.schedulers.background import BackgroundScheduler
        from apscheduler.triggers.cron import CronTrigger
        from apscheduler.triggers.interval import IntervalTrigger
        return BackgroundScheduler, CronTrigger, IntervalTrigger
    except ImportError:
        logger = get_logger()
        logger.warning('未安装 APScheduler，将使用简单内置调度器。请运行: pip install APScheduler')
        return None, None, None


class SimpleScheduler:
    def __init__(self):
        self.jobs = []
        self.running = False
        self._thread = None
        self.logger = get_logger()

    def add_job(self, func, trigger, args=None, kwargs=None, id=None, name=None):
        job = {
            'id': id or f"JOB{uuid.uuid4().hex[:8].upper()}",
            'name': name or func.__name__,
            'func': func,
            'trigger': trigger,
            'args': args or [],
            'kwargs': kwargs or {},
            'last_run': None,
        }
        self.jobs.append(job)
        self.logger.info('已添加调度任务: %s (%s)', job['id'], job['name'])
        return job

    def _should_run(self, job, now):
        trigger = job['trigger']

        if job['last_run'] is None:
            return True

        if hasattr(trigger, 'get_interval'):
            interval_sec = trigger.get_interval().total_seconds()
            return (now - job['last_run']).total_seconds() >= interval_sec

        if hasattr(trigger, 'fields'):
            from datetime import datetime as dt
            last_run_date = job['last_run'].replace(second=0, microsecond=0)
            now_date = now.replace(second=0, microsecond=0)
            return now_date > last_run_date

        return False

    def _run_loop(self):
        import time
        while self.running:
            now = datetime.now()
            for job in self.jobs:
                if self._should_run(job, now):
                    try:
                        self.logger.info('触发定时任务: %s', job['name'])
                        job['last_run'] = now
                        job['func'](*job['args'], **job['kwargs'])
                    except Exception as e:
                        self.logger.error('任务执行错误 [%s]: %s', job['name'], e, exc_info=True)
            time.sleep(30)

    def start(self):
        import threading
        if self.running:
            return
        self.running = True
        self._thread = threading.Thread(target=self._run_loop, daemon=True)
        self._thread.start()
        self.logger.info('简单调度器已启动')

    def shutdown(self):
        self.running = False
        if self._thread:
            self._thread.join(timeout=5)
        self.logger.info('简单调度器已停止')


def get_scheduler():
    BackgroundScheduler, CronTrigger, IntervalTrigger = _try_import_apscheduler()

    if BackgroundScheduler:
        scheduler = BackgroundScheduler(timezone='Asia/Shanghai')
        scheduler_type = 'apscheduler'
    else:
        scheduler = SimpleScheduler()
        CronTrigger = None
        IntervalTrigger = None
        scheduler_type = 'simple'

    return scheduler, scheduler_type, CronTrigger, IntervalTrigger


def start_scheduler(script_dir=None):
    if script_dir is None:
        script_dir = get_script_dir()

    logger = get_logger()
    logger.info('========== 定时调度器启动 ==========')

    config = load_scheduler_config(script_dir)
    jobs = [j for j in config.get('jobs', []) if j.get('enabled', True)]

    if not jobs:
        logger.warning('没有启用的调度任务，请先配置任务')
        return None

    init_processed_files_db(get_processed_files_db_path(script_dir))

    scheduler, scheduler_type, CronTrigger, IntervalTrigger = get_scheduler()

    for job_config in jobs:
        job_id = job_config['job_id']
        job_name = job_config.get('name', job_id)
        schedule_type = job_config.get('schedule_type', 'cron')

        try:
            if scheduler_type == 'apscheduler':
                if schedule_type == 'interval' and job_config.get('interval_minutes'):
                    trigger = IntervalTrigger(minutes=job_config['interval_minutes'])
                else:
                    cron_expr = job_config.get('cron_expression', '0 0 * * *')
                    parts = cron_expr.split()
                    if len(parts) == 5:
                        minute, hour, day, month, day_of_week = parts
                        trigger = CronTrigger(
                            minute=minute, hour=hour, day=day,
                            month=month, day_of_week=day_of_week,
                            timezone='Asia/Shanghai'
                        )
                    else:
                        logger.warning('cron 表达式格式错误，使用默认每日0点: %s', cron_expr)
                        trigger = CronTrigger(hour=0, minute=0, timezone='Asia/Shanghai')

                scheduler.add_job(
                    run_scheduled_pipeline,
                    trigger=trigger,
                    args=[job_config, script_dir],
                    id=job_id,
                    name=job_name,
                    misfire_grace_time=3600,
                    coalesce=True,
                    max_instances=1,
                )
            else:
                class SimpleIntervalTrigger:
                    def __init__(self, minutes):
                        self._minutes = minutes
                    def get_interval(self):
                        from datetime import timedelta
                        return timedelta(minutes=self._minutes)

                class SimpleCronTrigger:
                    def __init__(self, cron_expr):
                        self.fields = cron_expr.split()

                if schedule_type == 'interval' and job_config.get('interval_minutes'):
                    trigger = SimpleIntervalTrigger(job_config['interval_minutes'])
                else:
                    trigger = SimpleCronTrigger(job_config.get('cron_expression', '0 0 * * *'))

                scheduler.add_job(
                    run_scheduled_pipeline,
                    trigger=trigger,
                    args=[job_config, script_dir],
                    id=job_id,
                    name=job_name,
                )

            logger.info('已注册任务 [%s] %s: 类型=%s, 配置=%s',
                        job_id, job_name, schedule_type,
                        job_config.get('cron_expression') or f"{job_config.get('interval_minutes')}分钟")

        except Exception as e:
            logger.error('注册任务失败 [%s]: %s', job_id, e)

    try:
        scheduler.start()
        logger.info('调度器已启动 (%s)，共 %d 个任务', scheduler_type, len(jobs))
        logger.info('按 Ctrl+C 停止调度器')

        if scheduler_type == 'apscheduler':
            import time
            try:
                while True:
                    time.sleep(60)
            except (KeyboardInterrupt, SystemExit):
                logger.info('收到停止信号')
                scheduler.shutdown()
        else:
            import time
            try:
                while scheduler.running:
                    time.sleep(60)
            except (KeyboardInterrupt, SystemExit):
                logger.info('收到停止信号')
                scheduler.shutdown()

    except Exception as e:
        logger.error('调度器运行错误: %s', e)
        scheduler.shutdown()

    logger.info('========== 定时调度器停止 ==========')
    return scheduler


def generate_cron_script(job_config, script_path, script_dir=None):
    if script_dir is None:
        script_dir = get_script_dir()

    logger = get_logger()
    job_id = job_config['job_id']
    cron_expr = job_config.get('cron_expression', '0 0 * * *')

    python_path = sys.executable
    script_abs_path = os.path.abspath(script_path)

    log_path = get_scheduler_log_path(script_dir)

    cron_line = f'{cron_expr} {python_path} {script_abs_path} --run-job {job_id} >> {log_path} 2>&1'

    script_content = f'''#!/bin/bash
# 银行流水检验工具 - 定时任务 cron 配置
# 任务ID: {job_id}
# 任务名称: {job_config.get('name', '')}
# 监控目录: {job_config.get('watch_directory', '')}
# 调度时间: {cron_expr}
# 生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

# 添加到 crontab 的命令:
# crontab -e
# 然后添加以下行:

{cron_line}

# 或者运行以下命令自动添加:
# (crontab -l 2>/dev/null; echo "{cron_line}") | crontab -

# 查看当前 crontab:
# crontab -l

# 日志文件: {log_path}
'''

    output_path = os.path.join(script_dir, f'cron_job_{job_id}.sh')
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(script_content)

    os.chmod(output_path, 0o755)
    logger.info('已生成 cron 脚本: %s', output_path)
    return output_path, cron_line


def generate_windows_task_script(job_config, script_path, script_dir=None):
    if script_dir is None:
        script_dir = get_script_dir()

    logger = get_logger()
    job_id = job_config['job_id']
    job_name = job_config.get('name', f'bankcheck_{job_id}')
    cron_expr = job_config.get('cron_expression', '0 0 * * *')

    parts = cron_expr.split()
    start_time = '00:00'
    if len(parts) >= 2:
        start_time = f'{parts[1].zfill(2)}:{parts[0].zfill(2)}'

    python_path = sys.executable
    script_abs_path = os.path.abspath(script_path)
    log_path = get_scheduler_log_path(script_dir)

    schtasks_cmd = (
        f'SchTasks /Create /SC DAILY /TN "{job_name}" '
        f'/TR "\\"{python_path}\\" \\"{script_abs_path}\\" --run-job {job_id}" '
        f'/ST {start_time} /RL HIGHEST /F'
    )

    ps_content = f'''# 银行流水检验工具 - Windows 任务计划配置
# 任务ID: {job_id}
# 任务名称: {job_name}
# 监控目录: {job_config.get('watch_directory', '')}
# 调度时间: {cron_expr} (每日 {start_time})
# 生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

# 方法1: 以管理员身份运行此 PowerShell 脚本
$action = New-ScheduledTaskAction -Execute "{python_path}" -Argument "`"{script_abs_path}`" --run-job {job_id}"
$trigger = New-ScheduledTaskTrigger -Daily -At {start_time}
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
Register-ScheduledTask -TaskName "{job_name}" -Action $action -Trigger $trigger -Settings $settings -Description "银行流水检验工具定时任务" -Force
Write-Host "任务已创建: {job_name}"

# 方法2: 在命令行(cmd)中运行以下命令:
# {schtasks_cmd}

# 查看任务:
# Get-ScheduledTask -TaskName "{job_name}"

# 删除任务:
# Unregister-ScheduledTask -TaskName "{job_name}" -Confirm:$false

# 日志文件: {log_path}
'''

    bat_content = f'''@echo off
REM 银行流水检验工具 - Windows 任务计划配置
REM 任务ID: {job_id}
REM 任务名称: {job_name}
REM 监控目录: {job_config.get('watch_directory', '')}
REM 调度时间: {cron_expr} (每日 {start_time})
REM 生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

echo 创建 Windows 任务计划: {job_name}
{schtasks_cmd}

if %errorlevel%==0 (
    echo 任务创建成功！
    echo 查看任务: schtasks /Query /TN "{job_name}"
    echo 删除任务: schtasks /Delete /TN "{job_name}" /F
) else (
    echo 任务创建失败，请以管理员身份运行。
)

pause
'''

    ps_output = os.path.join(script_dir, f'task_job_{job_id}.ps1')
    bat_output = os.path.join(script_dir, f'task_job_{job_id}.bat')

    with open(ps_output, 'w', encoding='utf-8') as f:
        f.write(ps_content)

    with open(bat_output, 'w', encoding='gbk') as f:
        f.write(bat_content)

    logger.info('已生成 Windows 任务计划脚本: %s, %s', ps_output, bat_output)
    return ps_output, bat_output


def show_scheduler_menu(script_dir=None):
    if script_dir is None:
        script_dir = get_script_dir()
    logger = get_logger()

    while True:
        jobs = list_schedule_jobs(script_dir)

        print('\n' + '=' * 80)
        print('⏰ 定时批处理调度管理'.center(70))
        print('=' * 80)

        if jobs:
            print(f'{"ID":<14} {"名称":<20} {"类型":<10} {"调度":<20} {"状态":<8} {"目录":<30}')
            print('-' * 80)
            for job in jobs:
                status = '✅启用' if job.get('enabled', True) else '❌禁用'
                sched_type = job.get('schedule_type', 'cron')
                if sched_type == 'interval':
                    sched_display = f"每{job.get('interval_minutes', 60)}分钟"
                else:
                    sched_display = job.get('cron_expression', '0 0 * * *')
                watch_dir = job.get('watch_directory', '')[:28]
                print(f'{job["job_id"][:12]:<14} {job.get("name", "")[:18]:<20} {sched_type:<10} {sched_display[:18]:<20} {status:<8} {watch_dir:<30}')
        else:
            print('暂无调度任务，请先添加任务')

        print('\n请选择操作：')
        print('  1) ➕ 添加定时任务')
        print('  2) ✏️  编辑任务')
        print('  3) 🗑️  删除任务')
        print('  4) ▶️  启动调度器')
        print('  5) 🏃  立即运行指定任务')
        print('  6) 📜  查看任务运行历史')
        print('  7) 📄  生成系统任务脚本')
        print('  8) ⚙️  调度器设置')
        print('  0) 返回主菜单')

        choice = input('\n请输入选项: ').strip()

        if choice == '0':
            break
        elif choice == '1':
            _add_job_interactive(script_dir)
        elif choice == '2':
            _edit_job_interactive(script_dir)
        elif choice == '3':
            _remove_job_interactive(script_dir)
        elif choice == '4':
            start_scheduler(script_dir)
        elif choice == '5':
            _run_job_now_interactive(script_dir)
        elif choice == '6':
            _show_scheduler_history(script_dir)
        elif choice == '7':
            _generate_system_task_script(script_dir)
        elif choice == '8':
            _scheduler_settings_menu(script_dir)
        else:
            print('无效选项')
            input('按回车继续...')


def _add_job_interactive(script_dir=None):
    print('\n--- 添加定时任务 ---')

    name = input('任务名称: ').strip()
    if not name:
        print('名称不能为空')
        input('按回车继续...')
        return

    watch_directory = input('监控目录路径: ').strip().strip('"').strip("'")
    if not watch_directory or not os.path.isdir(watch_directory):
        print('目录不存在')
        input('按回车继续...')
        return

    print('\n调度类型:')
    print('  1) cron 表达式 (推荐)')
    print('  2) 固定间隔(分钟)')
    type_choice = input('请选择 (默认1): ').strip()

    job_config = {
        'name': name,
        'watch_directory': watch_directory,
        'enabled': True,
    }

    if type_choice == '2':
        interval_input = input('间隔分钟数 (默认60): ').strip()
        interval = int(interval_input) if interval_input.isdigit() else 60
        job_config['schedule_type'] = 'interval'
        job_config['interval_minutes'] = interval
    else:
        print('\ncron 表达式格式: 分 时 日 月 周')
        print('  示例:')
        print('    0 0 * * *    = 每天 00:00')
        print('    0 2 * * *    = 每天 02:00')
        print('    0 */6 * * *  = 每6小时')
        print('    0 8 * * 1-5  = 周一至周五 08:00')
        cron_expr = input('请输入 cron 表达式 (默认 0 0 * * *): ').strip() or '0 0 * * *'
        job_config['schedule_type'] = 'cron'
        job_config['cron_expression'] = cron_expr

    inc_choice = input('启用增量合并? (Y/n): ').strip().lower()
    job_config['incremental'] = inc_choice != 'n'

    job_config['description'] = input('任务描述 (可选): ').strip()

    job_id = add_schedule_job(job_config, script_dir)
    print(f'✅ 任务已添加，ID: {job_id}')
    input('按回车继续...')


def _edit_job_interactive(script_dir=None):
    jobs = list_schedule_jobs(script_dir)
    if not jobs:
        print('暂无任务')
        input('按回车继续...')
        return

    job_id = input('请输入要编辑的任务ID: ').strip()
    job = next((j for j in jobs if j['job_id'] == job_id), None)
    if not job:
        print('未找到该任务')
        input('按回车继续...')
        return

    print(f'\n当前配置:')
    print(f'  名称: {job.get("name", "")}')
    print(f'  目录: {job.get("watch_directory", "")}')
    print(f'  类型: {job.get("schedule_type", "cron")}')
    if job.get('schedule_type') == 'interval':
        print(f'  间隔: {job.get("interval_minutes", 60)} 分钟')
    else:
        print(f'  cron: {job.get("cron_expression", "")}')
    print(f'  增量: {"启用" if job.get("incremental", True) else "禁用"}')
    print(f'  状态: {"启用" if job.get("enabled", True) else "禁用"}')

    print(f'\n编辑 (留空保持当前值):')
    updates = {}

    new_name = input(f'新名称 [{job.get("name", "")}]: ').strip()
    if new_name:
        updates['name'] = new_name

    new_dir = input(f'新目录 [{job.get("watch_directory", "")}]: ').strip().strip('"').strip("'")
    if new_dir:
        if os.path.isdir(new_dir):
            updates['watch_directory'] = new_dir
        else:
            print('目录不存在，跳过')

    new_enabled = input(f'启用? (y/n) [{job.get("enabled", True)}]: ').strip().lower()
    if new_enabled in ['y', 'n']:
        updates['enabled'] = new_enabled == 'y'

    new_inc = input(f'增量合并? (y/n) [{job.get("incremental", True)}]: ').strip().lower()
    if new_inc in ['y', 'n']:
        updates['incremental'] = new_inc == 'y'

    if updates:
        update_schedule_job(job_id, updates, script_dir)
        print('✅ 任务已更新')
    else:
        print('未做修改')

    input('按回车继续...')


def _remove_job_interactive(script_dir=None):
    jobs = list_schedule_jobs(script_dir)
    if not jobs:
        print('暂无任务')
        input('按回车继续...')
        return

    job_id = input('请输入要删除的任务ID: ').strip()
    job = next((j for j in jobs if j['job_id'] == job_id), None)
    if not job:
        print('未找到该任务')
        input('按回车继续...')
        return

    confirm = input(f'确认删除任务 [{job.get("name", job_id)}]? (y/N): ').strip().lower()
    if confirm == 'y':
        remove_schedule_job(job_id, script_dir)
        print('✅ 任务已删除')
    else:
        print('已取消')

    input('按回车继续...')


def _run_job_now_interactive(script_dir=None):
    jobs = list_schedule_jobs(script_dir)
    if not jobs:
        print('暂无任务')
        input('按回车继续...')
        return

    job_id = input('请输入要立即运行的任务ID: ').strip()
    job = next((j for j in jobs if j['job_id'] == job_id), None)
    if not job:
        print('未找到该任务')
        input('按回车继续...')
        return

    print(f'🚀 立即运行任务: {job.get("name", job_id)}')
    result = run_scheduled_pipeline(job, script_dir)
    print(f'\n执行完成，状态: {result["status"]}')
    print(f'扫描文件: {result.get("files_scanned", 0)} 个')
    print(f'新增文件: {result.get("files_new", 0)} 个')
    print(f'处理文件: {result.get("files_processed", 0)} 个')
    print(f'提取记录: {result.get("records_extracted", 0)} 条')
    print(f'耗时: {result.get("duration_ms", 0)} ms')
    if result.get("output_path"):
        print(f'输出: {result["output_path"]}')

    input('按回车继续...')


def _show_scheduler_history(script_dir=None):
    if script_dir is None:
        script_dir = get_script_dir()

    db_path = get_processed_files_db_path(script_dir)
    if not os.path.exists(db_path):
        print('暂无运行历史')
        input('按回车继续...')
        return

    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute('''
        SELECT * FROM scheduler_runs
        ORDER BY started_at DESC
        LIMIT 50
    ''')
    rows = cursor.fetchall()
    conn.close()

    print('\n' + '=' * 80)
    print('📜 调度任务运行历史'.center(70))
    print('=' * 80)
    print(f'{"运行ID":<14} {"任务":<18} {"开始时间":<20} {"状态":<10} {"文件":<8} {"记录":<8}')
    print('-' * 80)

    for row in rows:
        run_id = row['run_id'][:12]
        job_name = row['job_name'][:16]
        started = row['started_at'][:19]
        status_sym = _get_status_symbol(row['status'])
        files_new = row['files_new'] or 0
        records = row['records_extracted'] or 0
        print(f'{run_id:<14} {job_name:<18} {started:<20} {status_sym} {row["status"]:<8} {files_new:<8} {records:<8}')

    if not rows:
        print('暂无运行记录')

    input('\n按回车继续...')


def _generate_system_task_script(script_dir=None):
    jobs = list_schedule_jobs(script_dir)
    if not jobs:
        print('暂无任务')
        input('按回车继续...')
        return

    job_id = input('请输入任务ID: ').strip()
    job = next((j for j in jobs if j['job_id'] == job_id), None)
    if not job:
        print('未找到该任务')
        input('按回车继续...')
        return

    script_path = os.path.abspath(__file__)

    print('\n选择目标系统:')
    print('  1) Linux/macOS (cron)')
    print('  2) Windows (任务计划)')
    sys_choice = input('请选择 (默认1): ').strip()

    if sys_choice == '2':
        ps_output, bat_output = generate_windows_task_script(job, script_path, script_dir)
        print(f'\n✅ 已生成 Windows 任务计划脚本:')
        print(f'  PowerShell: {ps_output}')
        print(f'  批处理: {bat_output}')
    else:
        output_path, cron_line = generate_cron_script(job, script_path, script_dir)
        print(f'\n✅ 已生成 cron 脚本: {output_path}')
        print(f'\n手动添加到 crontab:')
        print(f'  {cron_line}')

    input('按回车继续...')


def _scheduler_settings_menu(script_dir=None):
    config = load_scheduler_config(script_dir)
    settings = config.get('settings', {})

    while True:
        print('\n--- 调度器设置 ---')
        print(f'  1) 告警通知: {"启用" if settings.get("enable_alerts", True) else "禁用"}')
        print(f'  2) 最大并发任务数: {settings.get("max_concurrent_jobs", 1)}')
        print(f'  3) 检查间隔(秒): {settings.get("check_interval_seconds", 30)}')
        print(f'  4) 日志级别: {settings.get("log_level", "INFO")}')
        print('  0) 返回')

        choice = input('\n请选择: ').strip()

        if choice == '0':
            break
        elif choice == '1':
            settings['enable_alerts'] = not settings.get('enable_alerts', True)
        elif choice == '2':
            val = input('最大并发任务数: ').strip()
            if val.isdigit():
                settings['max_concurrent_jobs'] = int(val)
        elif choice == '3':
            val = input('检查间隔(秒): ').strip()
            if val.isdigit():
                settings['check_interval_seconds'] = int(val)
        elif choice == '4':
            level = input('日志级别 (DEBUG/INFO/WARNING/ERROR): ').strip().upper()
            if level in ['DEBUG', 'INFO', 'WARNING', 'ERROR']:
                settings['log_level'] = level

        config['settings'] = settings
        save_scheduler_config(config, script_dir)
        print('✅ 设置已保存')


def run_scheduler_flow(script_dir):
    logger = get_logger()
    logger.info('========== 定时调度管理启动 ==========')
    init_default_alert_rules(script_dir)
    init_processed_files_db(get_processed_files_db_path(script_dir))
    show_scheduler_menu(script_dir)
    logger.info('========== 定时调度管理关闭 ==========')


def parse_args_and_run():
    import argparse

    parser = argparse.ArgumentParser(description='银行流水检验工具')
    parser.add_argument('--scheduler', action='store_true', help='启动定时调度器')
    parser.add_argument('--scheduler-menu', action='store_true', help='打开调度管理菜单')
    parser.add_argument('--run-job', type=str, metavar='JOB_ID', help='立即运行指定ID的定时任务')
    parser.add_argument('--add-job', action='store_true', help='交互式添加定时任务')
    parser.add_argument('--list-jobs', action='store_true', help='列出所有定时任务')
    parser.add_argument('--watch-dir', type=str, metavar='DIR', help='单次运行: 监控目录路径')
    parser.add_argument('--once', action='store_true', help='单次运行后退出(配合--watch-dir使用)')
    parser.add_argument('--interval', type=int, metavar='MINUTES', help='单次运行模式下的间隔分钟数')
    parser.add_argument('--no-incremental', action='store_true', help='禁用增量合并，使用全量覆盖')
    parser.add_argument('--export', action='store_true', help='进入财务软件对接导出功能')
    parser.add_argument('--export-template', type=str, metavar='TEMPLATE',
                       help='指定导出模板类型: yonyou_voucher(用友), kingdee_voucher(金蝶), bank_journal(日记账)')
    parser.add_argument('--export-total', type=str, metavar='TOTAL_FILE',
                       help='指定总表文件路径用于导出')
    parser.add_argument('--export-output', type=str, metavar='OUTPUT_DIR',
                       help='指定导出文件输出目录')
    parser.add_argument('--export-operator', type=str, metavar='OPERATOR',
                       help='指定制单人名称')
    parser.add_argument('--preset-menu', action='store_true', help='打开预设管理菜单')
    parser.add_argument('--list-presets', action='store_true', help='列出所有预设')
    parser.add_argument('--apply-preset', type=str, metavar='PRESET_ID',
                       help='应用指定ID的预设，需配合--watch-dir使用')
    parser.add_argument('--save-preset', type=str, metavar='NAME',
                       help='保存当前配置为新预设')

    args = parser.parse_args()

    script_dir = get_script_dir()
    setup_logging()
    logger = get_logger()

    init_audit_db(get_audit_db_path(script_dir))
    init_default_alert_rules(script_dir)

    if args.list_jobs:
        jobs = list_schedule_jobs(script_dir)
        print(f'定时任务列表 (共 {len(jobs)} 个):')
        for job in jobs:
            status = '启用' if job.get('enabled', True) else '禁用'
            print(f'  [{job["job_id"]}] {job.get("name", "")} - {status}')
            print(f'    目录: {job.get("watch_directory", "")}')
            if job.get('schedule_type') == 'interval':
                print(f'    调度: 每 {job.get("interval_minutes", 60)} 分钟')
            else:
                print(f'    调度: {job.get("cron_expression", "")}')
        return True

    if args.list_presets:
        presets = list_presets(script_dir)
        print(f'预设列表 (共 {len(presets)} 个):')
        for preset in presets:
            banks = ', '.join(preset.get('enabled_banks', []))
            keep = KEEP_STRATEGIES.get(preset.get('keep_strategy'), '未知')
            print(f'  [{preset["preset_id"]}] {preset.get("name", "")}')
            print(f'    描述: {preset.get("description", "无")}')
            print(f'    银行: {banks}')
            print(f'    保留策略: {keep}')
            print(f'    增量: {"是" if preset.get("incremental", True) else "否"}')
            if preset.get('start_date') or preset.get('end_date'):
                print(f'    日期: {preset.get("start_date", "不限")} ~ {preset.get("end_date", "不限")}')
            print(f'    更新时间: {preset.get("updated_at", "")}')
        return True

    if args.apply_preset and args.watch_dir:
        preset = load_preset(args.apply_preset, script_dir)
        if not preset:
            logger.error('未找到预设ID: %s', args.apply_preset)
            print(f'错误: 未找到预设ID {args.apply_preset}')
            return True

        watch_dir = os.path.abspath(args.watch_dir)
        if not os.path.isdir(watch_dir):
            logger.error('目录不存在: %s', watch_dir)
            print(f'错误: 目录不存在 {watch_dir}')
            return True

        print(f'应用预设: {preset.get("name", "")} ({args.apply_preset})')
        print(f'处理目录: {watch_dir}')

        with AuditLogger('preset_pipeline', script_dir) as audit:
            audit.record_input(watch_dir)
            audit.set_extra_info({'preset_id': args.apply_preset, 'preset_name': preset.get('name', '')})
            result = apply_preset_to_pipeline(preset, watch_dir, script_dir)
            audit.record_result(result)

            if result.all_rows:
                print(f'\n✅ 处理完成！')
                print(f'   总记录数: {len(result.all_rows)}')
                print(f'   新增记录: {result.new_record_count}')
                print(f'   输出文件: {result.output_path}')
            else:
                print(f'\n⚠️  未提取到记录')
            return True

    if args.save_preset:
        preset_data = {
            'name': args.save_preset,
            'output_dir': args.export_output or '',
            'enabled_banks': list(BANK_PREFIXES),
            'keep_strategy': 'keep_unprocessed',
            'incremental': not args.no_incremental,
        }
        preset_id = save_preset(preset_data, script_dir)
        print(f'\n✅ 预设已保存，ID: {preset_id}')
        return True

    if args.add_job:
        _add_job_interactive(script_dir)
        return True

    if args.run_job:
        jobs = list_schedule_jobs(script_dir)
        job = next((j for j in jobs if j['job_id'] == args.run_job), None)
        if not job:
            logger.error('未找到任务ID: %s', args.run_job)
            return True
        run_scheduled_pipeline(job, script_dir)
        return True

    if args.watch_dir:
        watch_dir = os.path.abspath(args.watch_dir)
        if not os.path.isdir(watch_dir):
            logger.error('目录不存在: %s', watch_dir)
            return True

        job_config = {
            'job_id': 'ONETIME',
            'name': '单次运行任务',
            'watch_directory': watch_dir,
            'incremental': not args.no_incremental,
            'schedule_type': 'interval',
            'interval_minutes': args.interval or 60,
        }

        if args.once:
            logger.info('单次运行模式，目录: %s', watch_dir)
            run_scheduled_pipeline(job_config, script_dir)
            return True
        else:
            if not args.interval:
                logger.warning('未指定--interval，将使用默认60分钟间隔')

            logger.info('监控模式启动，目录: %s，间隔: %d 分钟，Ctrl+C 停止',
                        watch_dir, job_config['interval_minutes'])

            scheduler, scheduler_type, CronTrigger, IntervalTrigger = get_scheduler()

            if scheduler_type == 'apscheduler':
                trigger = IntervalTrigger(minutes=job_config['interval_minutes'])
                scheduler.add_job(
                    run_scheduled_pipeline,
                    trigger=trigger,
                    args=[job_config, script_dir],
                    id='watch_job',
                    name='监控目录任务',
                )
            else:
                class SimpleIntervalTrigger:
                    def __init__(self, minutes):
                        self._minutes = minutes
                    def get_interval(self):
                        from datetime import timedelta
                        return timedelta(minutes=self._minutes)
                scheduler.add_job(
                    run_scheduled_pipeline,
                    trigger=SimpleIntervalTrigger(job_config['interval_minutes']),
                    args=[job_config, script_dir],
                    id='watch_job',
                    name='监控目录任务',
                )

            try:
                run_scheduled_pipeline(job_config, script_dir)
                scheduler.start()
            except (KeyboardInterrupt, SystemExit):
                scheduler.shutdown()
            return True

    if args.scheduler:
        start_scheduler(script_dir)
        return True

    if args.scheduler_menu:
        run_scheduler_flow(script_dir)
        return True

    if args.preset_menu:
        run_preset_flow(script_dir)
        return True

    if args.export or args.export_template or args.export_total:
        if args.export_total and args.export_template:
            template_type = args.export_template
            if template_type not in FINANCIAL_EXPORT_TEMPLATES:
                logger.error('不支持的导出模板类型: %s', template_type)
                print(f'错误: 不支持的导出模板类型 "{template_type}"')
                print(f'支持的类型: {", ".join(FINANCIAL_EXPORT_TEMPLATES.keys())}')
                return True

            total_path = os.path.abspath(args.export_total)
            if not os.path.exists(total_path):
                logger.error('总表文件不存在: %s', total_path)
                print(f'错误: 总表文件不存在: {total_path}')
                return True

            result = export_financial_template(
                total_path=total_path,
                template_type=template_type,
                output_dir=args.export_output,
                operator=args.export_operator or '',
            )

            if result['success']:
                print(f'\n✅ 导出成功！')
                print(f'   模板: {result["template_name"]}')
                print(f'   凭证张数: {result["voucher_count"]}')
                print(f'   分录条数: {result["entry_count"]}')
                print(f'   输出文件: {result["output_path"]}\n')
            else:
                print(f'\n❌ 导出失败: {result.get("error", "未知错误")}\n')

            return True
        else:
            run_export_flow(script_dir)
            return True

    return None


def main():
    result = parse_args_and_run()
    if result is not None:
        return

    setup_logging()
    logger = get_logger()
    logger.info('========== 银行流水检验工具启动 ==========')

    script_dir = get_script_dir()

    init_audit_db(get_audit_db_path(script_dir))
    init_default_alert_rules(script_dir)
    run_alert_detection(script_dir)
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
    elif mode == 'monitor':
        run_monitor_flow(script_dir)
    elif mode == 'scheduler':
        run_scheduler_flow(script_dir)
    elif mode == 'export':
        run_export_flow(script_dir)
    elif mode == 'db_query':
        run_db_query_flow(script_dir)
    elif mode == 'db_stats':
        run_db_stats_flow(script_dir)
    elif mode == 'batch_history':
        run_batch_history_flow(script_dir)
    elif mode == 'preset':
        run_preset_flow(script_dir)

    logger.info('========== 银行流水检验工具运行结束 ==========')


# ──────────────────────────────────────────────
# 数据库查询与统计流程
# ──────────────────────────────────────────────

def run_db_query_flow(script_dir):
    """数据库查询流程：按主体/账号/时间范围查询流水记录"""
    logger = get_logger()

    if not HAS_DATABASE:
        show_warning('错误', '数据库模块不可用，请检查 database.py 文件是否存在。')
        logger.error('数据库模块不可用')
        return

    print('\n' + '=' * 60)
    print('数据库查询 - 支持多条件组合查询')
    print('=' * 60)
    print('提示：直接回车表示不使用该条件')
    print()

    subject = input('请输入主体名称（支持模糊匹配）: ').strip()
    account = input('请输入银行账号（支持模糊匹配）: ').strip()
    bank = input('请输入银行名称（如：北京银行、东亚银行）: ').strip()
    start_date = input('请输入开始日期（YYYY-MM-DD）: ').strip()
    end_date = input('请输入结束日期（YYYY-MM-DD）: ').strip()

    min_amount_str = input('请输入最小金额（单位：元）: ').strip()
    max_amount_str = input('请输入最大金额（单位：元）: ').strip()
    counterpart = input('请输入对方户名（支持模糊匹配）: ').strip()
    summary_keyword = input('请输入摘要关键词（支持模糊匹配）: ').strip()

    limit_str = input('请输入返回记录数上限（默认 1000）: ').strip()
    limit = int(limit_str) if limit_str.isdigit() else 1000

    min_amount = float(min_amount_str) if min_amount_str else None
    max_amount = float(max_amount_str) if max_amount_str else None

    print(f'\n正在查询数据库...')

    try:
        result = db_module.query_transactions(
            subject=subject if subject else None,
            account=account if account else None,
            bank=bank if bank else None,
            start_date=start_date if start_date else None,
            end_date=end_date if end_date else None,
            min_amount=min_amount,
            max_amount=max_amount,
            counterpart=counterpart if counterpart else None,
            summary_keyword=summary_keyword if summary_keyword else None,
            limit=limit,
            script_dir=script_dir,
        )

        print(f'\n查询完成！共找到 {result.total_count} 条记录')

        if result.records:
            print(f'\n返回前 {len(result.records)} 条记录：')
            print('-' * 100)
            print(f'{"序号":<6}{"日期":<20}{"主体":<20}{"银行":<10}{"金额(元)":<15}{"对方户名":<20}')
            print('-' * 100)

            for i, record in enumerate(result.records[:50], 1):
                amount = record.收款 if (record.收款 and record.收款 > 0) else (record.付款 or 0)
                counterpart = record.对方户名 or ''
                if len(counterpart) > 18:
                    counterpart = counterpart[:16] + '...'

                print(f'{i:<6}{str(record.交易日期 or "")[:19]:<20}'
                      f'{str(record.主体 or "")[:18]:<20}'
                      f'{str(record.银行 or "")[:8]:<10}'
                      f'{amount:>12,.2f}  '
                      f'{counterpart:<20}')

            if len(result.records) > 50:
                print(f'\n... 还有 {len(result.records) - 50} 条记录未显示')

        if result.summary:
            print(f'\n{"=" * 60}')
            print('查询结果汇总：')
            print('-' * 60)
            print(f'记录总数：{result.summary.get("记录总数", 0)}')
            print(f'付款笔数：{result.summary.get("付款笔数", 0)}')
            print(f'收款笔数：{result.summary.get("收款笔数", 0)}')
            print(f'付款总额：{result.summary.get("付款总额", 0):,.2f} 元')
            print(f'收款总额：{result.summary.get("收款总额", 0):,.2f} 元')
            print(f'净　　额：{result.summary.get("净额", 0):,.2f} 元')

            bank_dist = result.summary.get('银行分布', {})
            if bank_dist:
                print(f'\n银行分布：')
                for bank_name, cnt in bank_dist.items():
                    print(f'  {bank_name}: {cnt} 条')

            subject_dist = result.summary.get('主体分布', {})
            if subject_dist:
                print(f'\n主体分布（前 10）：')
                for i, (subj_name, cnt) in enumerate(list(subject_dist.items())[:10], 1):
                    print(f'  {i}. {subj_name}: {cnt} 条')

        export = input(f'\n是否导出查询结果到 Excel？(y/N): ').strip().lower()
        if export == 'y':
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            output_path = os.path.join(script_dir, f'查询结果_{timestamp}.xlsx')
            try:
                result.to_excel(output_path)
                show_info('导出成功', f'查询结果已导出到：\n{output_path}')
                logger.info('查询结果已导出: %s', output_path)
            except Exception as e:
                show_warning('导出失败', f'导出 Excel 失败：{e}')
                logger.error('导出查询结果失败: %s', e)

    except Exception as e:
        show_warning('查询失败', f'数据库查询失败：{e}')
        logger.error('数据库查询失败: %s', e, exc_info=True)


def run_db_stats_flow(script_dir):
    """数据库统计流程：查看数据汇总统计信息"""
    logger = get_logger()

    if not HAS_DATABASE:
        show_warning('错误', '数据库模块不可用，请检查 database.py 文件是否存在。')
        logger.error('数据库模块不可用')
        return

    print('\n' + '=' * 60)
    print('数据库统计信息')
    print('=' * 60)

    try:
        stats = db_module.get_db_statistics(script_dir=script_dir)

        print(f'\n【总体概览】')
        print(f'  总记录数：{stats.get("总记录数", 0):,} 条')
        print(f'  导入批次：{stats.get("导入批次数量", 0)} 次')

        date_range = stats.get('日期范围', {})
        if date_range:
            print(f'  最早交易：{date_range.get("最早交易日期", "无")}')
            print(f'  最晚交易：{date_range.get("最晚交易日期", "无")}')

        by_bank = stats.get('按银行统计', [])
        if by_bank:
            print(f'\n【按银行统计】')
            print('-' * 80)
            print(f'{"银行":<15}{"记录数":<12}{"付款总额(元)":<20}{"收款总额(元)":<20}')
            print('-' * 80)
            for row in by_bank:
                bank_name = row.get('银行', '未知')
                cnt = row.get('cnt', 0)
                payment = row.get('total_payment', 0) or 0
                receipt = row.get('total_receipt', 0) or 0
                print(f'{bank_name:<15}{cnt:<12,}{payment:>18,.2f}  {receipt:>18,.2f}')

        by_subject = stats.get('按主体统计', [])
        if by_subject:
            print(f'\n【按主体统计（前 15）】')
            print('-' * 90)
            print(f'{"序号":<6}{"主体":<30}{"记录数":<12}{"付款总额(元)":<20}{"收款总额(元)":<20}')
            print('-' * 90)
            for i, row in enumerate(by_subject[:15], 1):
                subject_name = str(row.get('主体', '未知'))[:28]
                cnt = row.get('cnt', 0)
                payment = row.get('total_payment', 0) or 0
                receipt = row.get('total_receipt', 0) or 0
                print(f'{i:<6}{subject_name:<30}{cnt:<12,}'
                      f'{payment:>18,.2f}  {receipt:>18,.2f}')

        by_account = stats.get('按账号统计', [])
        if by_account:
            print(f'\n【按账号统计（前 15）】')
            print('-' * 80)
            print(f'{"序号":<6}{"银行账号":<25}{"主体":<20}{"银行":<12}{"记录数":<10}')
            print('-' * 80)
            for i, row in enumerate(by_account[:15], 1):
                account = str(row.get('银行账号', '未知'))[:22]
                subject = str(row.get('主体', '未知'))[:18]
                bank = str(row.get('银行', ''))[:10]
                cnt = row.get('cnt', 0)
                print(f'{i:<6}{account:<25}{subject:<20}{bank:<12}{cnt:<10,}')

        by_date = stats.get('近30天交易趋势', [])
        if by_date:
            print(f'\n【近 30 天交易趋势】')
            print('-' * 60)
            print(f'{"日期":<15}{"记录数":<10}{"付款(元)":<18}{"收款(元)":<18}')
            print('-' * 60)
            for row in reversed(by_date):
                dt = str(row.get('dt', ''))[:10]
                cnt = row.get('cnt', 0)
                payment = row.get('total_payment', 0) or 0
                receipt = row.get('total_receipt', 0) or 0
                print(f'{dt:<15}{cnt:<10}{payment:>16,.2f}  {receipt:>16,.2f}')

        export = input(f'\n是否导出统计信息到 Excel？(y/N): ').strip().lower()
        if export == 'y':
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            output_path = os.path.join(script_dir, f'数据库统计_{timestamp}.xlsx')
            try:
                with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                    if by_bank:
                        pd.DataFrame(by_bank).to_excel(writer, sheet_name='按银行统计', index=False)
                    if by_subject:
                        pd.DataFrame(by_subject).to_excel(writer, sheet_name='按主体统计', index=False)
                    if by_account:
                        pd.DataFrame(by_account).to_excel(writer, sheet_name='按账号统计', index=False)
                    if by_date:
                        pd.DataFrame(by_date).to_excel(writer, sheet_name='近30天趋势', index=False)

                    overview_df = pd.DataFrame([{
                        '总记录数': stats.get('总记录数', 0),
                        '导入批次数量': stats.get('导入批次数量', 0),
                        '最早交易日期': date_range.get('最早交易日期', ''),
                        '最晚交易日期': date_range.get('最晚交易日期', ''),
                    }])
                    overview_df.to_excel(writer, sheet_name='总体概览', index=False)

                show_info('导出成功', f'统计信息已导出到：\n{output_path}')
                logger.info('数据库统计已导出: %s', output_path)
            except Exception as e:
                show_warning('导出失败', f'导出 Excel 失败：{e}')
                logger.error('导出统计信息失败: %s', e)

    except Exception as e:
        show_warning('统计失败', f'获取统计信息失败：{e}')
        logger.error('获取数据库统计失败: %s', e, exc_info=True)


# ──────────────────────────────────────────────
# 历史批次与版本管理流程
# ──────────────────────────────────────────────

def run_batch_history_flow(script_dir):
    """批次管理流程：查看历史批次与版本回溯"""
    logger = get_logger()

    if not HAS_BATCH_MANAGER:
        show_warning('错误', '批次管理模块不可用，请检查 batch_manager.py 文件是否存在。')
        logger.error('批次管理模块不可用')
        return

    try:
        batch_manager = batch_module.get_batch_manager(script_dir)
    except Exception as e:
        show_warning('错误', f'批次管理器初始化失败：{e}')
        logger.error('批次管理器初始化失败: %s', e, exc_info=True)
        return

    logger.info('========== 批次管理面板启动 ==========')

    while True:
        stats = batch_manager.get_statistics()

        print('\n' + '=' * 60)
        print('批次管理面板')
        print('=' * 60)
        print(f'总批次数: {stats["total_batches"]}')
        print(f'  ├─ 成功: {stats["success_batches"]}')
        print(f'  ├─ 失败: {stats["failed_batches"]}')
        print(f'  └─ 运行中: {stats["running_batches"]}')
        print(f'累计处理记录: {stats["total_records"]:,} 条')
        print(f'累计新增记录: {stats["total_new_records"]:,} 条')
        print('-' * 60)
        print('  1) 查看最近批次列表')
        print('  2) 按条件查询批次')
        print('  3) 查看批次详情')
        print('  4) 恢复批次文件')
        print('  5) 删除批次')
        print('  0) 返回主菜单')
        print('-' * 60)

        choice = input('请选择操作（0-5）: ').strip()

        if choice == '0':
            break
        elif choice == '1':
            _show_recent_batches(batch_manager)
        elif choice == '2':
            _query_batches(batch_manager)
        elif choice == '3':
            _show_batch_detail(batch_manager)
        elif choice == '4':
            _restore_batch(batch_manager)
        elif choice == '5':
            _delete_batch(batch_manager)
        else:
            print('无效选项，请重新选择。')

    logger.info('========== 批次管理面板关闭 ==========')


def _display_batch_list(batches):
    if not batches:
        print('\n未找到符合条件的批次。')
        return

    print('\n' + '-' * 100)
    print(f'{"序号":<5}{"批次号":<24}{"状态":<10}{"开始时间":<20}'
          f'{"记录数":<10}{"新增":<10}{"模式":<8}')
    print('-' * 100)

    for i, b in enumerate(batches, 1):
        mode = '增量' if b.incremental_mode else '全量'
        status = f'{b.status}'
        if status == 'success':
            status = '✅ 成功'
        elif status == 'warning':
            status = '⚠️ 警告'
        elif status == 'failed':
            status = '❌ 失败'
        elif status == 'running':
            status = '⏳ 运行中'
        print(f'{i:<5}{b.batch_id:<24}{status:<10}{b.start_time:<20}'
              f'{b.total_records:<10,}{b.new_records:<10,}{mode:<8}')
    print('-' * 100)


def _show_recent_batches(batch_manager):
    limit_str = input('请输入显示数量（默认 20）: ').strip()
    limit = int(limit_str) if limit_str.isdigit() else 20

    batches = batch_manager.query_batches(limit=limit)
    _display_batch_list(batches)


def _query_batches(batch_manager):
    print('\n查询条件（直接回车表示不使用该条件）:')
    start_date = input('开始日期（YYYY-MM-DD）: ').strip() or None
    end_date = input('结束日期（YYYY-MM-DD）: ').strip() or None
    status = input('状态（success/warning/failed/running）: ').strip() or None
    operator = input('操作员: ').strip() or None
    min_records_str = input('最小记录数: ').strip()
    min_records = int(min_records_str) if min_records_str.isdigit() else None

    batches = batch_manager.query_batches(
        start_date=start_date,
        end_date=end_date,
        status=status,
        operator=operator,
        min_records=min_records,
        limit=100,
    )
    _display_batch_list(batches)


def _show_batch_detail(batch_manager):
    batch_id = input('请输入批次号: ').strip()
    if not batch_id:
        print('未输入批次号。')
        return

    detail = batch_manager.get_batch_detail(batch_id)
    if not detail:
        print(f'\n批次不存在: {batch_id}')
        return

    info = detail['batch_info']
    metadata = detail['metadata']
    files = detail['files']

    status = info.get('status', '')
    if status == 'success':
        status_icon = '✅'
    elif status == 'warning':
        status_icon = '⚠️'
    elif status == 'failed':
        status_icon = '❌'
    else:
        status_icon = '⏳'

    mode = '增量合并' if info.get('incremental_mode') else '全量覆盖'

    print('\n' + '=' * 60)
    print(f'批次详情 - {batch_id}')
    print('=' * 60)
    print(f'状态: {status_icon} {status}')
    print(f'开始时间: {info.get("start_time", "")}')
    print(f'结束时间: {info.get("end_time", "")}')
    print(f'操作员: {info.get("operator", "未指定")}')
    print(f'输入文件夹: {info.get("input_folder", "")}')
    print(f'运行模式: {mode}')
    print('-' * 40)
    print('📊 处理统计:')
    print(f'  总记录数: {info.get("total_records", 0):,}')
    print(f'  新增记录: {info.get("new_records", 0):,}')
    print(f'  重复记录: {info.get("duplicate_records", 0):,}')
    print(f'  处理文件: {info.get("processed_files", 0)} 个')
    print(f'  未识别文件: {info.get("unprocessed_files", 0)} 个')
    print(f'  出错文件: {info.get("error_files", 0)} 个')
    print('-' * 40)

    processed_files = metadata.get('processed_files', [])
    if processed_files:
        print('📄 已处理文件:')
        for f in processed_files:
            print(f'  - {os.path.basename(f)}')

    unprocessed_files = metadata.get('unprocessed_files', [])
    if unprocessed_files:
        print('\n⚠️  未识别文件:')
        for f in unprocessed_files:
            print(f'  - {os.path.basename(f)}')

    error_files = metadata.get('error_files', [])
    if error_files:
        print('\n❌ 出错文件:')
        for f, err in error_files:
            print(f'  - {os.path.basename(f)}: {err}')

    if info.get('error_message'):
        print(f'\n❌ 错误信息:\n{info["error_message"]}')

    print('-' * 40)
    print('📁 归档文件:')
    for name, path in files.items():
        print(f'  - {name}')
        print(f'    路径: {path}')

    print(f'\n📂 归档目录: {info.get("batch_dir", "")}')
    print('=' * 60)


def _restore_batch(batch_manager):
    batch_id = input('请输入要恢复的批次号: ').strip()
    if not batch_id:
        print('未输入批次号。')
        return

    target_dir = input('请输入恢复目录（直接回车使用默认目录）: ').strip() or None

    try:
        restored = batch_manager.restore_batch(batch_id, target_dir)
        print(f'\n✅ 批次 {batch_id} 恢复成功！')
        print(f'恢复目录: {os.path.dirname(next(iter(restored.values()))) if restored else target_dir}')
        print('\n恢复的文件:')
        for name in restored.keys():
            print(f'  - {name}')
    except ValueError as e:
        print(f'\n❌ {e}')
    except Exception as e:
        print(f'\n❌ 恢复失败: {e}')
        get_logger().error('批次恢复失败: %s', e, exc_info=True)


def _delete_batch(batch_manager):
    batch_id = input('请输入要删除的批次号: ').strip()
    if not batch_id:
        print('未输入批次号。')
        return

    confirm = input(f'确认要删除批次 {batch_id} 吗？此操作不可恢复！(yes/N): ').strip().lower()
    if confirm != 'yes':
        print('已取消删除。')
        return

    try:
        success = batch_manager.delete_batch(batch_id)
        if success:
            print(f'\n✅ 批次 {batch_id} 已删除。')
        else:
            print(f'\n❌ 批次不存在: {batch_id}')
    except Exception as e:
        print(f'\n❌ 删除失败: {e}')
        get_logger().error('批次删除失败: %s', e, exc_info=True)


# ──────────────────────────────────────────────
# 任务配置预设管理模块
# ──────────────────────────────────────────────

PRESET_CONFIG_FILENAME = 'task_presets.json'


@dataclass
class TaskPreset:
    preset_id: str
    name: str
    description: str = ''
    output_dir: str = ''
    start_date: str = ''
    end_date: str = ''
    keep_strategy: str = 'keep_unprocessed'
    enabled_banks: List[str] = field(default_factory=list)
    incremental: bool = True
    created_at: str = ''
    updated_at: str = ''


KEEP_STRATEGIES = {
    'keep_unprocessed': '仅保留未处理文件',
    'keep_all': '保留所有文件',
    'delete_all': '删除所有已处理文件',
}


def get_preset_config_path(script_dir=None):
    if script_dir is None:
        script_dir = get_script_dir()
    return os.path.join(script_dir, PRESET_CONFIG_FILENAME)


def load_preset_config(script_dir=None):
    config_path = get_preset_config_path(script_dir)
    logger = get_logger()

    if not os.path.exists(config_path):
        logger.info('预设配置文件不存在，将创建默认配置: %s', config_path)
        default_config = {
            'presets': [],
            'settings': {
                'default_preset': '',
            }
        }
        save_preset_config(default_config, script_dir)
        return default_config

    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)
        logger.info('已加载预设配置，共 %d 个预设', len(config.get('presets', [])))
        return config
    except Exception as e:
        logger.error('加载预设配置失败: %s', e)
        return {'presets': [], 'settings': {}}


def save_preset_config(config, script_dir=None):
    config_path = get_preset_config_path(script_dir)
    try:
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
        logger = get_logger()
        logger.info('预设配置已保存: %s', config_path)
        return True
    except Exception as e:
        logger = get_logger()
        logger.error('保存预设配置失败: %s', e)
        return False


def _generate_preset_id():
    return f"PRESET{datetime.now().strftime('%Y%m%d%H%M%S')}{uuid.uuid4().hex[:6].upper()}"


def save_preset(preset_data, script_dir=None):
    config = load_preset_config(script_dir)
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    preset_id = preset_data.get('preset_id') or _generate_preset_id()
    preset_data['preset_id'] = preset_id

    if not preset_data.get('name'):
        preset_data['name'] = f'预设_{datetime.now().strftime("%Y%m%d_%H%M%S")}'

    preset_data['enabled_banks'] = preset_data.get('enabled_banks') or list(BANK_PREFIXES)
    preset_data['keep_strategy'] = preset_data.get('keep_strategy') or 'keep_unprocessed'
    preset_data['incremental'] = preset_data.get('incremental', True)

    existing = None
    for i, p in enumerate(config['presets']):
        if p['preset_id'] == preset_id:
            existing = i
            break

    if existing is not None:
        preset_data['created_at'] = config['presets'][existing].get('created_at', now)
        preset_data['updated_at'] = now
        config['presets'][existing] = preset_data
        action = '更新'
    else:
        preset_data['created_at'] = now
        preset_data['updated_at'] = now
        config['presets'].append(preset_data)
        action = '添加'

    save_preset_config(config, script_dir)

    logger = get_logger()
    logger.info('已%s预设 [%s] %s', action, preset_id, preset_data.get('name', ''))
    return preset_id


def load_preset(preset_id, script_dir=None):
    config = load_preset_config(script_dir)
    for preset in config.get('presets', []):
        if preset['preset_id'] == preset_id:
            return preset
    return None


def delete_preset(preset_id, script_dir=None):
    config = load_preset_config(script_dir)
    logger = get_logger()

    original_count = len(config['presets'])
    config['presets'] = [p for p in config['presets'] if p['preset_id'] != preset_id]

    if len(config['presets']) == original_count:
        logger.warning('未找到预设 [%s]', preset_id)
        return False

    if config.get('settings', {}).get('default_preset') == preset_id:
        config['settings']['default_preset'] = ''

    save_preset_config(config, script_dir)
    logger.info('已删除预设 [%s]', preset_id)
    return True


def list_presets(script_dir=None):
    config = load_preset_config(script_dir)
    return config.get('presets', [])


def set_default_preset(preset_id, script_dir=None):
    config = load_preset_config(script_dir)
    if 'settings' not in config:
        config['settings'] = {}
    config['settings']['default_preset'] = preset_id
    save_preset_config(config, script_dir)


def get_default_preset(script_dir=None):
    config = load_preset_config(script_dir)
    default_id = config.get('settings', {}).get('default_preset', '')
    if default_id:
        return load_preset(default_id, script_dir)
    return None


def apply_preset_to_pipeline(preset, folder, script_dir):
    logger = get_logger()

    if not preset:
        logger.warning('预设为空，使用默认配置运行')
        return run_pipeline(folder, script_dir)

    logger.info('应用预设 [%s] %s', preset.get('preset_id'), preset.get('name', ''))

    enabled_banks = preset.get('enabled_banks', BANK_PREFIXES)
    keep_strategy = preset.get('keep_strategy', 'keep_unprocessed')
    incremental = preset.get('incremental', True)
    start_date = preset.get('start_date', '')
    end_date = preset.get('end_date', '')

    result = run_pipeline_with_options(
        folder=folder,
        script_dir=script_dir,
        incremental=incremental,
        enabled_banks=enabled_banks,
        keep_strategy=keep_strategy,
        start_date=start_date,
        end_date=end_date,
    )

    return result


def run_pipeline_with_options(folder, script_dir, incremental=True,
                              enabled_banks=None, keep_strategy='keep_unprocessed',
                              start_date='', end_date='', batch_id=None):
    logger = get_logger()

    if enabled_banks is None:
        enabled_banks = BANK_PREFIXES

    lookup_file = find_lookup_file(script_dir)
    lookup_missing = lookup_file is None
    if lookup_missing:
        logger.warning('未找到主体查找表，"主体"列将为空')

    existing_keys = set()
    existing_records = []
    actual_incremental = False
    duplicate_count = 0
    new_record_count = 0

    if incremental:
        summary_path = get_summary_table_path(script_dir)
        existing_keys, existing_records = load_existing_keys(summary_path)
        actual_incremental = len(existing_records) > 0
        if actual_incremental:
            logger.info('===== 增量合并模式已启用 =====')
        else:
            logger.info('无历史数据，将以全量模式运行')

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
        return ProcessingResult(
            lookup_missing=lookup_missing,
            folder_empty=True,
            incremental_mode=actual_incremental,
            existing_record_count=len(existing_records),
        )

    def _parse_date(value):
        if not value:
            return None
        if isinstance(value, datetime):
            return value
        try:
            if isinstance(value, str):
                return datetime.strptime(value[:10], '%Y-%m-%d')
        except (ValueError, TypeError):
            pass
        return None

    start_dt = _parse_date(start_date)
    end_dt = _parse_date(end_date)

    if start_dt:
        logger.info('日期过滤 - 开始日期: %s', start_dt.strftime('%Y-%m-%d'))
    if end_dt:
        logger.info('日期过滤 - 结束日期: %s', end_dt.strftime('%Y-%m-%d'))

    def _is_date_in_range(trade_date):
        if trade_date is None:
            return True
        dt = _parse_date(trade_date)
        if dt is None:
            return True
        if start_dt and dt < start_dt:
            return False
        if end_dt and dt > end_dt:
            return False
        return True

    all_rows = []
    processed_files = []
    unprocessed_files = []
    error_files = []
    filtered_out_count = 0

    for filepath in excel_files:
        bank = identify_bank(filepath)
        if bank and bank in BANK_PROCESSORS and bank in enabled_banks:
            try:
                processor = BANK_PROCESSORS[bank]
                rows = processor(filepath, lookup_file)

                if start_dt or end_dt:
                    original_len = len(rows)
                    rows = [r for r in rows if _is_date_in_range(r.get('交易日期'))]
                    filtered_out_count += (original_len - len(rows))

                all_rows.extend(rows)
                processed_files.append(filepath)
                logger.info('成功处理文件: %s（%d 条记录）', filepath, len(rows))
            except Exception as e:
                error_files.append((filepath, str(e)))
                logger.error('处理文件「%s」时发生错误: %s', filepath, e, exc_info=True)
        else:
            unprocessed_files.append(filepath)
            if bank and bank not in enabled_banks:
                logger.info('文件「%s」所属银行「%s」不在启用列表中，跳过', filepath, bank)

    if filtered_out_count > 0:
        logger.info('日期过滤共排除 %d 条记录', filtered_out_count)

    error_file_paths = {f for f, _ in error_files}
    keep_set = set(unprocessed_files) | error_file_paths

    if keep_strategy == 'keep_all':
        keep_set = set(excel_files)
        logger.info('保留策略：保留所有文件')
    elif keep_strategy == 'delete_all':
        keep_set = set()
        logger.info('保留策略：删除所有已处理文件')
    else:
        logger.info('保留策略：仅保留未处理文件')

    delete_processed_files(excel_files, keep_set)

    output_path = None
    final_rows = []

    if all_rows:
        if actual_incremental:
            incremental_rows, duplicate_count = filter_incremental_records(all_rows, existing_keys)
            new_record_count = len(incremental_rows)
            output_path = merge_and_export_summary(existing_records, incremental_rows, script_dir)
            final_rows = existing_records + incremental_rows
        else:
            columns = [
                '唯一id', '银行', '银行账号', '主体', '交易日期',
                '付款', '收款', '摘要', '对方户名', '余额', '交易流水号',
            ]
            df = pd.DataFrame(all_rows, columns=columns)
            output_path = get_summary_table_path(script_dir)
            df.to_excel(output_path, index=False, engine='openpyxl')
            logger.info('总表输出完成: %s（共 %d 条记录）', output_path, len(all_rows))
            final_rows = all_rows
            new_record_count = len(all_rows)
    else:
        logger.warning('未提取到任何银行流水记录')
        if existing_records:
            output_path = merge_and_export_summary(existing_records, [], script_dir)
            final_rows = existing_records

    db_inserted = 0
    db_duplicates = 0
    if HAS_DATABASE and final_rows:
        try:
            if batch_id is None:
                batch_id = f"BATCH{datetime.now().strftime('%Y%m%d%H%M%S')}"
            db_inserted, db_duplicates = db_module.persist_transactions(
                final_rows,
                batch_id=batch_id,
                deduplicate=True,
                script_dir=script_dir,
            )
            logger.info(
                '数据库持久化完成: 批次 %s, 插入 %d 条, 去重跳过 %d 条',
                batch_id, db_inserted, db_duplicates,
            )
        except Exception as e:
            logger.error('数据库持久化失败: %s', e, exc_info=True)

    return ProcessingResult(
        all_rows=final_rows,
        processed_files=processed_files,
        unprocessed_files=unprocessed_files,
        error_files=error_files,
        output_path=output_path,
        lookup_missing=lookup_missing,
        incremental_mode=actual_incremental,
        existing_record_count=len(existing_records),
        new_record_count=new_record_count,
        duplicate_record_count=duplicate_count,
        db_inserted_count=db_inserted,
        db_duplicate_count=db_duplicates,
    )


def run_preset_flow(script_dir):
    """预设管理流程：管理和应用任务配置预设"""
    logger = get_logger()

    if HAS_TKINTER and tk is not None:
        try:
            gui_preset_manager(script_dir)
            return
        except Exception as e:
            logger.warning('GUI预设管理启动失败，将使用命令行模式: %s', e)

    print('\n' + '=' * 60)
    print('任务配置预设管理')
    print('=' * 60)

    while True:
        presets = list_presets(script_dir)
        default_preset = get_default_preset(script_dir)

        print(f'\n当前预设数量: {len(presets)}')
        if default_preset:
            print(f'默认预设: [{default_preset["preset_id"]}] {default_preset.get("name", "")}')

        print('\n请选择操作：')
        print('  1) 列出所有预设')
        print('  2) 保存当前配置为新预设')
        print('  3) 加载并应用预设')
        print('  4) 删除预设')
        print('  5) 设为默认预设')
        print('  6) 查看预设详情')
        print('  0) 返回主菜单')

        choice = input('\n请输入选项: ').strip()

        if choice == '1':
            _list_presets_cli(presets)
        elif choice == '2':
            _save_preset_cli(script_dir)
        elif choice == '3':
            _apply_preset_cli(script_dir)
        elif choice == '4':
            _delete_preset_cli(script_dir)
        elif choice == '5':
            _set_default_preset_cli(script_dir)
        elif choice == '6':
            _show_preset_detail_cli(script_dir)
        elif choice == '0':
            break
        else:
            print('无效选项，请重新输入')


def gui_preset_manager(script_dir):
    """GUI预设管理窗口"""
    if tk is None:
        raise RuntimeError('Tkinter not available')

    logger = get_logger()
    logger.info('启动GUI预设管理')

    root = tk.Tk()
    root.title('任务配置预设管理')
    root.geometry('750x600')
    root.minsize(700, 550)

    presets_list = []
    selected_preset_id = tk.StringVar()
    detail_text = None

    def refresh_presets():
        nonlocal presets_list
        presets_list = list_presets(script_dir)
        listbox.delete(0, tk.END)
        default_id = get_default_preset(script_dir)
        default_id = default_id['preset_id'] if default_id else ''
        for p in presets_list:
            marker = '★ ' if p['preset_id'] == default_id else '  '
            listbox.insert(tk.END, f"{marker}{p['preset_id']} - {p.get('name', '未命名')}")

    def on_select(event):
        selection = listbox.curselection()
        if not selection:
            return
        idx = selection[0]
        preset = presets_list[idx]
        selected_preset_id.set(preset['preset_id'])
        show_preset_detail(preset)

    def show_preset_detail(preset):
        if detail_text is None:
            return
        detail_text.config(state=tk.NORMAL)
        detail_text.delete(1.0, tk.END)

        banks = ', '.join(preset.get('enabled_banks', []))
        keep = KEEP_STRATEGIES.get(preset.get('keep_strategy'), '未知')
        incremental = '是' if preset.get('incremental', True) else '否'

        detail = f"""预设名称: {preset.get('name', '')}
预设ID: {preset.get('preset_id', '')}
描述: {preset.get('description', '无')}
{'-' * 50}
输出目录: {preset.get('output_dir', '默认')}
开始日期: {preset.get('start_date', '不限制')}
结束日期: {preset.get('end_date', '不限制')}
保留策略: {keep}
启用银行: {banks}
增量合并: {incremental}
{'-' * 50}
创建时间: {preset.get('created_at', '')}
更新时间: {preset.get('updated_at', '')}
"""
        detail_text.insert(tk.END, detail)
        detail_text.config(state=tk.DISABLED)

    def add_preset():
        PresetEditorDialog(root, script_dir, on_saved=refresh_presets)

    def edit_preset():
        pid = selected_preset_id.get()
        if not pid:
            messagebox.showwarning('提示', '请先选择一个预设')
            return
        preset = load_preset(pid, script_dir)
        if preset:
            PresetEditorDialog(root, script_dir, preset=preset, on_saved=refresh_presets)

    def delete_selected():
        pid = selected_preset_id.get()
        if not pid:
            messagebox.showwarning('提示', '请先选择一个预设')
            return
        preset = load_preset(pid, script_dir)
        if not preset:
            return
        if messagebox.askyesno('确认删除', f'确定要删除预设「{preset.get("name", "")}」吗？'):
            if delete_preset(pid, script_dir):
                messagebox.showinfo('成功', '预设已删除')
                refresh_presets()
                selected_preset_id.set('')
                if detail_text:
                    detail_text.config(state=tk.NORMAL)
                    detail_text.delete(1.0, tk.END)
                    detail_text.config(state=tk.DISABLED)
            else:
                messagebox.showerror('错误', '删除失败')

    def set_default():
        pid = selected_preset_id.get()
        if not pid:
            messagebox.showwarning('提示', '请先选择一个预设')
            return
        set_default_preset(pid, script_dir)
        messagebox.showinfo('成功', '已设为默认预设')
        refresh_presets()

    def apply_preset():
        pid = selected_preset_id.get()
        if not pid:
            messagebox.showwarning('提示', '请先选择一个预设')
            return
        preset = load_preset(pid, script_dir)
        if not preset:
            return

        folder = filedialog.askdirectory(title='请选择银行流水文件夹')
        if not folder:
            return

        if not messagebox.askyesno('确认应用',
                                   f'即将应用预设「{preset.get("name", "")}」\n处理文件夹: {folder}\n\n是否继续？'):
            return

        root.destroy()

        with AuditLogger('preset_pipeline', script_dir) as audit:
            audit.record_input(folder)
            audit.set_extra_info({'preset_id': pid, 'preset_name': preset.get('name', '')})
            result = apply_preset_to_pipeline(preset, folder, script_dir)
            audit.record_result(result)

            msg = format_result_message(result)
            msg += f'\n\n审计编号: {audit.audit_id}'
            msg += f'\n预设: {preset.get("name", "")} ({pid})'
            show_info('完成' if result.all_rows else '提示', msg)

    main_frame = tk.Frame(root)
    main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    left_frame = tk.Frame(main_frame)
    left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))

    tk.Label(left_frame, text='预设列表（★ 表示默认）', font=('Arial', 12, 'bold')).pack(anchor=tk.W)

    listbox_frame = tk.Frame(left_frame)
    listbox_frame.pack(fill=tk.BOTH, expand=True, pady=5)

    scrollbar = tk.Scrollbar(listbox_frame)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    listbox = tk.Listbox(listbox_frame, font=('Arial', 11), yscrollcommand=scrollbar.set)
    listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar.config(command=listbox.yview)
    listbox.bind('<<ListboxSelect>>', on_select)

    btn_frame = tk.Frame(left_frame)
    btn_frame.pack(fill=tk.X, pady=5)

    tk.Button(btn_frame, text='新增', width=8, command=add_preset, bg='#4CAF50', fg='white').pack(side=tk.LEFT, padx=2)
    tk.Button(btn_frame, text='编辑', width=8, command=edit_preset, bg='#2196F3', fg='white').pack(side=tk.LEFT, padx=2)
    tk.Button(btn_frame, text='删除', width=8, command=delete_selected, bg='#f44336', fg='white').pack(side=tk.LEFT, padx=2)
    tk.Button(btn_frame, text='设为默认', width=10, command=set_default, bg='#FF9800', fg='white').pack(side=tk.LEFT, padx=2)

    right_frame = tk.Frame(main_frame)
    right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))

    tk.Label(right_frame, text='预设详情', font=('Arial', 12, 'bold')).pack(anchor=tk.W)

    detail_text = tk.Text(right_frame, font=('Arial', 11), wrap=tk.WORD, height=15)
    detail_text.pack(fill=tk.BOTH, expand=True, pady=5)
    detail_text.config(state=tk.DISABLED)

    bottom_frame = tk.Frame(root)
    bottom_frame.pack(fill=tk.X, padx=10, pady=(0, 10))

    tk.Button(bottom_frame, text='应用预设处理文件', height=2, command=apply_preset,
              bg='#FF5722', fg='white', font=('Arial', 12, 'bold')).pack(fill=tk.X)
    tk.Button(bottom_frame, text='关闭', height=2, command=root.destroy,
              bg='#9E9E9E', fg='white', font=('Arial', 11)).pack(fill=tk.X, pady=(5, 0))

    refresh_presets()
    root.mainloop()


class PresetEditorDialog:
    """预设编辑器对话框"""

    def __init__(self, parent, script_dir, preset=None, on_saved=None):
        self.script_dir = script_dir
        self.preset = preset
        self.on_saved = on_saved

        self.dialog = tk.Toplevel(parent)
        self.dialog.title('编辑预设' if preset else '新增预设')
        self.dialog.geometry('500x560')
        self.dialog.transient(parent)
        self.dialog.grab_set()

        self._build_ui()

        if preset:
            self._load_preset(preset)

    def _build_ui(self):
        pad = {'padx': 10, 'pady': 5}

        tk.Label(self.dialog, text='预设名称:', font=('Arial', 11)).pack(anchor=tk.W, **pad)
        self.name_var = tk.StringVar()
        tk.Entry(self.dialog, textvariable=self.name_var, font=('Arial', 11)).pack(fill=tk.X, **pad)

        tk.Label(self.dialog, text='描述:', font=('Arial', 11)).pack(anchor=tk.W, **pad)
        self.desc_text = tk.Text(self.dialog, height=3, font=('Arial', 11))
        self.desc_text.pack(fill=tk.X, **pad)

        tk.Label(self.dialog, text='输出目录（可选）:', font=('Arial', 11)).pack(anchor=tk.W, **pad)
        output_frame = tk.Frame(self.dialog)
        output_frame.pack(fill=tk.X, **pad)
        self.output_var = tk.StringVar()
        tk.Entry(output_frame, textvariable=self.output_var, font=('Arial', 11)).pack(side=tk.LEFT, fill=tk.X, expand=True)

        def browse_output():
            d = filedialog.askdirectory(title='选择输出目录')
            if d:
                self.output_var.set(d)

        tk.Button(output_frame, text='浏览...', command=browse_output).pack(side=tk.LEFT, padx=5)

        date_frame = tk.Frame(self.dialog)
        date_frame.pack(fill=tk.X, **pad)

        tk.Label(date_frame, text='开始日期:', font=('Arial', 11)).pack(side=tk.LEFT)
        self.start_date_var = tk.StringVar()
        tk.Entry(date_frame, textvariable=self.start_date_var, width=15, font=('Arial', 11)).pack(side=tk.LEFT, padx=5)
        tk.Label(date_frame, text='(YYYY-MM-DD)', font=('Arial', 9), fg='#666').pack(side=tk.LEFT)

        date_frame2 = tk.Frame(self.dialog)
        date_frame2.pack(fill=tk.X, **pad)
        tk.Label(date_frame2, text='结束日期:', font=('Arial', 11)).pack(side=tk.LEFT)
        self.end_date_var = tk.StringVar()
        tk.Entry(date_frame2, textvariable=self.end_date_var, width=15, font=('Arial', 11)).pack(side=tk.LEFT, padx=5)
        tk.Label(date_frame2, text='(YYYY-MM-DD)', font=('Arial', 9), fg='#666').pack(side=tk.LEFT)

        tk.Label(self.dialog, text='保留策略:', font=('Arial', 11)).pack(anchor=tk.W, **pad)
        self.keep_var = tk.StringVar(value='keep_unprocessed')
        keep_frame = tk.Frame(self.dialog)
        keep_frame.pack(anchor=tk.W, **pad)
        for key, desc in KEEP_STRATEGIES.items():
            tk.Radiobutton(keep_frame, text=desc, variable=self.keep_var, value=key, font=('Arial', 10)).pack(anchor=tk.W)

        tk.Label(self.dialog, text='启用银行:', font=('Arial', 11)).pack(anchor=tk.W, **pad)
        self.bank_vars = {}
        bank_frame = tk.Frame(self.dialog)
        bank_frame.pack(anchor=tk.W, **pad)
        for bank in BANK_PREFIXES:
            var = tk.BooleanVar(value=True)
            self.bank_vars[bank] = var
            tk.Checkbutton(bank_frame, text=bank, variable=var, font=('Arial', 10)).pack(side=tk.LEFT, padx=5)

        tk.Label(self.dialog, text='增量合并:', font=('Arial', 11)).pack(anchor=tk.W, **pad)
        self.incremental_var = tk.BooleanVar(value=True)
        inc_frame = tk.Frame(self.dialog)
        inc_frame.pack(anchor=tk.W, **pad)
        tk.Radiobutton(inc_frame, text='启用（推荐）', variable=self.incremental_var, value=True, font=('Arial', 10)).pack(side=tk.LEFT, padx=10)
        tk.Radiobutton(inc_frame, text='禁用（全量覆盖）', variable=self.incremental_var, value=False, font=('Arial', 10)).pack(side=tk.LEFT)

        btn_frame = tk.Frame(self.dialog)
        btn_frame.pack(fill=tk.X, padx=10, pady=20)

        tk.Button(btn_frame, text='保存', width=12, command=self._save,
                  bg='#4CAF50', fg='white', font=('Arial', 11, 'bold')).pack(side=tk.RIGHT, padx=5)
        tk.Button(btn_frame, text='取消', width=12, command=self.dialog.destroy,
                  bg='#9E9E9E', fg='white', font=('Arial', 11)).pack(side=tk.RIGHT)

    def _load_preset(self, preset):
        self.name_var.set(preset.get('name', ''))
        self.desc_text.delete(1.0, tk.END)
        self.desc_text.insert(1.0, preset.get('description', ''))
        self.output_var.set(preset.get('output_dir', ''))
        self.start_date_var.set(preset.get('start_date', ''))
        self.end_date_var.set(preset.get('end_date', ''))
        self.keep_var.set(preset.get('keep_strategy', 'keep_unprocessed'))
        self.incremental_var.set(preset.get('incremental', True))

        enabled_banks = preset.get('enabled_banks', BANK_PREFIXES)
        for bank in BANK_PREFIXES:
            self.bank_vars[bank].set(bank in enabled_banks)

    def _save(self):
        name = self.name_var.get().strip()
        if not name:
            messagebox.showerror('错误', '请输入预设名称', parent=self.dialog)
            return

        start_date = self.start_date_var.get().strip()
        end_date = self.end_date_var.get().strip()

        def validate_date(d):
            if not d:
                return True
            try:
                datetime.strptime(d, '%Y-%m-%d')
                return True
            except ValueError:
                return False

        if start_date and not validate_date(start_date):
            messagebox.showerror('错误', '开始日期格式错误，请使用 YYYY-MM-DD 格式', parent=self.dialog)
            return

        if end_date and not validate_date(end_date):
            messagebox.showerror('错误', '结束日期格式错误，请使用 YYYY-MM-DD 格式', parent=self.dialog)
            return

        enabled_banks = [bank for bank, var in self.bank_vars.items() if var.get()]
        if not enabled_banks:
            messagebox.showerror('错误', '请至少选择一个银行', parent=self.dialog)
            return

        preset_data = {
            'name': name,
            'description': self.desc_text.get(1.0, tk.END).strip(),
            'output_dir': self.output_var.get().strip(),
            'start_date': start_date,
            'end_date': end_date,
            'keep_strategy': self.keep_var.get(),
            'enabled_banks': enabled_banks,
            'incremental': self.incremental_var.get(),
        }

        if self.preset:
            preset_data['preset_id'] = self.preset['preset_id']

        preset_id = save_preset(preset_data, self.script_dir)

        messagebox.showinfo('成功', f'预设已保存\nID: {preset_id}', parent=self.dialog)
        self.dialog.destroy()

        if self.on_saved:
            self.on_saved()


def _list_presets_cli(presets):
    if not presets:
        print('\n暂无预设配置')
        return

    print('\n' + '-' * 80)
    print(f'{"ID":<22}{"名称":<20}{"描述":<25}{"更新时间":<20}')
    print('-' * 80)
    for p in presets:
        name = (p.get('name', '')[:18] + '..') if len(p.get('name', '')) > 20 else p.get('name', '')
        desc = (p.get('description', '')[:23] + '..') if len(p.get('description', '')) > 25 else p.get('description', '')
        print(f'{p["preset_id"]:<22}{name:<20}{desc:<25}{p.get("updated_at", ""):<20}')
    print('-' * 80)


def _save_preset_cli(script_dir):
    print('\n--- 保存新预设 ---')

    name = input('预设名称: ').strip()
    if not name:
        print('预设名称不能为空')
        return

    description = input('预设描述（可选）: ').strip()
    output_dir = input('输出目录（可选，留空使用默认）: ').strip()
    start_date = input('开始日期（YYYY-MM-DD，可选）: ').strip()
    end_date = input('结束日期（YYYY-MM-DD，可选）: ').strip()

    print('\n保留策略选项：')
    for key, desc in KEEP_STRATEGIES.items():
        print(f'  {key}: {desc}')
    keep_strategy = input(f'保留策略（默认: keep_unprocessed）: ').strip() or 'keep_unprocessed'
    if keep_strategy not in KEEP_STRATEGIES:
        print(f'无效的保留策略，使用默认: keep_unprocessed')
        keep_strategy = 'keep_unprocessed'

    print('\n可用银行列表：')
    for i, bank in enumerate(BANK_PREFIXES, 1):
        print(f'  {i}) {bank}')
    bank_input = input('选择启用的银行（输入编号，逗号分隔，回车全选）: ').strip()
    if bank_input:
        try:
            indices = [int(x.strip()) - 1 for x in bank_input.split(',')]
            enabled_banks = [BANK_PREFIXES[i] for i in indices if 0 <= i < len(BANK_PREFIXES)]
        except (ValueError, IndexError):
            print('输入无效，将启用所有银行')
            enabled_banks = list(BANK_PREFIXES)
    else:
        enabled_banks = list(BANK_PREFIXES)

    incremental_input = input('是否启用增量合并? (y/N): ').strip().lower()
    incremental = incremental_input == 'y'

    preset_data = {
        'name': name,
        'description': description,
        'output_dir': output_dir,
        'start_date': start_date,
        'end_date': end_date,
        'keep_strategy': keep_strategy,
        'enabled_banks': enabled_banks,
        'incremental': incremental,
    }

    preset_id = save_preset(preset_data, script_dir)
    print(f'\n✅ 预设已保存，ID: {preset_id}')


def _apply_preset_cli(script_dir):
    presets = list_presets(script_dir)
    if not presets:
        print('\n暂无预设可用')
        return

    _list_presets_cli(presets)
    preset_id = input('\n请输入要应用的预设ID: ').strip()

    preset = load_preset(preset_id, script_dir)
    if not preset:
        print(f'❌ 未找到预设: {preset_id}')
        return

    print(f'\n即将应用预设: {preset.get("name", "")}')
    _print_preset_detail(preset)

    folder = ask_directory('请选择银行流水文件夹')
    if not folder:
        print('未选择文件夹，取消应用')
        return

    confirm = input(f'\n确认应用预设处理文件夹「{folder}」? (Y/n): ').strip().lower()
    if confirm and confirm != 'y':
        print('已取消')
        return

    with AuditLogger('preset_pipeline', script_dir) as audit:
        audit.record_input(folder)
        audit.set_extra_info({'preset_id': preset_id, 'preset_name': preset.get('name', '')})

        result = apply_preset_to_pipeline(preset, folder, script_dir)
        audit.record_result(result)

        msg = format_result_message(result)
        msg += f'\n\n审计编号: {audit.audit_id}'
        msg += f'\n预设: {preset.get("name", "")} ({preset_id})'
        show_info('完成' if result.all_rows else '提示', msg)


def _delete_preset_cli(script_dir):
    presets = list_presets(script_dir)
    if not presets:
        print('\n暂无预设可删除')
        return

    _list_presets_cli(presets)
    preset_id = input('\n请输入要删除的预设ID: ').strip()

    preset = load_preset(preset_id, script_dir)
    if not preset:
        print(f'❌ 未找到预设: {preset_id}')
        return

    confirm = input(f'确认删除预设「{preset.get("name", "")}」? (y/N): ').strip().lower()
    if confirm != 'y':
        print('已取消')
        return

    if delete_preset(preset_id, script_dir):
        print('✅ 预设已删除')
    else:
        print('❌ 删除失败')


def _set_default_preset_cli(script_dir):
    presets = list_presets(script_dir)
    if not presets:
        print('\n暂无预设')
        return

    _list_presets_cli(presets)
    preset_id = input('\n请输入要设为默认的预设ID: ').strip()

    preset = load_preset(preset_id, script_dir)
    if not preset:
        print(f'❌ 未找到预设: {preset_id}')
        return

    set_default_preset(preset_id, script_dir)
    print(f'✅ 已将「{preset.get("name", "")}」设为默认预设')


def _show_preset_detail_cli(script_dir):
    presets = list_presets(script_dir)
    if not presets:
        print('\n暂无预设')
        return

    _list_presets_cli(presets)
    preset_id = input('\n请输入要查看的预设ID: ').strip()

    preset = load_preset(preset_id, script_dir)
    if not preset:
        print(f'❌ 未找到预设: {preset_id}')
        return

    _print_preset_detail(preset)


def _print_preset_detail(preset):
    print('\n' + '=' * 50)
    print(f'预设名称: {preset.get("name", "")}')
    print(f'预设ID: {preset.get("preset_id", "")}')
    print(f'描述: {preset.get("description", "无")}')
    print('-' * 50)
    print(f'输出目录: {preset.get("output_dir", "默认")}')
    print(f'开始日期: {preset.get("start_date", "不限制")}')
    print(f'结束日期: {preset.get("end_date", "不限制")}')
    print(f'保留策略: {KEEP_STRATEGIES.get(preset.get("keep_strategy"), "未知")}')
    print(f'启用银行: {", ".join(preset.get("enabled_banks", []))}')
    print(f'增量合并: {"是" if preset.get("incremental", True) else "否"}')
    print('-' * 50)
    print(f'创建时间: {preset.get("created_at", "")}')
    print(f'更新时间: {preset.get("updated_at", "")}')
    print('=' * 50)


if __name__ == '__main__':
    main()
