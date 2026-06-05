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
from datetime import datetime

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


# 根据 tkinter 是否可用，选择交互方式
if HAS_TKINTER:
    ask_directory = gui_askdirectory
    show_info = gui_showinfo
    show_warning = gui_showwarning
else:
    ask_directory = cli_askdirectory
    show_info = cli_showinfo
    show_warning = cli_showwarning


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

    target = str(bank_account).strip()
    tmp_path = None
    try:
        wb, tmp_path = open_workbook_compat(lookup_file)
        ws = wb.active
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=2, max_col=2):
            cell = row[0]
            if cell.value is not None and str(cell.value).strip() == target:
                subject = ws.cell(row=cell.row, column=1).value
                wb.close()
                cleanup_temp_file(tmp_path)
                logger.debug('银行账号「%s」匹配到主体: %s', target, subject)
                return subject if subject else ''
        wb.close()
        logger.warning('银行账号「%s」在查找表中未找到对应主体', target)
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

def main():
    # ── 初始化日志 ──
    setup_logging()
    logger = get_logger()
    logger.info('========== 银行流水检验工具启动 ==========')

    # ── 选择文件夹（自动适配 GUI / 命令行） ──
    folder = ask_directory('请选择银行流水文件夹')
    if not folder:
        show_info('提示', '未选择文件夹，程序退出。')
        logger.info('用户未选择文件夹，程序退出')
        return

    logger.info('用户选择文件夹: %s', folder)
    script_dir = get_script_dir()

    # ── 查找主体映射表 ──
    lookup_file = find_lookup_file(script_dir)
    if not lookup_file:
        show_warning(
            '警告',
            '在程序所在目录下未找到主体查找表文件，\n"主体"列将为空。\n'
            '建议将查找表文件命名为"主体查找表.xlsx"并放在程序所在目录下。'
        )
        logger.warning('未找到主体查找表，"主体"列将为空')

    # ── 复制文件夹 ──
    folder_name = os.path.basename(folder.rstrip('/\\'))
    parent_dir = os.path.dirname(folder.rstrip('/\\'))
    new_folder = os.path.join(parent_dir, f"{folder_name}＋检验版")

    if os.path.exists(new_folder):
        logger.info('＋检验版文件夹已存在，先删除: %s', new_folder)
        shutil.rmtree(new_folder)
    shutil.copytree(folder, new_folder)
    logger.info('已复制文件夹为＋检验版: %s', new_folder)

    # ── 扫描 Excel 文件 ──
    excel_files = scan_excel_files(new_folder)
    if not excel_files:
        show_info('提示', '文件夹中未发现任何 Excel 文件。')
        logger.warning('检验版文件夹中未发现任何 Excel 文件')
        return

    # ── 逐文件处理 ──
    all_rows = []
    processed_files = []       # 成功处理的文件
    unprocessed_files = []     # 无法识别银行的文件 → 保留
    error_files = []           # 识别成功但处理出错的文件 → 也需要删除

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

    # ── 删除除无法识别银行外所有的文件 ──
    # 根据要求，只要文件名前缀被识别为对应银行（无论是成功处理还是出错），均需删除
    for filepath in excel_files:
        if filepath not in unprocessed_files:
            try:
                os.remove(filepath)
                logger.debug('已删除文件: %s', filepath)
            except OSError as e:
                logger.error('删除文件「%s」失败: %s', filepath, e)

    # ── 输出总表 ──
    if all_rows:
        columns = [
            '唯一id', '银行', '银行账号', '主体', '交易日期',
            '付款', '收款', '摘要', '对方户名', '余额', '交易流水号',
        ]
        df = pd.DataFrame(all_rows, columns=columns)
        output_path = os.path.join(script_dir, '银行流水总表.xlsx')
        df.to_excel(output_path, index=False, engine='openpyxl')
        logger.info('总表输出完成: %s（共 %d 条记录）', output_path, len(all_rows))

        # ── 汇总提示 ──
        msg = (
            f'处理完成！\n\n'
            f'已处理文件数：{len(processed_files)}\n'
            f'提取记录数：{len(all_rows)}\n'
            f'总表路径：{output_path}'
        )
        if unprocessed_files:
            names = '\n  '.join(os.path.basename(f) for f in unprocessed_files)
            msg += f'\n\n无法识别的文件（{len(unprocessed_files)} 个，已保留）：\n  {names}'
        if error_files:
            err_info = '\n  '.join(f'{os.path.basename(f)}: {e}' for f, e in error_files)
            msg += f'\n\n处理出错的文件（{len(error_files)} 个，已保留）：\n  {err_info}'

        show_info('完成', msg)
    else:
        msg = '未提取到任何银行流水记录。'
        if unprocessed_files:
            names = '\n  '.join(os.path.basename(f) for f in unprocessed_files)
            msg += f'\n\n无法识别的文件（{len(unprocessed_files)} 个，已保留）：\n  {names}'
        if error_files:
            err_info = '\n  '.join(f'{os.path.basename(f)}: {e}' for f, e in error_files)
            msg += f'\n\n处理出错的文件（{len(error_files)} 个，已保留）：\n  {err_info}'
        show_info('提示', msg)
        logger.warning('未提取到任何银行流水记录')

    logger.info('========== 银行流水检验工具运行结束 ==========')


if __name__ == '__main__':
    main()
