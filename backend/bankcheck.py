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

银行解析规则通过 bank_rules.yaml 配置文件管理，支持热更新。
业务人员可通过修改配置文件增删银行规则，无需修改 Python 代码。
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
import yaml

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


def _normalize_path(path):
    """规范化路径：展开用户目录、去除引号、转为绝对路径、规范化斜杠"""
    if not path:
        return ''
    path = path.strip().strip('"').strip("'")
    path = os.path.expanduser(path)
    path = os.path.abspath(path)
    path = os.path.normpath(path)
    return path


_cli_default_dir = None
_cli_default_file = None


def set_cli_default_dir(path):
    """设置全局默认文件夹路径，供 cli_askdirectory 使用"""
    global _cli_default_dir
    _cli_default_dir = _normalize_path(path) if path else None


def get_cli_default_dir():
    """获取全局默认文件夹路径"""
    return _cli_default_dir


def set_cli_default_file(path):
    """设置全局默认文件路径，供 cli_askfile 使用"""
    global _cli_default_file
    _cli_default_file = _normalize_path(path) if path else None


def get_cli_default_file():
    """获取全局默认文件路径"""
    return _cli_default_file


def cli_askdirectory(title='请选择文件夹', default_path=None, max_retries=3):
    """
    命令行模式下让用户输入文件夹路径。

    Args:
        title: 提示标题
        default_path: 默认路径，若提供且有效则直接返回，无需用户输入；
                      若为 None 则尝试使用全局默认路径（set_cli_default_dir 设置）
        max_retries: 最大重试次数

    Returns:
        有效的文件夹绝对路径，或空字符串（用户取消/重试次数耗尽）
    """
    if default_path is None:
        default_path = _cli_default_dir

    if default_path:
        normalized = _normalize_path(default_path)
        if os.path.isdir(normalized):
            return normalized
        else:
            print(f'\n[提示] 默认路径无效: {default_path}')

    print(f'\n{title}')
    print('（输入 q 取消，支持 ~ 和相对路径）')

    for attempt in range(max_retries):
        prompt = '请输入文件夹路径: ' if attempt == 0 else f'请重新输入文件夹路径（还可重试 {max_retries - attempt} 次）: '
        raw = input(prompt).strip()

        if raw.lower() == 'q':
            print('已取消选择。')
            return ''

        if not raw:
            print('❌ 路径不能为空，请重新输入。')
            continue

        normalized = _normalize_path(raw)

        if not os.path.exists(normalized):
            print(f'❌ 路径不存在: {normalized}')
            continue

        if not os.path.isdir(normalized):
            print(f'❌ 路径不是文件夹: {normalized}')
            continue

        return normalized

    print(f'❌ 已超过最大重试次数（{max_retries} 次）。')
    return ''


def cli_showinfo(title, message):
    """命令行模式下打印信息"""
    print(f'\n[{title}] {message}')


def cli_showwarning(title, message):
    """命令行模式下打印警告"""
    print(f'\n[警告 - {title}] {message}')


def cli_askfile(title='请选择文件', default_path=None, max_retries=3):
    """
    命令行模式下让用户输入文件路径。

    Args:
        title: 提示标题
        default_path: 默认路径，若提供且有效则直接返回，无需用户输入；
                      若为 None 则尝试使用全局默认路径（set_cli_default_file 设置）
        max_retries: 最大重试次数

    Returns:
        有效的文件绝对路径，或空字符串（用户取消/重试次数耗尽）
    """
    if default_path is None:
        default_path = _cli_default_file

    if default_path:
        normalized = _normalize_path(default_path)
        if os.path.isfile(normalized):
            return normalized
        else:
            print(f'\n[提示] 默认路径无效: {default_path}')

    print(f'\n{title}')
    print('（输入 q 取消，支持 ~ 和相对路径）')

    for attempt in range(max_retries):
        prompt = '请输入文件路径: ' if attempt == 0 else f'请重新输入文件路径（还可重试 {max_retries - attempt} 次）: '
        raw = input(prompt).strip()

        if raw.lower() == 'q':
            print('已取消选择。')
            return ''

        if not raw:
            print('❌ 路径不能为空，请重新输入。')
            continue

        normalized = _normalize_path(raw)

        if not os.path.exists(normalized):
            print(f'❌ 路径不存在: {normalized}')
            continue

        if not os.path.isfile(normalized):
            print(f'❌ 路径不是文件: {normalized}')
            continue

        return normalized

    print(f'❌ 已超过最大重试次数（{max_retries} 次）。')
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
    print('  10)主体汇总分析：按主体/银行/月份统计收支净额与笔数')
    print('  11)余额连续性校验：逐笔核对余额连续性，识别断裂或跳变')
    print('  12)重复交易检测：跨文件去重与疑似重复标记')
    choice = input('请输入选项（1-12，直接回车默认为 1）: ').strip()
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
    elif choice == '10':
        return 'subject_summary'
    elif choice == '11':
        return 'balance_check'
    elif choice == '12':
        return 'duplicate_check'
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
    root.geometry('480x680')
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
        ('主体汇总', 'subject_summary', '按主体/银行/月份统计', '#3F51B5'),
        ('余额校验', 'balance_check', '逐笔核对余额连续性', '#8BC34A'),
        ('重复检测', 'duplicate_check', '跨文件去重与疑似重复标记', '#F44336'),
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


# ──────────────────────────────────────────────
# 进度条窗口与进度回调机制
# ──────────────────────────────────────────────

@dataclass
class ProgressInfo:
    """进度信息数据类"""
    stage: str = ''
    stage_index: int = 0
    total_stages: int = 0
    percent: int = 0
    message: str = ''
    current_file: str = ''
    processed_files: int = 0
    total_files: int = 0
    processed_records: int = 0
    extra: Dict[str, Any] = field(default_factory=dict)


class ProgressWindow:
    """
    GUI 进度条窗口，在批量处理时显示当前进度和阶段状态。
    使用 after() 方法确保 UI 更新在主线程中执行，避免线程安全问题。
    """

    STAGES = [
        ('初始化', '正在初始化处理环境...'),
        ('扫描文件', '正在扫描文件夹中的 Excel 文件...'),
        ('识别银行', '正在识别各文件的银行类型...'),
        ('解析文件', '正在解析银行流水文件...'),
        ('合并数据', '正在合并并去重数据...'),
        ('导出总表', '正在导出总表文件...'),
        ('黑白名单', '正在应用对方户名黑白名单...'),
        ('数据库', '正在写入数据库...'),
        ('生成报告', '正在生成汇总与检验报告...'),
        ('完成', '处理完成！'),
    ]

    def __init__(self, title='处理进度', parent=None):
        if not HAS_TKINTER or tk is None:
            self.root = None
            return

        try:
            if parent:
                self.root = tk.Toplevel(parent)
            else:
                self.root = tk.Tk()
                self.root.title(title)
        except Exception:
            self.root = None
            return

        self.root.title(title)
        self.root.geometry('560x420')
        self.root.resizable(False, False)
        self.root.attributes('-topmost', True)

        try:
            self.root.option_add('*Font', 'Arial 10')
        except Exception:
            pass

        self._build_ui()
        self._cancelled = False
        self._closed = False

    def _build_ui(self):
        """构建 UI 组件"""
        main_frame = tk.Frame(self.root, padx=20, pady=15)
        main_frame.pack(fill=tk.BOTH, expand=True)

        title_label = tk.Label(
            main_frame,
            text='银行流水批量处理',
            font=('Arial', 16, 'bold'),
            fg='#2c3e50',
        )
        title_label.pack(pady=(0, 5))

        self.stage_label = tk.Label(
            main_frame,
            text='准备中...',
            font=('Arial', 12, 'bold'),
            fg='#3498db',
            anchor='w',
        )
        self.stage_label.pack(fill=tk.X, pady=(10, 3))

        from tkinter import ttk
        self.progress_bar = ttk.Progressbar(
            main_frame,
            orient='horizontal',
            length=520,
            mode='determinate',
            maximum=100,
        )
        self.progress_bar.pack(fill=tk.X, pady=5)

        self.percent_label = tk.Label(
            main_frame,
            text='0%',
            font=('Arial', 11),
            fg='#7f8c8d',
            anchor='e',
        )
        self.percent_label.pack(fill=tk.X, pady=(0, 10))

        self.message_label = tk.Label(
            main_frame,
            text='正在准备处理任务...',
            font=('Arial', 10),
            fg='#555',
            anchor='w',
            wraplength=520,
            justify='left',
        )
        self.message_label.pack(fill=tk.X, pady=(5, 5))

        file_frame = tk.Frame(main_frame)
        file_frame.pack(fill=tk.X, pady=(5, 5))
        tk.Label(
            file_frame,
            text='当前文件:',
            font=('Arial', 9, 'bold'),
            fg='#666',
        ).pack(side=tk.LEFT, anchor='w')
        self.current_file_label = tk.Label(
            file_frame,
            text='-',
            font=('Arial', 9),
            fg='#888',
            anchor='w',
        )
        self.current_file_label.pack(side=tk.LEFT, padx=(5, 0), anchor='w')

        stats_frame = tk.LabelFrame(main_frame, text='处理统计', padx=10, pady=8)
        stats_frame.pack(fill=tk.X, pady=(10, 5))

        self.stats_labels = {}
        stats = [
            ('files', '文件进度', '0 / 0'),
            ('records', '记录数', '0'),
            ('success', '成功文件', '0'),
            ('errors', '出错文件', '0'),
        ]
        for i, (key, name, default) in enumerate(stats):
            col = i % 2
            row = i // 2
            frm = tk.Frame(stats_frame)
            frm.grid(row=row, column=col, sticky='w', padx=10, pady=3)
            tk.Label(
                frm,
                text=f'{name}:',
                font=('Arial', 9, 'bold'),
                fg='#555',
                width=10,
                anchor='w',
            ).pack(side=tk.LEFT)
            lbl = tk.Label(
                frm,
                text=default,
                font=('Arial', 9),
                fg='#333',
                anchor='w',
            )
            lbl.pack(side=tk.LEFT)
            self.stats_labels[key] = lbl

        btn_frame = tk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(15, 0))
        self.cancel_btn = tk.Button(
            btn_frame,
            text='取消处理',
            width=12,
            command=self._on_cancel,
            bg='#e74c3c',
            fg='white',
            font=('Arial', 10, 'bold'),
        )
        self.cancel_btn.pack(side=tk.RIGHT)

        self.root.protocol('WM_DELETE_WINDOW', self._on_close)

    def _on_cancel(self):
        """点击取消按钮"""
        self._cancelled = True
        self.cancel_btn.config(state=tk.DISABLED, text='取消中...')
        self.message_label.config(text='正在取消，请稍候...', fg='#e74c3c')

    def _on_close(self):
        """关闭窗口"""
        self._cancelled = True
        self._closed = True
        try:
            self.root.destroy()
        except Exception:
            pass

    def is_cancelled(self) -> bool:
        return self._cancelled

    def update_progress(self, info: ProgressInfo):
        """
        线程安全地更新进度。通过 after() 将 UI 更新投递到主线程。
        """
        if self.root is None or self._closed:
            return
        try:
            self.root.after(0, self._do_update, info)
        except Exception:
            pass

    def _do_update(self, info: ProgressInfo):
        """实际执行 UI 更新（在主线程中）"""
        if self._closed:
            return

        try:
            if info.stage:
                self.stage_label.config(text=f'【{info.stage_index + 1}/{info.total_stages}】{info.stage}')

            self.progress_bar['value'] = info.percent
            self.percent_label.config(text=f'{info.percent}%')

            if info.message:
                self.message_label.config(text=info.message, fg='#555')

            if info.current_file:
                display_name = os.path.basename(info.current_file)
                if len(display_name) > 45:
                    display_name = display_name[:42] + '...'
                self.current_file_label.config(text=display_name, fg='#2980b9')

            if info.total_files > 0:
                self.stats_labels['files'].config(text=f'{info.processed_files} / {info.total_files}')

            if info.processed_records > 0:
                self.stats_labels['records'].config(text=f'{info.processed_records:,}')

            if 'success_count' in info.extra:
                self.stats_labels['success'].config(text=str(info.extra['success_count']))

            if 'error_count' in info.extra:
                self.stats_labels['errors'].config(text=str(info.extra['error_count']),
                                                    fg='#e74c3c' if info.extra['error_count'] > 0 else '#333')

            self.root.update_idletasks()
        except Exception:
            pass

    def set_completed(self, final_message: str = '处理完成！'):
        """标记为完成状态"""
        if self.root is None or self._closed:
            return
        try:
            self.root.after(0, self._do_completed, final_message)
        except Exception:
            pass

    def _do_completed(self, final_message: str):
        self.progress_bar['value'] = 100
        self.percent_label.config(text='100%')
        self.stage_label.config(text=f'【{len(self.STAGES)}/{len(self.STAGES)}】完成', fg='#27ae60')
        self.message_label.config(text=final_message, fg='#27ae60')
        self.cancel_btn.config(text='关闭', command=self._on_close, bg='#27ae60')
        try:
            self.root.update_idletasks()
        except Exception:
            pass

    def set_error(self, error_message: str):
        """标记为错误状态"""
        if self.root is None or self._closed:
            return
        try:
            self.root.after(0, self._do_error, error_message)
        except Exception:
            pass

    def _do_error(self, error_message: str):
        self.stage_label.config(text='处理出错', fg='#e74c3c')
        self.message_label.config(text=error_message, fg='#e74c3c')
        self.cancel_btn.config(text='关闭', command=self._on_close, bg='#e74c3c')
        try:
            self.root.update_idletasks()
        except Exception:
            pass

    def show(self):
        """显示窗口"""
        if self.root is None:
            return
        try:
            self.root.update()
            self.root.deiconify()
            self.root.lift()
        except Exception:
            pass

    def close(self):
        """关闭窗口"""
        self._closed = True
        if self.root is not None:
            try:
                self.root.after(0, self.root.destroy)
            except Exception:
                pass

    def wait(self, timeout_ms: int = 100):
        """等待指定时间，同时处理 UI 事件"""
        if self.root is None or self._closed:
            return
        try:
            self.root.update()
            self.root.after(timeout_ms)
        except Exception:
            pass


class ResultDetailWindow:
    """
    GUI 结果详情窗口，展示处理完成后以结构化方式展示统计摘要和文件清单。
    包含按银行、按主体统计摘要，以及成功/未处理/失败三类文件清单。
    """

    def __init__(self, result: ProcessingResult, title='处理结果详情', parent=None):
        if not HAS_TKINTER or tk is None:
            self.root = None
            return

        try:
            if parent:
                self.root = tk.Toplevel(parent)
            else:
                self.root = tk.Tk()
                self.root.title(title)
        except Exception:
            self.root = None
            return

        self.root.title(title)
        self.root.geometry('780x620')
        self.root.minsize(640, 480)
        self.root.attributes('-topmost', True)

        try:
            self.root.option_add('*Font', 'Arial 10')
        except Exception:
            pass

        self.result = result
        self.summary = build_processing_summary(result.file_process_details)
        self._closed = False

        self._build_ui()

    def _build_ui(self):
        """构建 UI 组件"""
        from tkinter import ttk

        main_frame = tk.Frame(self.root, padx=15, pady=12)
        main_frame.pack(fill=tk.BOTH, expand=True)

        title_label = tk.Label(
            main_frame,
            text='处理结果详情',
            font=('Arial', 16, 'bold'),
            fg='#2c3e50',
        )
        title_label.pack(pady=(0, 8))

        self._build_summary_section(main_frame)
        self._build_tabs(main_frame)

        btn_frame = tk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(10, 0))
        close_btn = tk.Button(
            btn_frame,
            text='关闭',
            width=12,
            command=self._on_close,
            bg='#3498db',
            fg='white',
            font=('Arial', 10, 'bold'),
        )
        close_btn.pack(side=tk.RIGHT)

        self.root.protocol('WM_DELETE_WINDOW', self._on_close)

    def _build_summary_section(self, parent):
        """构建统计摘要区域"""
        total = self.summary['total']
        by_bank = self.summary['by_bank']
        by_subject = self.summary['by_subject']

        summary_frame = tk.LabelFrame(parent, text='总体统计摘要', padx=12, pady=8)
        summary_frame.pack(fill=tk.X, pady=(0, 8))

        stats = [
            ('总文件数', total['files'], '#2c3e50'),
            ('银行数', total['banks'], '#3498db'),
            ('主体数', total['subjects'], '#9b59b6'),
            ('成功文件', total['success'], '#27ae60'),
            ('未处理文件', total['unprocessed'], '#f39c12'),
            ('失败文件', total['error'], '#e74c3c'),
            ('提取记录', f"{total['records']:,}", '#2c3e50'),
            ('跳过行', f"{total['skipped_rows']:,}", '#7f8c8d'),
        ]
        for i, (name, value, color) in enumerate(stats):
            col = i % 4
            row = i // 4
            frm = tk.Frame(summary_frame)
            frm.grid(row=row, column=col, sticky='w', padx=15, pady=3)
            tk.Label(
                frm,
                text=f'{name}:',
                font=('Arial', 9, 'bold'),
                fg='#555',
                width=10,
                anchor='w',
            ).pack(side=tk.LEFT)
            tk.Label(
                frm,
                text=str(value),
                font=('Arial', 10, 'bold'),
                fg=color,
                anchor='w',
            ).pack(side=tk.LEFT)

    def _build_tabs(self, parent):
        """构建文件清单标签页"""
        from tkinter import ttk

        tab_frame = tk.LabelFrame(parent, text='文件清单', padx=8, pady=6)
        tab_frame.pack(fill=tk.BOTH, expand=True, pady=(5, 0))

        notebook = ttk.Notebook(tab_frame)
        notebook.pack(fill=tk.BOTH, expand=True)

        success_files = self.summary['by_status']['success']
        unprocessed_files = self.summary['by_status']['unprocessed']
        error_files = self.summary['by_status']['error']

        self._build_file_tab(notebook, f'成功文件 ({len(success_files)})', success_files, 'success')
        self._build_file_tab(notebook, f'未处理文件 ({len(unprocessed_files)})', unprocessed_files, 'unprocessed')
        self._build_file_tab(notebook, f'失败文件 ({len(error_files)})', error_files, 'error')

    def _build_file_tab(self, notebook, tab_title, files, status_type):
        """构建单个文件清单标签页"""
        from tkinter import ttk

        frame = tk.Frame(notebook, padx=5, pady=5)
        notebook.add(frame, text=tab_title)

        columns = ('file_name', 'bank_name', 'subject', 'records', 'skipped_rows')
        tree = ttk.Treeview(frame, columns=columns, show='headings', height=12)

        tree.heading('file_name', text='文件名')
        tree.heading('bank_name', text='银行')
        tree.heading('subject', text='主体')
        tree.heading('records', text='提取记录')
        tree.heading('skipped_rows', text='跳过行')

        tree.column('file_name', width=280, anchor='w')
        tree.column('bank_name', width=120, anchor='center')
        tree.column('subject', width=140, anchor='center')
        tree.column('records', width=90, anchor='e')
        tree.column('skipped_rows', width=80, anchor='e')

        scrollbar = ttk.Scrollbar(frame, orient='vertical', command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)

        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        tag_colors = {
            'success': '#e8f5e9',
            'unprocessed': '#fff8e1',
            'error': '#ffebee',
        }
        tree.tag_configure(status_type, background=tag_colors.get(status_type, 'white'))

        for d in files:
            values = (
                d.file_name or os.path.basename(d.file_path),
                d.bank_name or '-',
                d.subject or '-',
                f"{d.extracted_records:,}" if d.extracted_records else '-',
                f"{d.skipped_rows:,}" if d.skipped_rows else '-',
            )
            tree.insert('', 'end', values=values, tags=(status_type,))

        if not files:
            tree.insert('', 'end', values=('（无）', '', '', '', ''), tags=('empty',))
            tree.tag_configure('empty', foreground='#999')

    def _on_close(self):
        """关闭窗口"""
        self._closed = True
        try:
            self.root.destroy()
        except Exception:
            pass

    def show(self):
        """显示窗口并进入事件循环"""
        if self.root is None:
            return
        try:
            self.root.update()
            self.root.deiconify()
            self.root.lift()
            self.root.mainloop()
        except Exception:
            pass

    def is_closed(self) -> bool:
        return self._closed


def create_progress_callback(progress_window: Optional[ProgressWindow]):
    """
    创建进度回调函数。
    返回一个可调用对象，接收 ProgressInfo 并更新窗口。
    """
    if progress_window is None:
        def _noop_callback(*args, **kwargs):
            pass
        return _noop_callback

    def _callback(info: ProgressInfo):
        progress_window.update_progress(info)
        if progress_window.is_cancelled():
            raise RuntimeError('用户取消了操作')

    return _callback


def show_result_detail_dialog(result: ProcessingResult, parent=None) -> bool:
    """
    显示结果详情对话框。GUI 模式下显示结构化窗口，CLI 模式下打印文本消息。

    Returns:
        True 表示成功显示了详情窗口，False 表示退化为文本消息
    """
    if not HAS_TKINTER or tk is None:
        return False

    if not result.file_process_details:
        return False

    try:
        win = ResultDetailWindow(result, title='处理结果详情', parent=parent)
        if win.root is not None:
            win.show()
            return True
    except Exception as e:
        logger = get_logger()
        logger.warning('结果详情窗口显示失败: %s', e)

    return False


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


def get_program_dir():
    """
    获取程序所在目录（只读目录，用于读取配置文件）。
    与 get_script_dir() 功能相同，但语义更清晰。
    """
    return get_script_dir()


def is_writable(dir_path):
    """
    检测目录是否具有写入权限。

    Args:
        dir_path: 待检测的目录路径

    Returns:
        bool: True 表示可写，False 表示不可写
    """
    if not os.path.isdir(dir_path):
        return False
    try:
        test_file = os.path.join(dir_path, '.bankcheck_write_test_' + uuid.uuid4().hex[:8])
        with open(test_file, 'w') as f:
            f.write('test')
        os.remove(test_file)
        return True
    except (OSError, IOError):
        return False


def get_user_data_dir():
    """
    获取用户可写的应用数据目录。
    跨平台策略：
    - Windows: %APPDATA%\\bankcheck 或 %USERPROFILE%\\AppData\\Roaming\\bankcheck
    - macOS: ~/Library/Application Support/bankcheck
    - Linux: ~/.bankcheck

    Returns:
        str: 用户数据目录的绝对路径
    """
    app_name = 'bankcheck'
    if sys.platform.startswith('win'):
        base_dir = os.environ.get('APPDATA')
        if not base_dir:
            base_dir = os.path.expanduser('~\\AppData\\Roaming')
        return os.path.join(base_dir, app_name)
    elif sys.platform == 'darwin':
        return os.path.join(os.path.expanduser('~/Library/Application Support'), app_name)
    else:
        return os.path.join(os.path.expanduser('~'), '.' + app_name)


def get_writable_dir():
    """
    获取可写的工作目录。
    策略：
    1. 优先尝试使用程序目录（get_program_dir()）
    2. 如果程序目录不可写（如安装在 Program Files、/Applications 等受保护目录），
       则使用用户数据目录（get_user_data_dir()）

    Returns:
        str: 可写目录的绝对路径
    """
    program_dir = get_program_dir()
    if is_writable(program_dir):
        return program_dir
    user_data_dir = get_user_data_dir()
    os.makedirs(user_data_dir, exist_ok=True)
    return user_data_dir


def get_output_dir(subdir=None):
    """
    获取输出文件目录，用于保存日志、总表、查找表、数据库等可写文件。

    Args:
        subdir: 可选子目录名称，如 'logs'、'history' 等

    Returns:
        str: 输出目录的绝对路径
    """
    base_dir = get_writable_dir()
    if subdir:
        output_dir = os.path.join(base_dir, subdir)
    else:
        output_dir = base_dir
    os.makedirs(output_dir, exist_ok=True)
    return output_dir


def setup_logging():
    """
    初始化日志系统。
    - 控制台输出 INFO 级别及以上日志
    - 日志文件（bankcheck.log）记录 DEBUG 级别及以上日志，
      文件保存在可写目录下（优先程序目录，否则用户数据目录）
    """
    log_dir = get_output_dir()
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
    fd, tmp_path = tempfile.mkstemp(suffix='.xlsx', prefix='bankcheck_')
    os.close(fd)
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
# 银行规则配置模块
# ──────────────────────────────────────────────

BANK_RULES_CONFIG_FILE = 'bank_rules.yaml'


@dataclass
class BankRule:
    """单个银行的解析规则"""
    bank_name: str
    account_cell: str
    start_row: int
    columns: Dict[str, int]
    payment_sign: str = 'negative'
    enabled: bool = True


class BankRuleConfig:
    """银行规则配置管理器 - 单例模式"""
    _instance = None
    _config_path = None
    _rules: Dict[str, BankRule] = field(default_factory=dict)
    _last_modified: float = 0.0

    def __new__(cls, config_path=None):
        if cls._instance is None:
            cls._instance = super(BankRuleConfig, cls).__new__(cls)
            cls._instance._initialized = False
        return cls._instance

    def __init__(self, config_path=None):
        if self._initialized:
            return
        self._initialized = True
        if config_path is None:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            config_path = os.path.join(script_dir, BANK_RULES_CONFIG_FILE)
        self._config_path = config_path
        self._rules = {}
        self._last_modified = 0.0
        self.load_config()

    def load_config(self):
        """加载配置文件，支持热更新"""
        logger = get_logger()
        if not os.path.exists(self._config_path):
            logger.error('银行规则配置文件不存在: %s', self._config_path)
            return False

        try:
            current_mtime = os.path.getmtime(self._config_path)
            if current_mtime == self._last_modified and self._rules:
                return True

            with open(self._config_path, 'r', encoding='utf-8') as f:
                config_data = yaml.safe_load(f)

            self._rules = {}
            banks_config = config_data.get('banks', [])
            for bank_config in banks_config:
                if not bank_config.get('enabled', True):
                    continue
                rule = BankRule(
                    bank_name=bank_config['bank_name'],
                    account_cell=bank_config['account_cell'],
                    start_row=bank_config['start_row'],
                    columns=bank_config['columns'],
                    payment_sign=bank_config.get('payment_sign', 'negative'),
                    enabled=bank_config.get('enabled', True),
                )
                self._rules[rule.bank_name] = rule

            self._last_modified = current_mtime
            logger.info('已加载 %d 个银行规则配置', len(self._rules))
            return True
        except Exception as e:
            logger.error('加载银行规则配置失败: %s', e, exc_info=True)
            return False

    def get_rule(self, bank_name: str) -> Optional[BankRule]:
        """根据银行名称获取规则，自动检查配置更新"""
        self.load_config()
        return self._rules.get(bank_name)

    def get_all_bank_names(self) -> List[str]:
        """获取所有已启用的银行名称列表，自动检查配置更新"""
        self.load_config()
        return list(self._rules.keys())

    def get_config_path(self) -> str:
        """获取配置文件路径"""
        return self._config_path


class GenericBankParser:
    """通用银行流水解析器 - 根据配置规则动态解析 Excel"""

    def __init__(self, rule: BankRule):
        self.rule = rule
        self.logger = get_logger()

    def parse(self, filepath: str, lookup_file: str) -> Tuple[List[Dict[str, Any]], FileProcessDetail]:
        """
        根据配置规则解析银行流水 Excel 文件

        Args:
            filepath: Excel 文件路径
            lookup_file: 主体查找表路径

        Returns:
            tuple: (解析后的记录列表, 文件处理详情)
        """
        self.logger.info('开始处理%s文件: %s', self.rule.bank_name, filepath)

        detail = FileProcessDetail(
            file_path=filepath,
            file_name=os.path.basename(filepath),
            bank_name=self.rule.bank_name,
            process_status='处理中',
        )

        wb, tmp_path = open_workbook_compat(filepath)
        try:
            ws = wb.active

            bank_account = ws[self.rule.account_cell].value
            if bank_account is None:
                self.logger.warning('文件「%s」%s 单元格为空，银行账号缺失',
                                    filepath, self.rule.account_cell)

            detail.bank_account = str(bank_account) if bank_account is not None else ''
            subject = get_subject(bank_account, lookup_file)
            detail.subject = subject

            rows = []
            columns = self.rule.columns
            start_row = self.rule.start_row
            detail.total_rows_in_excel = ws.max_row

            for row_idx in range(start_row, ws.max_row + 1):
                trade_date = ws.cell(row=row_idx, column=columns['trade_date']).value
                if trade_date is None:
                    raw_parts = []
                    for col_name, col_idx in columns.items():
                        try:
                            v = ws.cell(row=row_idx, column=col_idx).value
                            if v is not None:
                                raw_parts.append(f"{col_name}={v}")
                        except Exception:
                            pass
                    detail.skipped_rows += 1
                    detail.skipped_details.append(SkippedRowDetail(
                        file_path=filepath,
                        file_name=os.path.basename(filepath),
                        row_number=row_idx,
                        reason='交易日期为空',
                        raw_content='; '.join(raw_parts) if raw_parts else '(整行为空)',
                    ))
                    continue

                payment_val = ws.cell(row=row_idx, column=columns['payment']).value
                if is_numeric(payment_val):
                    payment = to_float(payment_val)
                    if self.rule.payment_sign == 'negative':
                        payment = -abs(payment)
                else:
                    payment = None

                receipt_val = ws.cell(row=row_idx, column=columns['receipt']).value
                receipt = to_float(receipt_val) if is_numeric(receipt_val) else None

                summary = ws.cell(row=row_idx, column=columns['summary']).value
                counterpart = ws.cell(row=row_idx, column=columns['counterpart']).value
                balance = ws.cell(row=row_idx, column=columns['balance']).value
                transaction_id = ws.cell(row=row_idx, column=columns['transaction_id']).value

                rows.append({
                    '唯一id': generate_unique_id(),
                    '银行': self.rule.bank_name,
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

            detail.extracted_records = len(rows)
            detail.process_status = '成功'
            wb.close()
            self.logger.info('%s文件处理完成，提取 %d 条记录，跳过 %d 行',
                             self.rule.bank_name, len(rows), detail.skipped_rows)
            return rows, detail
        except Exception as e:
            detail.process_status = '失败'
            detail.error_message = str(e)
            raise
        finally:
            cleanup_temp_file(tmp_path)


def _create_bank_processor(bank_name: str):
    """创建基于配置的银行处理器函数"""
    def processor(filepath, lookup_file):
        config = BankRuleConfig()
        rule = config.get_rule(bank_name)
        if rule is None:
            logger = get_logger()
            logger.error('未找到银行「%s」的解析规则', bank_name)
            return [], FileProcessDetail(
                file_path=filepath,
                file_name=os.path.basename(filepath),
                bank_name=bank_name,
                process_status='失败',
                error_message=f'未找到银行「{bank_name}」的解析规则',
            )
        parser = GenericBankParser(rule)
        return parser.parse(filepath, lookup_file)
    return processor


_bank_config_singleton = None


def get_bank_config():
    """获取银行规则配置单例"""
    global _bank_config_singleton
    if _bank_config_singleton is None:
        _bank_config_singleton = BankRuleConfig()
    return _bank_config_singleton


# ──────────────────────────────────────────────
# 银行识别
# ──────────────────────────────────────────────

def _get_bank_prefixes():
    """从配置获取银行前缀列表"""
    config = get_bank_config()
    return config.get_all_bank_names()


BANK_PREFIXES = _get_bank_prefixes()


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


def _find_lookup_in_dir(directory):
    """
    在指定目录下查找主体查找表 Excel 文件（内部辅助函数）。

    查找策略（按优先级）：
    1. 优先按文件名精确匹配 "主体查找表.xlsx" 或 "主体查找表.xls"
    2. 若未精确匹配到，回退到查找目录下唯一的 Excel 文件（排除输出总表和临时文件）
    """
    if not directory or not os.path.isdir(directory):
        return None

    # ── 策略 1：按文件名精确匹配 ──
    for name in LOOKUP_FILE_NAMES:
        candidate = os.path.join(directory, name)
        if os.path.isfile(candidate):
            return candidate

    # ── 策略 2：回退到唯一 Excel 文件 ──
    excel_exts = ('.xlsx', '.xls')
    exclude_names = {'银行流水总表.xlsx', '银行流水总表.xls'}
    excel_files = []
    try:
        for f in os.listdir(directory):
            if f.startswith('~$'):
                continue
            if f in exclude_names:
                continue
            if f.lower().endswith(excel_exts):
                excel_files.append(os.path.join(directory, f))
    except OSError:
        return None

    if len(excel_files) == 1:
        return excel_files[0]
    return None


def _copy_lookup_to_output(program_lookup_path, output_dir):
    """
    将程序目录下的查找表复制到用户可写的输出目录。

    Args:
        program_lookup_path: 程序目录下的查找表路径
        output_dir: 输出目录路径

    Returns:
        str: 复制后的目标路径，失败返回 None
    """
    logger = get_logger()
    if not program_lookup_path or not os.path.isfile(program_lookup_path):
        return None

    try:
        filename = os.path.basename(program_lookup_path)
        target_path = os.path.join(output_dir, filename)
        if os.path.exists(target_path):
            return target_path
        os.makedirs(output_dir, exist_ok=True)
        shutil.copy2(program_lookup_path, target_path)
        logger.info('已将查找表从程序目录复制到输出目录: %s -> %s',
                    program_lookup_path, target_path)
        return target_path
    except Exception as e:
        logger.warning('复制查找表到输出目录失败: %s', e)
        return None


def find_lookup_file(script_dir=None):
    """
    智能查找主体查找表 Excel 文件。

    查找策略（按优先级）：
    1. 如果传入了 script_dir 参数，只在该目录查找（向后兼容，测试专用）
    2. 否则优先在输出目录（可写目录）查找
    3. 如果在输出目录没找到，在程序目录查找
    4. 如果在程序目录找到但输出目录没找到，自动复制到输出目录
    5. 在每个目录内：
       - 优先按文件名精确匹配 "主体查找表.xlsx" 或 "主体查找表.xls"
       - 若未精确匹配到，回退到查找目录下唯一的 Excel 文件

    Args:
        script_dir: 可选，指定的脚本目录（兼容旧接口）

    Returns:
        Optional[str]: 查找表文件路径，未找到返回 None
    """
    logger = get_logger()
    output_dir = get_output_dir()
    program_dir = get_program_dir()

    # ── 策略 0：如果传入了 script_dir，只在该目录查找（向后兼容） ──
    if script_dir is not None:
        custom_lookup = _find_lookup_in_dir(script_dir)
        if custom_lookup:
            logger.info('在指定目录找到主体查找表: %s', custom_lookup)
            return custom_lookup
        logger.warning('在指定目录未找到主体查找表: %s', script_dir)
        return None

    # ── 策略 1：优先在输出目录查找 ──
    output_lookup = _find_lookup_in_dir(output_dir)
    if output_lookup:
        logger.info('在输出目录找到主体查找表: %s', output_lookup)
        return output_lookup

    # ── 策略 2：在程序目录查找，并尝试复制到输出目录 ──
    program_lookup = _find_lookup_in_dir(program_dir)
    if program_lookup:
        logger.info('在程序目录找到主体查找表: %s', program_lookup)
        copied_path = _copy_lookup_to_output(program_lookup, output_dir)
        if copied_path:
            return copied_path
        return program_lookup

    logger.warning('未找到主体查找表文件')
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
# 银行处理器（基于配置动态生成）
# ──────────────────────────────────────────────

def _build_bank_processors():
    """从配置文件动态构建银行处理器注册表"""
    config = get_bank_config()
    bank_names = config.get_all_bank_names()
    processors = {}
    for bank_name in bank_names:
        processors[bank_name] = _create_bank_processor(bank_name)
    return processors


# 银行处理器注册表（从配置动态生成）
BANK_PROCESSORS = _build_bank_processors()


def reload_bank_processors():
    """重新加载银行配置并重建处理器注册表（热更新）"""
    global BANK_PROCESSORS, BANK_PREFIXES, _bank_config_singleton
    _bank_config_singleton = None
    config = get_bank_config()
    config.load_config()
    BANK_PREFIXES = _get_bank_prefixes()
    BANK_PROCESSORS = _build_bank_processors()
    logger = get_logger()
    logger.info('已重新加载 %d 个银行处理器', len(BANK_PROCESSORS))


# 向后兼容：保留旧函数名作为别名，确保现有代码和测试正常运行
# 直接从 BANK_PROCESSORS 中获取，确保是同一个函数实例
process_beijing_bank = BANK_PROCESSORS.get('北京银行', _create_bank_processor('北京银行'))
process_east_asia_bank = BANK_PROCESSORS.get('东亚银行', _create_bank_processor('东亚银行'))


# ──────────────────────────────────────────────
# 主流程
# ──────────────────────────────────────────────

@dataclass
class ProcessingResult:
    all_rows: List[dict] = field(default_factory=list)
    processed_files: List[str] = field(default_factory=list)
    unprocessed_files: List[str] = field(default_factory=list)
    error_files: List[Tuple[str, str]] = field(default_factory=list)
    file_process_details: List[FileProcessDetail] = field(default_factory=list)
    output_path: Optional[str] = None
    output_paths: Dict[str, Any] = field(default_factory=dict)
    subject_summary_path: Optional[str] = None
    balance_check_path: Optional[str] = None
    duplicate_check_path: Optional[str] = None
    lookup_missing: bool = False
    folder_empty: bool = False
    incremental_mode: bool = False
    existing_record_count: int = 0
    new_record_count: int = 0
    duplicate_record_count: int = 0
    db_inserted_count: int = 0
    db_duplicate_count: int = 0
    verification_report_path: Optional[str] = None
    verification_report_md_path: Optional[str] = None


@dataclass
class SkippedRowDetail:
    """跳过行明细"""
    file_path: str = ''
    file_name: str = ''
    row_number: int = 0
    reason: str = ''
    raw_content: str = ''


@dataclass
class FileProcessDetail:
    """单文件处理详情"""
    file_path: str = ''
    file_name: str = ''
    bank_name: str = ''
    bank_account: str = ''
    subject: str = ''
    total_rows_in_excel: int = 0
    extracted_records: int = 0
    skipped_rows: int = 0
    skipped_details: List[SkippedRowDetail] = field(default_factory=list)
    process_status: str = ''
    error_message: str = ''


@dataclass
class UnmatchedAccount:
    """主体未匹配账号"""
    bank_account: str = ''
    bank_name: str = ''
    record_count: int = 0
    total_income: float = 0.0
    total_expense: float = 0.0
    file_sources: List[str] = field(default_factory=list)


@dataclass
class VerificationReportData:
    """检验报告完整数据"""
    source_info: Dict[str, Any] = field(default_factory=dict)
    file_details: List[FileProcessDetail] = field(default_factory=list)
    skipped_rows: List[SkippedRowDetail] = field(default_factory=list)
    unmatched_accounts: List[UnmatchedAccount] = field(default_factory=list)
    amount_summary: Dict[str, Any] = field(default_factory=dict)
    by_subject_summary: List[Dict[str, Any]] = field(default_factory=list)
    by_bank_summary: List[Dict[str, Any]] = field(default_factory=list)


# ──────────────────────────────────────────────
# 增量合并模块
# ──────────────────────────────────────────────

SUMMARY_TABLE_FILENAME = '银行流水总表.xlsx'

OUTPUT_FORMAT_XLSX = 'xlsx'
OUTPUT_FORMAT_CSV = 'csv'
OUTPUT_FORMAT_SPLIT_BY_BANK = 'split_by_bank'

OUTPUT_FORMATS = {
    OUTPUT_FORMAT_XLSX: 'Excel 总表 (.xlsx)',
    OUTPUT_FORMAT_CSV: 'CSV 总表 (.csv)',
    OUTPUT_FORMAT_SPLIT_BY_BANK: '按银行拆分多表',
}

DEFAULT_OUTPUT_FORMATS = [OUTPUT_FORMAT_XLSX]


def get_summary_table_path(script_dir=None, output_dir=None, format_type=None):
    """
    获取总表文件路径。

    路径策略：
    - 如果指定了 output_dir，使用 output_dir
    - 否则使用可写目录（get_output_dir()）

    Args:
        script_dir: 可选，兼容旧接口，实际不使用
        output_dir: 可选，指定输出目录
        format_type: 可选，输出格式类型，用于生成不同后缀

    Returns:
        str: 总表文件的绝对路径
    """
    if output_dir:
        base_dir = output_dir
    else:
        base_dir = get_output_dir()

    if format_type == OUTPUT_FORMAT_CSV:
        return os.path.join(base_dir, '银行流水总表.csv')
    return os.path.join(base_dir, SUMMARY_TABLE_FILENAME)


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


def _sanitize_filename(name):
    """清理文件名中的非法字符"""
    import re
    return re.sub(r'[\\/:*?"<>|]', '_', name).strip() or '未知'


def backup_existing_file(file_path):
    """
    如果目标文件已存在，按时间戳重命名进行备份。

    备份命名规则：原文件名_YYYYMMDD_HHMMSS.原扩展名
    例如：银行流水总表.xlsx -> 银行流水总表_20260611_143052.xlsx

    Args:
        file_path: 待检查的目标文件路径

    Returns:
        str or None: 如果执行了备份，返回备份文件路径；否则返回 None
    """
    logger = get_logger()

    if not file_path or not os.path.exists(file_path):
        return None

    try:
        dir_name = os.path.dirname(file_path)
        base_name = os.path.basename(file_path)
        name_part, ext_part = os.path.splitext(base_name)

        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        backup_name = f'{name_part}_{timestamp}{ext_part}'
        backup_path = os.path.join(dir_name, backup_name)

        counter = 1
        while os.path.exists(backup_path):
            backup_name = f'{name_part}_{timestamp}_{counter}{ext_part}'
            backup_path = os.path.join(dir_name, backup_name)
            counter += 1

        shutil.copy2(file_path, backup_path)
        logger.info('已备份历史文件: %s -> %s', file_path, backup_path)
        return backup_path

    except Exception as e:
        logger.warning('备份历史文件失败，将继续覆盖写入: %s, 错误: %s', file_path, e)
        return None


def export_summary_to_csv(records, output_dir=None, columns=None):
    """
    导出总表为 CSV 格式。

    Args:
        records: 记录列表
        output_dir: 输出目录
        columns: 列名列表，默认使用标准列

    Returns:
        str: 输出文件路径
    """
    logger = get_logger()

    if columns is None:
        columns = [
            '唯一id', '银行', '银行账号', '主体', '交易日期',
            '付款', '收款', '摘要', '对方户名', '余额', '交易流水号',
        ]

    if not records:
        logger.warning('无任何记录可输出')
        return None

    df = pd.DataFrame(records, columns=columns)
    output_path = get_summary_table_path(output_dir=output_dir, format_type=OUTPUT_FORMAT_CSV)

    base_dir = os.path.dirname(output_path)
    os.makedirs(base_dir, exist_ok=True)

    backup_existing_file(output_path)

    df.to_csv(output_path, index=False, encoding='utf-8-sig')

    logger.info('CSV 总表输出完成: %s（共 %d 条记录）', output_path, len(records))
    return output_path


def export_summary_by_bank(records, output_dir=None, columns=None, format_type=OUTPUT_FORMAT_XLSX):
    """
    按银行拆分为多个子表文件导出。

    Args:
        records: 记录列表
        output_dir: 输出目录
        columns: 列名列表，默认使用标准列
        format_type: 子表格式，支持 xlsx 或 csv

    Returns:
        List[str]: 输出文件路径列表
    """
    logger = get_logger()

    if columns is None:
        columns = [
            '唯一id', '银行', '银行账号', '主体', '交易日期',
            '付款', '收款', '摘要', '对方户名', '余额', '交易流水号',
        ]

    if not records:
        logger.warning('无任何记录可输出')
        return []

    if output_dir is None:
        output_dir = get_output_dir()

    bank_output_dir = os.path.join(output_dir, '按银行拆分')
    os.makedirs(bank_output_dir, exist_ok=True)

    bank_groups: Dict[str, List[Dict[str, Any]]] = {}
    for rec in records:
        bank = str(rec.get('银行') or '').strip() or '未知银行'
        if bank not in bank_groups:
            bank_groups[bank] = []
        bank_groups[bank].append(rec)

    output_paths = []
    for bank, bank_records in bank_groups.items():
        safe_bank_name = _sanitize_filename(bank)
        df = pd.DataFrame(bank_records, columns=columns)

        if format_type == OUTPUT_FORMAT_CSV:
            file_path = os.path.join(bank_output_dir, f'{safe_bank_name}_流水.csv')
            df.to_csv(file_path, index=False, encoding='utf-8-sig')
        else:
            file_path = os.path.join(bank_output_dir, f'{safe_bank_name}_流水.xlsx')
            df.to_excel(file_path, index=False, engine='openpyxl')

        output_paths.append(file_path)
        logger.info('银行子表输出完成: %s（%d 条记录）', file_path, len(bank_records))

    logger.info('按银行拆分导出完成，共生成 %d 个文件', len(output_paths))
    return output_paths


def merge_and_export_summary(existing_records, incremental_rows, script_dir=None, output_dir=None, output_formats=None):
    """
    合并历史记录与增量记录，并按指定格式输出到总表。

    Args:
        existing_records: 历史记录列表
        incremental_rows: 新增记录列表
        script_dir: 可选，脚本目录（兼容旧接口）
        output_dir: 可选，输出目录，默认使用可写目录
        output_formats: 可选，输出格式列表，如 [OUTPUT_FORMAT_XLSX, OUTPUT_FORMAT_CSV]

    Returns:
        Dict: 各格式输出文件路径映射
    """
    logger = get_logger()

    columns = [
        '唯一id', '银行', '银行账号', '主体', '交易日期',
        '付款', '收款', '摘要', '对方户名', '余额', '交易流水号',
    ]

    merged_records = existing_records + incremental_rows

    if not merged_records:
        logger.warning('无任何记录可输出')
        return {}

    if output_formats is None:
        output_formats = DEFAULT_OUTPUT_FORMATS

    df = pd.DataFrame(merged_records, columns=columns)
    base_dir = output_dir or get_output_dir()
    os.makedirs(base_dir, exist_ok=True)

    output_paths: Dict[str, Any] = {}

    if OUTPUT_FORMAT_XLSX in output_formats:
        xlsx_path = get_summary_table_path(script_dir, output_dir)
        backup_existing_file(xlsx_path)
        df.to_excel(xlsx_path, index=False, engine='openpyxl')
        output_paths[OUTPUT_FORMAT_XLSX] = xlsx_path
        logger.info('Excel 总表输出完成: %s', xlsx_path)

    if OUTPUT_FORMAT_CSV in output_formats:
        csv_path = export_summary_to_csv(merged_records, output_dir, columns)
        output_paths[OUTPUT_FORMAT_CSV] = csv_path

    if OUTPUT_FORMAT_SPLIT_BY_BANK in output_formats:
        split_paths = export_summary_by_bank(merged_records, output_dir, columns)
        output_paths[OUTPUT_FORMAT_SPLIT_BY_BANK] = split_paths

    logger.info('总表多格式输出完成: 历史 %d 条 + 新增 %d 条 = 共 %d 条，格式: %s',
                len(existing_records), len(incremental_rows), len(merged_records),
                ', '.join(output_formats))
    return output_paths


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


def cli_ask_output_format():
    """命令行模式下询问用户输出格式"""
    print('\n请选择输出格式（可多选，用逗号分隔，直接回车默认仅导出Excel）：')
    for i, (key, desc) in enumerate(OUTPUT_FORMATS.items(), 1):
        default_mark = ' (默认)' if key == OUTPUT_FORMAT_XLSX else ''
        print(f'  {i}) {desc}{default_mark}')
    choice = input('请输入选项（例如: 1 或 1,2 或 1,2,3）: ').strip()

    if not choice:
        return DEFAULT_OUTPUT_FORMATS

    selected = []
    for c in choice.replace(',', ' ').split():
        if c.isdigit():
            idx = int(c) - 1
            keys = list(OUTPUT_FORMATS.keys())
            if 0 <= idx < len(keys):
                selected.append(keys[idx])

    if not selected:
        return DEFAULT_OUTPUT_FORMATS
    return selected


def gui_ask_output_format():
    """GUI 模式下询问用户输出格式"""
    if not HAS_TKINTER or tk is None:
        return cli_ask_output_format()

    root = tk.Tk()
    root.withdraw()

    options = list(OUTPUT_FORMATS.items())
    choices = []

    top = tk.Toplevel(root)
    top.title('选择输出格式')
    top.geometry('400x300')

    tk.Label(top, text='请选择输出格式（可多选）：', font=('Arial', 12)).pack(pady=10)

    vars = []
    for i, (key, desc) in enumerate(options):
        var = tk.BooleanVar(value=(key == OUTPUT_FORMAT_XLSX))
        vars.append(var)
        tk.Checkbutton(top, text=desc, variable=var).pack(anchor='w', padx=30)

    result = {'selected': []}

    def on_ok():
        selected = []
        for i, var in enumerate(vars):
            if var.get():
                selected.append(options[i][0])
        result['selected'] = selected if selected else DEFAULT_OUTPUT_FORMATS
        top.destroy()

    tk.Button(top, text='确定', command=on_ok, width=10).pack(pady=20)
    top.grab_set()
    top.wait_window()

    root.destroy()
    return result['selected'] or DEFAULT_OUTPUT_FORMATS


if HAS_TKINTER:
    ask_incremental_mode = gui_ask_incremental_mode
    ask_output_format = gui_ask_output_format
else:
    ask_incremental_mode = cli_ask_incremental_mode
    ask_output_format = cli_ask_output_format


def run_pipeline(folder, script_dir=None, incremental=True, batch_id=None,
                 keep_strategy='keep_unprocessed', output_formats=None,
                 progress_callback=None):
    logger = get_logger()

    total_stages = 10

    def _report(stage_idx, percent, message='', current_file='',
                processed_files=0, total_files=0, processed_records=0, extra=None):
        if progress_callback is None:
            return
        try:
            info = ProgressInfo(
                stage=ProgressWindow.STAGES[stage_idx][0] if stage_idx < len(ProgressWindow.STAGES) else '',
                stage_index=stage_idx,
                total_stages=total_stages,
                percent=percent,
                message=message or (ProgressWindow.STAGES[stage_idx][1] if stage_idx < len(ProgressWindow.STAGES) else ''),
                current_file=current_file,
                processed_files=processed_files,
                total_files=total_files,
                processed_records=processed_records,
                extra=extra or {},
            )
            progress_callback(info)
        except Exception:
            pass

    _report(0, 3, message='正在查找主体查找表并初始化...')

    lookup_file = find_lookup_file(script_dir)
    lookup_missing = lookup_file is None
    if lookup_missing:
        logger.warning('未找到主体查找表，"主体"列将为空')

    existing_keys = set()
    existing_records = []
    actual_incremental = False
    duplicate_count = 0
    new_record_count = 0

    if output_formats is None:
        output_formats = DEFAULT_OUTPUT_FORMATS

    _report(0, 7, message='正在检查增量合并模式...')
    if incremental:
        summary_path = get_summary_table_path(script_dir)
        existing_keys, existing_records = load_existing_keys(summary_path)
        actual_incremental = len(existing_records) > 0
        if actual_incremental:
            logger.info('===== 增量合并模式已启用 =====')
        else:
            logger.info('无历史数据，将以全量模式运行')

    _report(0, 10, message='初始化完成，准备复制文件夹...')

    folder_name = os.path.basename(folder.rstrip('/\\'))
    parent_dir = os.path.dirname(folder.rstrip('/\\'))
    new_folder = os.path.join(parent_dir, f"{folder_name}＋检验版")

    _report(1, 12, message=f'正在复制文件夹为「{folder_name}＋检验版」...')
    if os.path.exists(new_folder):
        logger.info('＋检验版文件夹已存在，先删除: %s', new_folder)
        shutil.rmtree(new_folder)
    shutil.copytree(folder, new_folder)
    logger.info('已复制文件夹为＋检验版: %s', new_folder)

    _report(1, 18, message='正在扫描文件夹中的 Excel 文件...')
    excel_files = scan_excel_files(new_folder)
    if not excel_files:
        logger.warning('检验版文件夹中未发现任何 Excel 文件')
        _report(9, 100, message='文件夹中未发现任何 Excel 文件')
        return ProcessingResult(
            lookup_missing=lookup_missing,
            folder_empty=True,
            incremental_mode=actual_incremental,
            existing_record_count=len(existing_records),
        )

    _report(1, 20, message=f'扫描完成，共发现 {len(excel_files)} 个 Excel 文件',
            total_files=len(excel_files))

    all_rows = []
    processed_files = []
    unprocessed_files = []
    error_files = []
    file_process_details: List[FileProcessDetail] = []
    success_count = 0
    error_count = 0

    _report(3, 20, message='开始解析银行流水文件...',
            total_files=len(excel_files),
            extra={'success_count': 0, 'error_count': 0})

    for idx, filepath in enumerate(excel_files):
        _report(3, 20 + int((idx / len(excel_files)) * 40),
                message=f'正在解析文件 {idx + 1}/{len(excel_files)}: {os.path.basename(filepath)}',
                current_file=filepath,
                processed_files=idx,
                total_files=len(excel_files),
                processed_records=len(all_rows),
                extra={'success_count': success_count, 'error_count': error_count})

        bank = identify_bank(filepath)
        if bank and bank in BANK_PROCESSORS:
            try:
                processor = BANK_PROCESSORS[bank]
                rows, detail = processor(filepath, lookup_file)
                all_rows.extend(rows)
                processed_files.append(filepath)
                file_process_details.append(detail)
                success_count += 1
                logger.info('成功处理文件: %s（%d 条记录，跳过 %d 行）',
                            filepath, len(rows), detail.skipped_rows)
            except Exception as e:
                error_files.append((filepath, str(e)))
                error_count += 1
                file_process_details.append(FileProcessDetail(
                    file_path=filepath,
                    file_name=os.path.basename(filepath),
                    bank_name=bank or '',
                    process_status='失败',
                    error_message=str(e),
                ))
                logger.error('处理文件「%s」时发生错误: %s', filepath, e, exc_info=True)
        else:
            unprocessed_files.append(filepath)
            file_process_details.append(FileProcessDetail(
                file_path=filepath,
                file_name=os.path.basename(filepath),
                bank_name=bank or '未识别',
                process_status='未处理',
                error_message='无法识别银行类型' if not bank else f'银行「{bank}」无可用解析规则',
            ))

    _report(3, 60, message=f'文件解析完成：成功 {success_count} 个，失败 {error_count} 个，未处理 {len(unprocessed_files)} 个',
            processed_files=len(excel_files),
            total_files=len(excel_files),
            processed_records=len(all_rows),
            extra={'success_count': success_count, 'error_count': error_count})

    _report(4, 63, message='正在清理已处理文件...')
    delete_processed_files(excel_files, processed_files, error_files, unprocessed_files, strategy=keep_strategy)

    output_path = None
    output_paths: Dict[str, Any] = {}
    final_rows = []

    _report(4, 66, message='正在合并数据并去重...')
    if all_rows:
        if actual_incremental:
            incremental_rows, duplicate_count = filter_incremental_records(all_rows, existing_keys)
            new_record_count = len(incremental_rows)
            _report(5, 68, message='正在增量合并并导出总表...')
            output_paths = merge_and_export_summary(
                existing_records, incremental_rows, script_dir, output_formats=output_formats
            )
            final_rows = existing_records + incremental_rows
        else:
            columns = [
                '唯一id', '银行', '银行账号', '主体', '交易日期',
                '付款', '收款', '摘要', '对方户名', '余额', '交易流水号',
            ]
            merged_records = all_rows
            df = pd.DataFrame(merged_records, columns=columns)
            base_dir = script_dir or get_output_dir()
            os.makedirs(base_dir, exist_ok=True)

            _report(5, 68, message='正在导出 Excel 总表...')
            if OUTPUT_FORMAT_XLSX in output_formats:
                xlsx_path = get_summary_table_path(script_dir)
                backup_existing_file(xlsx_path)
                df.to_excel(xlsx_path, index=False, engine='openpyxl')
                output_paths[OUTPUT_FORMAT_XLSX] = xlsx_path
                logger.info('Excel 总表输出完成: %s', xlsx_path)

            _report(5, 73, message='正在导出 CSV 总表...')
            if OUTPUT_FORMAT_CSV in output_formats:
                csv_path = export_summary_to_csv(merged_records, base_dir, columns)
                output_paths[OUTPUT_FORMAT_CSV] = csv_path

            _report(5, 76, message='正在按银行拆分子表...')
            if OUTPUT_FORMAT_SPLIT_BY_BANK in output_formats:
                split_paths = export_summary_by_bank(merged_records, base_dir, columns)
                output_paths[OUTPUT_FORMAT_SPLIT_BY_BANK] = split_paths

            logger.info('总表多格式输出完成: 共 %d 条记录，格式: %s',
                        len(merged_records), ', '.join(output_formats))
            final_rows = all_rows
            new_record_count = len(all_rows)
    else:
        logger.warning('未提取到任何银行流水记录')
        if existing_records:
            _report(5, 70, message='无新数据，仅导出历史总表...')
            output_paths = merge_and_export_summary(
                existing_records, [], script_dir, output_formats=output_formats
            )
            final_rows = existing_records

    _report(5, 78, message='总表导出完成', processed_records=len(final_rows))

    if output_paths.get(OUTPUT_FORMAT_XLSX):
        output_path = output_paths[OUTPUT_FORMAT_XLSX]
    elif output_paths.get(OUTPUT_FORMAT_CSV):
        output_path = output_paths[OUTPUT_FORMAT_CSV]

    _report(6, 80, message='正在应用对方户名黑白名单...')
    if final_rows:
        final_rows, _cp_tag_summary = apply_counterparty_rules(final_rows, script_dir)
        if _cp_tag_summary.get('tagged_count', 0) > 0:
            logger.info('对方户名黑白名单打标: 总记录 %d, 命中 %d (黑名单 %d, 白名单 %d)',
                        _cp_tag_summary.get('total_records', 0),
                        _cp_tag_summary.get('tagged_count', 0),
                        _cp_tag_summary.get('blacklist_hits', 0),
                        _cp_tag_summary.get('whitelist_hits', 0))
            _cp_columns = [
                '唯一id', '银行', '银行账号', '主体', '交易日期',
                '付款', '收款', '摘要', '对方户名', '余额', '交易流水号',
                '黑白名单标签', '命中规则名称', '命中关键词',
            ]
            _cp_df = pd.DataFrame(final_rows, columns=_cp_columns)
            if output_paths.get(OUTPUT_FORMAT_XLSX):
                _cp_df.to_excel(output_paths[OUTPUT_FORMAT_XLSX], index=False, engine='openpyxl')
                logger.info('已将黑白名单打标结果回写到Excel总表: %s', output_paths[OUTPUT_FORMAT_XLSX])
            if output_paths.get(OUTPUT_FORMAT_CSV):
                _cp_df.to_csv(output_paths[OUTPUT_FORMAT_CSV], index=False, encoding='utf-8-sig')
                logger.info('已将黑白名单打标结果回写到CSV总表: %s', output_paths[OUTPUT_FORMAT_CSV])
            if output_paths.get(OUTPUT_FORMAT_SPLIT_BY_BANK):
                bank_groups: Dict[str, List[Dict[str, Any]]] = {}
                for rec in final_rows:
                    bank = str(rec.get('银行') or '').strip() or '未知银行'
                    if bank not in bank_groups:
                        bank_groups[bank] = []
                    bank_groups[bank].append(rec)
                for bank, bank_records in bank_groups.items():
                    safe_bank_name = _sanitize_filename(bank)
                    for sp in output_paths[OUTPUT_FORMAT_SPLIT_BY_BANK]:
                        if safe_bank_name in os.path.basename(sp):
                            pd.DataFrame(bank_records, columns=_cp_columns).to_excel(
                                sp, index=False, engine='openpyxl')
                            logger.info('已将黑白名单打标结果回写到银行子表: %s', sp)
                            break

    _report(6, 82,
            message=f'黑白名单处理完成：命中 {_cp_tag_summary.get("tagged_count", 0) if final_rows else 0} 条记录',
            processed_records=len(final_rows))

    db_inserted = 0
    db_duplicates = 0
    _report(7, 84, message='正在写入数据库...')
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

    _report(7, 87,
            message=f'数据库写入完成：新增 {db_inserted} 条，跳过重复 {db_duplicates} 条')

    subject_summary_path = None
    balance_check_path = None
    duplicate_check_path = None
    verification_report_path = None
    verification_report_md_path = None

    _report(8, 88, message='正在生成各类汇总与检验报告...')
    report_count = 0
    total_reports = 4

    if final_rows:
        try:
            output_dir = script_dir
            if output_path:
                output_dir = os.path.dirname(output_path) or script_dir
            source_info = {
                '数据来源': '主流程自动生成',
                '总表文件': os.path.basename(output_path) if output_path else '内存数据',
                '记录数': len(final_rows),
                '运行模式': '增量合并' if actual_incremental else '全量覆盖',
                '生成时间': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            }
            _report(8, 89, message='[1/4] 正在生成主体维度汇总分析...')
            subject_summary_path = generate_subject_summary_from_records(
                final_rows, output_dir, source_info
            )
            if subject_summary_path:
                report_count += 1
                logger.info('主体维度汇总分析已自动生成: %s', subject_summary_path)
        except Exception as e:
            logger.error('自动生成主体汇总分析失败: %s', e, exc_info=True)
            subject_summary_path = None

        try:
            output_dir = script_dir
            if output_path:
                output_dir = os.path.dirname(output_path) or script_dir
            source_info = {
                '数据来源': '主流程自动生成',
                '总表文件': os.path.basename(output_path) if output_path else '内存数据',
                '记录数': len(final_rows),
                '运行模式': '增量合并' if actual_incremental else '全量覆盖',
                '生成时间': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            }
            _report(8, 91, message='[2/4] 正在生成余额连续性校验报告...')
            balance_check_path = generate_balance_check_from_records(
                final_rows, output_dir, source_info
            )
            if balance_check_path:
                report_count += 1
                logger.info('余额连续性校验报告已自动生成: %s', balance_check_path)
        except Exception as e:
            logger.error('自动生成余额连续性校验报告失败: %s', e, exc_info=True)
            balance_check_path = None

    if final_rows:
        try:
            output_dir = script_dir
            if output_path:
                output_dir = os.path.dirname(output_path) or script_dir
            source_info = {
                '数据来源': '主流程自动生成',
                '总表文件': os.path.basename(output_path) if output_path else '内存数据',
                '记录数': len(final_rows),
                '运行模式': '增量合并' if actual_incremental else '全量覆盖',
                '生成时间': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            }
            _report(8, 93, message='[3/4] 正在生成重复交易检测报告...')
            duplicate_check_path = generate_duplicate_check_from_records(
                final_rows, output_dir, source_info
            )
            if duplicate_check_path:
                report_count += 1
                logger.info('重复交易检测报告已自动生成: %s', duplicate_check_path)
        except Exception as e:
            logger.error('自动生成重复交易检测报告失败: %s', e, exc_info=True)
            duplicate_check_path = None

    try:
        output_dir = script_dir
        if output_path:
            output_dir = os.path.dirname(output_path) or script_dir
        source_info = {
            '数据来源': '主流程自动生成',
            '总表文件': os.path.basename(output_path) if output_path else '内存数据',
            '输入文件夹': folder,
            '运行模式': '增量合并' if actual_incremental else '全量覆盖',
            '生成时间': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            '操作人': get_current_user(),
        }
        _report(8, 96, message='[4/4] 正在生成流水检验报告...')
        verification_report_path, verification_report_md_path = generate_verification_report_from_records(
            final_rows, file_process_details, output_dir, source_info
        )
        if verification_report_path:
            report_count += 1
            logger.info('流水检验报告(Excel)已自动生成: %s', verification_report_path)
        if verification_report_md_path:
            logger.info('流水检验报告(Markdown)已自动生成: %s', verification_report_md_path)
    except Exception as e:
        logger.error('自动生成流水检验报告失败: %s', e, exc_info=True)
        verification_report_path = None
        verification_report_md_path = None

    _report(8, 98, message=f'报告生成完成：共生成 {report_count}/{total_reports} 份报告')
    _report(9, 100, message='全部处理完成！')

    log_processing_summary(file_process_details)

    return ProcessingResult(
        all_rows=final_rows,
        processed_files=processed_files,
        unprocessed_files=unprocessed_files,
        error_files=error_files,
        file_process_details=file_process_details,
        output_path=output_path,
        output_paths=output_paths,
        subject_summary_path=subject_summary_path,
        balance_check_path=balance_check_path,
        duplicate_check_path=duplicate_check_path,
        verification_report_path=verification_report_path,
        verification_report_md_path=verification_report_md_path,
        lookup_missing=lookup_missing,
        incremental_mode=actual_incremental,
        existing_record_count=len(existing_records),
        new_record_count=new_record_count,
        duplicate_record_count=duplicate_count,
        db_inserted_count=db_inserted,
        db_duplicate_count=db_duplicates,
    )


def build_processing_summary(file_details: List[FileProcessDetail]) -> Dict[str, Any]:
    """
    构建处理统计摘要，按银行、按主体、按文件状态三个维度统计。

    返回结构：
    {
        'by_bank': [{'银行': str, '文件数': int, '成功文件': int, '失败文件': int, '未处理文件': int, '提取记录数': int, '跳过行数': int}, ...],
        'by_subject': [{'主体': str, '文件数': int, '银行数': int, '提取记录数': int, '跳过行数': int}, ...],
        'by_status': {
            'success': [FileProcessDetail, ...],
            'unprocessed': [FileProcessDetail, ...],
            'error': [FileProcessDetail, ...],
        },
        'total': {
            'files': int,
            'success': int,
            'unprocessed': int,
            'error': int,
            'records': int,
            'skipped_rows': int,
            'banks': int,
            'subjects': int,
        }
    }
    """
    if not file_details:
        return {
            'by_bank': [],
            'by_subject': [],
            'by_status': {'success': [], 'unprocessed': [], 'error': []},
            'total': {'files': 0, 'success': 0, 'unprocessed': 0, 'error': 0,
                      'records': 0, 'skipped_rows': 0, 'banks': 0, 'subjects': 0},
        }

    by_bank_map: Dict[str, Dict[str, Any]] = {}
    by_subject_map: Dict[str, Dict[str, Any]] = {}
    success_files: List[FileProcessDetail] = []
    unprocessed_files: List[FileProcessDetail] = []
    error_files: List[FileProcessDetail] = []

    for d in file_details:
        bank = d.bank_name or '未知银行'
        subject = d.subject or '未指定主体'
        status = d.process_status

        if bank not in by_bank_map:
            by_bank_map[bank] = {
                '银行': bank,
                '文件数': 0,
                '成功文件': 0,
                '失败文件': 0,
                '未处理文件': 0,
                '提取记录数': 0,
                '跳过行数': 0,
            }
        by_bank_map[bank]['文件数'] += 1
        if status == '成功':
            by_bank_map[bank]['成功文件'] += 1
            by_bank_map[bank]['提取记录数'] += d.extracted_records
            by_bank_map[bank]['跳过行数'] += d.skipped_rows
        elif status == '失败':
            by_bank_map[bank]['失败文件'] += 1
        elif status == '未处理':
            by_bank_map[bank]['未处理文件'] += 1

        if subject not in by_subject_map:
            by_subject_map[subject] = {
                '主体': subject,
                '文件数': 0,
                '银行数': 0,
                '银行集合': set(),
                '提取记录数': 0,
                '跳过行数': 0,
            }
        by_subject_map[subject]['文件数'] += 1
        by_subject_map[subject]['银行集合'].add(bank)
        if status == '成功':
            by_subject_map[subject]['提取记录数'] += d.extracted_records
            by_subject_map[subject]['跳过行数'] += d.skipped_rows

        if status == '成功':
            success_files.append(d)
        elif status == '未处理':
            unprocessed_files.append(d)
        elif status == '失败':
            error_files.append(d)

    for s in by_subject_map.values():
        s['银行数'] = len(s['银行集合'])
        del s['银行集合']

    by_bank_list = sorted(by_bank_map.values(), key=lambda x: x['提取记录数'], reverse=True)
    by_subject_list = sorted(by_subject_map.values(), key=lambda x: x['提取记录数'], reverse=True)

    total_files = len(file_details)
    total_success = len(success_files)
    total_unprocessed = len(unprocessed_files)
    total_error = len(error_files)
    total_records = sum(d.extracted_records for d in success_files)
    total_skipped = sum(d.skipped_rows for d in success_files)
    total_banks = len(by_bank_map)
    total_subjects = len(by_subject_map)

    return {
        'by_bank': by_bank_list,
        'by_subject': by_subject_list,
        'by_status': {
            'success': success_files,
            'unprocessed': unprocessed_files,
            'error': error_files,
        },
        'total': {
            'files': total_files,
            'success': total_success,
            'unprocessed': total_unprocessed,
            'error': total_error,
            'records': total_records,
            'skipped_rows': total_skipped,
            'banks': total_banks,
            'subjects': total_subjects,
        },
    }


def log_processing_summary(file_details: List[FileProcessDetail]):
    """将处理统计摘要输出到日志"""
    logger = get_logger()
    if not file_details:
        logger.info('处理统计：无文件处理详情')
        return

    summary = build_processing_summary(file_details)
    total = summary['total']
    by_bank = summary['by_bank']
    by_subject = summary['by_subject']

    logger.info('=' * 60)
    logger.info('处理统计摘要')
    logger.info('=' * 60)
    logger.info('文件总数: %d  银行数: %d  主体数: %d',
                total['files'], total['banks'], total['subjects'])
    logger.info('成功: %d  未处理: %d  失败: %d',
                total['success'], total['unprocessed'], total['error'])
    logger.info('提取记录: %d  跳过行: %d',
                total['records'], total['skipped_rows'])

    if by_bank:
        logger.info('─' * 40)
        logger.info('按银行统计：')
        for item in by_bank:
            logger.info('  %-15s 文件:%-3d 成功:%-3d 失败:%-3d 未处理:%-3d 记录:%d',
                        item['银行'], item['文件数'], item['成功文件'],
                        item['失败文件'], item['未处理文件'], item['提取记录数'])

    if by_subject:
        logger.info('─' * 40)
        logger.info('按主体统计：')
        for item in by_subject:
            logger.info('  %-20s 文件:%-3d 银行数:%-3d 记录:%d',
                        item['主体'], item['文件数'], item['银行数'], item['提取记录数'])

    logger.info('=' * 60)


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
            )
        else:
            msg = (
                f'处理完成！\n\n'
                f'运行模式：全量覆盖\n'
                f'已处理文件数：{len(result.processed_files)}\n'
                f'提取记录数：{len(result.all_rows)}\n'
            )

        output_files = []
        if result.output_paths.get(OUTPUT_FORMAT_XLSX):
            output_files.append(f'Excel 总表：{result.output_paths[OUTPUT_FORMAT_XLSX]}')
        if result.output_paths.get(OUTPUT_FORMAT_CSV):
            output_files.append(f'CSV 总表：{result.output_paths[OUTPUT_FORMAT_CSV]}')
        if result.output_paths.get(OUTPUT_FORMAT_SPLIT_BY_BANK):
            split_dir = os.path.dirname(result.output_paths[OUTPUT_FORMAT_SPLIT_BY_BANK][0])
            output_files.append(f'按银行拆分（{len(result.output_paths[OUTPUT_FORMAT_SPLIT_BY_BANK])} 个文件）：{split_dir}')
        elif result.output_path:
            output_files.append(f'总表路径：{result.output_path}')

        msg += '\n'.join(output_files)

        if HAS_DATABASE and (result.db_inserted_count > 0 or result.db_duplicate_count > 0):
            msg += (
                f'\n\n数据库持久化：\n'
                f'├─ 新增入库：{result.db_inserted_count} 条\n'
                f'└─ 重复跳过：{result.db_duplicate_count} 条'
            )

        if result.subject_summary_path:
            msg += f'\n\n主体汇总分析：{result.subject_summary_path}'

        if result.balance_check_path:
            msg += f'\n\n余额连续性校验：{result.balance_check_path}'

        if result.duplicate_check_path:
            msg += f'\n\n重复交易检测：{result.duplicate_check_path}'

        if result.verification_report_path:
            msg += f'\n\n检验报告(Excel)：{result.verification_report_path}'

        if result.verification_report_md_path:
            msg += f'\n检验报告(Markdown)：{result.verification_report_md_path}'
    else:
        if result.incremental_mode and result.existing_record_count > 0:
            msg = (
                f'本次未提取到任何新增银行流水记录。\n\n'
                f'运行模式：增量合并\n'
                f'历史记录保留：{result.existing_record_count} 条\n'
            )
            output_files = []
            if result.output_paths.get(OUTPUT_FORMAT_XLSX):
                output_files.append(f'Excel 总表：{result.output_paths[OUTPUT_FORMAT_XLSX]}')
            if result.output_paths.get(OUTPUT_FORMAT_CSV):
                output_files.append(f'CSV 总表：{result.output_paths[OUTPUT_FORMAT_CSV]}')
            if result.output_paths.get(OUTPUT_FORMAT_SPLIT_BY_BANK):
                split_dir = os.path.dirname(result.output_paths[OUTPUT_FORMAT_SPLIT_BY_BANK][0])
                output_files.append(f'按银行拆分（{len(result.output_paths[OUTPUT_FORMAT_SPLIT_BY_BANK])} 个文件）：{split_dir}')
            elif result.output_path:
                output_files.append(f'总表路径：{result.output_path}')
            msg += '\n'.join(output_files)
        else:
            msg = '未提取到任何银行流水记录。'

        if result.verification_report_path:
            msg += f'\n\n检验报告(Excel)：{result.verification_report_path}'

        if result.verification_report_md_path:
            msg += f'\n检验报告(Markdown)：{result.verification_report_md_path}'

    if result.unprocessed_files:
        names = '\n  '.join(os.path.basename(f) for f in result.unprocessed_files)
        msg += f'\n\n无法识别的文件（{len(result.unprocessed_files)} 个，已保留）：\n  {names}'
    if result.error_files:
        err_info = '\n  '.join(f'{os.path.basename(f)}: {e}' for f, e in result.error_files)
        msg += f'\n\n处理出错的文件（{len(result.error_files)} 个，已保留）：\n  {err_info}'

    if result.file_process_details:
        summary = build_processing_summary(result.file_process_details)
        total = summary['total']
        by_bank = summary['by_bank']
        by_subject = summary['by_subject']

        if by_bank:
            msg += '\n\n━━━ 按银行统计 ━━━'
            for item in by_bank:
                msg += (
                    f'\n  ● {item["银行"]}'
                    f'  文件:{item["文件数"]}'
                    f'  成功:{item["成功文件"]}'
                    f'  失败:{item["失败文件"]}'
                    f'  未处理:{item["未处理文件"]}'
                    f'  记录:{item["提取记录数"]:,}'
                )

        if by_subject:
            msg += '\n\n━━━ 按主体统计 ━━━'
            for item in by_subject:
                msg += (
                    f'\n  ● {item["主体"]}'
                    f'  文件:{item["文件数"]}'
                    f'  银行数:{item["银行数"]}'
                    f'  记录:{item["提取记录数"]:,}'
                )

        msg += (
            f'\n\n━━━ 文件维度汇总 ━━━'
            f'\n  总文件数: {total["files"]}'
            f'  银行数: {total["banks"]}'
            f'  主体数: {total["subjects"]}'
            f'\n  成功: {total["success"]}'
            f'  未处理: {total["unprocessed"]}'
            f'  失败: {total["error"]}'
            f'  提取记录: {total["records"]:,}'
            f'  跳过行: {total["skipped_rows"]:,}'
        )

    return msg


def delete_processed_files(excel_files, processed_files, error_files, unprocessed_files,
                           strategy='keep_unprocessed', archive_dir_name='已处理归档'):
    logger = get_logger()

    if strategy == 'keep_all':
        logger.info('保留策略：保留所有文件')
        return

    error_file_paths = {f for f, _ in error_files}
    processed_set = set(processed_files)

    if strategy == 'keep_unprocessed':
        logger.info('保留策略：仅保留未处理文件')
        for filepath in excel_files:
            if filepath in processed_set:
                try:
                    os.remove(filepath)
                    logger.debug('已删除文件: %s', filepath)
                except OSError as e:
                    logger.error('删除文件「%s」失败: %s', filepath, e)

    elif strategy == 'delete_all':
        logger.info('保留策略：删除所有已处理文件')
        for filepath in excel_files:
            try:
                os.remove(filepath)
                logger.debug('已删除文件: %s', filepath)
            except OSError as e:
                logger.error('删除文件「%s」失败: %s', filepath, e)

    elif strategy == 'move_to_archive':
        logger.info('保留策略：移动到已处理归档子目录')
        if not excel_files:
            return
        parent_dir = os.path.dirname(excel_files[0])
        archive_dir = os.path.join(parent_dir, archive_dir_name)
        os.makedirs(archive_dir, exist_ok=True)
        for filepath in processed_files:
            try:
                dest = os.path.join(archive_dir, os.path.basename(filepath))
                counter = 1
                base, ext = os.path.splitext(dest)
                while os.path.exists(dest):
                    dest = f"{base}_{counter}{ext}"
                    counter += 1
                shutil.move(filepath, dest)
                logger.debug('已移动文件到归档: %s -> %s', filepath, dest)
            except (OSError, shutil.Error) as e:
                logger.error('移动文件「%s」到归档失败: %s', filepath, e)

    else:
        logger.warning('未知保留策略「%s」，回退为仅保留未处理文件', strategy)
        for filepath in excel_files:
            if filepath in processed_set:
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
    """主流程：处理银行流水文件夹，输出总表（带进度条，简化版本）"""
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

    keep_strategy = ask_keep_strategy()
    if keep_strategy is None:
        logger.info('用户取消保留策略选择，返回主菜单')
        return
    logger.info('用户选择保留策略: %s', KEEP_STRATEGIES.get(keep_strategy, keep_strategy))

    output_formats = ask_output_format()
    logger.info('用户选择输出格式: %s', ', '.join(output_formats))

    progress_win = None
    pipeline_error = None
    try:
        if HAS_TKINTER and tk is not None:
            try:
                progress_win = ProgressWindow(title='银行流水处理进度')
                progress_win.show()
                logger.info('进度条窗口已创建')
            except Exception as e:
                logger.warning('进度条窗口创建失败: %s', e)
                progress_win = None
    except Exception:
        progress_win = None

    progress_cb = create_progress_callback(progress_win)

    try:
        result = run_pipeline(
            folder, script_dir,
            incremental=incremental,
            keep_strategy=keep_strategy,
            output_formats=output_formats,
            progress_callback=progress_cb,
        )

        if result.lookup_missing and progress_win is None:
            show_warning(
                '警告',
                '在程序所在目录下未找到主体查找表文件，\n"主体"列将为空。\n'
                '建议将查找表文件命名为"主体查找表.xlsx"并放在程序所在目录下。'
            )

        msg = format_result_message(result)
        if progress_win:
            progress_win.set_completed(f'处理完成！共 {len(result.all_rows):,} 条记录'
                                       if result.all_rows else '处理完成')
            try:
                for _ in range(30):
                    progress_win.wait(50)
                    if progress_win.is_cancelled() or progress_win._closed:
                        break
            except Exception:
                pass
            show_result_detail_dialog(result, progress_win.root)
        else:
            if not show_result_detail_dialog(result):
                show_info('完成' if result.all_rows else '提示', msg)

    except RuntimeError as e:
        if '用户取消了操作' in str(e):
            pipeline_error = '用户已取消处理'
            logger.info('用户取消了处理操作')
            if progress_win:
                progress_win.set_error('处理已取消')
            else:
                show_warning('已取消', '用户已取消处理操作')
        else:
            pipeline_error = str(e)
            logger.error('处理失败: %s', e, exc_info=True)
            if progress_win:
                progress_win.set_error(f'处理出错: {str(e)[:80]}')
            else:
                show_warning('错误', f'处理出错：\n{e}')
    except Exception as e:
        pipeline_error = str(e)
        logger.error('处理失败: %s', e, exc_info=True)
        if progress_win:
            progress_win.set_error(f'处理出错: {str(e)[:80]}')
        else:
            show_warning('错误', f'处理出错：\n{e}')
    finally:
        if progress_win:
            try:
                if pipeline_error:
                    try:
                        for _ in range(60):
                            progress_win.wait(50)
                            if progress_win._closed:
                                break
                    except Exception:
                        pass
                progress_win.close()
            except Exception:
                pass


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
    """
    获取审计数据库文件路径。

    路径策略：
    - 如果指定了 script_dir 且可写，使用 script_dir
    - 否则使用可写目录（get_output_dir()）

    Args:
        script_dir: 可选，指定的目录

    Returns:
        str: 审计数据库文件的绝对路径
    """
    if script_dir and is_writable(script_dir):
        base_dir = script_dir
    else:
        base_dir = get_output_dir()
    return os.path.join(base_dir, AUDIT_DB_FILENAME)


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
    """主流程：处理银行流水文件夹，输出总表（带进度条）"""
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

    keep_strategy = ask_keep_strategy()
    if keep_strategy is None:
        logger.info('用户取消保留策略选择，返回主菜单')
        return
    logger.info('用户选择保留策略: %s', KEEP_STRATEGIES.get(keep_strategy, keep_strategy))

    progress_win = None
    try:
        if HAS_TKINTER and tk is not None:
            try:
                progress_win = ProgressWindow(title='银行流水处理进度')
                progress_win.show()
                logger.info('进度条窗口已创建')
            except Exception as e:
                logger.warning('进度条窗口创建失败，将继续后台处理: %s', e)
                progress_win = None
    except Exception:
        progress_win = None

    progress_cb = create_progress_callback(progress_win)

    batch_id = None
    batch_manager = None
    result = None
    pipeline_error = None

    try:
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

            result = run_pipeline(
                folder, script_dir,
                incremental=incremental,
                batch_id=batch_id,
                keep_strategy=keep_strategy,
                progress_callback=progress_cb,
            )
            audit.record_result(result)

            if result.lookup_missing and progress_win is None:
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

            if progress_win:
                progress_win.set_completed(f'处理完成！共 {len(result.all_rows):,} 条记录'
                                           if result.all_rows else '处理完成')
                try:
                    for _ in range(30):
                        progress_win.wait(50)
                        if progress_win.is_cancelled() or progress_win._closed:
                            break
                except Exception:
                    pass
                show_result_detail_dialog(result, progress_win.root)
            else:
                if not show_result_detail_dialog(result):
                    show_info('完成' if result.all_rows else '提示', msg)

    except RuntimeError as e:
        if '用户取消了操作' in str(e):
            pipeline_error = '用户已取消处理'
            logger.info('用户取消了处理操作')
            if progress_win:
                progress_win.set_error('处理已取消')
                try:
                    for _ in range(20):
                        progress_win.wait(50)
                        if progress_win._closed:
                            break
                except Exception:
                    pass
            else:
                show_warning('已取消', '用户已取消处理操作')
        else:
            pipeline_error = str(e)
            logger.error('主流程执行失败: %s', e, exc_info=True)
            if progress_win:
                progress_win.set_error(f'处理出错: {str(e)[:80]}')
                try:
                    for _ in range(30):
                        progress_win.wait(50)
                        if progress_win._closed:
                            break
                except Exception:
                    pass
            else:
                show_warning('错误', f'处理出错：\n{e}')

    except Exception as e:
        pipeline_error = str(e)
        logger.error('主流程执行失败: %s', e, exc_info=True)
        if progress_win:
            progress_win.set_error(f'处理出错: {str(e)[:80]}')
            try:
                for _ in range(30):
                    progress_win.wait(50)
                    if progress_win._closed:
                        break
            except Exception:
                pass
        else:
            show_warning('错误', f'处理出错：\n{e}')

    finally:
        if progress_win:
            try:
                close_timer = None

                def _schedule_close():
                    nonlocal close_timer
                    try:
                        if progress_win and not progress_win._closed:
                            progress_win.close()
                    except Exception:
                        pass

                if pipeline_error:
                    try:
                        close_timer = progress_win.root.after(3000, _schedule_close)
                        for _ in range(60):
                            progress_win.wait(50)
                            if progress_win._closed:
                                break
                    except Exception:
                        pass
                progress_win.close()
            except Exception:
                pass
        logger.info('主流程结束')


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
    keep_strategy = job_config.get('keep_strategy', 'keep_unprocessed')
    output_formats = job_config.get('output_formats', DEFAULT_OUTPUT_FORMATS)

    logger.info('========== 定时任务启动 [%s] %s ==========', job_id, job_name)
    logger.info('监控目录: %s', watch_directory)
    logger.info('运行模式: %s', '增量合并' if incremental else '全量覆盖')
    logger.info('保留策略: %s', KEEP_STRATEGIES.get(keep_strategy, keep_strategy))
    logger.info('输出格式: %s', ', '.join(output_formats))

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

            result = run_pipeline(watch_directory, script_dir, incremental=incremental, keep_strategy=keep_strategy, output_formats=output_formats)
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
            run_data['output_paths'] = result.output_paths
            run_data['status'] = 'success' if not result.error_files else 'partial'

            if result.lookup_missing:
                logger.warning('未找到主体查找表，主体列为空')

            logger.info('定时任务完成: 新增文件 %d 个，处理 %d 个，错误 %d 个，提取记录 %d 条',
                        len(new_files), len(result.processed_files),
                        len(result.error_files), result.new_record_count)

            if result.output_path:
                logger.info('输出总表: %s', result.output_path)
            if result.output_paths.get(OUTPUT_FORMAT_CSV):
                logger.info('输出CSV总表: %s', result.output_paths[OUTPUT_FORMAT_CSV])
            if result.output_paths.get(OUTPUT_FORMAT_SPLIT_BY_BANK):
                logger.info('输出银行子表 %d 个，目录: %s/按银行拆分/',
                            len(result.output_paths[OUTPUT_FORMAT_SPLIT_BY_BANK]),
                            os.path.dirname(result.output_path) if result.output_path else get_output_dir())

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

    print('\n文件保留策略:')
    for i, (key, desc) in enumerate(KEEP_STRATEGIES.items(), 1):
        print(f'  {i}) {desc}')
    keep_choice = input('请选择 (默认1-仅保留未处理文件): ').strip()
    keep_keys = list(KEEP_STRATEGIES.keys())
    if keep_choice.isdigit() and 1 <= int(keep_choice) <= len(keep_keys):
        job_config['keep_strategy'] = keep_keys[int(keep_choice) - 1]
    else:
        job_config['keep_strategy'] = 'keep_unprocessed'

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

    current_keep = job.get('keep_strategy', 'keep_unprocessed')
    print(f'\n当前配置:')
    print(f'  名称: {job.get("name", "")}')
    print(f'  目录: {job.get("watch_directory", "")}')
    print(f'  类型: {job.get("schedule_type", "cron")}')
    if job.get('schedule_type') == 'interval':
        print(f'  间隔: {job.get("interval_minutes", 60)} 分钟')
    else:
        print(f'  cron: {job.get("cron_expression", "")}')
    print(f'  增量: {"启用" if job.get("incremental", True) else "禁用"}')
    print(f'  保留策略: {KEEP_STRATEGIES.get(current_keep, current_keep)}')
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

    print('\n文件保留策略:')
    for i, (key, desc) in enumerate(KEEP_STRATEGIES.items(), 1):
        marker = ' *' if key == current_keep else ''
        print(f'  {i}) {desc}{marker}')
    keep_choice = input('请选择新策略编号 (留空保持当前): ').strip()
    keep_keys = list(KEEP_STRATEGIES.keys())
    if keep_choice.isdigit() and 1 <= int(keep_choice) <= len(keep_keys):
        updates['keep_strategy'] = keep_keys[int(keep_choice) - 1]

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
    parser.add_argument('--keep-strategy', type=str, metavar='STRATEGY',
                       help='文件保留策略: keep_all(保留所有文件), keep_unprocessed(仅保留未处理/失败文件), delete_all(删除所有文件), move_to_archive(移动到已处理归档子目录)')
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
    parser.add_argument('--subject-summary', action='store_true',
                       help='进入主体维度汇总分析功能')
    parser.add_argument('--summary-total', type=str, metavar='TOTAL_FILE',
                       help='指定总表文件直接生成主体维度汇总分析')
    parser.add_argument('--summary-output', type=str, metavar='OUTPUT_DIR',
                       help='指定主体汇总分析输出目录')
    parser.add_argument('--balance-check', action='store_true',
                       help='进入余额连续性校验功能')
    parser.add_argument('--balance-total', type=str, metavar='TOTAL_FILE',
                       help='指定总表文件直接生成余额连续性校验报告')
    parser.add_argument('--balance-output', type=str, metavar='OUTPUT_DIR',
                       help='指定余额连续性校验报告输出目录')
    parser.add_argument('--balance-tolerance', type=float, metavar='TOLERANCE',
                       help='指定余额校验容差（元），默认 0.01')
    parser.add_argument('--duplicate-check', action='store_true',
                       help='进入重复交易检测功能')
    parser.add_argument('--duplicate-total', type=str, metavar='TOTAL_FILE',
                       help='指定总表文件直接生成重复交易检测报告')
    parser.add_argument('--duplicate-output', type=str, metavar='OUTPUT_DIR',
                       help='指定重复交易检测报告输出目录')
    parser.add_argument('--output-format', type=str, action='append', metavar='FORMAT',
                       help='指定输出格式，可多次指定。可选值: xlsx(默认), csv, split_by_bank。'
                            '例如: --output-format xlsx --output-format csv')
    parser.add_argument('--input-dir', type=str, metavar='DIR',
                       help='默认输入文件夹路径，用于交互式菜单的默认值，支持脚本化批处理')
    parser.add_argument('--input-file', type=str, metavar='FILE',
                       help='默认输入文件路径，用于交互式菜单的默认值，支持脚本化批处理')

    args = parser.parse_args()

    if args.input_dir:
        set_cli_default_dir(args.input_dir)
    elif args.watch_dir:
        set_cli_default_dir(args.watch_dir)

    if args.input_file:
        set_cli_default_file(args.input_file)
    elif args.export_total:
        set_cli_default_file(args.export_total)
    elif args.summary_total:
        set_cli_default_file(args.summary_total)
    elif args.balance_total:
        set_cli_default_file(args.balance_total)
    elif args.duplicate_total:
        set_cli_default_file(args.duplicate_total)

    output_formats = DEFAULT_OUTPUT_FORMATS
    if args.output_format:
        output_formats = []
        for fmt in args.output_format:
            fmt_lower = fmt.lower()
            if fmt_lower in OUTPUT_FORMATS:
                output_formats.append(fmt_lower)
            else:
                logger.warning('忽略无效的输出格式: %s，支持的格式: %s', fmt, ', '.join(OUTPUT_FORMATS.keys()))
        if not output_formats:
            output_formats = DEFAULT_OUTPUT_FORMATS
            logger.warning('没有有效的输出格式，将使用默认格式: %s', ', '.join(output_formats))

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
            keep = KEEP_STRATEGIES.get(job.get('keep_strategy', 'keep_unprocessed'), '未知')
            print(f'    保留策略: {keep}')
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
            'keep_strategy': args.keep_strategy or 'keep_unprocessed',
            'schedule_type': 'interval',
            'interval_minutes': args.interval or 60,
            'output_formats': output_formats,
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

    if args.subject_summary or args.summary_total:
        if args.summary_total:
            total_path = os.path.abspath(args.summary_total)
            if not os.path.exists(total_path):
                logger.error('总表文件不存在: %s', total_path)
                print(f'错误: 总表文件不存在: {total_path}')
                return True

            output_dir = None
            if args.summary_output:
                output_dir = os.path.abspath(args.summary_output)
                os.makedirs(output_dir, exist_ok=True)

            result_path = generate_subject_summary_from_total(total_path, output_dir)
            if result_path:
                print(f'\n✅ 主体维度汇总分析已生成！')
                print(f'   输出文件: {result_path}\n')
            else:
                print(f'\n❌ 生成失败，请检查总表文件是否有数据\n')
            return True
        else:
            run_subject_summary_flow(script_dir)
            return True

    if args.balance_check or args.balance_total:
        if args.balance_total:
            total_path = os.path.abspath(args.balance_total)
            if not os.path.exists(total_path):
                logger.error('总表文件不存在: %s', total_path)
                print(f'错误: 总表文件不存在: {total_path}')
                return True

            output_dir = None
            if args.balance_output:
                output_dir = os.path.abspath(args.balance_output)
                os.makedirs(output_dir, exist_ok=True)

            tolerance = args.balance_tolerance if args.balance_tolerance is not None else 0.01

            result_path = generate_balance_check_from_total(total_path, output_dir, tolerance)
            if result_path:
                print(f'\n✅ 余额连续性校验报告已生成！')
                print(f'   输出文件: {result_path}\n')
            else:
                print(f'\n❌ 生成失败，请检查总表文件是否有数据\n')
            return True
        else:
            run_balance_check_flow(script_dir)
            return True

    if args.duplicate_check or args.duplicate_total:
        if args.duplicate_total:
            total_path = os.path.abspath(args.duplicate_total)
            if not os.path.exists(total_path):
                logger.error('总表文件不存在: %s', total_path)
                print(f'错误: 总表文件不存在: {total_path}')
                return True

            output_dir = None
            if args.duplicate_output:
                output_dir = os.path.abspath(args.duplicate_output)
                os.makedirs(output_dir, exist_ok=True)

            result_path = generate_duplicate_check_from_total(total_path, output_dir)
            if result_path:
                print(f'\n✅ 重复交易检测报告已生成！')
                print(f'   输出文件: {result_path}\n')
            else:
                print(f'\n❌ 生成失败，请检查总表文件是否有数据\n')
            return True
        else:
            run_duplicate_check_flow(script_dir)
            return True

    return None


# ──────────────────────────────────────────────
# 对方户名黑名单/白名单匹配模块
# ──────────────────────────────────────────────

@dataclass
class CounterpartyRule:
    rule_id: str
    name: str
    rule_type: str
    keywords: List[str]
    match_mode: str = 'contains'
    category: str = ''
    severity: str = 'medium'
    enabled: bool = True
    description: Optional[str] = None
    created_at: str = ''
    updated_at: str = ''
    created_by: str = ''


class CounterpartyRuleConfig:

    def __init__(self, script_dir=None):
        self.script_dir = script_dir or get_script_dir()
        self.config_path = os.path.join(self.script_dir, 'counterparty_rules.json')
        self._rules: List[CounterpartyRule] = []
        self.load_config()

    def load_config(self):
        logger = get_logger()
        if os.path.exists(self.config_path):
            try:
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                self._rules = [CounterpartyRule(**r) for r in data.get('rules', [])]
                logger.info('对方户名规则配置已加载: %d 条规则', len(self._rules))
            except Exception as e:
                logger.error('加载对方户名规则配置失败: %s', e)
                self._rules = []
        else:
            self._rules = []
            self.save_config()

    def save_config(self):
        logger = get_logger()
        try:
            data = {
                'rules': [vars(r) for r in self._rules],
            }
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            logger.info('对方户名规则配置已保存: %s', self.config_path)
        except Exception as e:
            logger.error('保存对方户名规则配置失败: %s', e)

    def get_rules(self, rule_type=None, enabled=None) -> List[CounterpartyRule]:
        result = self._rules
        if rule_type is not None:
            result = [r for r in result if r.rule_type == rule_type]
        if enabled is not None:
            result = [r for r in result if r.enabled == enabled]
        return result

    def add_rule(self, rule: CounterpartyRule) -> str:
        if not rule.rule_id:
            rule.rule_id = f"CPR{datetime.now().strftime('%Y%m%d%H%M%S')}{uuid.uuid4().hex[:6].upper()}"
        now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        if not rule.created_at:
            rule.created_at = now
        rule.updated_at = now
        self._rules.append(rule)
        self.save_config()
        return rule.rule_id

    def update_rule(self, rule_id: str, updates: dict) -> bool:
        for i, r in enumerate(self._rules):
            if r.rule_id == rule_id:
                for k, v in updates.items():
                    if hasattr(r, k):
                        setattr(r, k, v)
                r.updated_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                self.save_config()
                return True
        return False

    def delete_rule(self, rule_id: str) -> bool:
        original_len = len(self._rules)
        self._rules = [r for r in self._rules if r.rule_id != rule_id]
        if len(self._rules) < original_len:
            self.save_config()
            return True
        return False

    def toggle_rule(self, rule_id: str, enabled: bool) -> bool:
        for r in self._rules:
            if r.rule_id == rule_id:
                r.enabled = enabled
                r.updated_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                self.save_config()
                return True
        return False


_counterparty_rule_config_instance = None


def get_counterparty_rule_config(script_dir=None) -> CounterpartyRuleConfig:
    global _counterparty_rule_config_instance
    if _counterparty_rule_config_instance is None:
        _counterparty_rule_config_instance = CounterpartyRuleConfig(script_dir)
    return _counterparty_rule_config_instance


def _match_counterparty(name: str, rule: CounterpartyRule) -> Optional[str]:
    if not name:
        return None
    import re
    name = str(name).strip()
    for kw in rule.keywords:
        kw = str(kw).strip()
        if not kw:
            continue
        if rule.match_mode == 'exact':
            if name == kw:
                return kw
        elif rule.match_mode == 'startswith':
            if name.startswith(kw):
                return kw
        elif rule.match_mode == 'endswith':
            if name.endswith(kw):
                return kw
        elif rule.match_mode == 'regex':
            try:
                if re.search(kw, name):
                    return kw
            except re.error:
                pass
        else:
            if kw in name:
                return kw
    return None


def apply_counterparty_rules(records: List[Dict], script_dir=None) -> Tuple[List[Dict], Dict[str, Any]]:
    logger = get_logger()
    config = get_counterparty_rule_config(script_dir)
    rules = config.get_rules(enabled=True)

    blacklist_hits = 0
    whitelist_hits = 0
    rule_hit_counts: Dict[str, int] = {}
    tagged_count = 0

    for rec in records:
        tags = []
        hit_names = []
        hit_kw = []
        counterparty = rec.get('对方户名', '')
        for rule in rules:
            matched_kw = _match_counterparty(counterparty, rule)
            if matched_kw:
                prefix = '黑名单' if rule.rule_type == 'blacklist' else '白名单'
                cat = f"-{rule.category}" if rule.category else ""
                label = f"{prefix}:{cat}{rule.name}"
                tags.append(label)
                hit_names.append(rule.name)
                hit_kw.append(matched_kw)
                rule_hit_counts[rule.name] = rule_hit_counts.get(rule.name, 0) + 1
                if rule.rule_type == 'blacklist':
                    blacklist_hits += 1
                else:
                    whitelist_hits += 1

        if tags:
            rec['黑白名单标签'] = ','.join(tags)
            rec['命中规则名称'] = ','.join(hit_names)
            rec['命中关键词'] = ','.join(hit_kw)
            tagged_count += 1
        else:
            rec['黑白名单标签'] = ''
            rec['命中规则名称'] = ''
            rec['命中关键词'] = ''

    summary = {
        'total_records': len(records),
        'tagged_count': tagged_count,
        'blacklist_hits': blacklist_hits,
        'whitelist_hits': whitelist_hits,
        'rule_hit_counts': rule_hit_counts,
    }
    logger.info(
        '对方户名规则匹配完成: 总记录 %d, 命中 %d, 黑名单 %d, 白名单 %d',
        summary['total_records'], summary['tagged_count'],
        summary['blacklist_hits'], summary['whitelist_hits'],
    )
    return records, summary


def export_counterparty_tags(records: List[Dict], output_path: str) -> str:
    logger = get_logger()
    tagged = [r for r in records if r.get('黑白名单标签')]
    if not tagged:
        logger.info('没有命中对方户名规则的记录，跳过导出')
        return ''

    import pandas as pd
    df = pd.DataFrame(tagged)

    cols = list(df.columns)
    priority_cols = ['黑白名单标签', '命中规则名称', '命中关键词']
    for pc in reversed(priority_cols):
        if pc in cols:
            cols.remove(pc)
            cols.insert(0, pc)
    df = df[cols]

    df.to_excel(output_path, index=False, engine='openpyxl')

    wb = openpyxl.load_workbook(output_path)
    ws = wb.active

    tag_col = None
    for col_idx in range(1, ws.max_column + 1):
        if ws.cell(row=1, column=col_idx).value == '黑白名单标签':
            tag_col = col_idx
            break

    if tag_col:
        red_fill = openpyxl.styles.PatternFill(start_color='FFCCCC', end_color='FFCCCC', fill_type='solid')
        green_fill = openpyxl.styles.PatternFill(start_color='CCFFCC', end_color='CCFFCC', fill_type='solid')
        for row_idx in range(2, ws.max_row + 1):
            cell_val = str(ws.cell(row=row_idx, column=tag_col).value or '')
            if '黑名单' in cell_val:
                for col_idx in range(1, ws.max_column + 1):
                    ws.cell(row=row_idx, column=col_idx).fill = red_fill
            elif '白名单' in cell_val:
                for col_idx in range(1, ws.max_column + 1):
                    ws.cell(row=row_idx, column=col_idx).fill = green_fill

    wb.save(output_path)
    wb.close()
    logger.info('对方户名标签导出完成: %s (%d 条记录)', output_path, len(tagged))
    return output_path


def add_counterparty_keyword_rule(name, rule_type, keywords, match_mode='contains',
                                   category='', severity='medium', description=None,
                                   script_dir=None, username=None) -> str:
    config = get_counterparty_rule_config(script_dir)
    rule = CounterpartyRule(
        rule_id='',
        name=name,
        rule_type=rule_type,
        keywords=keywords,
        match_mode=match_mode,
        category=category,
        severity=severity,
        enabled=True,
        description=description,
        created_at='',
        updated_at='',
        created_by=username or get_current_user(),
    )
    return config.add_rule(rule)


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
    elif mode == 'subject_summary':
        run_subject_summary_flow(script_dir)
    elif mode == 'balance_check':
        run_balance_check_flow(script_dir)
    elif mode == 'duplicate_check':
        run_duplicate_check_flow(script_dir)

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
    'move_to_archive': '移动到已处理归档子目录',
}


def cli_ask_keep_strategy():
    """命令行模式下询问用户文件保留策略"""
    print('\n请选择文件保留策略：')
    for i, (key, desc) in enumerate(KEEP_STRATEGIES.items(), 1):
        marker = '（推荐）' if key == 'keep_unprocessed' else ''
        print(f'  {i}) {desc}{marker}')
    choice = input('请输入选项（直接回车默认为 1 仅保留未处理文件）: ').strip()
    keep_keys = list(KEEP_STRATEGIES.keys())
    if choice.isdigit() and 1 <= int(choice) <= len(keep_keys):
        return keep_keys[int(choice) - 1]
    return 'keep_unprocessed'


def gui_ask_keep_strategy():
    """GUI 模式下询问用户文件保留策略"""
    if tk is None:
        return cli_ask_keep_strategy()

    try:
        root = tk.Tk()
        root.title('选择文件保留策略')
        root.geometry('460x400')
        root.resizable(False, False)

        result = {'strategy': None}

        def select_strategy(strategy):
            result['strategy'] = strategy
            root.destroy()

        tk.Label(root, text='请选择文件保留策略', font=('Arial', 14, 'bold')).pack(pady=15)

        strategies = [
            ('keep_unprocessed', '仅保留未处理文件（推荐）',
             '删除处理成功的文件，保留出错和无法识别的文件',
             '#4CAF50'),
            ('keep_all', '保留所有文件',
             '不删除任何原始文件',
             '#2196F3'),
            ('move_to_archive', '移动到已处理归档子目录',
             '将处理成功的文件移动到「已处理归档」子目录',
             '#FF9800'),
            ('delete_all', '删除所有已处理文件',
             '删除所有 Excel 文件（包括成功、失败和未识别的）',
             '#f44336'),
        ]

        for key, name, desc, color in strategies:
            frame = tk.Frame(root)
            frame.pack(fill='x', padx=20, pady=4)
            btn = tk.Button(
                frame,
                text=name,
                bg=color,
                fg='white',
                font=('Arial', 10, 'bold'),
                width=30,
                command=lambda k=key: select_strategy(k),
            )
            btn.pack(side='left')
            tk.Label(frame, text=desc, font=('Arial', 8), fg='#666').pack(
                side='left', padx=10
            )

        tk.Button(
            root,
            text='取消',
            width=12,
            command=lambda: select_strategy(None),
            bg='#9E9E9E',
            fg='white',
            font=('Arial', 10),
        ).pack(pady=15)

        root.mainloop()
        return result['strategy']
    except Exception:
        return cli_ask_keep_strategy()


if HAS_TKINTER:
    ask_keep_strategy = gui_ask_keep_strategy
else:
    ask_keep_strategy = cli_ask_keep_strategy


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
    output_dir = preset.get('output_dir', '') or None

    result = run_pipeline_with_options(
        folder=folder,
        script_dir=script_dir,
        incremental=incremental,
        enabled_banks=enabled_banks,
        keep_strategy=keep_strategy,
        start_date=start_date,
        end_date=end_date,
        output_dir=output_dir,
    )

    return result


def run_pipeline_with_options(folder, script_dir, incremental=True,
                              enabled_banks=None, keep_strategy='keep_unprocessed',
                              start_date='', end_date='', batch_id=None, output_dir=None,
                              output_formats=None):
    logger = get_logger()

    if enabled_banks is None:
        enabled_banks = BANK_PREFIXES

    if output_formats is None:
        output_formats = DEFAULT_OUTPUT_FORMATS

    lookup_file = find_lookup_file(script_dir)
    lookup_missing = lookup_file is None
    if lookup_missing:
        logger.warning('未找到主体查找表，"主体"列将为空')

    existing_keys = set()
    existing_records = []
    actual_incremental = False
    duplicate_count = 0
    new_record_count = 0

    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
        logger.info('使用自定义输出目录: %s', output_dir)

    if incremental:
        summary_path = get_summary_table_path(script_dir, output_dir)
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
    file_process_details: List[FileProcessDetail] = []

    for filepath in excel_files:
        bank = identify_bank(filepath)
        if bank and bank in BANK_PROCESSORS and bank in enabled_banks:
            try:
                processor = BANK_PROCESSORS[bank]
                rows, detail = processor(filepath, lookup_file)

                if start_dt or end_dt:
                    original_len = len(rows)
                    rows = [r for r in rows if _is_date_in_range(r.get('交易日期'))]
                    filtered_out_count += (original_len - len(rows))

                all_rows.extend(rows)
                processed_files.append(filepath)
                file_process_details.append(detail)
                logger.info('成功处理文件: %s（%d 条记录，跳过 %d 行）',
                            filepath, len(rows), detail.skipped_rows)
            except Exception as e:
                error_files.append((filepath, str(e)))
                file_process_details.append(FileProcessDetail(
                    file_path=filepath,
                    file_name=os.path.basename(filepath),
                    bank_name=bank or '',
                    process_status='失败',
                    error_message=str(e),
                ))
                logger.error('处理文件「%s」时发生错误: %s', filepath, e, exc_info=True)
        else:
            unprocessed_files.append(filepath)
            reason = ''
            if bank and bank not in enabled_banks:
                reason = f'银行「{bank}」不在启用列表中'
                logger.info('文件「%s」所属银行「%s」不在启用列表中，跳过', filepath, bank)
            elif not bank:
                reason = '无法识别银行类型'
            else:
                reason = f'银行「{bank}」无可用解析规则'
            file_process_details.append(FileProcessDetail(
                file_path=filepath,
                file_name=os.path.basename(filepath),
                bank_name=bank or '未识别',
                process_status='未处理',
                error_message=reason,
            ))

    if filtered_out_count > 0:
        logger.info('日期过滤共排除 %d 条记录', filtered_out_count)

    delete_processed_files(excel_files, processed_files, error_files, unprocessed_files, strategy=keep_strategy)

    output_path = None
    output_paths: Dict[str, Any] = {}
    final_rows = []

    if all_rows:
        if actual_incremental:
            incremental_rows, duplicate_count = filter_incremental_records(all_rows, existing_keys)
            new_record_count = len(incremental_rows)
            output_paths = merge_and_export_summary(
                existing_records, incremental_rows, script_dir, output_dir, output_formats=output_formats
            )
            final_rows = existing_records + incremental_rows
        else:
            columns = [
                '唯一id', '银行', '银行账号', '主体', '交易日期',
                '付款', '收款', '摘要', '对方户名', '余额', '交易流水号',
            ]
            merged_records = all_rows
            df = pd.DataFrame(merged_records, columns=columns)
            base_dir = output_dir or script_dir or get_output_dir()
            os.makedirs(base_dir, exist_ok=True)

            if OUTPUT_FORMAT_XLSX in output_formats:
                xlsx_path = get_summary_table_path(script_dir, output_dir)
                df.to_excel(xlsx_path, index=False, engine='openpyxl')
                output_paths[OUTPUT_FORMAT_XLSX] = xlsx_path
                logger.info('Excel 总表输出完成: %s', xlsx_path)

            if OUTPUT_FORMAT_CSV in output_formats:
                csv_path = export_summary_to_csv(merged_records, base_dir, columns)
                output_paths[OUTPUT_FORMAT_CSV] = csv_path

            if OUTPUT_FORMAT_SPLIT_BY_BANK in output_formats:
                split_paths = export_summary_by_bank(merged_records, base_dir, columns)
                output_paths[OUTPUT_FORMAT_SPLIT_BY_BANK] = split_paths

            logger.info('总表多格式输出完成: 共 %d 条记录，格式: %s',
                        len(merged_records), ', '.join(output_formats))
            final_rows = all_rows
            new_record_count = len(all_rows)
    else:
        logger.warning('未提取到任何银行流水记录')
        if existing_records:
            output_paths = merge_and_export_summary(
                existing_records, [], script_dir, output_dir, output_formats=output_formats
            )
            final_rows = existing_records

    if output_paths.get(OUTPUT_FORMAT_XLSX):
        output_path = output_paths[OUTPUT_FORMAT_XLSX]
    elif output_paths.get(OUTPUT_FORMAT_CSV):
        output_path = output_paths[OUTPUT_FORMAT_CSV]

    if final_rows:
        final_rows, _cp_tag_summary = apply_counterparty_rules(final_rows, script_dir)
        if _cp_tag_summary.get('tagged_count', 0) > 0:
            logger.info('对方户名黑白名单打标: 总记录 %d, 命中 %d (黑名单 %d, 白名单 %d)',
                        _cp_tag_summary.get('total_records', 0),
                        _cp_tag_summary.get('tagged_count', 0),
                        _cp_tag_summary.get('blacklist_hits', 0),
                        _cp_tag_summary.get('whitelist_hits', 0))
            _cp_columns = [
                '唯一id', '银行', '银行账号', '主体', '交易日期',
                '付款', '收款', '摘要', '对方户名', '余额', '交易流水号',
                '黑白名单标签', '命中规则名称', '命中关键词',
            ]
            _cp_df = pd.DataFrame(final_rows, columns=_cp_columns)
            if output_paths.get(OUTPUT_FORMAT_XLSX):
                _cp_df.to_excel(output_paths[OUTPUT_FORMAT_XLSX], index=False, engine='openpyxl')
                logger.info('已将黑白名单打标结果回写到Excel总表: %s', output_paths[OUTPUT_FORMAT_XLSX])
            if output_paths.get(OUTPUT_FORMAT_CSV):
                _cp_df.to_csv(output_paths[OUTPUT_FORMAT_CSV], index=False, encoding='utf-8-sig')
                logger.info('已将黑白名单打标结果回写到CSV总表: %s', output_paths[OUTPUT_FORMAT_CSV])
            if output_paths.get(OUTPUT_FORMAT_SPLIT_BY_BANK):
                bank_groups: Dict[str, List[Dict[str, Any]]] = {}
                for rec in final_rows:
                    bank = str(rec.get('银行') or '').strip() or '未知银行'
                    if bank not in bank_groups:
                        bank_groups[bank] = []
                    bank_groups[bank].append(rec)
                for bank, bank_records in bank_groups.items():
                    safe_bank_name = _sanitize_filename(bank)
                    for sp in output_paths[OUTPUT_FORMAT_SPLIT_BY_BANK]:
                        if safe_bank_name in os.path.basename(sp):
                            pd.DataFrame(bank_records, columns=_cp_columns).to_excel(
                                sp, index=False, engine='openpyxl')
                            logger.info('已将黑白名单打标结果回写到银行子表: %s', sp)
                            break

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

    subject_summary_path = None
    balance_check_path = None
    if final_rows:
        try:
            output_dir_for_summary = output_dir or script_dir
            if output_path:
                output_dir_for_summary = os.path.dirname(output_path) or output_dir_for_summary
            source_info = {
                '数据来源': '主流程自动生成(预设)',
                '总表文件': os.path.basename(output_path) if output_path else '内存数据',
                '记录数': len(final_rows),
                '运行模式': '增量合并' if actual_incremental else '全量覆盖',
                '启用银行': ', '.join(enabled_banks) if enabled_banks else '全部',
                '日期范围': f'{start_date or "不限"} ~ {end_date or "不限"}',
                '生成时间': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            }
            subject_summary_path = generate_subject_summary_from_records(
                final_rows, output_dir_for_summary, source_info
            )
            if subject_summary_path:
                logger.info('主体维度汇总分析已自动生成: %s', subject_summary_path)
        except Exception as e:
            logger.error('自动生成主体汇总分析失败: %s', e, exc_info=True)
            subject_summary_path = None

        try:
            output_dir_for_check = output_dir or script_dir
            if output_path:
                output_dir_for_check = os.path.dirname(output_path) or output_dir_for_check
            source_info = {
                '数据来源': '主流程自动生成(预设)',
                '总表文件': os.path.basename(output_path) if output_path else '内存数据',
                '记录数': len(final_rows),
                '运行模式': '增量合并' if actual_incremental else '全量覆盖',
                '启用银行': ', '.join(enabled_banks) if enabled_banks else '全部',
                '日期范围': f'{start_date or "不限"} ~ {end_date or "不限"}',
                '生成时间': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            }
            balance_check_path = generate_balance_check_from_records(
                final_rows, output_dir_for_check, source_info
            )
            if balance_check_path:
                logger.info('余额连续性校验报告已自动生成: %s', balance_check_path)
        except Exception as e:
            logger.error('自动生成余额连续性校验报告失败: %s', e, exc_info=True)
            balance_check_path = None

    duplicate_check_path = None
    if final_rows:
        try:
            output_dir_for_check = output_dir or script_dir
            if output_path:
                output_dir_for_check = os.path.dirname(output_path) or output_dir_for_check
            source_info = {
                '数据来源': '主流程自动生成(预设)',
                '总表文件': os.path.basename(output_path) if output_path else '内存数据',
                '记录数': len(final_rows),
                '运行模式': '增量合并' if actual_incremental else '全量覆盖',
                '启用银行': ', '.join(enabled_banks) if enabled_banks else '全部',
                '日期范围': f'{start_date or "不限"} ~ {end_date or "不限"}',
                '生成时间': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            }
            duplicate_check_path = generate_duplicate_check_from_records(
                final_rows, output_dir_for_check, source_info
            )
            if duplicate_check_path:
                logger.info('重复交易检测报告已自动生成: %s', duplicate_check_path)
        except Exception as e:
            logger.error('自动生成重复交易检测报告失败: %s', e, exc_info=True)
            duplicate_check_path = None

    verification_report_path = None
    verification_report_md_path = None
    try:
        output_dir_for_report = output_dir or script_dir
        if output_path:
            output_dir_for_report = os.path.dirname(output_path) or output_dir_for_report
        source_info = {
            '数据来源': '主流程自动生成(预设)',
            '总表文件': os.path.basename(output_path) if output_path else '内存数据',
            '输入文件夹': folder,
            '运行模式': '增量合并' if actual_incremental else '全量覆盖',
            '启用银行': ', '.join(enabled_banks) if enabled_banks else '全部',
            '日期范围': f'{start_date or "不限"} ~ {end_date or "不限"}',
            '生成时间': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            '操作人': get_current_user(),
        }
        verification_report_path, verification_report_md_path = generate_verification_report_from_records(
            final_rows, file_process_details, output_dir_for_report, source_info
        )
        if verification_report_path:
            logger.info('流水检验报告(Excel)已自动生成: %s', verification_report_path)
        if verification_report_md_path:
            logger.info('流水检验报告(Markdown)已自动生成: %s', verification_report_md_path)
    except Exception as e:
        logger.error('自动生成流水检验报告失败: %s', e, exc_info=True)
        verification_report_path = None
        verification_report_md_path = None

    log_processing_summary(file_process_details)

    return ProcessingResult(
        all_rows=final_rows,
        processed_files=processed_files,
        unprocessed_files=unprocessed_files,
        error_files=error_files,
        file_process_details=file_process_details,
        output_path=output_path,
        output_paths=output_paths,
        subject_summary_path=subject_summary_path,
        balance_check_path=balance_check_path,
        duplicate_check_path=duplicate_check_path,
        verification_report_path=verification_report_path,
        verification_report_md_path=verification_report_md_path,
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
            if not show_result_detail_dialog(result):
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


# ──────────────────────────────────────────────
# 主体维度汇总分析模块
# ──────────────────────────────────────────────

SUBJECT_SUMMARY_FILENAME = '主体维度汇总分析.xlsx'


@dataclass
class SubjectDimensionSummary:
    """主体维度汇总数据"""
    subject: str = ''
    bank: str = ''
    year_month: str = ''
    total_income: float = 0.0
    total_expense: float = 0.0
    net_amount: float = 0.0
    transaction_count: int = 0
    income_count: int = 0
    expense_count: int = 0


@dataclass
class SubjectSummaryResult:
    """汇总分析结果容器"""
    by_subject: List[Dict[str, Any]] = field(default_factory=list)
    by_subject_bank: List[Dict[str, Any]] = field(default_factory=list)
    by_subject_month: List[Dict[str, Any]] = field(default_factory=list)
    by_subject_bank_month: List[Dict[str, Any]] = field(default_factory=list)
    by_bank: List[Dict[str, Any]] = field(default_factory=list)
    by_month: List[Dict[str, Any]] = field(default_factory=list)
    overall_summary: Dict[str, Any] = field(default_factory=dict)


def _extract_year_month(trade_date) -> str:
    """从交易日期提取年月，格式 YYYY-MM"""
    if trade_date is None:
        return '未知'

    s = str(trade_date).strip()
    if not s:
        return '未知'

    dt = _normalize_date(trade_date)
    if dt is None:
        if len(s) >= 7:
            if s[4] in '-/.' and len(s) >= 7:
                return s[:7]
            elif s.isdigit() and len(s) >= 6:
                return f'{s[:4]}-{s[4:6]}'
        return '未知'

    try:
        return dt.strftime('%Y-%m')
    except Exception:
        if len(s) >= 7 and s[4] in '-/.' and len(s) >= 7:
            return s[:7]
        return '未知'


def summarize_transactions(records: List[Dict[str, Any]]) -> SubjectSummaryResult:
    """
    对交易记录进行多维度汇总分析。

    维度组合：
    1. 按主体
    2. 按主体 + 银行
    3. 按主体 + 月份
    4. 按主体 + 银行 + 月份
    5. 按银行
    6. 按月份

    每个维度统计：收入总额、支出总额、净额、交易笔数、收入笔数、支出笔数

    Args:
        records: 交易记录列表，每条包含 '主体', '银行', '交易日期', '付款', '收款' 等字段

    Returns:
        SubjectSummaryResult: 多维度汇总结果
    """
    logger = get_logger()

    empty_overall = {
        'total_income': 0.0, 'total_expense': 0.0, 'net_amount': 0.0,
        'transaction_count': 0, 'income_count': 0, 'expense_count': 0,
        'subject_count': 0, 'bank_count': 0, 'month_count': 0,
    }

    if not records:
        logger.warning('无交易记录可汇总')
        return SubjectSummaryResult(overall_summary=empty_overall)

    agg_3d: Dict[Tuple[str, str, str], SubjectDimensionSummary] = {}
    agg_subject: Dict[str, SubjectDimensionSummary] = {}
    agg_subject_bank: Dict[Tuple[str, str], SubjectDimensionSummary] = {}
    agg_subject_month: Dict[Tuple[str, str], SubjectDimensionSummary] = {}
    agg_bank: Dict[str, SubjectDimensionSummary] = {}
    agg_month: Dict[str, SubjectDimensionSummary] = {}

    total_income = 0.0
    total_expense = 0.0
    total_count = 0
    total_income_count = 0
    total_expense_count = 0

    def _update(entry, income, expense, is_income, is_expense):
        entry.total_income += income
        entry.total_expense += expense
        entry.transaction_count += 1
        if is_income:
            entry.income_count += 1
        if is_expense:
            entry.expense_count += 1

    for rec in records:
        subject = str(rec.get('主体') or '').strip() or '未指定主体'
        bank = str(rec.get('银行') or '').strip() or '未知银行'
        year_month = _extract_year_month(rec.get('交易日期'))

        payment = to_float(rec.get('付款'))
        receipt = to_float(rec.get('收款'))

        income = 0.0
        expense = 0.0
        is_income = False
        is_expense = False

        if receipt is not None and receipt > 0:
            income = receipt
            is_income = True
        if payment is not None and payment < 0:
            expense = abs(payment)
            is_expense = True

        if not is_income and not is_expense:
            continue

        total_income += income
        total_expense += expense
        total_count += 1
        if is_income:
            total_income_count += 1
        if is_expense:
            total_expense_count += 1

        if subject not in agg_subject:
            agg_subject[subject] = SubjectDimensionSummary(subject=subject)
        _update(agg_subject[subject], income, expense, is_income, is_expense)

        if bank not in agg_bank:
            agg_bank[bank] = SubjectDimensionSummary(bank=bank)
        _update(agg_bank[bank], income, expense, is_income, is_expense)

        if year_month not in agg_month:
            agg_month[year_month] = SubjectDimensionSummary(year_month=year_month)
        _update(agg_month[year_month], income, expense, is_income, is_expense)

        key_sb = (subject, bank)
        if key_sb not in agg_subject_bank:
            agg_subject_bank[key_sb] = SubjectDimensionSummary(subject=subject, bank=bank)
        _update(agg_subject_bank[key_sb], income, expense, is_income, is_expense)

        key_sm = (subject, year_month)
        if key_sm not in agg_subject_month:
            agg_subject_month[key_sm] = SubjectDimensionSummary(subject=subject, year_month=year_month)
        _update(agg_subject_month[key_sm], income, expense, is_income, is_expense)

        key_3d = (subject, bank, year_month)
        if key_3d not in agg_3d:
            agg_3d[key_3d] = SubjectDimensionSummary(subject=subject, bank=bank, year_month=year_month)
        _update(agg_3d[key_3d], income, expense, is_income, is_expense)

    def finalize(entries):
        for e in entries:
            e.net_amount = round(e.total_income - e.total_expense, 2)
            e.total_income = round(e.total_income, 2)
            e.total_expense = round(e.total_expense, 2)

    finalize(agg_3d.values())
    finalize(agg_subject.values())
    finalize(agg_subject_bank.values())
    finalize(agg_subject_month.values())
    finalize(agg_bank.values())
    finalize(agg_month.values())

    def to_dict_list(entries, fields):
        result = []
        for e in entries:
            d = {}
            for f in fields:
                d[f] = getattr(e, f)
            result.append(d)
        return result

    fields_all = ['subject', 'bank', 'year_month', 'total_income', 'total_expense',
                  'net_amount', 'transaction_count', 'income_count', 'expense_count']
    fields_s = ['subject', 'total_income', 'total_expense', 'net_amount',
                'transaction_count', 'income_count', 'expense_count']
    fields_sb = ['subject', 'bank', 'total_income', 'total_expense', 'net_amount',
                 'transaction_count', 'income_count', 'expense_count']
    fields_sm = ['subject', 'year_month', 'total_income', 'total_expense', 'net_amount',
                 'transaction_count', 'income_count', 'expense_count']
    fields_b = ['bank', 'total_income', 'total_expense', 'net_amount',
                'transaction_count', 'income_count', 'expense_count']
    fields_m = ['year_month', 'total_income', 'total_expense', 'net_amount',
                'transaction_count', 'income_count', 'expense_count']

    result = SubjectSummaryResult(
        by_subject=sorted(
            to_dict_list(agg_subject.values(), fields_s),
            key=lambda x: x['net_amount'], reverse=True
        ),
        by_subject_bank=sorted(
            to_dict_list(agg_subject_bank.values(), fields_sb),
            key=lambda x: (x['subject'], x['net_amount']), reverse=True
        ),
        by_subject_month=sorted(
            to_dict_list(agg_subject_month.values(), fields_sm),
            key=lambda x: (x['subject'], x['year_month'])
        ),
        by_subject_bank_month=sorted(
            to_dict_list(agg_3d.values(), fields_all),
            key=lambda x: (x['subject'], x['bank'], x['year_month'])
        ),
        by_bank=sorted(
            to_dict_list(agg_bank.values(), fields_b),
            key=lambda x: x['net_amount'], reverse=True
        ),
        by_month=sorted(
            to_dict_list(agg_month.values(), fields_m),
            key=lambda x: x['year_month']
        ),
        overall_summary={
            'total_income': round(total_income, 2),
            'total_expense': round(total_expense, 2),
            'net_amount': round(total_income - total_expense, 2),
            'transaction_count': total_count,
            'income_count': total_income_count,
            'expense_count': total_expense_count,
            'subject_count': len(agg_subject),
            'bank_count': len(agg_bank),
            'month_count': len(agg_month),
        },
    )

    logger.info(
        '主体维度汇总完成: %d 条记录, %d 个主体, %d 家银行, %d 个月份',
        total_count, len(agg_subject), len(agg_bank), len(agg_month)
    )

    return result


CN_COLUMNS_S = {
    'subject': '主体',
    'total_income': '收入总额(元)',
    'total_expense': '支出总额(元)',
    'net_amount': '净额(元)',
    'transaction_count': '交易笔数',
    'income_count': '收入笔数',
    'expense_count': '支出笔数',
}

CN_COLUMNS_SB = {
    **CN_COLUMNS_S,
    'bank': '开户银行',
}

CN_COLUMNS_SM = {
    **CN_COLUMNS_S,
    'year_month': '月份',
}

CN_COLUMNS_ALL = {
    **CN_COLUMNS_SB,
    'year_month': '月份',
}

CN_COLUMNS_B = {
    'bank': '开户银行',
    'total_income': '收入总额(元)',
    'total_expense': '支出总额(元)',
    'net_amount': '净额(元)',
    'transaction_count': '交易笔数',
    'income_count': '收入笔数',
    'expense_count': '支出笔数',
}

CN_COLUMNS_M = {
    'year_month': '月份',
    'total_income': '收入总额(元)',
    'total_expense': '支出总额(元)',
    'net_amount': '净额(元)',
    'transaction_count': '交易笔数',
    'income_count': '收入笔数',
    'expense_count': '支出笔数',
}


def _rename_columns(df, col_map):
    """重命名 DataFrame 列名为中文，并调整列顺序"""
    existing = [c for c in col_map.keys() if c in df.columns]
    df = df[existing]
    df = df.rename(columns=col_map)
    return df


def _apply_number_format(ws, amount_cols, count_cols):
    """对工作表应用数字格式"""
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            col_letter = cell.column_letter
            if col_letter in amount_cols:
                cell.number_format = '#,##0.00'
            elif col_letter in count_cols:
                cell.number_format = '#,##0'


def export_subject_summary(summary_result: SubjectSummaryResult,
                           output_path: str,
                           source_info: Optional[Dict[str, Any]] = None) -> str:
    """
    将主体维度汇总分析结果导出为多 Sheet Excel 文件。

    输出的 Sheet 包括：
    1. 汇总总览 - 整体统计信息
    2. 按主体汇总 - 各主体收支净额统计
    3. 按主体+银行汇总 - 各主体在各银行的收支统计
    4. 按主体+月份汇总 - 各主体月度收支趋势
    5. 按主体+银行+月份汇总 - 最细粒度多维分析
    6. 按银行汇总 - 各银行总体收支
    7. 按月份汇总 - 全量月度收支趋势

    Args:
        summary_result: summarize_transactions 返回的汇总结果
        output_path: 输出 Excel 文件路径
        source_info: 可选，数据源信息（如总表路径、记录数等）

    Returns:
        str: 输出文件路径
    """
    logger = get_logger()

    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            overview_data = []
            overall = summary_result.overall_summary

            overview_items = [
                ('统计项', '数值'),
                ('交易总笔数', overall.get('transaction_count', 0)),
                ('收入笔数', overall.get('income_count', 0)),
                ('支出笔数', overall.get('expense_count', 0)),
                ('收入总额(元)', overall.get('total_income', 0)),
                ('支出总额(元)', overall.get('total_expense', 0)),
                ('净额(元)', overall.get('net_amount', 0)),
                ('涉及主体数', overall.get('subject_count', 0)),
                ('涉及银行数', overall.get('bank_count', 0)),
                ('覆盖月份数', overall.get('month_count', 0)),
            ]
            if source_info:
                for k, v in source_info.items():
                    overview_items.append((k, v))

            overview_df = pd.DataFrame(overview_items[1:], columns=overview_items[0])
            overview_df.to_excel(writer, sheet_name='汇总总览', index=False)

            sheet_configs = [
                ('按主体汇总', summary_result.by_subject, CN_COLUMNS_S),
                ('按主体+银行', summary_result.by_subject_bank, CN_COLUMNS_SB),
                ('按主体+月份', summary_result.by_subject_month, CN_COLUMNS_SM),
                ('主体+银行+月份', summary_result.by_subject_bank_month, CN_COLUMNS_ALL),
                ('按银行汇总', summary_result.by_bank, CN_COLUMNS_B),
                ('按月份汇总', summary_result.by_month, CN_COLUMNS_M),
            ]

            for sheet_name, data, col_map in sheet_configs:
                if not data:
                    continue
                df = pd.DataFrame(data)
                df = _rename_columns(df, col_map)
                df.to_excel(writer, sheet_name=sheet_name, index=False)

                ws = writer.sheets[sheet_name]
                amount_cols = set()
                count_cols = set()
                for idx, col_name in enumerate(df.columns, 1):
                    col_letter = openpyxl.utils.get_column_letter(idx)
                    if '元' in str(col_name):
                        amount_cols.add(col_letter)
                    elif '笔数' in str(col_name):
                        count_cols.add(col_letter)
                    max_len = max(
                        len(str(col_name)),
                        max((len(str(v)) for v in df.iloc[:, idx - 1].astype(str)), default=0)
                    )
                    ws.column_dimensions[col_letter].width = min(max_len + 4, 40)

                _apply_number_format(ws, amount_cols, count_cols)

            ws_overview = writer.sheets['汇总总览']
            for col_idx in range(1, 3):
                col_letter = openpyxl.utils.get_column_letter(col_idx)
                ws_overview.column_dimensions[col_letter].width = 25

            for row in ws_overview.iter_rows(min_row=2):
                for cell in row:
                    if cell.column == 2:
                        val = cell.value
                        if isinstance(val, (int, float)):
                            if isinstance(val, float):
                                cell.number_format = '#,##0.00'
                            else:
                                cell.number_format = '#,##0'

        logger.info('主体维度汇总分析已导出: %s', output_path)
        return output_path

    except Exception as e:
        logger.error('导出主体维度汇总分析失败: %s', e, exc_info=True)
        raise


def generate_subject_summary_from_records(records: List[Dict[str, Any]],
                                          output_dir: Optional[str] = None,
                                          source_info: Optional[Dict[str, Any]] = None) -> Optional[str]:
    """
    从交易记录列表直接生成汇总分析 Excel 文件。

    Args:
        records: 交易记录列表
        output_dir: 输出目录，默认为当前脚本目录
        source_info: 数据源信息，会写入"汇总总览"Sheet

    Returns:
        str: 生成的文件路径，如无数据则返回 None
    """
    logger = get_logger()

    if not records:
        logger.warning('无交易记录，跳过汇总分析生成')
        return None

    if output_dir is None:
        output_dir = get_script_dir()

    os.makedirs(output_dir, exist_ok=True)

    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f'主体维度汇总分析_{timestamp}.xlsx'
    output_path = os.path.join(output_dir, filename)

    summary_result = summarize_transactions(records)

    if not summary_result.overall_summary.get('transaction_count'):
        logger.warning('汇总结果为空，跳过导出')
        return None

    return export_subject_summary(summary_result, output_path, source_info)


def generate_subject_summary_from_total(total_path: str,
                                        output_dir: Optional[str] = None) -> Optional[str]:
    """
    从银行流水总表文件生成汇总分析 Excel。

    Args:
        total_path: 银行流水总表 Excel 文件路径
        output_dir: 输出目录，默认为总表所在目录

    Returns:
        str: 生成的文件路径，失败则返回 None
    """
    logger = get_logger()

    records = load_total_table(total_path)
    if not records:
        logger.warning('总表无数据: %s', total_path)
        return None

    if output_dir is None:
        output_dir = os.path.dirname(total_path) or get_script_dir()

    source_info = {
        '数据来源文件': os.path.basename(total_path),
        '总表记录数': len(records),
        '生成时间': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
    }

    return generate_subject_summary_from_records(records, output_dir, source_info)


def run_subject_summary_flow(script_dir):
    """主体维度汇总分析 CLI 流程"""
    logger = get_logger()
    logger.info('========== 主体维度汇总分析开始 ==========')

    print('\n' + '=' * 70)
    print('主体维度汇总分析 - 按主体/银行/月份统计收支与笔数')
    print('=' * 70)
    print('\n请选择数据来源：')
    print('  1) 从银行流水总表文件（Excel）')
    print('  2) 从数据库（按条件查询后汇总）')
    print('  0) 返回主菜单')

    choice = input('\n请输入选项（默认 1）: ').strip() or '1'

    records = []
    source_info = {}

    if choice == '0':
        return
    elif choice == '1':
        total_path = ask_file('请选择【银行流水总表】文件')
        if not total_path:
            show_info('提示', '未选择总表文件，返回。')
            return
        logger.info('用户选择总表文件: %s', total_path)
        records = load_total_table(total_path)
        if not records:
            show_warning('错误', '总表文件无数据或读取失败。')
            return
        source_info = {
            '数据来源文件': os.path.basename(total_path),
            '总表记录数': len(records),
            '生成时间': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        }
    elif choice == '2':
        if not HAS_DATABASE:
            show_warning('错误', '数据库模块不可用。')
            return

        print('\n输入查询条件（直接回车表示不限制）：')
        subject = input('主体名称: ').strip() or None
        bank = input('银行名称: ').strip() or None
        start_date = input('开始日期 (YYYY-MM-DD): ').strip() or None
        end_date = input('结束日期 (YYYY-MM-DD): ').strip() or None

        try:
            qr = db_module.query_transactions(
                subject=subject, bank=bank,
                start_date=start_date, end_date=end_date,
                limit=999999, script_dir=script_dir
            )
            records = [r.to_dict() for r in qr.records]
        except Exception as e:
            show_warning('错误', f'数据库查询失败: {e}')
            logger.error('数据库查询失败: %s', e, exc_info=True)
            return

        if not records:
            show_info('提示', '查询结果为空。')
            return

        source_info = {
            '数据来源': '数据库查询',
            '查询主体': subject or '全部',
            '查询银行': bank or '全部',
            '日期范围': f'{start_date or "不限"} ~ {end_date or "不限"}',
            '记录数': len(records),
            '生成时间': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        }
    else:
        print('无效选项')
        return

    summary_result = summarize_transactions(records)
    overall = summary_result.overall_summary

    print('\n' + '=' * 70)
    print('汇总总览')
    print('=' * 70)
    print(f'  交易总笔数: {overall.get("transaction_count", 0):,}')
    print(f'  收入笔数:   {overall.get("income_count", 0):,}')
    print(f'  支出笔数:   {overall.get("expense_count", 0):,}')
    print(f'  收入总额:   {overall.get("total_income", 0):>15,.2f} 元')
    print(f'  支出总额:   {overall.get("total_expense", 0):>15,.2f} 元')
    print(f'  净　　额:   {overall.get("net_amount", 0):>15,.2f} 元')
    print(f'  涉及主体:   {overall.get("subject_count", 0)} 个')
    print(f'  涉及银行:   {overall.get("bank_count", 0)} 家')
    print(f'  覆盖月份:   {overall.get("month_count", 0)} 个月')

    if summary_result.by_subject:
        print('\n' + '-' * 70)
        print('按主体汇总（前 10）')
        print('-' * 70)
        print(f'{"主体":<25}{"收入(元)":>18}{"支出(元)":>18}{"净额(元)":>18}{"笔数":>8}')
        for row in summary_result.by_subject[:10]:
            print(
                f'{str(row["subject"])[:23]:<25}'
                f'{row["total_income"]:>18,.2f}'
                f'{row["total_expense"]:>18,.2f}'
                f'{row["net_amount"]:>18,.2f}'
                f'{row["transaction_count"]:>8,}'
            )

    output_dir = input(f'\n请输入输出目录（回车使用当前目录）: ').strip()
    if not output_dir:
        output_dir = script_dir

    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_path = os.path.join(output_dir, f'主体维度汇总分析_{timestamp}.xlsx')

    try:
        export_subject_summary(summary_result, output_path, source_info)
        msg = f'汇总分析已导出！\n\n输出文件：{output_path}'
        show_info('导出成功', msg)
        logger.info('主体维度汇总分析导出完成: %s', output_path)
    except Exception as e:
        msg = f'导出失败：{e}'
        show_warning('导出失败', msg)
        logger.error('主体维度汇总分析导出失败: %s', e, exc_info=True)

    logger.info('========== 主体维度汇总分析结束 ==========')


# ──────────────────────────────────────────────
# 余额连续性校验模块
# ──────────────────────────────────────────────

@dataclass
class BalanceBreakRecord:
    """余额断裂记录"""
    bank_account: str
    subject: str
    bank: str
    transaction_date: Optional[datetime]
    prev_balance: float
    receipt: float
    payment: float
    expected_balance: float
    actual_balance: float
    diff_amount: float
    transaction_id: str
    summary: str


@dataclass
class BalanceCheckResult:
    """余额连续性校验结果"""
    total_accounts: int = 0
    checked_accounts: int = 0
    skipped_accounts: int = 0
    break_count: int = 0
    break_records: List[BalanceBreakRecord] = field(default_factory=list)
    accounts_with_breaks: List[str] = field(default_factory=list)
    check_summary: Dict[str, Any] = field(default_factory=dict)


def _parse_transaction_date(value) -> Optional[datetime]:
    """解析交易日期，支持多种格式，无效值返回 None"""
    if value is None:
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, pd.Timestamp):
        if pd.isna(value):
            return None
        return value.to_pydatetime()
    parsed = _normalize_date(value)
    if parsed is None:
        return None
    if isinstance(parsed, pd.Timestamp) and pd.isna(parsed):
        return None
    if pd.isna(parsed):
        return None
    return parsed


def _safe_float(value) -> float:
    """安全转换为 float，空值或无效值返回 0.0"""
    result = to_float(value)
    return result if result is not None else 0.0


def check_balance_continuity(records: List[Dict[str, Any]],
                             tolerance: float = 0.01) -> BalanceCheckResult:
    """
    余额连续性校验。

    校验逻辑：
    1. 按银行账号分组
    2. 每组内按交易日期排序
    3. 逐笔核对：上期余额 + 收款 + 付款 = 当期余额
       （付款字段约定为负数，与银行解析 payment_sign=negative 一致）
    4. 对断裂或跳变的记录生成异常清单

    Args:
        records: 交易记录列表
        tolerance: 容差，默认 0.01 元

    Returns:
        BalanceCheckResult: 校验结果
    """
    logger = get_logger()
    result = BalanceCheckResult()

    if not records:
        logger.warning('无交易记录可校验')
        result.check_summary = {'status': '无数据'}
        return result

    account_groups: Dict[str, List[Dict[str, Any]]] = {}
    for record in records:
        account = str(record.get('银行账号', '')).strip()
        if not account:
            continue
        if account not in account_groups:
            account_groups[account] = []
        account_groups[account].append(record)

    result.total_accounts = len(account_groups)
    logger.info('共 %d 个账号待校验', result.total_accounts)

    for account, account_records in account_groups.items():
        logger.debug('校验账号: %s, 记录数: %d', account, len(account_records))

        parsed_records = []
        for r in account_records:
            parsed = r.copy()
            parsed['_date'] = _parse_transaction_date(r.get('交易日期'))
            parsed['_balance'] = _safe_float(r.get('余额'))
            parsed['_receipt'] = _safe_float(r.get('收款'))
            parsed['_payment'] = _safe_float(r.get('付款'))
            parsed_records.append(parsed)

        has_valid_date = any(r['_date'] is not None for r in parsed_records)
        if not has_valid_date:
            logger.warning('账号 %s 无有效交易日期，跳过', account)
            result.skipped_accounts += 1
            continue

        parsed_records.sort(key=lambda r: (r['_date'] or datetime.min, str(r.get('交易流水号', ''))))
        result.checked_accounts += 1

        has_break = False
        for i in range(1, len(parsed_records)):
            prev = parsed_records[i - 1]
            curr = parsed_records[i]

            prev_balance = prev['_balance']
            receipt = curr['_receipt']
            payment = curr['_payment']
            actual_balance = curr['_balance']

            expected_balance = prev_balance + receipt + payment
            diff = abs(actual_balance - expected_balance)

            if diff > tolerance:
                has_break = True
                result.break_count += 1

                break_record = BalanceBreakRecord(
                    bank_account=account,
                    subject=str(curr.get('主体', '')),
                    bank=str(curr.get('银行', '')),
                    transaction_date=curr['_date'],
                    prev_balance=prev_balance,
                    receipt=receipt,
                    payment=payment,
                    expected_balance=expected_balance,
                    actual_balance=actual_balance,
                    diff_amount=actual_balance - expected_balance,
                    transaction_id=str(curr.get('交易流水号', '')),
                    summary=str(curr.get('摘要', ''))
                )
                result.break_records.append(break_record)
                logger.debug(
                    '余额断裂 - 账号: %s, 日期: %s, 预期: %.2f, 实际: %.2f, 差异: %.2f',
                    account, curr['_date'], expected_balance, actual_balance, actual_balance - expected_balance
                )

        if has_break and account not in result.accounts_with_breaks:
            result.accounts_with_breaks.append(account)

    result.check_summary = {
        'total_accounts': result.total_accounts,
        'checked_accounts': result.checked_accounts,
        'skipped_accounts': result.skipped_accounts,
        'break_count': result.break_count,
        'accounts_with_breaks_count': len(result.accounts_with_breaks),
        'tolerance': tolerance,
        'check_time': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
    }

    logger.info(
        '余额连续性校验完成 - 总账号: %d, 已校验: %d, 跳过: %d, 异常笔数: %d, 异常账号: %d',
        result.total_accounts, result.checked_accounts, result.skipped_accounts,
        result.break_count, len(result.accounts_with_breaks)
    )

    return result


def export_balance_check_result(check_result: BalanceCheckResult,
                                output_path: str,
                                source_info: Optional[Dict[str, Any]] = None) -> str:
    """
    导出余额连续性校验结果为 Excel 文件。

    输出的 Sheet 包括：
    1. 校验总览 - 整体统计信息
    2. 异常明细 - 所有余额断裂的交易记录
    3. 异常账号清单 - 存在余额断裂的账号列表

    Args:
        check_result: check_balance_continuity 返回的校验结果
        output_path: 输出 Excel 文件路径
        source_info: 可选，数据源信息

    Returns:
        str: 输出文件路径
    """
    logger = get_logger()

    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            overview_data = []
            summary = check_result.check_summary

            overview_items = [
                ('校验项', '数值'),
                ('账号总数', summary.get('total_accounts', 0)),
                ('已校验账号数', summary.get('checked_accounts', 0)),
                ('跳过账号数', summary.get('skipped_accounts', 0)),
                ('余额断裂笔数', summary.get('break_count', 0)),
                ('异常账号数', summary.get('accounts_with_breaks_count', 0)),
                ('容差(元)', summary.get('tolerance', 0.01)),
                ('校验时间', summary.get('check_time', '')),
            ]
            if source_info:
                for k, v in source_info.items():
                    overview_items.append((k, v))

            overview_df = pd.DataFrame(overview_items[1:], columns=overview_items[0])
            overview_df.to_excel(writer, sheet_name='校验总览', index=False)

            if check_result.break_records:
                break_data = []
                for br in check_result.break_records:
                    break_data.append({
                        '主体': br.subject,
                        '银行': br.bank,
                        '银行账号': br.bank_account,
                        '交易日期': br.transaction_date.strftime('%Y-%m-%d') if br.transaction_date else '',
                        '上期余额(元)': br.prev_balance,
                        '本期收款(元)': br.receipt,
                        '本期付款(元)': br.payment,
                        '预期余额(元)': br.expected_balance,
                        '实际余额(元)': br.actual_balance,
                        '差异(元)': br.diff_amount,
                        '交易流水号': br.transaction_id,
                        '摘要': br.summary,
                    })

                break_df = pd.DataFrame(break_data)
                break_df = break_df[[
                    '主体', '银行', '银行账号', '交易日期', '上期余额(元)',
                    '本期收款(元)', '本期付款(元)', '预期余额(元)', '实际余额(元)',
                    '差异(元)', '交易流水号', '摘要'
                ]]
                break_df.to_excel(writer, sheet_name='异常明细', index=False)

                ws = writer.sheets['异常明细']
                amount_cols = set()
                for idx, col_name in enumerate(break_df.columns, 1):
                    col_letter = openpyxl.utils.get_column_letter(idx)
                    if '元' in str(col_name):
                        amount_cols.add(col_letter)
                    max_len = max(
                        len(str(col_name)),
                        max((len(str(v)) for v in break_df.iloc[:, idx - 1].astype(str)), default=0)
                    )
                    ws.column_dimensions[col_letter].width = min(max_len + 4, 40)

                for row in ws.iter_rows(min_row=2):
                    for cell in row:
                        col_letter = cell.column_letter
                        if col_letter in amount_cols:
                            cell.number_format = '#,##0.00'

            if check_result.accounts_with_breaks:
                account_data = []
                for account in check_result.accounts_with_breaks:
                    account_breaks = [br for br in check_result.break_records if br.bank_account == account]
                    sample = account_breaks[0] if account_breaks else None
                    max_diff = max((abs(br.diff_amount) for br in account_breaks), default=0)
                    account_data.append({
                        '序号': len(account_data) + 1,
                        '主体': sample.subject if sample else '',
                        '银行': sample.bank if sample else '',
                        '银行账号': account,
                        '异常笔数': len(account_breaks),
                        '最大差异(元)': max_diff,
                        '首笔异常日期': account_breaks[0].transaction_date.strftime('%Y-%m-%d')
                        if account_breaks and account_breaks[0].transaction_date else '',
                    })

                account_df = pd.DataFrame(account_data)
                account_df.to_excel(writer, sheet_name='异常账号清单', index=False)

                ws = writer.sheets['异常账号清单']
                amount_cols = set()
                count_cols = set()
                for idx, col_name in enumerate(account_df.columns, 1):
                    col_letter = openpyxl.utils.get_column_letter(idx)
                    if '元' in str(col_name):
                        amount_cols.add(col_letter)
                    elif '笔数' in str(col_name) or '序号' in str(col_name):
                        count_cols.add(col_letter)
                    max_len = max(
                        len(str(col_name)),
                        max((len(str(v)) for v in account_df.iloc[:, idx - 1].astype(str)), default=0)
                    )
                    ws.column_dimensions[col_letter].width = min(max_len + 4, 40)

                for row in ws.iter_rows(min_row=2):
                    for cell in row:
                        col_letter = cell.column_letter
                        if col_letter in amount_cols:
                            cell.number_format = '#,##0.00'
                        elif col_letter in count_cols:
                            cell.number_format = '#,##0'

            ws_overview = writer.sheets['校验总览']
            for col_idx in range(1, 3):
                col_letter = openpyxl.utils.get_column_letter(col_idx)
                ws_overview.column_dimensions[col_letter].width = 25

            for row in ws_overview.iter_rows(min_row=2):
                for cell in row:
                    if cell.column == 2:
                        val = cell.value
                        if isinstance(val, (int, float)):
                            if isinstance(val, float):
                                cell.number_format = '#,##0.00'
                            else:
                                cell.number_format = '#,##0'

        logger.info('余额连续性校验结果已导出: %s', output_path)
        return output_path

    except Exception as e:
        logger.error('导出余额连续性校验结果失败: %s', e, exc_info=True)
        raise


def generate_balance_check_from_records(records: List[Dict[str, Any]],
                                        output_dir: Optional[str] = None,
                                        source_info: Optional[Dict[str, Any]] = None,
                                        tolerance: float = 0.01) -> Optional[str]:
    """
    从交易记录列表直接生成余额连续性校验报告。

    Args:
        records: 交易记录列表
        output_dir: 输出目录
        source_info: 数据源信息
        tolerance: 容差

    Returns:
        str: 生成的文件路径，如无数据则返回 None
    """
    logger = get_logger()

    if not records:
        logger.warning('无交易记录，跳过余额连续性校验')
        return None

    if output_dir is None:
        output_dir = get_script_dir()

    os.makedirs(output_dir, exist_ok=True)

    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f'余额连续性校验报告_{timestamp}.xlsx'
    output_path = os.path.join(output_dir, filename)

    check_result = check_balance_continuity(records, tolerance=tolerance)

    return export_balance_check_result(check_result, output_path, source_info)


def generate_balance_check_from_total(total_path: str,
                                      output_dir: Optional[str] = None,
                                      tolerance: float = 0.01) -> Optional[str]:
    """
    从银行流水总表文件生成余额连续性校验报告。

    Args:
        total_path: 银行流水总表 Excel 文件路径
        output_dir: 输出目录
        tolerance: 容差

    Returns:
        str: 生成的文件路径，失败则返回 None
    """
    logger = get_logger()

    records = load_total_table(total_path)
    if not records:
        logger.warning('总表无数据: %s', total_path)
        return None

    if output_dir is None:
        output_dir = os.path.dirname(total_path) or get_script_dir()

    source_info = {
        '数据来源文件': os.path.basename(total_path),
        '总表记录数': len(records),
        '生成时间': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
    }

    return generate_balance_check_from_records(records, output_dir, source_info, tolerance)


def run_balance_check_flow(script_dir):
    """余额连续性校验 CLI 流程"""
    logger = get_logger()
    logger.info('========== 余额连续性校验开始 ==========')

    print('\n' + '=' * 70)
    print('余额连续性校验 - 逐笔核对余额连续性，识别断裂或跳变')
    print('=' * 70)
    print('\n请选择数据来源：')
    print('  1) 从银行流水总表文件（Excel）')
    print('  2) 从数据库（按条件查询后校验）')
    print('  0) 返回主菜单')

    choice = input('\n请输入选项（默认 1）: ').strip() or '1'

    records = []
    source_info = {}

    if choice == '0':
        return
    elif choice == '1':
        total_path = ask_file('请选择【银行流水总表】文件')
        if not total_path:
            show_info('提示', '未选择总表文件，返回。')
            return
        logger.info('用户选择总表文件: %s', total_path)
        records = load_total_table(total_path)
        if not records:
            show_warning('错误', '总表文件无数据或读取失败。')
            return
        source_info = {
            '数据来源文件': os.path.basename(total_path),
            '总表记录数': len(records),
            '生成时间': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        }
    elif choice == '2':
        if not HAS_DATABASE:
            show_warning('错误', '数据库模块不可用。')
            return

        print('\n输入查询条件（直接回车表示不限制）：')
        subject = input('主体名称: ').strip() or None
        bank = input('银行名称: ').strip() or None
        account = input('银行账号: ').strip() or None
        start_date = input('开始日期 (YYYY-MM-DD): ').strip() or None
        end_date = input('结束日期 (YYYY-MM-DD): ').strip() or None

        try:
            qr = db_module.query_transactions(
                subject=subject, bank=bank, account=account,
                start_date=start_date, end_date=end_date,
                limit=999999, script_dir=script_dir
            )
            records = [r.to_dict() for r in qr.records]
        except Exception as e:
            show_warning('错误', f'数据库查询失败: {e}')
            logger.error('数据库查询失败: %s', e, exc_info=True)
            return

        if not records:
            show_info('提示', '查询结果为空。')
            return

        source_info = {
            '数据来源': '数据库查询',
            '查询主体': subject or '全部',
            '查询银行': bank or '全部',
            '查询账号': account or '全部',
            '日期范围': f'{start_date or "不限"} ~ {end_date or "不限"}',
            '记录数': len(records),
            '生成时间': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        }
    else:
        print('无效选项')
        return

    tolerance_input = input('\n请输入容差（元，直接回车默认 0.01）: ').strip()
    tolerance = 0.01
    if tolerance_input:
        try:
            tolerance = float(tolerance_input)
        except ValueError:
            print('输入无效，使用默认容差 0.01 元')
            tolerance = 0.01

    print(f'\n开始校验，容差: {tolerance} 元...')
    check_result = check_balance_continuity(records, tolerance=tolerance)
    summary = check_result.check_summary

    print('\n' + '=' * 70)
    print('校验结果总览')
    print('=' * 70)
    print(f'  账号总数:     {summary.get("total_accounts", 0):,}')
    print(f'  已校验账号:   {summary.get("checked_accounts", 0):,}')
    print(f'  跳过账号:     {summary.get("skipped_accounts", 0):,}')
    print(f'  余额断裂笔数: {summary.get("break_count", 0):,}')
    print(f'  异常账号数:   {summary.get("accounts_with_breaks_count", 0):,}')

    if check_result.break_count > 0:
        print(f'\n  ⚠️  发现 {check_result.break_count} 笔余额异常，涉及 {len(check_result.accounts_with_breaks)} 个账号')
        for account in check_result.accounts_with_breaks[:10]:
            account_breaks = [br for br in check_result.break_records if br.bank_account == account]
            sample = account_breaks[0] if account_breaks else None
            print(f'    - {account} ({sample.subject if sample else "未知主体"}): {len(account_breaks)} 笔异常')
        if len(check_result.accounts_with_breaks) > 10:
            print(f'    ... 还有 {len(check_result.accounts_with_breaks) - 10} 个账号，详见导出文件')
    else:
        print(f'\n  ✅ 所有账号余额连续性校验通过！')

    output_dir = input('\n请输入输出目录（直接回车默认当前目录）: ').strip()
    if not output_dir:
        output_dir = script_dir

    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_path = os.path.join(output_dir, f'余额连续性校验报告_{timestamp}.xlsx')

    try:
        export_balance_check_result(check_result, output_path, source_info)
        msg = f'校验报告已导出！\n\n输出文件：{output_path}'
        show_info('导出成功', msg)
        logger.info('余额连续性校验报告导出完成: %s', output_path)
    except Exception as e:
        msg = f'导出失败：{e}'
        show_warning('导出失败', msg)
        logger.error('余额连续性校验报告导出失败: %s', e, exc_info=True)

    logger.info('========== 余额连续性校验结束 ==========')


# ──────────────────────────────────────────────
# 重复交易检测模块
# ──────────────────────────────────────────────

DUPLICATE_CHECK_FILENAME = '重复交易检测报告.xlsx'


@dataclass
class DuplicateRecord:
    """疑似重复记录"""
    group_id: int
    record_index: int
    bank_account: str
    trade_date: str
    payment: float
    receipt: float
    counterpart: str
    transaction_id: str
    bank: str
    subject: str
    summary: str
    balance: float
    unique_id: str
    import_batch: str
    match_type: str
    match_key: str


@dataclass
class DuplicateGroup:
    """重复组：包含两条及以上疑似重复的交易"""
    group_id: int
    match_type: str
    match_key: str
    records: List[DuplicateRecord] = field(default_factory=list)
    record_count: int = 0

    def __post_init__(self):
        self.record_count = len(self.records)


@dataclass
class DuplicateCheckResult:
    """重复交易检测结果"""
    total_records: int = 0
    duplicate_group_count: int = 0
    duplicate_record_count: int = 0
    groups: List[DuplicateGroup] = field(default_factory=list)
    match_type_stats: Dict[str, int] = field(default_factory=dict)
    check_summary: Dict[str, Any] = field(default_factory=dict)


def _normalize_for_key(value) -> str:
    if value is None:
        return ''
    if isinstance(value, float) and value != value:
        return ''
    if isinstance(value, float) and value == int(value):
        return str(int(value))
    s = str(value).strip()
    if not s:
        return ''
    if '.' in s:
        try:
            f = float(s)
            if f == int(f):
                return str(int(f))
        except (ValueError, TypeError, OverflowError):
            pass
    return s


def _make_transaction_id_key(row) -> str:
    bank_account = _account_key(row.get('银行账号'))
    transaction_id = _normalize_for_key(row.get('交易流水号'))
    if transaction_id:
        return f"tid::{bank_account}::{transaction_id}"
    return ''


def _make_amount_date_key(row) -> str:
    bank_account = _account_key(row.get('银行账号'))
    trade_date = _normalize_for_key(row.get('交易日期'))
    payment = to_float(row.get('付款'))
    receipt = to_float(row.get('收款'))
    p = f"{payment:.2f}" if payment is not None else ''
    r = f"{receipt:.2f}" if receipt is not None else ''
    return f"amt::{bank_account}::{trade_date}::{p}::{r}"


def _make_full_key(row) -> str:
    bank_account = _account_key(row.get('银行账号'))
    trade_date = _normalize_for_key(row.get('交易日期'))
    payment = to_float(row.get('付款'))
    receipt = to_float(row.get('收款'))
    counterpart = _normalize_for_key(row.get('对方户名'))
    p = f"{payment:.2f}" if payment is not None else ''
    r = f"{receipt:.2f}" if receipt is not None else ''
    return f"full::{bank_account}::{trade_date}::{p}::{r}::{counterpart}"


def detect_duplicates(records: List[Dict[str, Any]]) -> DuplicateCheckResult:
    """
    重复交易检测。

    检测逻辑（三级匹配，由严到宽）：
    1. 交易流水号匹配（match_type=transaction_id）：
       同一银行账号 + 同一交易流水号 → 银行重复记账或重复导入
    2. 金额日期匹配（match_type=amount_date）：
       同一银行账号 + 同一交易日期 + 同一付款金额 + 同一收款金额
       → 疑似重复（不同流水号但金额/日期/账号完全一致）
    3. 完全匹配（match_type=full）：
       同一银行账号 + 同一交易日期 + 同一付款 + 同一收款 + 同一对方户名
       → 最高疑似度重复

    同一条记录只归入最高优先级的匹配组，不会重复计数。

    Args:
        records: 交易记录列表

    Returns:
        DuplicateCheckResult: 检测结果
    """
    logger = get_logger()
    result = DuplicateCheckResult()
    result.total_records = len(records)

    if not records:
        logger.warning('无交易记录可检测重复')
        result.check_summary = {'status': '无数据'}
        return result

    tid_groups: Dict[str, List[int]] = {}
    amt_groups: Dict[str, List[int]] = {}
    full_groups: Dict[str, List[int]] = {}

    for idx, rec in enumerate(records):
        tid_key = _make_transaction_id_key(rec)
        if tid_key:
            tid_groups.setdefault(tid_key, []).append(idx)

        amt_key = _make_amount_date_key(rec)
        if amt_key:
            amt_groups.setdefault(amt_key, []).append(idx)

        full_key = _make_full_key(rec)
        if full_key:
            full_groups.setdefault(full_key, []).append(idx)

    claimed: set = set()
    group_id_counter = 0
    match_type_stats = {'transaction_id': 0, 'amount_date': 0, 'full': 0}

    for key, indices in tid_groups.items():
        if len(indices) < 2:
            continue
        group_id_counter += 1
        group = DuplicateGroup(
            group_id=group_id_counter,
            match_type='transaction_id',
            match_key=key,
        )
        for idx in indices:
            rec = records[idx]
            dup_rec = DuplicateRecord(
                group_id=group_id_counter,
                record_index=idx + 1,
                bank_account=str(rec.get('银行账号', '')),
                trade_date=str(rec.get('交易日期', '')),
                payment=_safe_float(rec.get('付款')),
                receipt=_safe_float(rec.get('收款')),
                counterpart=str(rec.get('对方户名', '')),
                transaction_id=str(rec.get('交易流水号', '')),
                bank=str(rec.get('银行', '')),
                subject=str(rec.get('主体', '')),
                summary=str(rec.get('摘要', '')),
                balance=_safe_float(rec.get('余额')),
                unique_id=str(rec.get('唯一id', '')),
                import_batch=str(rec.get('导入批次号', '')),
                match_type='transaction_id',
                match_key=key,
            )
            group.records.append(dup_rec)
            claimed.add(idx)
        group.record_count = len(group.records)
        result.groups.append(group)
        match_type_stats['transaction_id'] += 1

    for key, indices in amt_groups.items():
        if len(indices) < 2:
            continue
        unclaimed = [i for i in indices if i not in claimed]
        if len(unclaimed) < 2:
            continue
        group_id_counter += 1
        group = DuplicateGroup(
            group_id=group_id_counter,
            match_type='amount_date',
            match_key=key,
        )
        for idx in unclaimed:
            rec = records[idx]
            dup_rec = DuplicateRecord(
                group_id=group_id_counter,
                record_index=idx + 1,
                bank_account=str(rec.get('银行账号', '')),
                trade_date=str(rec.get('交易日期', '')),
                payment=_safe_float(rec.get('付款')),
                receipt=_safe_float(rec.get('收款')),
                counterpart=str(rec.get('对方户名', '')),
                transaction_id=str(rec.get('交易流水号', '')),
                bank=str(rec.get('银行', '')),
                subject=str(rec.get('主体', '')),
                summary=str(rec.get('摘要', '')),
                balance=_safe_float(rec.get('余额')),
                unique_id=str(rec.get('唯一id', '')),
                import_batch=str(rec.get('导入批次号', '')),
                match_type='amount_date',
                match_key=key,
            )
            group.records.append(dup_rec)
            claimed.add(idx)
        group.record_count = len(group.records)
        result.groups.append(group)
        match_type_stats['amount_date'] += 1

    for key, indices in full_groups.items():
        if len(indices) < 2:
            continue
        unclaimed = [i for i in indices if i not in claimed]
        if len(unclaimed) < 2:
            continue
        group_id_counter += 1
        group = DuplicateGroup(
            group_id=group_id_counter,
            match_type='full',
            match_key=key,
        )
        for idx in unclaimed:
            rec = records[idx]
            dup_rec = DuplicateRecord(
                group_id=group_id_counter,
                record_index=idx + 1,
                bank_account=str(rec.get('银行账号', '')),
                trade_date=str(rec.get('交易日期', '')),
                payment=_safe_float(rec.get('付款')),
                receipt=_safe_float(rec.get('收款')),
                counterpart=str(rec.get('对方户名', '')),
                transaction_id=str(rec.get('交易流水号', '')),
                bank=str(rec.get('银行', '')),
                subject=str(rec.get('主体', '')),
                summary=str(rec.get('摘要', '')),
                balance=_safe_float(rec.get('余额')),
                unique_id=str(rec.get('唯一id', '')),
                import_batch=str(rec.get('导入批次号', '')),
                match_type='full',
                match_key=key,
            )
            group.records.append(dup_rec)
            claimed.add(idx)
        group.record_count = len(group.records)
        result.groups.append(group)
        match_type_stats['full'] += 1

    result.duplicate_group_count = len(result.groups)
    result.duplicate_record_count = sum(g.record_count for g in result.groups)
    result.match_type_stats = match_type_stats

    tid_rec_count = sum(g.record_count for g in result.groups if g.match_type == 'transaction_id')
    amt_rec_count = sum(g.record_count for g in result.groups if g.match_type == 'amount_date')
    full_rec_count = sum(g.record_count for g in result.groups if g.match_type == 'full')

    result.check_summary = {
        'total_records': result.total_records,
        'duplicate_group_count': result.duplicate_group_count,
        'duplicate_record_count': result.duplicate_record_count,
        'duplicate_rate': round(result.duplicate_record_count / result.total_records * 100, 2) if result.total_records else 0,
        'transaction_id_groups': match_type_stats.get('transaction_id', 0),
        'transaction_id_records': tid_rec_count,
        'amount_date_groups': match_type_stats.get('amount_date', 0),
        'amount_date_records': amt_rec_count,
        'full_groups': match_type_stats.get('full', 0),
        'full_records': full_rec_count,
        'check_time': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
    }

    logger.info(
        '重复交易检测完成 - 总记录: %d, 疑似重复组: %d, 疑似重复记录: %d, 重复率: %.2f%%',
        result.total_records, result.duplicate_group_count,
        result.duplicate_record_count, result.check_summary['duplicate_rate']
    )

    return result


def export_duplicate_check_result(check_result: DuplicateCheckResult,
                                  output_path: str,
                                  source_info: Optional[Dict[str, Any]] = None) -> str:
    """
    导出重复交易检测结果为 Excel 文件。

    输出的 Sheet 包括：
    1. 检测总览 - 整体统计信息
    2. 疑似重复明细 - 所有疑似重复记录的详细列表
    3. 重复组汇总 - 按重复组汇总信息

    Args:
        check_result: detect_duplicates 返回的检测结果
        output_path: 输出 Excel 文件路径
        source_info: 可选，数据源信息

    Returns:
        str: 输出文件路径
    """
    logger = get_logger()

    MATCH_TYPE_LABELS = {
        'transaction_id': '交易流水号匹配',
        'amount_date': '金额+日期匹配',
        'full': '完全匹配(含对方户名)',
    }

    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            overview_items = [
                ('检测项', '数值'),
                ('总记录数', check_result.check_summary.get('total_records', 0)),
                ('疑似重复组数', check_result.check_summary.get('duplicate_group_count', 0)),
                ('疑似重复记录数', check_result.check_summary.get('duplicate_record_count', 0)),
                ('重复率(%)', check_result.check_summary.get('duplicate_rate', 0)),
                ('流水号匹配组数', check_result.check_summary.get('transaction_id_groups', 0)),
                ('流水号匹配记录数', check_result.check_summary.get('transaction_id_records', 0)),
                ('金额日期匹配组数', check_result.check_summary.get('amount_date_groups', 0)),
                ('金额日期匹配记录数', check_result.check_summary.get('amount_date_records', 0)),
                ('完全匹配组数', check_result.check_summary.get('full_groups', 0)),
                ('完全匹配记录数', check_result.check_summary.get('full_records', 0)),
                ('检测时间', check_result.check_summary.get('check_time', '')),
            ]
            if source_info:
                for k, v in source_info.items():
                    overview_items.append((k, v))

            overview_df = pd.DataFrame(overview_items[1:], columns=overview_items[0])
            overview_df.to_excel(writer, sheet_name='检测总览', index=False)

            if check_result.groups:
                detail_data = []
                for group in check_result.groups:
                    for rec in group.records:
                        detail_data.append({
                            '重复组ID': rec.group_id,
                            '匹配类型': MATCH_TYPE_LABELS.get(rec.match_type, rec.match_type),
                            '序号(总表)': rec.record_index,
                            '唯一ID': rec.unique_id,
                            '主体': rec.subject,
                            '银行': rec.bank,
                            '银行账号': rec.bank_account,
                            '交易日期': rec.trade_date,
                            '付款(元)': rec.payment,
                            '收款(元)': rec.receipt,
                            '对方户名': rec.counterpart,
                            '余额(元)': rec.balance,
                            '交易流水号': rec.transaction_id,
                            '摘要': rec.summary,
                            '导入批次号': rec.import_batch,
                        })

                detail_df = pd.DataFrame(detail_data)
                detail_cols = [
                    '重复组ID', '匹配类型', '序号(总表)', '唯一ID',
                    '主体', '银行', '银行账号', '交易日期',
                    '付款(元)', '收款(元)', '对方户名', '余额(元)',
                    '交易流水号', '摘要', '导入批次号',
                ]
                detail_df = detail_df[[c for c in detail_cols if c in detail_df.columns]]
                detail_df.to_excel(writer, sheet_name='疑似重复明细', index=False)

                ws = writer.sheets['疑似重复明细']
                amount_cols = set()
                for idx, col_name in enumerate(detail_df.columns, 1):
                    col_letter = openpyxl.utils.get_column_letter(idx)
                    if '元' in str(col_name):
                        amount_cols.add(col_letter)
                    max_len = max(
                        len(str(col_name)),
                        max((len(str(v)) for v in detail_df.iloc[:, idx - 1].astype(str)), default=0)
                    )
                    ws.column_dimensions[col_letter].width = min(max_len + 4, 40)

                for row in ws.iter_rows(min_row=2):
                    for cell in row:
                        if cell.column_letter in amount_cols:
                            cell.number_format = '#,##0.00'

                group_data = []
                for group in check_result.groups:
                    first_rec = group.records[0] if group.records else None
                    same_batch = all(
                        r.import_batch == group.records[0].import_batch
                        for r in group.records
                    ) if group.records else True

                    duplicate_source = '同一批次导入' if same_batch else '跨批次导入'
                    if group.match_type == 'transaction_id':
                        duplicate_source += ' / 银行重复记账可能性高'
                    elif group.match_type == 'full':
                        duplicate_source += ' / 对方户名一致，疑似重复'
                    else:
                        duplicate_source += ' / 流水号不同但金额日期一致'

                    group_data.append({
                        '重复组ID': group.group_id,
                        '匹配类型': MATCH_TYPE_LABELS.get(group.match_type, group.match_type),
                        '重复记录数': group.record_count,
                        '主体': first_rec.subject if first_rec else '',
                        '银行': first_rec.bank if first_rec else '',
                        '银行账号': first_rec.bank_account if first_rec else '',
                        '交易日期': first_rec.trade_date if first_rec else '',
                        '付款(元)': first_rec.payment if first_rec else 0,
                        '收款(元)': first_rec.receipt if first_rec else 0,
                        '对方户名': first_rec.counterpart if first_rec else '',
                        '交易流水号': first_rec.transaction_id if first_rec else '',
                        '重复来源分析': duplicate_source,
                    })

                group_df = pd.DataFrame(group_data)
                group_df.to_excel(writer, sheet_name='重复组汇总', index=False)

                ws_group = writer.sheets['重复组汇总']
                amount_cols_g = set()
                count_cols_g = set()
                for idx, col_name in enumerate(group_df.columns, 1):
                    col_letter = openpyxl.utils.get_column_letter(idx)
                    if '元' in str(col_name):
                        amount_cols_g.add(col_letter)
                    elif '数' in str(col_name) or 'ID' in str(col_name):
                        count_cols_g.add(col_letter)
                    max_len = max(
                        len(str(col_name)),
                        max((len(str(v)) for v in group_df.iloc[:, idx - 1].astype(str)), default=0)
                    )
                    ws_group.column_dimensions[col_letter].width = min(max_len + 4, 50)

                for row in ws_group.iter_rows(min_row=2):
                    for cell in row:
                        if cell.column_letter in amount_cols_g:
                            cell.number_format = '#,##0.00'
                        elif cell.column_letter in count_cols_g:
                            cell.number_format = '#,##0'

            ws_overview = writer.sheets['检测总览']
            for col_idx in range(1, 3):
                col_letter = openpyxl.utils.get_column_letter(col_idx)
                ws_overview.column_dimensions[col_letter].width = 25

            for row in ws_overview.iter_rows(min_row=2):
                for cell in row:
                    if cell.column == 2:
                        val = cell.value
                        if isinstance(val, (int, float)):
                            if isinstance(val, float):
                                cell.number_format = '#,##0.00'
                            else:
                                cell.number_format = '#,##0'

        logger.info('重复交易检测结果已导出: %s', output_path)
        return output_path

    except Exception as e:
        logger.error('导出重复交易检测结果失败: %s', e, exc_info=True)
        raise


def generate_duplicate_check_from_records(records: List[Dict[str, Any]],
                                          output_dir: Optional[str] = None,
                                          source_info: Optional[Dict[str, Any]] = None) -> Optional[str]:
    """
    从交易记录列表直接生成重复交易检测报告。

    Args:
        records: 交易记录列表
        output_dir: 输出目录
        source_info: 数据源信息

    Returns:
        str: 生成的文件路径，如无数据则返回 None
    """
    logger = get_logger()

    if not records:
        logger.warning('无交易记录，跳过重复交易检测')
        return None

    if output_dir is None:
        output_dir = get_script_dir()

    os.makedirs(output_dir, exist_ok=True)

    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f'重复交易检测报告_{timestamp}.xlsx'
    output_path = os.path.join(output_dir, filename)

    check_result = detect_duplicates(records)

    if not check_result.duplicate_group_count:
        logger.info('未检测到重复交易，仍导出报告')
        return export_duplicate_check_result(check_result, output_path, source_info)

    return export_duplicate_check_result(check_result, output_path, source_info)


def generate_duplicate_check_from_total(total_path: str,
                                        output_dir: Optional[str] = None) -> Optional[str]:
    """
    从银行流水总表文件生成重复交易检测报告。

    Args:
        total_path: 银行流水总表 Excel 文件路径
        output_dir: 输出目录

    Returns:
        str: 生成的文件路径，失败则返回 None
    """
    logger = get_logger()

    records = load_total_table(total_path)
    if not records:
        logger.warning('总表无数据: %s', total_path)
        return None

    if output_dir is None:
        output_dir = os.path.dirname(total_path) or get_script_dir()

    source_info = {
        '数据来源文件': os.path.basename(total_path),
        '总表记录数': len(records),
        '生成时间': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
    }

    return generate_duplicate_check_from_records(records, output_dir, source_info)


def run_duplicate_check_flow(script_dir):
    """重复交易检测 CLI 流程"""
    logger = get_logger()
    logger.info('========== 重复交易检测开始 ==========')

    print('\n' + '=' * 70)
    print('重复交易检测 - 跨文件去重与疑似重复标记')
    print('=' * 70)
    print('\n请选择数据来源：')
    print('  1) 从银行流水总表文件（Excel）')
    print('  2) 从数据库（按条件查询后检测）')
    print('  0) 返回主菜单')

    choice = input('\n请输入选项（默认 1）: ').strip() or '1'

    records = []
    source_info = {}

    if choice == '0':
        return
    elif choice == '1':
        total_path = ask_file('请选择【银行流水总表】文件')
        if not total_path:
            show_info('提示', '未选择总表文件，返回。')
            return
        logger.info('用户选择总表文件: %s', total_path)
        records = load_total_table(total_path)
        if not records:
            show_warning('错误', '总表文件无数据或读取失败。')
            return
        source_info = {
            '数据来源文件': os.path.basename(total_path),
            '总表记录数': len(records),
            '生成时间': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        }
    elif choice == '2':
        if not HAS_DATABASE:
            show_warning('错误', '数据库模块不可用。')
            return

        print('\n输入查询条件（直接回车表示不限制）：')
        subject = input('主体名称: ').strip() or None
        bank = input('银行名称: ').strip() or None
        account = input('银行账号: ').strip() or None
        start_date = input('开始日期 (YYYY-MM-DD): ').strip() or None
        end_date = input('结束日期 (YYYY-MM-DD): ').strip() or None

        try:
            qr = db_module.query_transactions(
                subject=subject, bank=bank, account=account,
                start_date=start_date, end_date=end_date,
                limit=999999, script_dir=script_dir
            )
            records = [r.to_dict() for r in qr.records]
        except Exception as e:
            show_warning('错误', f'数据库查询失败: {e}')
            logger.error('数据库查询失败: %s', e, exc_info=True)
            return

        if not records:
            show_info('提示', '查询结果为空。')
            return

        source_info = {
            '数据来源': '数据库查询',
            '查询主体': subject or '全部',
            '查询银行': bank or '全部',
            '查询账号': account or '全部',
            '日期范围': f'{start_date or "不限"} ~ {end_date or "不限"}',
            '记录数': len(records),
            '生成时间': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        }
    else:
        print('无效选项')
        return

    print(f'\n开始检测，共 {len(records)} 条记录...')
    check_result = detect_duplicates(records)
    summary = check_result.check_summary

    MATCH_TYPE_LABELS = {
        'transaction_id': '交易流水号匹配',
        'amount_date': '金额+日期匹配',
        'full': '完全匹配(含对方户名)',
    }

    print('\n' + '=' * 70)
    print('检测结果总览')
    print('=' * 70)
    print(f'  总记录数:         {summary.get("total_records", 0):,}')
    print(f'  疑似重复组数:     {summary.get("duplicate_group_count", 0):,}')
    print(f'  疑似重复记录数:   {summary.get("duplicate_record_count", 0):,}')
    print(f'  重复率:           {summary.get("duplicate_rate", 0):.2f}%')
    print()
    print(f'  流水号匹配:       {summary.get("transaction_id_groups", 0)} 组 / {summary.get("transaction_id_records", 0)} 条')
    print(f'  金额日期匹配:     {summary.get("amount_date_groups", 0)} 组 / {summary.get("amount_date_records", 0)} 条')
    print(f'  完全匹配:         {summary.get("full_groups", 0)} 组 / {summary.get("full_records", 0)} 条')

    if check_result.duplicate_group_count > 0:
        print(f'\n  ⚠️  发现 {check_result.duplicate_group_count} 组疑似重复交易，涉及 {check_result.duplicate_record_count} 条记录')
        for group in check_result.groups[:10]:
            first = group.records[0] if group.records else None
            print(f'    - 组{group.group_id} [{MATCH_TYPE_LABELS.get(group.match_type, group.match_type)}]: '
                  f'{first.bank_account if first else ""} / {first.trade_date if first else ""} / '
                  f'{group.record_count}条')
        if len(check_result.groups) > 10:
            print(f'    ... 还有 {len(check_result.groups) - 10} 组，详见导出文件')
    else:
        print(f'\n  ✅ 未检测到重复交易！')

    output_dir = input('\n请输入输出目录（直接回车默认当前目录）: ').strip()
    if not output_dir:
        output_dir = script_dir

    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_path = os.path.join(output_dir, f'重复交易检测报告_{timestamp}.xlsx')

    try:
        export_duplicate_check_result(check_result, output_path, source_info)
        msg = f'检测报告已导出！\n\n输出文件：{output_path}'
        show_info('导出成功', msg)
        logger.info('重复交易检测报告导出完成: %s', output_path)
    except Exception as e:
        msg = f'导出失败：{e}'
        show_warning('导出失败', msg)
        logger.error('重复交易检测报告导出失败: %s', e, exc_info=True)

    logger.info('========== 重复交易检测结束 ==========')


# ──────────────────────────────────────────────
# 流水检验报告模块
# ──────────────────────────────────────────────

def _collect_unmatched_accounts(
    records: List[Dict[str, Any]],
    file_details: List[FileProcessDetail],
) -> List[UnmatchedAccount]:
    """从交易记录和文件处理详情中收集主体未匹配的账号"""
    logger = get_logger()

    account_map: Dict[str, UnmatchedAccount] = {}

    for rec in records:
        subject = str(rec.get('主体') or '').strip()
        bank_account = str(rec.get('银行账号') or '').strip()
        bank_name = str(rec.get('银行') or '').strip()

        if not bank_account:
            continue
        if subject:
            continue

        key = f"{bank_name}||{bank_account}"
        if key not in account_map:
            account_map[key] = UnmatchedAccount(
                bank_account=bank_account,
                bank_name=bank_name,
            )

        payment = to_float(rec.get('付款'))
        receipt = to_float(rec.get('收款'))

        if receipt is not None and receipt > 0:
            account_map[key].total_income += receipt
        if payment is not None and payment < 0:
            account_map[key].total_expense += abs(payment)

        account_map[key].record_count += 1

    for detail in file_details:
        if detail.subject or not detail.bank_account:
            continue
        key = f"{detail.bank_name}||{detail.bank_account}"
        if key not in account_map:
            account_map[key] = UnmatchedAccount(
                bank_account=detail.bank_account,
                bank_name=detail.bank_name,
            )
        if detail.file_name and detail.file_name not in account_map[key].file_sources:
            account_map[key].file_sources.append(detail.file_name)

    result = sorted(
        account_map.values(),
        key=lambda x: (x.bank_name, x.bank_account),
    )
    logger.info('收集到 %d 个主体未匹配的账号', len(result))
    return result


def _collect_all_skipped_rows(
    file_details: List[FileProcessDetail],
) -> List[SkippedRowDetail]:
    """从所有文件处理详情中收集全部跳过行"""
    all_skipped: List[SkippedRowDetail] = []
    for detail in file_details:
        all_skipped.extend(detail.skipped_details)
    return all_skipped


def _build_amount_summary(
    records: List[Dict[str, Any]],
) -> Tuple[Dict[str, Any], List[Dict[str, Any]], List[Dict[str, Any]]]:
    """构建金额汇总数据（总体、按主体、按银行）"""
    logger = get_logger()

    total_income = 0.0
    total_expense = 0.0
    total_count = 0
    income_count = 0
    expense_count = 0

    by_subject: Dict[str, Dict[str, Any]] = {}
    by_bank: Dict[str, Dict[str, Any]] = {}

    for rec in records:
        subject = str(rec.get('主体') or '').strip() or '未指定主体'
        bank = str(rec.get('银行') or '').strip() or '未知银行'

        payment = to_float(rec.get('付款'))
        receipt = to_float(rec.get('收款'))

        income = 0.0
        expense = 0.0
        is_income = False
        is_expense = False

        if receipt is not None and receipt > 0:
            income = receipt
            is_income = True
        if payment is not None and payment < 0:
            expense = abs(payment)
            is_expense = True

        if not is_income and not is_expense:
            continue

        total_income += income
        total_expense += expense
        total_count += 1
        if is_income:
            income_count += 1
        if is_expense:
            expense_count += 1

        if subject not in by_subject:
            by_subject[subject] = {
                '主体': subject,
                '收入合计': 0.0,
                '支出合计': 0.0,
                '净额': 0.0,
                '交易笔数': 0,
                '收入笔数': 0,
                '支出笔数': 0,
            }
        s = by_subject[subject]
        s['收入合计'] += income
        s['支出合计'] += expense
        s['净额'] += income - expense
        s['交易笔数'] += 1
        if is_income:
            s['收入笔数'] += 1
        if is_expense:
            s['支出笔数'] += 1

        if bank not in by_bank:
            by_bank[bank] = {
                '银行': bank,
                '收入合计': 0.0,
                '支出合计': 0.0,
                '净额': 0.0,
                '交易笔数': 0,
                '收入笔数': 0,
                '支出笔数': 0,
            }
        b = by_bank[bank]
        b['收入合计'] += income
        b['支出合计'] += expense
        b['净额'] += income - expense
        b['交易笔数'] += 1
        if is_income:
            b['收入笔数'] += 1
        if is_expense:
            b['支出笔数'] += 1

    overall = {
        '总记录数': len(records),
        '有效交易笔数': total_count,
        '收入笔数': income_count,
        '支出笔数': expense_count,
        '收入合计': total_income,
        '支出合计': total_expense,
        '净额': total_income - total_expense,
        '主体数': len(by_subject),
        '银行数': len(by_bank),
    }

    by_subject_list = sorted(
        by_subject.values(),
        key=lambda x: abs(x['净额']),
        reverse=True,
    )
    by_bank_list = sorted(
        by_bank.values(),
        key=lambda x: abs(x['净额']),
        reverse=True,
    )

    logger.info(
        '金额汇总完成: %d 条记录, 收入 %.2f, 支出 %.2f, 净额 %.2f',
        total_count, total_income, total_expense, total_income - total_expense,
    )
    return overall, by_subject_list, by_bank_list


def _build_verification_report_data(
    records: List[Dict[str, Any]],
    file_details: List[FileProcessDetail],
    source_info: Optional[Dict[str, Any]] = None,
) -> VerificationReportData:
    """构建完整的检验报告数据"""
    logger = get_logger()

    skipped_rows = _collect_all_skipped_rows(file_details)
    unmatched_accounts = _collect_unmatched_accounts(records, file_details)
    amount_summary, by_subject_summary, by_bank_summary = _build_amount_summary(records)

    total_files = len(file_details)
    success_count = sum(1 for d in file_details if d.process_status == '成功')
    failed_count = sum(1 for d in file_details if d.process_status == '失败')
    unprocessed_count = sum(1 for d in file_details if d.process_status == '未处理')
    total_extracted = sum(d.extracted_records for d in file_details)
    total_skipped = sum(d.skipped_rows for d in file_details)

    source = dict(source_info or {})
    source.update({
        '文件总数': total_files,
        '成功处理': success_count,
        '处理失败': failed_count,
        '未处理': unprocessed_count,
        '提取记录总数': total_extracted,
        '跳过行总数': total_skipped,
        '未匹配账号数': len(unmatched_accounts),
    })

    data = VerificationReportData(
        source_info=source,
        file_details=file_details,
        skipped_rows=skipped_rows,
        unmatched_accounts=unmatched_accounts,
        amount_summary=amount_summary,
        by_subject_summary=by_subject_summary,
        by_bank_summary=by_bank_summary,
    )
    logger.info('检验报告数据构建完成')
    return data


def _fmt_amount(val: float) -> str:
    """格式化金额"""
    if val is None:
        return '-'
    try:
        return f"{val:,.2f}"
    except Exception:
        return str(val)


def export_verification_report_markdown(
    report_data: VerificationReportData,
    output_path: str,
) -> str:
    """导出检验报告为 Markdown 格式"""
    logger = get_logger()
    os.makedirs(os.path.dirname(output_path) or '.', exist_ok=True)

    lines: List[str] = []

    lines.append('# 银行流水检验报告')
    lines.append('')
    lines.append(f'**生成时间**: {report_data.source_info.get("生成时间", "-")}')
    lines.append(f'**操作人**: {report_data.source_info.get("操作人", "-")}')
    lines.append(f'**数据来源**: {report_data.source_info.get("数据来源", "-")}')
    lines.append(f'**输入文件夹**: {report_data.source_info.get("输入文件夹", "-")}')
    lines.append(f'**总表文件**: {report_data.source_info.get("总表文件", "-")}')
    lines.append(f'**运行模式**: {report_data.source_info.get("运行模式", "-")}')
    lines.append('')

    lines.append('## 一、处理总览')
    lines.append('')
    lines.append('| 指标 | 数值 |')
    lines.append('|------|------|')
    lines.append(f'| 文件总数 | {report_data.source_info.get("文件总数", 0)} |')
    lines.append(f'| 成功处理 | {report_data.source_info.get("成功处理", 0)} |')
    lines.append(f'| 处理失败 | {report_data.source_info.get("处理失败", 0)} |')
    lines.append(f'| 未处理 | {report_data.source_info.get("未处理", 0)} |')
    lines.append(f'| 提取记录总数 | {report_data.source_info.get("提取记录总数", 0):,} |')
    lines.append(f'| 跳过行总数 | {report_data.source_info.get("跳过行总数", 0):,} |')
    lines.append(f'| 未匹配账号数 | {report_data.source_info.get("未匹配账号数", 0)} |')
    lines.append('')

    lines.append('## 二、金额汇总')
    lines.append('')
    s = report_data.amount_summary
    lines.append('| 指标 | 数值 |')
    lines.append('|------|------|')
    lines.append(f"| 总记录数 | {s.get('总记录数', 0):,} |")
    lines.append(f"| 有效交易笔数 | {s.get('有效交易笔数', 0):,} |")
    lines.append(f"| 收入笔数 | {s.get('收入笔数', 0):,} |")
    lines.append(f"| 支出笔数 | {s.get('支出笔数', 0):,} |")
    lines.append(f"| 收入合计 | {_fmt_amount(s.get('收入合计', 0.0))} |")
    lines.append(f"| 支出合计 | {_fmt_amount(s.get('支出合计', 0.0))} |")
    lines.append(f"| 净额 | {_fmt_amount(s.get('净额', 0.0))} |")
    lines.append(f"| 主体数 | {s.get('主体数', 0)} |")
    lines.append(f"| 银行数 | {s.get('银行数', 0)} |")
    lines.append('')

    if report_data.by_subject_summary:
        lines.append('### 按主体汇总')
        lines.append('')
        lines.append('| 主体 | 收入合计 | 支出合计 | 净额 | 交易笔数 | 收入笔数 | 支出笔数 |')
        lines.append('|------|----------|----------|------|----------|----------|----------|')
        for row in report_data.by_subject_summary:
            lines.append(
                f"| {row['主体']} | {_fmt_amount(row['收入合计'])} | "
                f"{_fmt_amount(row['支出合计'])} | {_fmt_amount(row['净额'])} | "
                f"{row['交易笔数']:,} | {row['收入笔数']:,} | {row['支出笔数']:,} |"
            )
        lines.append('')

    if report_data.by_bank_summary:
        lines.append('### 按银行汇总')
        lines.append('')
        lines.append('| 银行 | 收入合计 | 支出合计 | 净额 | 交易笔数 | 收入笔数 | 支出笔数 |')
        lines.append('|------|----------|----------|------|----------|----------|----------|')
        for row in report_data.by_bank_summary:
            lines.append(
                f"| {row['银行']} | {_fmt_amount(row['收入合计'])} | "
                f"{_fmt_amount(row['支出合计'])} | {_fmt_amount(row['净额'])} | "
                f"{row['交易笔数']:,} | {row['收入笔数']:,} | {row['支出笔数']:,} |"
            )
        lines.append('')

    lines.append('## 三、各文件处理状态')
    lines.append('')
    lines.append('| 文件名 | 银行 | 银行账号 | 主体 | 处理状态 | 总行数 | 提取记录 | 跳过行 | 错误信息 |')
    lines.append('|--------|------|----------|------|----------|--------|----------|--------|----------|')
    for d in report_data.file_details:
        err = d.error_message.replace('|', '\\|') if d.error_message else ''
        lines.append(
            f"| {d.file_name} | {d.bank_name} | {d.bank_account} | {d.subject or '-'} | "
            f"{d.process_status} | {d.total_rows_in_excel:,} | {d.extracted_records:,} | "
            f"{d.skipped_rows:,} | {err} |"
        )
    lines.append('')

    lines.append('## 四、跳过行明细')
    lines.append('')
    if report_data.skipped_rows:
        lines.append(f'共 **{len(report_data.skipped_rows)}** 行被跳过。')
        lines.append('')
        lines.append('| 文件名 | 行号 | 跳过原因 | 行内容摘要 |')
        lines.append('|--------|------|----------|------------|')
        for sr in report_data.skipped_rows:
            content = (sr.raw_content or '').replace('|', '\\|').replace('\n', ' ')
            if len(content) > 100:
                content = content[:100] + '...'
            lines.append(
                f"| {sr.file_name} | {sr.row_number} | {sr.reason} | {content} |"
            )
    else:
        lines.append('无跳过行。')
    lines.append('')

    lines.append('## 五、主体未匹配账号列表')
    lines.append('')
    if report_data.unmatched_accounts:
        lines.append(f'共 **{len(report_data.unmatched_accounts)}** 个账号未匹配到主体。')
        lines.append('')
        lines.append('| 银行 | 银行账号 | 记录数 | 收入合计 | 支出合计 | 来源文件 |')
        lines.append('|------|----------|--------|----------|----------|----------|')
        for ua in report_data.unmatched_accounts:
            sources = ', '.join(ua.file_sources) if ua.file_sources else '-'
            lines.append(
                f"| {ua.bank_name} | {ua.bank_account} | {ua.record_count:,} | "
                f"{_fmt_amount(ua.total_income)} | {_fmt_amount(ua.total_expense)} | {sources} |"
            )
    else:
        lines.append('所有账号均已成功匹配主体。')
    lines.append('')

    lines.append('---')
    lines.append(f'*本报告由 bankcheck 工具自动生成，生成时间 {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}*')

    content = '\n'.join(lines)
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(content)

    logger.info('检验报告(Markdown)已导出: %s', output_path)
    return output_path


def export_verification_report_excel(
    report_data: VerificationReportData,
    output_path: str,
    source_info: Optional[Dict[str, Any]] = None,
) -> str:
    """导出检验报告为 Excel 格式（多 Sheet）"""
    logger = get_logger()
    os.makedirs(os.path.dirname(output_path) or '.', exist_ok=True)

    wb = openpyxl.Workbook()

    ws_info = wb.active
    ws_info.title = '报告信息'
    info_title = [
        ('项目', '内容'),
    ]
    info_rows = list(info_title)
    si = report_data.source_info
    info_rows.extend([
        ('报告类型', '银行流水检验报告'),
        ('生成时间', si.get('生成时间', '')),
        ('操作人', si.get('操作人', '')),
        ('数据来源', si.get('数据来源', '')),
        ('输入文件夹', si.get('输入文件夹', '')),
        ('总表文件', si.get('总表文件', '')),
        ('运行模式', si.get('运行模式', '')),
        ('', ''),
        ('文件总数', si.get('文件总数', 0)),
        ('成功处理', si.get('成功处理', 0)),
        ('处理失败', si.get('处理失败', 0)),
        ('未处理', si.get('未处理', 0)),
        ('提取记录总数', si.get('提取记录总数', 0)),
        ('跳过行总数', si.get('跳过行总数', 0)),
        ('未匹配账号数', si.get('未匹配账号数', 0)),
    ])
    for r_idx, row in enumerate(info_rows, 1):
        for c_idx, val in enumerate(row, 1):
            cell = ws_info.cell(row=r_idx, column=c_idx, value=val)
            if r_idx == 1:
                cell.font = openpyxl.styles.Font(bold=True)
    for col_idx, width in enumerate([20, 60], 1):
        ws_info.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = width

    ws_amount = wb.create_sheet('金额汇总')
    amount_header = ['指标', '数值']
    ws_amount.append(amount_header)
    s = report_data.amount_summary
    amount_rows = [
        ('总记录数', s.get('总记录数', 0)),
        ('有效交易笔数', s.get('有效交易笔数', 0)),
        ('收入笔数', s.get('收入笔数', 0)),
        ('支出笔数', s.get('支出笔数', 0)),
        ('收入合计', s.get('收入合计', 0.0)),
        ('支出合计', s.get('支出合计', 0.0)),
        ('净额', s.get('净额', 0.0)),
        ('主体数', s.get('主体数', 0)),
        ('银行数', s.get('银行数', 0)),
    ]
    for row in amount_rows:
        ws_amount.append(list(row))
    for cell in ws_amount[1]:
        cell.font = openpyxl.styles.Font(bold=True)
    for row in ws_amount.iter_rows(min_row=2):
        for cell in row:
            if isinstance(cell.value, float):
                cell.number_format = '#,##0.00'
            elif isinstance(cell.value, int):
                cell.number_format = '#,##0'
    ws_amount.column_dimensions['A'].width = 20
    ws_amount.column_dimensions['B'].width = 20

    if report_data.by_subject_summary:
        ws_subject = wb.create_sheet('按主体汇总')
        subject_header = ['主体', '收入合计', '支出合计', '净额', '交易笔数', '收入笔数', '支出笔数']
        ws_subject.append(subject_header)
        for row in report_data.by_subject_summary:
            ws_subject.append([
                row['主体'], row['收入合计'], row['支出合计'], row['净额'],
                row['交易笔数'], row['收入笔数'], row['支出笔数'],
            ])
        for cell in ws_subject[1]:
            cell.font = openpyxl.styles.Font(bold=True)
        for row in ws_subject.iter_rows(min_row=2):
            for idx, cell in enumerate(row):
                if idx in (1, 2, 3):
                    cell.number_format = '#,##0.00'
                elif idx in (4, 5, 6):
                    cell.number_format = '#,##0'
        widths = [25, 18, 18, 18, 12, 12, 12]
        for i, w in enumerate(widths, 1):
            ws_subject.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

    if report_data.by_bank_summary:
        ws_bank = wb.create_sheet('按银行汇总')
        bank_header = ['银行', '收入合计', '支出合计', '净额', '交易笔数', '收入笔数', '支出笔数']
        ws_bank.append(bank_header)
        for row in report_data.by_bank_summary:
            ws_bank.append([
                row['银行'], row['收入合计'], row['支出合计'], row['净额'],
                row['交易笔数'], row['收入笔数'], row['支出笔数'],
            ])
        for cell in ws_bank[1]:
            cell.font = openpyxl.styles.Font(bold=True)
        for row in ws_bank.iter_rows(min_row=2):
            for idx, cell in enumerate(row):
                if idx in (1, 2, 3):
                    cell.number_format = '#,##0.00'
                elif idx in (4, 5, 6):
                    cell.number_format = '#,##0'
        widths = [20, 18, 18, 18, 12, 12, 12]
        for i, w in enumerate(widths, 1):
            ws_bank.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

    ws_files = wb.create_sheet('文件处理状态')
    file_header = ['文件名', '银行', '银行账号', '主体', '处理状态',
                   '总行数', '提取记录', '跳过行', '错误信息']
    ws_files.append(file_header)
    for d in report_data.file_details:
        ws_files.append([
            d.file_name, d.bank_name, d.bank_account, d.subject or '',
            d.process_status, d.total_rows_in_excel, d.extracted_records,
            d.skipped_rows, d.error_message or '',
        ])
    for cell in ws_files[1]:
        cell.font = openpyxl.styles.Font(bold=True)
    for row in ws_files.iter_rows(min_row=2):
        for idx, cell in enumerate(row):
            if idx in (5, 6, 7):
                cell.number_format = '#,##0'
    widths = [40, 15, 25, 20, 10, 10, 12, 10, 40]
    for i, w in enumerate(widths, 1):
        ws_files.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

    ws_skip = wb.create_sheet('跳过行明细')
    skip_header = ['文件名', '行号', '跳过原因', '行内容摘要']
    ws_skip.append(skip_header)
    for sr in report_data.skipped_rows:
        ws_skip.append([sr.file_name, sr.row_number, sr.reason, sr.raw_content or ''])
    for cell in ws_skip[1]:
        cell.font = openpyxl.styles.Font(bold=True)
    widths = [40, 10, 15, 80]
    for i, w in enumerate(widths, 1):
        ws_skip.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

    ws_unmatched = wb.create_sheet('未匹配账号')
    unmatched_header = ['银行', '银行账号', '记录数', '收入合计', '支出合计', '来源文件']
    ws_unmatched.append(unmatched_header)
    for ua in report_data.unmatched_accounts:
        ws_unmatched.append([
            ua.bank_name, ua.bank_account, ua.record_count,
            ua.total_income, ua.total_expense,
            ', '.join(ua.file_sources) if ua.file_sources else '',
        ])
    for cell in ws_unmatched[1]:
        cell.font = openpyxl.styles.Font(bold=True)
    for row in ws_unmatched.iter_rows(min_row=2):
        for idx, cell in enumerate(row):
            if idx == 2:
                cell.number_format = '#,##0'
            elif idx in (3, 4):
                cell.number_format = '#,##0.00'
    widths = [15, 25, 12, 18, 18, 50]
    for i, w in enumerate(widths, 1):
        ws_unmatched.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

    wb.save(output_path)
    logger.info('检验报告(Excel)已导出: %s', output_path)
    return output_path


def generate_verification_report_from_records(
    records: List[Dict[str, Any]],
    file_details: List[FileProcessDetail],
    output_dir: Optional[str] = None,
    source_info: Optional[Dict[str, Any]] = None,
) -> Tuple[Optional[str], Optional[str]]:
    """
    从交易记录和文件处理详情生成流水检验报告（Excel + Markdown）。

    Args:
        records: 交易记录列表
        file_details: 各文件处理详情列表
        output_dir: 输出目录
        source_info: 数据源信息

    Returns:
        tuple: (Excel报告路径, Markdown报告路径)
    """
    logger = get_logger()

    if output_dir is None:
        output_dir = get_script_dir()

    os.makedirs(output_dir, exist_ok=True)

    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    excel_path = os.path.join(output_dir, f'流水检验报告_{timestamp}.xlsx')
    md_path = os.path.join(output_dir, f'流水检验报告_{timestamp}.md')

    try:
        report_data = _build_verification_report_data(records, file_details, source_info)

        excel_result = export_verification_report_excel(report_data, excel_path, source_info)
        md_result = export_verification_report_markdown(report_data, md_path)

        return excel_result, md_result
    except Exception as e:
        logger.error('生成检验报告失败: %s', e, exc_info=True)
        return None, None


def generate_verification_report_from_total(
    total_path: str,
    output_dir: Optional[str] = None,
) -> Tuple[Optional[str], Optional[str]]:
    """
    从银行流水总表文件生成检验报告（无文件级详情时退化为汇总模式）。

    Args:
        total_path: 银行流水总表 Excel 文件路径
        output_dir: 输出目录

    Returns:
        tuple: (Excel报告路径, Markdown报告路径)
    """
    logger = get_logger()

    records = load_total_table(total_path)
    if not records:
        logger.warning('总表无数据: %s', total_path)
        return None, None

    if output_dir is None:
        output_dir = os.path.dirname(total_path) or get_script_dir()

    source_info = {
        '数据来源文件': os.path.basename(total_path),
        '总表记录数': len(records),
        '生成时间': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
    }

    file_details: List[FileProcessDetail] = []
    account_file_map: Dict[Tuple[str, str], set] = {}
    for rec in records:
        bank = str(rec.get('银行') or '').strip() or '未知'
        account = str(rec.get('银行账号') or '').strip()
        if not account:
            continue
        key = (bank, account)
        if key not in account_file_map:
            account_file_map[key] = set()
        account_file_map[key].add(bank)

    for (bank, account), _ in account_file_map.items():
        subject = ''
        for rec in records:
            if str(rec.get('银行账号') or '').strip() == account:
                subject = str(rec.get('主体') or '').strip()
                break
        rec_count = sum(
            1 for r in records
            if str(r.get('银行账号') or '').strip() == account
        )
        file_details.append(FileProcessDetail(
            file_name=os.path.basename(total_path),
            file_path=total_path,
            bank_name=bank,
            bank_account=account,
            subject=subject,
            total_rows_in_excel=rec_count,
            extracted_records=rec_count,
            process_status='汇总',
        ))

    return generate_verification_report_from_records(records, file_details, output_dir, source_info)


if __name__ == '__main__':
    main()
