# -*- coding: utf-8 -*-
"""
运行性能剖析报告模块
记录每个文件的打开耗时、行遍历耗时、查找表命中耗时，
输出性能瓶颈分析供优化大文件夹场景时参考。
"""

import os
import sys
import time
import json
import logging
from dataclasses import dataclass, field, asdict
from typing import List, Optional, Dict, Any
from datetime import datetime
from contextlib import contextmanager


def get_script_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def get_logger():
    return logging.getLogger('bankcheck')


@dataclass
class FileOpenRecord:
    filepath: str
    duration_ms: float
    file_size_kb: float = 0.0
    bank_name: str = ''
    is_xls_convert: bool = False

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


@dataclass
class RowTraversalRecord:
    filepath: str
    sheet_name: str
    row_count: int
    duration_ms: float
    bank_name: str = ''
    extracted_count: int = 0

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


@dataclass
class LookupHitRecord:
    bank_account: str
    duration_ms: float
    hit: bool = False
    fuzzy: bool = False
    similarity: float = 0.0

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


@dataclass
class PhaseRecord:
    phase_name: str
    duration_ms: float
    detail: str = ''

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


class PerfProfiler:
    def __init__(self):
        self.file_open_records: List[FileOpenRecord] = []
        self.row_traversal_records: List[RowTraversalRecord] = []
        self.lookup_hit_records: List[LookupHitRecord] = []
        self.phase_records: List[PhaseRecord] = []
        self._start_time: Optional[float] = None
        self._total_duration_ms: float = 0.0

    def start(self):
        self._start_time = time.perf_counter()

    def stop(self):
        if self._start_time is not None:
            self._total_duration_ms = (time.perf_counter() - self._start_time) * 1000
            self._start_time = None

    def record_file_open(self, filepath: str, duration_ms: float,
                         bank_name: str = '', is_xls_convert: bool = False):
        file_size_kb = 0.0
        try:
            file_size_kb = os.path.getsize(filepath) / 1024.0
        except (OSError, TypeError):
            pass
        rec = FileOpenRecord(
            filepath=filepath,
            duration_ms=duration_ms,
            file_size_kb=round(file_size_kb, 2),
            bank_name=bank_name,
            is_xls_convert=is_xls_convert,
        )
        self.file_open_records.append(rec)

    def record_row_traversal(self, filepath: str, sheet_name: str,
                             row_count: int, duration_ms: float,
                             bank_name: str = '', extracted_count: int = 0):
        rec = RowTraversalRecord(
            filepath=filepath,
            sheet_name=sheet_name,
            row_count=row_count,
            duration_ms=duration_ms,
            bank_name=bank_name,
            extracted_count=extracted_count,
        )
        self.row_traversal_records.append(rec)

    def record_lookup_hit(self, bank_account: str, duration_ms: float,
                          hit: bool = False, fuzzy: bool = False,
                          similarity: float = 0.0):
        rec = LookupHitRecord(
            bank_account=bank_account,
            duration_ms=duration_ms,
            hit=hit,
            fuzzy=fuzzy,
            similarity=similarity,
        )
        self.lookup_hit_records.append(rec)

    def record_phase(self, phase_name: str, duration_ms: float, detail: str = ''):
        rec = PhaseRecord(
            phase_name=phase_name,
            duration_ms=duration_ms,
            detail=detail,
        )
        self.phase_records.append(rec)

    @contextmanager
    def measure_file_open(self, filepath: str, bank_name: str = '',
                          is_xls_convert: bool = False):
        start = time.perf_counter()
        try:
            yield
        finally:
            duration_ms = (time.perf_counter() - start) * 1000
            self.record_file_open(filepath, duration_ms, bank_name, is_xls_convert)

    @contextmanager
    def measure_row_traversal(self, filepath: str, sheet_name: str,
                              bank_name: str = ''):
        start = time.perf_counter()
        result_holder = {'row_count': 0, 'extracted_count': 0}
        try:
            yield result_holder
        finally:
            duration_ms = (time.perf_counter() - start) * 1000
            self.record_row_traversal(
                filepath, sheet_name,
                result_holder['row_count'], duration_ms,
                bank_name, result_holder['extracted_count'],
            )

    @contextmanager
    def measure_lookup_hit(self, bank_account: str):
        start = time.perf_counter()
        result_holder = {'hit': False, 'fuzzy': False, 'similarity': 0.0}
        try:
            yield result_holder
        finally:
            duration_ms = (time.perf_counter() - start) * 1000
            self.record_lookup_hit(
                bank_account, duration_ms,
                result_holder['hit'], result_holder['fuzzy'],
                result_holder['similarity'],
            )

    @contextmanager
    def measure_phase(self, phase_name: str, detail: str = ''):
        start = time.perf_counter()
        try:
            yield
        finally:
            duration_ms = (time.perf_counter() - start) * 1000
            self.record_phase(phase_name, duration_ms, detail)

    def get_summary(self) -> Dict[str, Any]:
        total_file_open_ms = sum(r.duration_ms for r in self.file_open_records)
        total_traversal_ms = sum(r.duration_ms for r in self.row_traversal_records)
        total_lookup_ms = sum(r.duration_ms for r in self.lookup_hit_records)
        total_phase_ms = sum(r.duration_ms for r in self.phase_records)

        file_open_count = len(self.file_open_records)
        traversal_count = len(self.row_traversal_records)
        lookup_count = len(self.lookup_hit_records)

        avg_file_open_ms = total_file_open_ms / file_open_count if file_open_count else 0
        avg_traversal_ms = total_traversal_ms / traversal_count if traversal_count else 0
        avg_lookup_ms = total_lookup_ms / lookup_count if lookup_count else 0

        lookup_hit_count = sum(1 for r in self.lookup_hit_records if r.hit)
        lookup_miss_count = lookup_count - lookup_hit_count
        lookup_fuzzy_count = sum(1 for r in self.lookup_hit_records if r.fuzzy)
        lookup_hit_rate = lookup_hit_count / lookup_count if lookup_count else 0

        total_rows_traversed = sum(r.row_count for r in self.row_traversal_records)
        total_rows_extracted = sum(r.extracted_count for r in self.row_traversal_records)

        slowest_file_opens = sorted(
            self.file_open_records, key=lambda r: r.duration_ms, reverse=True
        )[:5]
        slowest_traversals = sorted(
            self.row_traversal_records, key=lambda r: r.duration_ms, reverse=True
        )[:5]

        bottlenecks = self._analyze_bottlenecks(
            total_file_open_ms, total_traversal_ms, total_lookup_ms,
            avg_file_open_ms, avg_traversal_ms, avg_lookup_ms,
            lookup_hit_rate, file_open_count,
        )

        return {
            'total_duration_ms': round(self._total_duration_ms, 2),
            'file_open': {
                'count': file_open_count,
                'total_ms': round(total_file_open_ms, 2),
                'avg_ms': round(avg_file_open_ms, 2),
                'max_ms': round(max((r.duration_ms for r in self.file_open_records), default=0), 2),
                'min_ms': round(min((r.duration_ms for r in self.file_open_records), default=0), 2),
                'slowest': [
                    {
                        'filepath': os.path.basename(r.filepath),
                        'duration_ms': round(r.duration_ms, 2),
                        'size_kb': r.file_size_kb,
                        'bank': r.bank_name,
                        'xls_convert': r.is_xls_convert,
                    }
                    for r in slowest_file_opens
                ],
            },
            'row_traversal': {
                'count': traversal_count,
                'total_ms': round(total_traversal_ms, 2),
                'avg_ms': round(avg_traversal_ms, 2),
                'max_ms': round(max((r.duration_ms for r in self.row_traversal_records), default=0), 2),
                'min_ms': round(min((r.duration_ms for r in self.row_traversal_records), default=0), 2),
                'total_rows': total_rows_traversed,
                'total_extracted': total_rows_extracted,
                'slowest': [
                    {
                        'filepath': os.path.basename(r.filepath),
                        'sheet': r.sheet_name,
                        'duration_ms': round(r.duration_ms, 2),
                        'row_count': r.row_count,
                        'bank': r.bank_name,
                    }
                    for r in slowest_traversals
                ],
            },
            'lookup_hit': {
                'count': lookup_count,
                'total_ms': round(total_lookup_ms, 2),
                'avg_ms': round(avg_lookup_ms, 2),
                'hit_count': lookup_hit_count,
                'miss_count': lookup_miss_count,
                'fuzzy_count': lookup_fuzzy_count,
                'hit_rate': round(lookup_hit_rate, 4),
            },
            'phases': [
                {
                    'name': r.phase_name,
                    'duration_ms': round(r.duration_ms, 2),
                    'detail': r.detail,
                }
                for r in self.phase_records
            ],
            'bottlenecks': bottlenecks,
        }

    def _analyze_bottlenecks(self, total_file_open_ms, total_traversal_ms,
                             total_lookup_ms, avg_file_open_ms, avg_traversal_ms,
                             avg_lookup_ms, lookup_hit_rate, file_open_count):
        bottlenecks = []
        measured_total = total_file_open_ms + total_traversal_ms + total_lookup_ms
        if measured_total == 0:
            return bottlenecks

        file_open_pct = total_file_open_ms / measured_total * 100
        traversal_pct = total_traversal_ms / measured_total * 100
        lookup_pct = total_lookup_ms / measured_total * 100

        if file_open_pct > 40:
            bottlenecks.append({
                'phase': 'file_open',
                'severity': 'high',
                'pct': round(file_open_pct, 1),
                'suggestion': '文件打开耗时占比 %.1f%%，建议：1) 减少单次打开文件数量 2) 使用 read_only 模式 3) 检查网络磁盘 IO 延迟' % file_open_pct,
            })
        elif file_open_pct > 25:
            bottlenecks.append({
                'phase': 'file_open',
                'severity': 'medium',
                'pct': round(file_open_pct, 1),
                'suggestion': '文件打开耗时占比 %.1f%%，可考虑使用 read_only 模式优化大文件打开速度' % file_open_pct,
            })

        if traversal_pct > 40:
            bottlenecks.append({
                'phase': 'row_traversal',
                'severity': 'high',
                'pct': round(traversal_pct, 1),
                'suggestion': '行遍历耗时占比 %.1f%%，建议：1) 检查大文件是否可拆分 2) 减少无关行读取 3) 评估是否需要多线程并行处理' % traversal_pct,
            })
        elif traversal_pct > 25:
            bottlenecks.append({
                'phase': 'row_traversal',
                'severity': 'medium',
                'pct': round(traversal_pct, 1),
                'suggestion': '行遍历耗时占比 %.1f%%，可考虑优化行遍历逻辑减少不必要的单元格访问' % traversal_pct,
            })

        if lookup_pct > 20:
            bottlenecks.append({
                'phase': 'lookup_hit',
                'severity': 'high' if lookup_pct > 35 else 'medium',
                'pct': round(lookup_pct, 1),
                'suggestion': '查找表命中耗时占比 %.1f%%，建议：1) 预加载查找表到内存 2) 使用哈希表替代线性查找 3) 检查是否误启用了模糊匹配' % lookup_pct,
            })

        if avg_file_open_ms > 2000:
            bottlenecks.append({
                'phase': 'file_open',
                'severity': 'high',
                'pct': round(file_open_pct, 1),
                'suggestion': '单文件平均打开耗时 %.0fms，超过 2000ms 阈值，建议检查文件大小和磁盘性能' % avg_file_open_ms,
            })

        if lookup_hit_rate < 0.5 and file_open_count > 5:
            bottlenecks.append({
                'phase': 'lookup_hit',
                'severity': 'medium',
                'pct': round(lookup_pct, 1),
                'suggestion': '查找表命中率仅 %.1f%%，大量账号未匹配，建议补充查找表数据以减少无效查找' % (lookup_hit_rate * 100),
            })

        if file_open_count > 50:
            bottlenecks.append({
                'phase': 'file_open',
                'severity': 'low',
                'pct': round(file_open_pct, 1),
                'suggestion': '共处理 %d 个文件，文件数量较多，建议评估是否需要批量并行处理' % file_open_count,
            })

        return bottlenecks

    def generate_report(self) -> str:
        summary = self.get_summary()
        now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        lines = []
        lines.append('# 运行性能剖析报告')
        lines.append('')
        lines.append(f'**生成时间**: {now}')
        lines.append(f'**总耗时**: {summary["total_duration_ms"]:.1f} ms'
                     f' ({summary["total_duration_ms"] / 1000:.2f} s)')
        lines.append('')

        fo = summary['file_open']
        lines.append('---')
        lines.append('')
        lines.append('## 1. 文件打开耗时')
        lines.append('')
        lines.append(f'| 指标 | 值 |')
        lines.append(f'|------|------|')
        lines.append(f'| 文件数量 | {fo["count"]} |')
        lines.append(f'| 总耗时 | {fo["total_ms"]:.1f} ms |')
        lines.append(f'| 平均耗时 | {fo["avg_ms"]:.1f} ms |')
        lines.append(f'| 最大耗时 | {fo["max_ms"]:.1f} ms |')
        lines.append(f'| 最小耗时 | {fo["min_ms"]:.1f} ms |')
        lines.append('')

        if fo['slowest']:
            lines.append('### 最慢的 5 个文件')
            lines.append('')
            lines.append('| 文件名 | 耗时(ms) | 大小(KB) | 银行 | XLS转换 |')
            lines.append('|--------|----------|----------|------|---------|')
            for s in fo['slowest']:
                lines.append(
                    f'| {s["filepath"]} | {s["duration_ms"]:.1f} | '
                    f'{s["size_kb"]:.1f} | {s["bank"]} | '
                    f'{"是" if s["xls_convert"] else "否"} |'
                )
            lines.append('')

        rt = summary['row_traversal']
        lines.append('---')
        lines.append('')
        lines.append('## 2. 行遍历耗时')
        lines.append('')
        lines.append(f'| 指标 | 值 |')
        lines.append(f'|------|------|')
        lines.append(f'| 工作表数量 | {rt["count"]} |')
        lines.append(f'| 总耗时 | {rt["total_ms"]:.1f} ms |')
        lines.append(f'| 平均耗时 | {rt["avg_ms"]:.1f} ms |')
        lines.append(f'| 最大耗时 | {rt["max_ms"]:.1f} ms |')
        lines.append(f'| 最小耗时 | {rt["min_ms"]:.1f} ms |')
        lines.append(f'| 遍历总行数 | {rt["total_rows"]} |')
        lines.append(f'| 提取记录数 | {rt["total_extracted"]} |')
        lines.append('')

        if rt['slowest']:
            lines.append('### 最慢的 5 个工作表')
            lines.append('')
            lines.append('| 文件名 | 工作表 | 耗时(ms) | 行数 | 银行 |')
            lines.append('|--------|--------|----------|------|------|')
            for s in rt['slowest']:
                lines.append(
                    f'| {s["filepath"]} | {s["sheet"]} | '
                    f'{s["duration_ms"]:.1f} | {s["row_count"]} | {s["bank"]} |'
                )
            lines.append('')

        lk = summary['lookup_hit']
        lines.append('---')
        lines.append('')
        lines.append('## 3. 查找表命中耗时')
        lines.append('')
        lines.append(f'| 指标 | 值 |')
        lines.append(f'|------|------|')
        lines.append(f'| 查找次数 | {lk["count"]} |')
        lines.append(f'| 总耗时 | {lk["total_ms"]:.1f} ms |')
        lines.append(f'| 平均耗时 | {lk["avg_ms"]:.3f} ms |')
        lines.append(f'| 命中次数 | {lk["hit_count"]} |')
        lines.append(f'| 未命中次数 | {lk["miss_count"]} |')
        lines.append(f'| 模糊匹配次数 | {lk["fuzzy_count"]} |')
        lines.append(f'| 命中率 | {lk["hit_rate"] * 100:.1f}% |')
        lines.append('')

        if summary['phases']:
            lines.append('---')
            lines.append('')
            lines.append('## 4. 阶段耗时')
            lines.append('')
            lines.append('| 阶段 | 耗时(ms) | 说明 |')
            lines.append('|------|----------|------|')
            for p in summary['phases']:
                lines.append(f'| {p["name"]} | {p["duration_ms"]:.1f} | {p["detail"]} |')
            lines.append('')

        if summary['bottlenecks']:
            lines.append('---')
            lines.append('')
            lines.append('## 5. 性能瓶颈分析')
            lines.append('')
            for b in summary['bottlenecks']:
                severity_icon = {'high': '🔴', 'medium': '🟡', 'low': '🟢'}.get(
                    b['severity'], '⚪'
                )
                lines.append(f'### {severity_icon} {b["phase"]}（占比 {b["pct"]}%，严重度: {b["severity"]}）')
                lines.append('')
                lines.append(f'{b["suggestion"]}')
                lines.append('')
        else:
            lines.append('---')
            lines.append('')
            lines.append('## 5. 性能瓶颈分析')
            lines.append('')
            lines.append('未发现明显性能瓶颈。')
            lines.append('')

        lines.append('---')
        lines.append(f'*报告由 perf_profiler 自动生成 @ {now}*')

        return '\n'.join(lines)

    def save_report(self, output_dir: str) -> str:
        os.makedirs(output_dir, exist_ok=True)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f'性能剖析报告_{timestamp}.md'
        report_path = os.path.join(output_dir, filename)

        report_content = self.generate_report()
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write(report_content)

        summary_data = self.get_summary()
        json_filename = f'性能剖析数据_{timestamp}.json'
        json_path = os.path.join(output_dir, json_filename)
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(summary_data, f, ensure_ascii=False, indent=2)

        logger = get_logger()
        logger.info('性能剖析报告已保存: %s', report_path)
        logger.info('性能剖析数据已保存: %s', json_path)
        return report_path

    def reset(self):
        self.file_open_records.clear()
        self.row_traversal_records.clear()
        self.lookup_hit_records.clear()
        self.phase_records.clear()
        self._start_time = None
        self._total_duration_ms = 0.0


_global_profiler: Optional[PerfProfiler] = None


def get_profiler() -> PerfProfiler:
    global _global_profiler
    if _global_profiler is None:
        _global_profiler = PerfProfiler()
    return _global_profiler


def reset_profiler():
    global _global_profiler
    _global_profiler = None
