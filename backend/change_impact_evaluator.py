# -*- coding: utf-8 -*-
"""
变更影响评估模块

功能：
1. 当用户修改银行列映射（bank_rules.yaml）时，自动用历史样本试跑
2. 当用户修改查找表（主体映射）时，自动用历史样本试跑
3. 输出影响记录数、字段差异示例报告，降低配置变更风险
"""

import os
import sys
import json
import logging
import tempfile
import hashlib
import shutil
from typing import List, Dict, Optional, Tuple, Any
from dataclasses import dataclass, field, asdict
from datetime import datetime

import pandas as pd
import openpyxl
import yaml


def get_script_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def get_logger():
    return logging.getLogger('bankcheck')


HISTORY_SAMPLES_DIR_NAME = 'history_samples'
IMPACT_REPORT_DIR_NAME = 'impact_reports'
MAX_HISTORY_SAMPLES = 50
MAX_FIELD_DIFF_EXAMPLES = 20


def get_history_samples_dir(script_dir=None):
    if script_dir is None:
        script_dir = get_script_dir()
    history_dir = os.path.join(script_dir, HISTORY_SAMPLES_DIR_NAME)
    os.makedirs(history_dir, exist_ok=True)
    return history_dir


def get_impact_report_dir(script_dir=None):
    if script_dir is None:
        script_dir = get_script_dir()
    report_dir = os.path.join(script_dir, IMPACT_REPORT_DIR_NAME)
    os.makedirs(report_dir, exist_ok=True)
    return report_dir


def _normalize_value(val):
    if val is None:
        return None
    if isinstance(val, float):
        if pd.isna(val):
            return None
        if val == int(val):
            return int(val)
    if isinstance(val, str):
        stripped = val.strip()
        return stripped if stripped else None
    return val


def _values_equal(v1, v2):
    nv1 = _normalize_value(v1)
    nv2 = _normalize_value(v2)
    if nv1 is None and nv2 is None:
        return True
    if nv1 is None or nv2 is None:
        return False
    return str(nv1) == str(nv2)


@dataclass
class FieldDiff:
    field_name: str
    old_value: Any
    new_value: Any
    record_id: str = ''


@dataclass
class RecordDiff:
    record_id: str
    field_diffs: List[FieldDiff] = field(default_factory=list)


@dataclass
class ImpactReport:
    change_type: str
    change_target: str
    timestamp: str = ''
    total_records: int = 0
    affected_records: int = 0
    unaffected_records: int = 0
    added_records: int = 0
    removed_records: int = 0
    field_diff_summary: Dict[str, int] = field(default_factory=dict)
    field_diff_examples: List[Dict[str, Any]] = field(default_factory=list)
    details: str = ''

    def to_dict(self):
        return {
            'change_type': self.change_type,
            'change_target': self.change_target,
            'timestamp': self.timestamp,
            'total_records': self.total_records,
            'affected_records': self.affected_records,
            'unaffected_records': self.unaffected_records,
            'added_records': self.added_records,
            'removed_records': self.removed_records,
            'field_diff_summary': self.field_diff_summary,
            'field_diff_examples': self.field_diff_examples,
            'details': self.details,
        }

    def to_markdown(self) -> str:
        lines = [
            f'# 变更影响评估报告',
            f'',
            f'- **变更类型**: {self.change_type}',
            f'- **变更目标**: {self.change_target}',
            f'- **评估时间**: {self.timestamp}',
            f'',
            f'## 总体统计',
            f'',
            f'| 指标 | 数量 |',
            f'|------|------|',
            f'| 总记录数 | {self.total_records} |',
            f'| 受影响记录数 | {self.affected_records} |',
            f'| 未受影响记录数 | {self.unaffected_records} |',
            f'| 新增记录数 | {self.added_records} |',
            f'| 减少记录数 | {self.removed_records} |',
            f'',
        ]

        if self.field_diff_summary:
            lines.append('## 字段差异统计')
            lines.append('')
            lines.append('| 字段名 | 差异记录数 |')
            lines.append('|--------|------------|')
            for field_name, count in sorted(
                self.field_diff_summary.items(), key=lambda x: -x[1]
            ):
                lines.append(f'| {field_name} | {count} |')
            lines.append('')

        if self.field_diff_examples:
            lines.append('## 字段差异示例')
            lines.append('')
            for i, example in enumerate(self.field_diff_examples[:MAX_FIELD_DIFF_EXAMPLES], 1):
                lines.append(f'### 示例 {i}')
                lines.append('')
                lines.append(f'- **记录标识**: {example.get("record_id", "")}')
                lines.append(f'- **字段**: {example.get("field_name", "")}')
                lines.append(f'- **变更前**: `{example.get("old_value", "")}`')
                lines.append(f'- **变更后**: `{example.get("new_value", "")}`')
                lines.append('')

        if self.details:
            lines.append('## 详细说明')
            lines.append('')
            lines.append(self.details)
            lines.append('')

        if self.affected_records > 0:
            lines.append('## ⚠️ 风险提示')
            lines.append('')
            lines.append(
                f'本次变更将影响 **{self.affected_records}** 条记录，请确认变更是否符合预期。'
            )
            lines.append('')
        else:
            lines.append('## ✅ 评估结论')
            lines.append('')
            lines.append('本次变更未影响任何记录，可以安全应用。')
            lines.append('')

        return '\n'.join(lines)

    def save_report(self, output_dir=None) -> str:
        if not self.timestamp:
            self.timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        if output_dir is None:
            output_dir = get_impact_report_dir()
        os.makedirs(output_dir, exist_ok=True)

        safe_target = self.change_target.replace('/', '_').replace('\\', '_')
        base_name = f'impact_report_{self.change_type}_{safe_target}_{self.timestamp}'

        md_path = os.path.join(output_dir, base_name + '.md')
        with open(md_path, 'w', encoding='utf-8') as f:
            f.write(self.to_markdown())

        json_path = os.path.join(output_dir, base_name + '.json')
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(self.to_dict(), f, ensure_ascii=False, indent=2)

        return md_path


class HistorySampleManager:
    """历史样本数据管理器"""

    def __init__(self, samples_dir=None):
        self.samples_dir = samples_dir or get_history_samples_dir()
        self.logger = get_logger()

    def _get_sample_metadata_path(self):
        return os.path.join(self.samples_dir, 'samples_metadata.json')

    def _load_metadata(self) -> List[Dict[str, Any]]:
        meta_path = self._get_sample_metadata_path()
        if not os.path.exists(meta_path):
            return []
        try:
            with open(meta_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return []

    def _save_metadata(self, metadata: List[Dict[str, Any]]):
        meta_path = self._get_sample_metadata_path()
        try:
            with open(meta_path, 'w', encoding='utf-8') as f:
                json.dump(metadata, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self.logger.warning('保存样本元数据失败: %s', e)

    def _file_hash(self, filepath: str) -> str:
        h = hashlib.sha256()
        with open(filepath, 'rb') as f:
            while True:
                chunk = f.read(8192)
                if not chunk:
                    break
                h.update(chunk)
        return h.hexdigest()

    def save_sample(
        self,
        source_filepath: str,
        bank_name: str,
        records: List[Dict[str, Any]],
    ) -> Optional[str]:
        """
        保存一份处理结果作为历史样本

        Args:
            source_filepath: 原始银行流水文件路径
            bank_name: 银行名称
            records: 解析后的记录列表

        Returns:
            保存的样本文件路径，失败返回 None
        """
        try:
            if not records:
                return None

            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            safe_bank = bank_name.replace('/', '_').replace('\\', '_')
            sample_filename = f'{safe_bank}_{timestamp}.json'
            sample_path = os.path.join(self.samples_dir, sample_filename)

            sample_data = {
                'source_file': os.path.basename(source_filepath),
                'source_full_path': os.path.abspath(source_filepath),
                'source_hash': self._file_hash(source_filepath) if os.path.exists(source_filepath) else '',
                'bank_name': bank_name,
                'saved_at': timestamp,
                'record_count': len(records),
                'records': records,
            }

            with open(sample_path, 'w', encoding='utf-8') as f:
                json.dump(sample_data, f, ensure_ascii=False, indent=2, default=str)

            metadata = self._load_metadata()
            metadata.insert(0, {
                'filename': sample_filename,
                'bank_name': bank_name,
                'source_file': os.path.basename(source_filepath),
                'source_full_path': os.path.abspath(source_filepath),
                'source_hash': sample_data['source_hash'],
                'saved_at': timestamp,
                'record_count': len(records),
                'sample_path': sample_path,
            })

            while len(metadata) > MAX_HISTORY_SAMPLES:
                old = metadata.pop()
                old_path = os.path.join(self.samples_dir, old['filename'])
                if os.path.exists(old_path):
                    try:
                        os.remove(old_path)
                    except OSError:
                        pass

            self._save_metadata(metadata)
            self.logger.info('已保存历史样本: %s (%d 条记录)', sample_filename, len(records))
            return sample_path

        except Exception as e:
            self.logger.warning('保存历史样本失败: %s', e)
            return None

    def list_samples(self, bank_name: Optional[str] = None) -> List[Dict[str, Any]]:
        """列出所有历史样本，可选按银行筛选"""
        metadata = self._load_metadata()
        if bank_name:
            return [m for m in metadata if m.get('bank_name') == bank_name]
        return metadata

    def load_sample(self, sample_path: str) -> Optional[Dict[str, Any]]:
        """加载单个样本数据"""
        try:
            if not os.path.exists(sample_path):
                return None
            with open(sample_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            self.logger.warning('加载样本失败 %s: %s', sample_path, e)
            return None

    def load_samples_for_bank(self, bank_name: str, max_samples: int = 5) -> List[Dict[str, Any]]:
        """加载指定银行的最近 N 个样本"""
        samples = []
        for meta in self.list_samples(bank_name)[:max_samples]:
            data = self.load_sample(meta['sample_path'])
            if data:
                samples.append(data)
        return samples

    def load_all_samples(self, max_per_bank: int = 3) -> List[Dict[str, Any]]:
        """加载所有银行的最近样本"""
        samples = []
        metadata = self._load_metadata()
        bank_counts: Dict[str, int] = {}
        for meta in metadata:
            bank = meta.get('bank_name', '')
            if bank_counts.get(bank, 0) >= max_per_bank:
                continue
            data = self.load_sample(meta['sample_path'])
            if data:
                samples.append(data)
                bank_counts[bank] = bank_counts.get(bank, 0) + 1
        return samples

    def import_excel_sample(
        self,
        excel_path: str,
        bank_name: str,
        bank_rule=None,
    ) -> Optional[str]:
        """
        从 Excel 文件导入样本（通过 bankcheck 解析）
        """
        try:
            from bankcheck import (
                open_workbook_compat,
                cleanup_temp_file,
                GenericBankParser,
                BankRule,
                BankRuleConfig,
                get_subject_info,
            )

            if bank_rule is None:
                config = BankRuleConfig()
                bank_rule = config.get_rule(bank_name)

            if bank_rule is None:
                self.logger.warning('未找到银行「%s」的规则配置', bank_name)
                return None

            parser = GenericBankParser(bank_rule)
            wb, tmp_path = open_workbook_compat(excel_path)
            try:
                all_records = []
                for ws in wb.worksheets:
                    if ws.title in (bank_rule.skip_sheets or []):
                        continue
                    try:
                        account_cell = ws[bank_rule.account_cell]
                        account_val = str(account_cell.value).strip() if account_cell.value else ''
                    except Exception:
                        account_val = ''
                    lookup_file = None
                    sheet_records = parser._parse_sheet(ws, excel_path, ws.title, lookup_file)
                    all_records.extend(sheet_records)
                wb.close()
            finally:
                cleanup_temp_file(tmp_path)

            if all_records:
                return self.save_sample(excel_path, bank_name, all_records)
            return None

        except Exception as e:
            self.logger.warning('从 Excel 导入样本失败: %s', e)
            return None


def _compare_records(
    old_records: List[Dict[str, Any]],
    new_records: List[Dict[str, Any]],
    id_fields: Optional[List[str]] = None,
) -> Tuple[List[RecordDiff], int, int, int, int]:
    """
    对比两组记录，返回差异详情

    Returns:
        (record_diffs, total, affected, added, removed)
    """
    if id_fields is None:
        id_fields = ['交易流水号', '交易日期', '付款', '收款', '摘要']

    def make_key(rec):
        parts = []
        for f in id_fields:
            v = _normalize_value(rec.get(f))
            parts.append('' if v is None else str(v))
        return '|'.join(parts)

    old_map = {make_key(r): r for r in old_records if make_key(r)}
    new_map = {make_key(r): r for r in new_records if make_key(r)}

    record_diffs: List[RecordDiff] = []
    added = 0
    removed = 0
    affected = 0

    all_keys = set(old_map.keys()) | set(new_map.keys())

    for key in all_keys:
        in_old = key in old_map
        in_new = key in new_map

        if in_old and not in_new:
            removed += 1
            continue

        if in_new and not in_old:
            added += 1
            continue

        old_rec = old_map[key]
        new_rec = new_map[key]

        field_diffs: List[FieldDiff] = []
        all_fields = set(old_rec.keys()) | set(new_rec.keys())

        compare_fields = [
            f for f in all_fields
            if f not in {'唯一id', '来源文件名', '来源相对路径', '处理时间'}
        ]

        for field in compare_fields:
            ov = old_rec.get(field)
            nv = new_rec.get(field)
            if not _values_equal(ov, nv):
                field_diffs.append(FieldDiff(
                    field_name=field,
                    old_value=ov,
                    new_value=nv,
                    record_id=key,
                ))

        if field_diffs:
            affected += 1
            record_diffs.append(RecordDiff(record_id=key, field_diffs=field_diffs))

    total = len(all_keys)
    unaffected = total - affected - added - removed

    return record_diffs, total, affected, unaffected, added, removed


class ChangeImpactEvaluator:
    """变更影响评估器"""

    def __init__(self, samples_dir=None):
        self.sample_manager = HistorySampleManager(samples_dir)
        self.logger = get_logger()

    def _build_diff_report(
        self,
        record_diffs: List[RecordDiff],
        total: int,
        affected: int,
        unaffected: int,
        added: int,
        removed: int,
        change_type: str,
        change_target: str,
    ) -> ImpactReport:
        field_diff_summary: Dict[str, int] = {}
        field_diff_examples: List[Dict[str, Any]] = []

        for rd in record_diffs:
            for fd in rd.field_diffs:
                field_diff_summary[fd.field_name] = field_diff_summary.get(fd.field_name, 0) + 1
                if len(field_diff_examples) < MAX_FIELD_DIFF_EXAMPLES * 3:
                    field_diff_examples.append({
                        'record_id': rd.record_id,
                        'field_name': fd.field_name,
                        'old_value': '' if fd.old_value is None else str(fd.old_value),
                        'new_value': '' if fd.new_value is None else str(fd.new_value),
                    })

        return ImpactReport(
            change_type=change_type,
            change_target=change_target,
            timestamp=datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            total_records=total,
            affected_records=affected,
            unaffected_records=unaffected,
            added_records=added,
            removed_records=removed,
            field_diff_summary=field_diff_summary,
            field_diff_examples=field_diff_examples,
        )

    def evaluate_bank_rule_change(
        self,
        bank_name: str,
        old_rule_data: Dict[str, Any],
        new_rule_data: Dict[str, Any],
        samples: Optional[List[Dict[str, Any]]] = None,
    ) -> ImpactReport:
        """
        评估银行列映射变更的影响

        Args:
            bank_name: 银行名称
            old_rule_data: 旧规则配置字典
            new_rule_data: 新规则配置字典
            samples: 指定样本列表，None 时自动加载历史样本

        Returns:
            ImpactReport 评估报告
        """
        self.logger.info('开始评估银行「%s」规则变更影响', bank_name)

        try:
            from bankcheck import (
                open_workbook_compat,
                cleanup_temp_file,
                GenericBankParser,
                BankRule,
                load_lookup_table,
            )
        except ImportError as e:
            self.logger.error('导入 bankcheck 模块失败: %s', e)
            return ImpactReport(
                change_type='bank_rule',
                change_target=bank_name,
                timestamp=datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                details=f'评估失败：无法导入 bankcheck 模块 ({e})',
            )

        if samples is None:
            samples = self.sample_manager.load_samples_for_bank(bank_name, max_samples=5)

        if not samples:
            return ImpactReport(
                change_type='bank_rule',
                change_target=bank_name,
                timestamp=datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                details='未找到该银行的历史样本，无法进行影响评估。建议先处理一些该银行的流水文件以积累样本。',
            )

        all_diffs: List[RecordDiff] = []
        total_all = 0
        affected_all = 0
        unaffected_all = 0
        added_all = 0
        removed_all = 0
        real_reparse_count = 0
        simulated_reparse_count = 0
        failed_reparse_count = 0

        for sample in samples:
            old_records = sample.get('records', [])
            if not old_records:
                continue

            try:
                new_rule = BankRule(
                    bank_name=new_rule_data.get('bank_name', bank_name),
                    account_cell=new_rule_data.get('account_cell', 'A1'),
                    start_row=int(new_rule_data.get('start_row', 1)),
                    columns={k: int(v) for k, v in new_rule_data.get('columns', {}).items()},
                    payment_sign=new_rule_data.get('payment_sign', 'negative'),
                    enabled=True,
                    skip_sheets=new_rule_data.get('skip_sheets', []),
                    expected_headers=new_rule_data.get('expected_headers', {}),
                    header_validation='off',
                    multi_account=bool(new_rule_data.get('multi_account', False)),
                )
                parser = GenericBankParser(new_rule)

                sample_full_path = sample.get('source_full_path', '')
                sample_hash = sample.get('source_hash', '')
                sample_src = sample.get('source_file', '')

                new_records = None
                used_real_reparse = False

                if sample_full_path and os.path.exists(sample_full_path):
                    try:
                        current_hash = self.sample_manager._file_hash(sample_full_path)
                        if sample_hash and current_hash != sample_hash:
                            self.logger.warning(
                                '样本文件 %s 哈希校验失败，文件可能已被篡改，使用模拟解析',
                                sample_full_path
                            )
                        else:
                            self.logger.info(
                                '使用原始样本文件真实重跑: %s', sample_full_path
                            )
                            lookup_data = load_lookup_table(None)
                            new_records = parser.parse(
                                sample_full_path, lookup_data, base_dir=None
                            )
                            used_real_reparse = True
                            real_reparse_count += 1
                    except Exception as parse_e:
                        self.logger.warning(
                            '真实重跑样本 %s 失败，降级使用模拟解析: %s',
                            sample_full_path, parse_e
                        )
                        new_records = None

                if new_records is None:
                    new_records = self._reparse_with_rule(
                        parser, new_rule, sample_src, old_records
                    )
                    if not used_real_reparse:
                        simulated_reparse_count += 1

                diffs, total, affected, unaffected, added, removed = _compare_records(
                    old_records, new_records
                )
                all_diffs.extend(diffs)
                total_all += total
                affected_all += affected
                unaffected_all += unaffected
                added_all += added
                removed_all += removed

            except Exception as e:
                failed_reparse_count += 1
                self.logger.warning('样本 %s 重新解析失败: %s', sample.get('source_file'), e)
                continue

        report = self._build_diff_report(
            all_diffs, total_all, affected_all, unaffected_all, added_all, removed_all,
            'bank_rule', bank_name
        )

        changed_fields = []
        old_cols = old_rule_data.get('columns', {})
        new_cols = new_rule_data.get('columns', {})
        all_col_keys = set(old_cols.keys()) | set(new_cols.keys())
        for k in sorted(all_col_keys):
            ov = old_cols.get(k)
            nv = new_cols.get(k)
            if ov != nv:
                changed_fields.append(f'{k}: {ov} → {nv}')

        details_parts = []
        if changed_fields:
            details_parts.append('列映射变更内容:\n' + '\n'.join(f'  - {f}' for f in changed_fields))
        reparse_info = (
            f'解析方式: 真实重跑 {real_reparse_count} 个样本，'
            f'模拟解析 {simulated_reparse_count} 个样本'
        )
        if failed_reparse_count > 0:
            reparse_info += f'，失败 {failed_reparse_count} 个样本'
        details_parts.append(reparse_info)
        report.details = '\n\n'.join(details_parts)

        self.logger.info(
            '银行「%s」规则变更评估完成: 总计 %d 条，受影响 %d 条（真实重跑 %d，模拟 %d）',
            bank_name, total_all, affected_all, real_reparse_count, simulated_reparse_count
        )
        return report

    def _reparse_with_rule(
        self,
        parser: 'GenericBankParser',
        rule: 'BankRule',
        sample_source_file: str,
        sample_records: List[Dict[str, Any]],
    ) -> List[Dict[str, Any]]:
        """
        使用新规则重新解析（降级模拟方案）。
        当原始 Excel 文件不存在或无法使用时，
        从旧记录中反向构造数据进行模拟解析。
        优先使用真实样本文件重跑（evaluate_bank_rule_change 中实现）。
        """
        if not sample_records:
            return []

        try:
            from bankcheck import get_subject_info, generate_unique_id
        except ImportError:
            return []

        columns = rule.columns
        payment_sign = rule.payment_sign

        new_records = []
        for old_rec in sample_records:
            try:
                account_val = old_rec.get('银行账号', '')
                subject_info = get_subject_info(account_val, None)
                subject = subject_info.get('subject', '')
                extra_fields = subject_info.get('extra_fields', {})

                new_rec = {
                    '唯一id': generate_unique_id(),
                    '银行': rule.bank_name,
                    '银行账号': account_val,
                    '主体': subject,
                }

                trade_date = old_rec.get('交易日期')
                if 'trade_date' in columns and trade_date is not None:
                    new_rec['交易日期'] = trade_date
                elif 'trade_date' in columns:
                    new_rec['交易日期'] = None

                if 'payment' in columns:
                    old_payment = old_rec.get('付款')
                    if old_payment is not None and payment_sign == 'negative':
                        new_rec['付款'] = -abs(float(old_payment)) if old_payment else None
                    else:
                        new_rec['付款'] = old_payment
                else:
                    new_rec['付款'] = None

                if 'receipt' in columns:
                    new_rec['收款'] = old_rec.get('收款')
                else:
                    new_rec['收款'] = None

                for field_key, rec_key in [
                    ('summary', '摘要'),
                    ('counterpart', '对方户名'),
                    ('balance', '余额'),
                    ('transaction_id', '交易流水号'),
                ]:
                    if field_key in columns:
                        new_rec[rec_key] = old_rec.get(rec_key)
                    else:
                        new_rec[rec_key] = None

                for key, val in extra_fields.items():
                    new_rec[key] = val

                new_records.append(new_rec)
            except Exception:
                new_records.append(dict(old_rec))

        return new_records

    def evaluate_lookup_change(
        self,
        old_lookup_file: Optional[str],
        new_lookup_entries: List[Any],
        samples: Optional[List[Dict[str, Any]]] = None,
        max_samples: int = 10,
    ) -> ImpactReport:
        """
        评估查找表变更的影响

        Args:
            old_lookup_file: 旧查找表文件路径（用于读取旧条目）
            new_lookup_entries: 新的查找表条目列表（LookupEntry 对象）
            samples: 指定样本列表，None 时自动加载所有历史样本
            max_samples: 最多使用多少个样本

        Returns:
            ImpactReport 评估报告
        """
        self.logger.info('开始评估查找表变更影响')

        try:
            from lookup_manager import (
                read_lookup_entries,
                get_subject_info,
                LookupEntry,
                save_lookup_entries,
                _account_key,
            )
        except ImportError as e:
            self.logger.error('导入 lookup_manager 模块失败: %s', e)
            return ImpactReport(
                change_type='lookup_table',
                change_target='主体查找表',
                timestamp=datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                details=f'评估失败：无法导入 lookup_manager 模块 ({e})',
            )

        if samples is None:
            samples = self.sample_manager.load_all_samples(max_per_bank=3)
            samples = samples[:max_samples]

        if not samples:
            return ImpactReport(
                change_type='lookup_table',
                change_target='主体查找表',
                timestamp=datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                details='未找到历史样本，无法进行影响评估。建议先处理一些银行流水文件以积累样本。',
            )

        old_entries = []
        if old_lookup_file and os.path.exists(old_lookup_file):
            old_entries = read_lookup_entries(old_lookup_file)

        old_account_map = {}
        for e in old_entries:
            key = _account_key(e.account)
            if key not in old_account_map or old_account_map[key].priority < e.priority:
                old_account_map[key] = e

        new_account_map = {}
        for e in new_lookup_entries:
            key = _account_key(e.account)
            if key not in new_account_map or new_account_map[key].priority < e.priority:
                new_account_map[key] = e

        all_diffs: List[RecordDiff] = []
        total_all = 0
        affected_all = 0
        unaffected_all = 0
        added_all = 0
        removed_all = 0

        for sample in samples:
            old_records = sample.get('records', [])
            if not old_records:
                continue

            new_records = []
            for rec in old_records:
                new_rec = dict(rec)
                account = rec.get('银行账号', '')
                key = _account_key(account)

                new_entry = new_account_map.get(key)
                if new_entry:
                    new_rec['主体'] = new_entry.subject
                    if new_entry.extra_fields:
                        for k, v in new_entry.extra_fields.items():
                            new_rec[k] = v
                else:
                    new_rec['主体'] = ''

                new_records.append(new_rec)

            diffs, total, affected, unaffected, added, removed = _compare_records(
                old_records, new_records
            )
            all_diffs.extend(diffs)
            total_all += total
            affected_all += affected
            unaffected_all += unaffected
            added_all += added
            removed_all += removed

        report = self._build_diff_report(
            all_diffs, total_all, affected_all, unaffected_all, added_all, removed_all,
            'lookup_table', '主体查找表'
        )

        changed_accounts = []
        all_keys = set(old_account_map.keys()) | set(new_account_map.keys())
        for key in sorted(all_keys):
            old_e = old_account_map.get(key)
            new_e = new_account_map.get(key)

            old_subject = old_e.subject if old_e else '(未设置)'
            new_subject = new_e.subject if new_e else '(已删除)'
            acct = new_e.account if new_e else (old_e.account if old_e else key)

            if old_subject != new_subject:
                changed_accounts.append(f'{acct}: {old_subject} → {new_subject}')

        if changed_accounts:
            details_lines = ['查找表变更内容（前10条）:']
            for line in changed_accounts[:10]:
                details_lines.append(f'  - {line}')
            if len(changed_accounts) > 10:
                details_lines.append(f'  ... 另有 {len(changed_accounts) - 10} 个账号变更')
            report.details = '\n'.join(details_lines)

        self.logger.info(
            '查找表变更评估完成: 总计 %d 条，受影响 %d 条',
            total_all, affected_all
        )
        return report

    def save_sample_from_records(
        self,
        source_filepath: str,
        bank_name: str,
        records: List[Dict[str, Any]],
    ) -> Optional[str]:
        """保存处理结果为历史样本（便捷方法）"""
        return self.sample_manager.save_sample(source_filepath, bank_name, records)
