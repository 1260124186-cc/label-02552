# -*- coding: utf-8 -*-
"""
智能排障助手模块
功能：
  1. 读取最近一次 bankcheck.log 日志文件
  2. 结合 ProcessingResult 处理结果
  3. 使用规则引擎自动归纳最可能的故障原因
  4. 给出详细的修复步骤建议

支持的诊断类型：
  - 查找表缺失
  - 文件名不符（无法识别银行类型）
  - xls 文件缺少依赖库
  - 文件格式错误
  - 配置文件缺失
  - 空文件夹
  - 表头不匹配
"""

import os
import sys
import re
import logging
from dataclasses import dataclass, field
from typing import List, Optional, Dict, Any, Tuple
from enum import Enum
from datetime import datetime


try:
    from bank_knowledge_base import (
        get_knowledge_base,
        diagnose_from_knowledge_base,
        SearchResult,
    )
    HAS_KNOWLEDGE_BASE = True
except ImportError:
    HAS_KNOWLEDGE_BASE = False
    get_knowledge_base = None
    diagnose_from_knowledge_base = None
    SearchResult = None


def get_script_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def get_logger():
    return logging.getLogger('bankcheck')


class SeverityLevel(Enum):
    CRITICAL = 'critical'
    ERROR = 'error'
    WARNING = 'warning'
    INFO = 'info'


@dataclass
class TroubleshootingIssue:
    issue_id: str
    title: str
    severity: SeverityLevel
    description: str
    evidence: List[str] = field(default_factory=list)
    fix_steps: List[str] = field(default_factory=list)
    confidence: float = 0.0

    def to_dict(self) -> Dict[str, Any]:
        return {
            'issue_id': self.issue_id,
            'title': self.title,
            'severity': self.severity.value,
            'description': self.description,
            'evidence': self.evidence,
            'fix_steps': self.fix_steps,
            'confidence': round(self.confidence, 2),
        }


@dataclass
class TroubleshootingReport:
    timestamp: str = ''
    log_file: str = ''
    total_issues: int = 0
    issues: List[TroubleshootingIssue] = field(default_factory=list)
    summary: str = ''
    knowledge_base_matches: List[Dict[str, Any]] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        return {
            'timestamp': self.timestamp,
            'log_file': self.log_file,
            'total_issues': self.total_issues,
            'issues': [issue.to_dict() for issue in self.issues],
            'summary': self.summary,
            'knowledge_base_matches': self.knowledge_base_matches,
        }


class RuleEngine:
    def __init__(self):
        self.rules: List[Dict[str, Any]] = []
        self._register_rules()

    def _register_rules(self):
        self.rules = [
            {
                'id': 'lookup_missing',
                'title': '主体查找表缺失',
                'severity': SeverityLevel.WARNING,
                'check_func': self._check_lookup_missing,
            },
            {
                'id': 'unrecognized_files',
                'title': '文件名不符，无法识别银行类型',
                'severity': SeverityLevel.WARNING,
                'check_func': self._check_unrecognized_files,
            },
            {
                'id': 'xls_missing_dependency',
                'title': 'XLS 文件缺少依赖库',
                'severity': SeverityLevel.ERROR,
                'check_func': self._check_xls_dependency,
            },
            {
                'id': 'empty_folder',
                'title': '处理文件夹为空',
                'severity': SeverityLevel.WARNING,
                'check_func': self._check_empty_folder,
            },
            {
                'id': 'config_missing',
                'title': '银行规则配置文件缺失',
                'severity': SeverityLevel.ERROR,
                'check_func': self._check_config_missing,
            },
            {
                'id': 'file_format_error',
                'title': '文件格式错误或损坏',
                'severity': SeverityLevel.ERROR,
                'check_func': self._check_file_format_error,
            },
            {
                'id': 'header_mismatch',
                'title': '银行模板表头不匹配',
                'severity': SeverityLevel.WARNING,
                'check_func': self._check_header_mismatch,
            },
            {
                'id': 'permission_error',
                'title': '文件权限不足',
                'severity': SeverityLevel.ERROR,
                'check_func': self._check_permission_error,
            },
            {
                'id': 'knowledge_base_match',
                'title': '知识库已知问题匹配',
                'severity': SeverityLevel.INFO,
                'check_func': self._check_knowledge_base,
            },
        ]

    def _check_lookup_missing(self, log_lines: List[str],
                               processing_result: Optional[Any]) -> Optional[TroubleshootingIssue]:
        evidence = []
        confidence = 0.0

        if processing_result is not None and getattr(processing_result, 'lookup_missing', False):
            confidence += 0.6
            evidence.append('ProcessingResult.lookup_missing = True')

        lookup_patterns = [
            r'未找到主体查找表',
            r'主体查找表.*不存在',
            r'查找表.*为空',
            r'lookup.*missing',
        ]
        for line in log_lines:
            for pattern in lookup_patterns:
                if re.search(pattern, line, re.IGNORECASE):
                    confidence += 0.2
                    if len(evidence) < 5:
                        evidence.append(line.strip())
                    break

        if confidence < 0.1:
            return None

        confidence = min(confidence, 1.0)

        return TroubleshootingIssue(
            issue_id='lookup_missing',
            title='主体查找表缺失',
            severity=SeverityLevel.WARNING,
            description='程序未找到主体查找表文件，导致"主体"列将为空。查找表用于将银行账号映射到对应的主体名称。',
            evidence=evidence,
            fix_steps=[
                '确认主体查找表文件已准备好',
                '将查找表文件命名为"主体查找表.xlsx"（推荐）或"主体查找表.xls"',
                '将查找表文件放置在程序所在目录下',
                '重新运行流水处理程序',
                '若查找表文件名称不同，可重命名为标准名称或在程序目录下只保留这一个Excel文件',
            ],
            confidence=confidence,
        )

    def _check_unrecognized_files(self, log_lines: List[str],
                                    processing_result: Optional[Any]) -> Optional[TroubleshootingIssue]:
        evidence = []
        confidence = 0.0
        unprocessed_count = 0

        if processing_result is not None:
            unprocessed_files = getattr(processing_result, 'unprocessed_files', [])
            unprocessed_count = len(unprocessed_files)
            if unprocessed_count > 0:
                confidence += 0.5
                evidence.append(f'ProcessingResult.unprocessed_files 包含 {unprocessed_count} 个文件')
                for f in unprocessed_files[:3]:
                    evidence.append(f'  - {os.path.basename(f)}')

        unrec_patterns = [
            r'无法识别银行类型',
            r'unrecognized.*bank',
            r'cannot.*identify.*bank',
        ]
        for line in log_lines:
            for pattern in unrec_patterns:
                if re.search(pattern, line, re.IGNORECASE):
                    confidence += 0.15
                    if len(evidence) < 8:
                        evidence.append(line.strip())
                    break

        if confidence < 0.1:
            return None

        confidence = min(confidence, 1.0)

        return TroubleshootingIssue(
            issue_id='unrecognized_files',
            title=f'文件名不符，无法识别银行类型（{unprocessed_count} 个文件）',
            severity=SeverityLevel.WARNING,
            description='部分Excel文件的文件名不符合银行识别规则，导致程序无法确定其所属银行类型，这些文件将被保留不处理。',
            evidence=evidence,
            fix_steps=[
                '检查无法识别的文件名，确认银行名称',
                '将文件重命名为以银行名称开头的格式，如："北京银行_流水.xlsx"',
                '支持的银行名称可查看 bank_rules.yaml 配置文件',
                '或在 bank_rules.yaml 中添加新的银行解析规则',
                '重新运行流水处理程序',
            ],
            confidence=confidence,
        )

    def _check_xls_dependency(self, log_lines: List[str],
                                processing_result: Optional[Any]) -> Optional[TroubleshootingIssue]:
        evidence = []
        confidence = 0.0

        xls_patterns = [
            r'缺少 xlrd 库',
            r'xlrd.*import',
            r'处理.*\.xls.*需要.*xlrd',
            r'ImportError.*xlrd',
            r'No module named.*xlrd',
        ]
        for line in log_lines:
            for pattern in xls_patterns:
                if re.search(pattern, line, re.IGNORECASE):
                    confidence += 0.3
                    if len(evidence) < 5:
                        evidence.append(line.strip())
                    break

        if processing_result is not None:
            error_files = getattr(processing_result, 'error_files', [])
            for filepath, error_msg in error_files:
                if '.xls' in filepath.lower() and ('xlrd' in error_msg.lower() or 'import' in error_msg.lower()):
                    confidence += 0.3
                    evidence.append(f'错误文件: {os.path.basename(filepath)} - {error_msg}')
                    break

        if confidence < 0.1:
            return None

        confidence = min(confidence, 1.0)

        return TroubleshootingIssue(
            issue_id='xls_missing_dependency',
            title='XLS 文件缺少依赖库',
            severity=SeverityLevel.ERROR,
            description='处理 .xls 格式的 Excel 文件需要 xlrd 库，但当前环境中未安装该库。',
            evidence=evidence,
            fix_steps=[
                '打开命令行终端',
                '运行命令安装 xlrd 库: pip install xlrd==1.2.0',
                '注意：xlrd 2.0+ 版本不再支持 xls 格式，请安装 1.2.0 版本',
                '如果使用虚拟环境，请确保在正确的环境中安装',
                '安装完成后重新运行程序',
                '或者将 .xls 文件另存为 .xlsx 格式后再处理',
            ],
            confidence=confidence,
        )

    def _check_empty_folder(self, log_lines: List[str],
                             processing_result: Optional[Any]) -> Optional[TroubleshootingIssue]:
        evidence = []
        confidence = 0.0

        if processing_result is not None and getattr(processing_result, 'folder_empty', False):
            confidence += 0.6
            evidence.append('ProcessingResult.folder_empty = True')

        empty_patterns = [
            r'文件夹为空',
            r'no.*excel.*file',
            r'未找到.*Excel',
            r'empty.*folder',
        ]
        for line in log_lines:
            for pattern in empty_patterns:
                if re.search(pattern, line, re.IGNORECASE):
                    confidence += 0.2
                    if len(evidence) < 5:
                        evidence.append(line.strip())
                    break

        if confidence < 0.1:
            return None

        confidence = min(confidence, 1.0)

        return TroubleshootingIssue(
            issue_id='empty_folder',
            title='处理文件夹为空',
            severity=SeverityLevel.WARNING,
            description='选择的处理文件夹中没有找到任何 Excel 文件，程序无法进行处理。',
            evidence=evidence,
            fix_steps=[
                '确认选择的文件夹路径正确',
                '检查文件夹中是否包含 .xlsx 或 .xls 格式的银行流水文件',
                '确认文件没有被隐藏或被其他程序占用',
                '确认文件扩展名正确（不是 .csv 或其他格式）',
                '重新选择正确的文件夹后再次运行',
            ],
            confidence=confidence,
        )

    def _check_config_missing(self, log_lines: List[str],
                                processing_result: Optional[Any]) -> Optional[TroubleshootingIssue]:
        evidence = []
        confidence = 0.0

        config_patterns = [
            r'银行规则配置文件不存在',
            r'bank_rules.*yaml.*not.*found',
            r'配置文件.*不存在',
            r'config.*missing',
        ]
        for line in log_lines:
            for pattern in config_patterns:
                if re.search(pattern, line, re.IGNORECASE):
                    confidence += 0.3
                    if len(evidence) < 5:
                        evidence.append(line.strip())
                    break

        script_dir = get_script_dir()
        config_path = os.path.join(script_dir, 'bank_rules.yaml')
        if not os.path.exists(config_path):
            confidence += 0.4
            evidence.append(f'配置文件不存在: {config_path}')

        if confidence < 0.1:
            return None

        confidence = min(confidence, 1.0)

        return TroubleshootingIssue(
            issue_id='config_missing',
            title='银行规则配置文件缺失',
            severity=SeverityLevel.ERROR,
            description='银行解析规则配置文件 bank_rules.yaml 不存在，程序无法进行银行流水解析。',
            evidence=evidence,
            fix_steps=[
                '确认 bank_rules.yaml 文件是否在程序目录下',
                '如果文件被误删，可以从备份或安装包中恢复',
                '确认配置文件的文件名是否正确（区分大小写）',
                '如果需要，可以手动创建 bank_rules.yaml 并配置银行规则',
                '配置文件格式请参考项目文档',
            ],
            confidence=confidence,
        )

    def _check_file_format_error(self, log_lines: List[str],
                                   processing_result: Optional[Any]) -> Optional[TroubleshootingIssue]:
        evidence = []
        confidence = 0.0
        error_count = 0

        format_patterns = [
            r'文件格式错误',
            r'format.*error',
            r'无法打开.*文件',
            r'损坏',
            r'corrupt',
            r'BadZipFile',
            r'InvalidFileException',
            r'openpyxl.*exceptions',
        ]
        for line in log_lines:
            for pattern in format_patterns:
                if re.search(pattern, line, re.IGNORECASE):
                    confidence += 0.15
                    error_count += 1
                    if len(evidence) < 5:
                        evidence.append(line.strip())
                    break

        if processing_result is not None:
            error_files = getattr(processing_result, 'error_files', [])
            for filepath, error_msg in error_files:
                error_lower = error_msg.lower()
                if any(kw in error_lower for kw in ['format', 'corrupt', 'invalid', 'badzip', 'open']):
                    if '.xls' in filepath.lower() or '.xlsx' in filepath.lower():
                        confidence += 0.15
                        evidence.append(f'错误文件: {os.path.basename(filepath)} - {error_msg}')
                        break

        if confidence < 0.1:
            return None

        confidence = min(confidence, 1.0)

        return TroubleshootingIssue(
            issue_id='file_format_error',
            title=f'文件格式错误或损坏（{error_count} 处错误）',
            severity=SeverityLevel.ERROR,
            description='部分 Excel 文件格式错误或文件损坏，导致程序无法正常读取。',
            evidence=evidence,
            fix_steps=[
                '确认文件是有效的 Excel 文件，不是其他格式改后缀而来',
                '尝试用 Excel 打开文件，确认文件可以正常打开',
                '如果文件损坏，尝试从原始来源重新获取',
                '如果是 .xls 文件，确认安装了正确版本的 xlrd 库',
                '可以尝试将文件另存为 .xlsx 格式后再处理',
                '检查文件是否被其他程序占用（如正在编辑中）',
            ],
            confidence=confidence,
        )

    def _check_header_mismatch(self, log_lines: List[str],
                                 processing_result: Optional[Any]) -> Optional[TroubleshootingIssue]:
        evidence = []
        confidence = 0.0

        header_patterns = [
            r'表头.*不匹配',
            r'header.*mismatch',
            r'expected.*header',
            r'表头校验',
            r'银行模板.*变更',
        ]
        for line in log_lines:
            for pattern in header_patterns:
                if re.search(pattern, line, re.IGNORECASE):
                    confidence += 0.2
                    if len(evidence) < 5:
                        evidence.append(line.strip())
                    break

        if confidence < 0.1:
            return None

        confidence = min(confidence, 1.0)

        return TroubleshootingIssue(
            issue_id='header_mismatch',
            title='银行模板表头不匹配',
            severity=SeverityLevel.WARNING,
            description='检测到银行流水的表头与预期不符，可能是银行更新了导出模板。',
            evidence=evidence,
            fix_steps=[
                '打开银行流水文件，核对各列的表头名称',
                '对比 bank_rules.yaml 中配置的 expected_headers',
                '如果银行模板已更新，修改 bank_rules.yaml 中的列映射',
                '如果只是表头文字微调，可以在 expected_headers 中添加可接受的备选名称',
                '修改配置后重新运行程序',
            ],
            confidence=confidence,
        )

    def _check_permission_error(self, log_lines: List[str],
                                  processing_result: Optional[Any]) -> Optional[TroubleshootingIssue]:
        evidence = []
        confidence = 0.0

        perm_patterns = [
            r'权限不足',
            r'permission.*denied',
            r'PermissionError',
            r'无法写入',
            r'无法读取',
            r'access.*denied',
        ]
        for line in log_lines:
            for pattern in perm_patterns:
                if re.search(pattern, line, re.IGNORECASE):
                    confidence += 0.25
                    if len(evidence) < 5:
                        evidence.append(line.strip())
                    break

        if confidence < 0.1:
            return None

        confidence = min(confidence, 1.0)

        return TroubleshootingIssue(
            issue_id='permission_error',
            title='文件权限不足',
            severity=SeverityLevel.ERROR,
            description='程序在读取或写入文件时遇到权限不足的问题。',
            evidence=evidence,
            fix_steps=[
                '确认您有读取输入文件夹的权限',
                '确认您有写入输出目录的权限',
                '在 Windows 上，尝试以管理员身份运行程序',
                '在 macOS/Linux 上，检查文件和目录的读写权限',
                '确认文件没有被其他程序以独占方式打开',
                '如果文件在受保护的系统目录中，将其移动到用户目录下再处理',
            ],
            confidence=confidence,
        )

    def _check_knowledge_base(self, log_lines: List[str],
                               processing_result: Optional[Any]) -> Optional[TroubleshootingIssue]:
        if not HAS_KNOWLEDGE_BASE or diagnose_from_knowledge_base is None:
            return None

        bank_name = None
        if processing_result is not None:
            bank_name = getattr(processing_result, 'bank_name', None)
            if not bank_name:
                bank_name = getattr(processing_result, 'last_bank_name', None)

        try:
            results = diagnose_from_knowledge_base(log_lines, bank_name)
        except Exception:
            return None

        if not results:
            return None

        evidence = []
        fix_steps = []
        best_score = 0.0

        for r in results[:5]:
            best_score = max(best_score, r.match_score)
            evidence.append(
                f'[{r.bank_name}] {r.issue.title} (匹配度:{int(r.match_score * 100)}%)'
            )
            if r.issue.root_cause:
                evidence.append(f'  根因: {r.issue.root_cause}')
            for step_idx, step in enumerate(r.issue.fix_steps, 1):
                step_text = f'[{r.bank_name}] {step}'
                if step_text not in fix_steps:
                    fix_steps.append(step_text)

        if not evidence:
            return None

        confidence = min(best_score * 0.8 + 0.2, 1.0)

        return TroubleshootingIssue(
            issue_id='knowledge_base_match',
            title=f'知识库匹配到 {len(results)} 个已知问题',
            severity=SeverityLevel.INFO,
            description='从银行格式知识库中检索到与当前错误匹配的已知问题和解决方案。',
            evidence=evidence,
            fix_steps=fix_steps,
            confidence=confidence,
        )

    def analyze(self, log_lines: List[str],
                processing_result: Optional[Any] = None) -> List[TroubleshootingIssue]:
        issues = []
        for rule in self.rules:
            try:
                issue = rule['check_func'](log_lines, processing_result)
                if issue is not None:
                    issues.append(issue)
            except Exception as e:
                logger = get_logger()
                logger.warning('规则 %s 执行失败: %s', rule['id'], e)

        issues.sort(key=lambda x: (
            {'critical': 0, 'error': 1, 'warning': 2, 'info': 3}[x.severity.value],
            -x.confidence,
        ))

        return issues


def read_log_file(log_path: str, max_lines: int = 2000) -> List[str]:
    if not os.path.exists(log_path):
        return []

    try:
        with open(log_path, 'r', encoding='utf-8', errors='ignore') as f:
            lines = f.readlines()
        if len(lines) > max_lines:
            lines = lines[-max_lines:]
        return lines
    except Exception as e:
        logger = get_logger()
        logger.error('读取日志文件失败 %s: %s', log_path, e)
        return []


def find_latest_log_file(script_dir: Optional[str] = None) -> Optional[str]:
    if script_dir is None:
        script_dir = get_script_dir()

    log_dir = os.path.join(script_dir, 'logs')
    if os.path.isdir(log_dir):
        import re
        pattern = re.compile(r'^bankcheck_\d{8}_\d{6}\.log$')
        candidates = []
        try:
            for filename in os.listdir(log_dir):
                if pattern.match(filename):
                    filepath = os.path.join(log_dir, filename)
                    try:
                        mtime = os.path.getmtime(filepath)
                        candidates.append((mtime, filepath))
                    except OSError:
                        continue
            if candidates:
                candidates.sort(key=lambda x: x[0], reverse=True)
                return candidates[0][1]
        except OSError:
            pass

    try:
        import bankcheck as bc
        log_dir_bc = bc.get_log_dir(script_dir)
        latest = bc.find_latest_log_file_in_dir(log_dir_bc, prefix='bankcheck')
        if latest:
            return latest
        current = bc.get_current_log_file()
        if current and os.path.exists(current):
            current_parent = os.path.dirname(os.path.dirname(current)) if 'logs' in current else os.path.dirname(current)
            script_dir_abs = os.path.abspath(script_dir)
            current_parent_abs = os.path.abspath(current_parent)
            if current_parent_abs == script_dir_abs:
                return current
    except (ImportError, AttributeError):
        pass

    log_file = os.path.join(script_dir, 'bankcheck.log')
    if os.path.exists(log_file):
        return log_file

    for root, dirs, files in os.walk(script_dir):
        for f in files:
            if f == 'bankcheck.log' or (f.startswith('bankcheck_') and f.endswith('.log')):
                return os.path.join(root, f)

    return None


def get_latest_processing_result(script_dir: Optional[str] = None) -> Optional[Any]:
    try:
        import batch_manager as bm
        manager = bm.get_batch_manager(script_dir)
        batches = manager.query_batches(limit=1)
        if batches:
            latest_batch = batches[0]
            detail = manager.get_batch_detail(latest_batch.batch_id)
            if detail and detail.get('metadata'):
                return detail['metadata']
    except Exception as e:
        logger = get_logger()
        logger.debug('获取最近处理结果失败: %s', e)
    return None


def run_troubleshooting(log_path: Optional[str] = None,
                        processing_result: Optional[Any] = None,
                        script_dir: Optional[str] = None) -> TroubleshootingReport:
    logger = get_logger()
    logger.info('开始智能排障分析')

    if script_dir is None:
        script_dir = get_script_dir()

    if log_path is None:
        log_path = find_latest_log_file(script_dir)

    log_lines = []
    if log_path and os.path.exists(log_path):
        log_lines = read_log_file(log_path)
        logger.info('读取日志文件: %s (%d 行)', log_path, len(log_lines))
    else:
        logger.warning('未找到日志文件')

    if processing_result is None:
        processing_result = get_latest_processing_result(script_dir)

    engine = RuleEngine()
    issues = engine.analyze(log_lines, processing_result)

    kb_matches = []
    if HAS_KNOWLEDGE_BASE and diagnose_from_knowledge_base is not None and log_lines:
        try:
            bank_name = None
            if processing_result is not None:
                bank_name = getattr(processing_result, 'bank_name', None)
                if not bank_name:
                    bank_name = getattr(processing_result, 'last_bank_name', None)
            kb_results = diagnose_from_knowledge_base(log_lines, bank_name)
            kb_matches = [r.to_dict() for r in kb_results]
        except Exception as e:
            logger.warning('知识库检索失败: %s', e)

    report = TroubleshootingReport(
        timestamp=datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        log_file=log_path or '',
        total_issues=len(issues),
        issues=issues,
        knowledge_base_matches=kb_matches,
    )

    if not issues:
        report.summary = '未检测到明显问题，系统运行正常。'
    else:
        critical_count = sum(1 for i in issues if i.severity == SeverityLevel.CRITICAL)
        error_count = sum(1 for i in issues if i.severity == SeverityLevel.ERROR)
        warning_count = sum(1 for i in issues if i.severity == SeverityLevel.WARNING)
        report.summary = (
            f'检测到 {len(issues)} 个问题：'
            f'{critical_count} 个严重，'
            f'{error_count} 个错误，'
            f'{warning_count} 个警告。'
            f'最可能的原因：{issues[0].title}'
        )

    logger.info('排障分析完成，发现 %d 个问题', len(issues))
    return report


def print_report(report: TroubleshootingReport):
    print('\n' + '=' * 60)
    print('  智能排障助手 - 诊断报告')
    print('=' * 60)
    print(f'  分析时间: {report.timestamp}')
    print(f'  日志文件: {report.log_file or "未找到"}')
    print(f'  问题数量: {report.total_issues}')
    print('-' * 60)

    if not report.issues:
        print('  ✅ 未检测到明显问题')
    else:
        for idx, issue in enumerate(report.issues, 1):
            severity_icon = {
                'critical': '🔴',
                'error': '🔴',
                'warning': '🟡',
                'info': '🔵',
            }.get(issue.severity.value, '⚪')

            print(f'\n  {idx}. {severity_icon} {issue.title}')
            print(f'     置信度: {int(issue.confidence * 100)}%')
            print(f'     描述: {issue.description}')

            if issue.evidence:
                print(f'     证据:')
                for ev in issue.evidence[:3]:
                    print(f'       - {ev[:80]}')

            if issue.fix_steps:
                print(f'     修复步骤:')
                for step_idx, step in enumerate(issue.fix_steps, 1):
                    print(f'       {step_idx}. {step}')

    print('\n' + '-' * 60)
    print(f'  总结: {report.summary}')

    if report.knowledge_base_matches:
        print('\n' + '-' * 60)
        print(f'  📚 知识库匹配 ({len(report.knowledge_base_matches)} 条)')
        print('-' * 60)
        for idx, match in enumerate(report.knowledge_base_matches[:5], 1):
            issue_data = match.get('issue', {})
            print(f'\n  {idx}. [{match.get("bank_name", "")}] {issue_data.get("title", "")}')
            print(f'     匹配度: {int(match.get("match_score", 0) * 100)}%')
            if issue_data.get('root_cause'):
                print(f'     根因: {issue_data["root_cause"]}')
            if issue_data.get('fix_steps'):
                print(f'     修复步骤:')
                for step_idx, step in enumerate(issue_data['fix_steps'], 1):
                    print(f'       {step_idx}. {step}')

    print('=' * 60 + '\n')
