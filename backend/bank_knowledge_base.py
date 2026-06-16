# -*- coding: utf-8 -*-
"""
银行格式知识库模块
==================
以 Wiki 结构维护各银行已验证的模板截图、列说明、常见坑与对应配置版本，
处理失败时可检索匹配已知问题与解法。

核心功能：
  1. 按 Wiki 结构组织银行格式知识（模板截图、列说明、常见坑、配置版本）
  2. 支持 YAML 持久化存储，支持热更新
  3. 处理失败时通过关键词/错误类型检索匹配已知问题与解法
  4. 支持按银行名称、错误类型、关键词搜索
  5. 与 troubleshooter 集成，自动匹配知识库中的已知问题
"""

import os
import sys
import re
import logging
import hashlib
from dataclasses import dataclass, field, asdict
from datetime import datetime
from typing import List, Optional, Dict, Any, Tuple


def get_script_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def get_logger():
    return logging.getLogger('bankcheck')


KNOWLEDGE_BASE_FILE = 'bank_knowledge.yaml'


@dataclass
class ColumnDescription:
    field_key: str
    display_name: str
    description: str = ''
    example_values: List[str] = field(default_factory=list)
    pitfalls: List[str] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        return {
            'field_key': self.field_key,
            'display_name': self.display_name,
            'description': self.description,
            'example_values': list(self.example_values),
            'pitfalls': list(self.pitfalls),
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'ColumnDescription':
        return cls(
            field_key=data.get('field_key', ''),
            display_name=data.get('display_name', ''),
            description=data.get('description', ''),
            example_values=data.get('example_values', []),
            pitfalls=data.get('pitfalls', []),
        )


@dataclass
class KnownIssue:
    issue_id: str
    title: str
    description: str = ''
    symptoms: List[str] = field(default_factory=list)
    root_cause: str = ''
    fix_steps: List[str] = field(default_factory=list)
    error_patterns: List[str] = field(default_factory=list)
    severity: str = 'warning'
    config_version: str = ''
    last_seen_date: str = ''
    tags: List[str] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        return {
            'issue_id': self.issue_id,
            'title': self.title,
            'description': self.description,
            'symptoms': list(self.symptoms),
            'root_cause': self.root_cause,
            'fix_steps': list(self.fix_steps),
            'error_patterns': list(self.error_patterns),
            'severity': self.severity,
            'config_version': self.config_version,
            'last_seen_date': self.last_seen_date,
            'tags': list(self.tags),
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'KnownIssue':
        return cls(
            issue_id=data.get('issue_id', ''),
            title=data.get('title', ''),
            description=data.get('description', ''),
            symptoms=data.get('symptoms', []),
            root_cause=data.get('root_cause', ''),
            fix_steps=data.get('fix_steps', []),
            error_patterns=data.get('error_patterns', []),
            severity=data.get('severity', 'warning'),
            config_version=data.get('config_version', ''),
            last_seen_date=data.get('last_seen_date', ''),
            tags=data.get('tags', []),
        )


@dataclass
class TemplateScreenshot:
    name: str
    description: str = ''
    file_path: str = ''
    upload_date: str = ''
    config_version: str = ''
    notes: str = ''

    def to_dict(self) -> Dict[str, Any]:
        return {
            'name': self.name,
            'description': self.description,
            'file_path': self.file_path,
            'upload_date': self.upload_date,
            'config_version': self.config_version,
            'notes': self.notes,
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'TemplateScreenshot':
        return cls(
            name=data.get('name', ''),
            description=data.get('description', ''),
            file_path=data.get('file_path', ''),
            upload_date=data.get('upload_date', ''),
            config_version=data.get('config_version', ''),
            notes=data.get('notes', ''),
        )


@dataclass
class BankWikiEntry:
    bank_name: str
    display_name: str = ''
    description: str = ''
    config_version: str = '1.0'
    last_verified_date: str = ''
    verified_by: str = ''
    template_screenshots: List[TemplateScreenshot] = field(default_factory=list)
    column_descriptions: List[ColumnDescription] = field(default_factory=list)
    known_issues: List[KnownIssue] = field(default_factory=list)
    general_pitfalls: List[str] = field(default_factory=list)
    notes: str = ''
    tags: List[str] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        return {
            'bank_name': self.bank_name,
            'display_name': self.display_name,
            'description': self.description,
            'config_version': self.config_version,
            'last_verified_date': self.last_verified_date,
            'verified_by': self.verified_by,
            'template_screenshots': [s.to_dict() for s in self.template_screenshots],
            'column_descriptions': [c.to_dict() for c in self.column_descriptions],
            'known_issues': [i.to_dict() for i in self.known_issues],
            'general_pitfalls': list(self.general_pitfalls),
            'notes': self.notes,
            'tags': list(self.tags),
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'BankWikiEntry':
        return cls(
            bank_name=data.get('bank_name', ''),
            display_name=data.get('display_name', ''),
            description=data.get('description', ''),
            config_version=data.get('config_version', '1.0'),
            last_verified_date=data.get('last_verified_date', ''),
            verified_by=data.get('verified_by', ''),
            template_screenshots=[
                TemplateScreenshot.from_dict(s) for s in data.get('template_screenshots', [])
            ],
            column_descriptions=[
                ColumnDescription.from_dict(c) for c in data.get('column_descriptions', [])
            ],
            known_issues=[
                KnownIssue.from_dict(i) for i in data.get('known_issues', [])
            ],
            general_pitfalls=data.get('general_pitfalls', []),
            notes=data.get('notes', ''),
            tags=data.get('tags', []),
        )


@dataclass
class SearchResult:
    bank_name: str
    issue: KnownIssue
    match_score: float
    matched_patterns: List[str] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        return {
            'bank_name': self.bank_name,
            'issue': self.issue.to_dict(),
            'match_score': round(self.match_score, 2),
            'matched_patterns': self.matched_patterns,
        }


class BankKnowledgeBase:
    _instance = None
    _knowledge_path = None
    _entries: Dict[str, BankWikiEntry] = field(default_factory=dict)
    _last_modified: float = 0.0

    def __new__(cls, knowledge_path=None):
        if cls._instance is None:
            cls._instance = super(BankKnowledgeBase, cls).__new__(cls)
            cls._instance._initialized = False
        return cls._instance

    def __init__(self, knowledge_path=None):
        if self._initialized:
            return
        self._initialized = True
        if knowledge_path is None:
            knowledge_path = os.path.join(get_script_dir(), KNOWLEDGE_BASE_FILE)
        self._knowledge_path = knowledge_path
        self._entries = {}
        self._last_modified = 0.0
        self.load()

    def load(self) -> bool:
        logger = get_logger()
        if not os.path.exists(self._knowledge_path):
            logger.info('知识库文件不存在，将创建空白知识库: %s', self._knowledge_path)
            self._entries = {}
            return True

        try:
            import yaml
            current_mtime = os.path.getmtime(self._knowledge_path)
            if current_mtime == self._last_modified and self._entries:
                return True

            with open(self._knowledge_path, 'r', encoding='utf-8') as f:
                data = yaml.safe_load(f)

            if not data or 'banks' not in data:
                self._entries = {}
                self._last_modified = current_mtime
                return True

            self._entries = {}
            for bank_data in data['banks']:
                entry = BankWikiEntry.from_dict(bank_data)
                self._entries[entry.bank_name] = entry

            self._last_modified = current_mtime
            logger.info('知识库已加载，共 %d 个银行条目', len(self._entries))
            return True
        except Exception as e:
            logger.error('加载知识库失败: %s', e, exc_info=True)
            return False

    def save(self) -> bool:
        logger = get_logger()
        try:
            import yaml
            data = {
                'banks': [entry.to_dict() for entry in self._entries.values()],
                'metadata': {
                    'last_updated': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    'version': '1.0',
                },
            }

            os.makedirs(os.path.dirname(self._knowledge_path) or '.', exist_ok=True)
            with open(self._knowledge_path, 'w', encoding='utf-8') as f:
                yaml.safe_dump(data, f, allow_unicode=True, default_flow_style=False, sort_keys=False)

            self._last_modified = 0.0
            self.load()
            logger.info('知识库已保存，共 %d 个银行条目', len(self._entries))
            return True
        except Exception as e:
            logger.error('保存知识库失败: %s', e, exc_info=True)
            return False

    def get_entry(self, bank_name: str) -> Optional[BankWikiEntry]:
        self.load()
        return self._entries.get(bank_name)

    def get_all_entries(self) -> Dict[str, BankWikiEntry]:
        self.load()
        return dict(self._entries)

    def get_all_bank_names(self) -> List[str]:
        self.load()
        return list(self._entries.keys())

    def add_or_update_entry(self, entry: BankWikiEntry) -> bool:
        self._entries[entry.bank_name] = entry
        return self.save()

    def remove_entry(self, bank_name: str) -> bool:
        if bank_name in self._entries:
            del self._entries[bank_name]
            return self.save()
        return False

    def add_known_issue(self, bank_name: str, issue: KnownIssue) -> bool:
        self.load()
        entry = self._entries.get(bank_name)
        if entry is None:
            entry = BankWikiEntry(bank_name=bank_name)
            self._entries[bank_name] = entry

        for existing in entry.known_issues:
            if existing.issue_id == issue.issue_id:
                idx = entry.known_issues.index(existing)
                entry.known_issues[idx] = issue
                return self.save()

        entry.known_issues.append(issue)
        return self.save()

    def add_template_screenshot(self, bank_name: str, screenshot: TemplateScreenshot) -> bool:
        self.load()
        entry = self._entries.get(bank_name)
        if entry is None:
            entry = BankWikiEntry(bank_name=bank_name)
            self._entries[bank_name] = entry

        for i, existing in enumerate(entry.template_screenshots):
            if existing.name == screenshot.name:
                entry.template_screenshots[i] = screenshot
                return self.save()

        entry.template_screenshots.append(screenshot)
        return self.save()

    def add_column_description(self, bank_name: str, col_desc: ColumnDescription) -> bool:
        self.load()
        entry = self._entries.get(bank_name)
        if entry is None:
            entry = BankWikiEntry(bank_name=bank_name)
            self._entries[bank_name] = entry

        for i, existing in enumerate(entry.column_descriptions):
            if existing.field_key == col_desc.field_key:
                entry.column_descriptions[i] = col_desc
                return self.save()

        entry.column_descriptions.append(col_desc)
        return self.save()

    def search_issues_by_error(self, error_message: str,
                               bank_name: Optional[str] = None,
                               min_score: float = 0.3) -> List[SearchResult]:
        """
        根据错误信息检索匹配的已知问题。

        匹配策略：
          1. 错误消息与 known_issue.error_patterns 正则匹配
          2. 错误消息与 known_issue.symptoms 关键词匹配
          3. 错误消息与 known_issue.title 模糊匹配
          4. 可选按银行名称过滤

        Args:
            error_message: 错误消息文本
            bank_name: 可选，限定搜索到特定银行
            min_score: 最低匹配分数阈值

        Returns:
            按匹配分数降序排列的 SearchResult 列表
        """
        self.load()
        results = []

        entries_to_search = {}
        if bank_name:
            entry = self._entries.get(bank_name)
            if entry:
                entries_to_search[bank_name] = entry
        else:
            entries_to_search = self._entries

        for bk_name, entry in entries_to_search.items():
            for issue in entry.known_issues:
                score, matched = self._compute_match_score(error_message, issue)
                if score >= min_score:
                    results.append(SearchResult(
                        bank_name=bk_name,
                        issue=issue,
                        match_score=score,
                        matched_patterns=matched,
                    ))

        results.sort(key=lambda r: r.match_score, reverse=True)
        return results

    def search_issues_by_keywords(self, keywords: List[str],
                                   bank_name: Optional[str] = None,
                                   min_score: float = 0.3) -> List[SearchResult]:
        """
        根据关键词列表检索匹配的已知问题。

        Args:
            keywords: 搜索关键词列表
            bank_name: 可选，限定搜索到特定银行
            min_score: 最低匹配分数阈值

        Returns:
            按匹配分数降序排列的 SearchResult 列表
        """
        self.load()
        results = []
        query_text = ' '.join(keywords)

        entries_to_search = {}
        if bank_name:
            entry = self._entries.get(bank_name)
            if entry:
                entries_to_search[bank_name] = entry
        else:
            entries_to_search = self._entries

        for bk_name, entry in entries_to_search.items():
            for issue in entry.known_issues:
                score, matched = self._compute_keyword_match_score(query_text, issue)
                if score >= min_score:
                    results.append(SearchResult(
                        bank_name=bk_name,
                        issue=issue,
                        match_score=score,
                        matched_patterns=matched,
                    ))

        results.sort(key=lambda r: r.match_score, reverse=True)
        return results

    def search_all(self, query: str,
                   bank_name: Optional[str] = None,
                   min_score: float = 0.3) -> List[SearchResult]:
        """
        综合搜索：同时按错误模式和关键词匹配。

        Args:
            query: 搜索文本（可以是错误消息或关键词）
            bank_name: 可选，限定搜索到特定银行
            min_score: 最低匹配分数阈值

        Returns:
            按匹配分数降序排列的 SearchResult 列表
        """
        self.load()
        error_results = self.search_issues_by_error(query, bank_name, min_score)
        keyword_results = self.search_issues_by_keywords(query.split(), bank_name, min_score)

        seen_issue_ids = set()
        merged = []
        for r in error_results:
            key = (r.bank_name, r.issue.issue_id)
            if key not in seen_issue_ids:
                seen_issue_ids.add(key)
                merged.append(r)

        for r in keyword_results:
            key = (r.bank_name, r.issue.issue_id)
            if key not in seen_issue_ids:
                seen_issue_ids.add(key)
                merged.append(r)
            else:
                for existing in merged:
                    if (existing.bank_name, existing.issue.issue_id) == key:
                        existing.match_score = max(existing.match_score, r.match_score)
                        existing.matched_patterns = list(set(
                            existing.matched_patterns + r.matched_patterns
                        ))
                        break

        merged.sort(key=lambda r: r.match_score, reverse=True)
        return merged

    def _compute_match_score(self, error_message: str,
                             issue: KnownIssue) -> Tuple[float, List[str]]:
        score = 0.0
        matched = []
        error_lower = error_message.lower()

        for pattern in issue.error_patterns:
            try:
                if re.search(pattern, error_message, re.IGNORECASE):
                    score += 0.4
                    matched.append(f'error_pattern:{pattern}')
            except re.error:
                if pattern.lower() in error_lower:
                    score += 0.35
                    matched.append(f'error_pattern(fallback):{pattern}')

        for symptom in issue.symptoms:
            symptom_lower = symptom.lower()
            symptom_tokens = symptom_lower.split()
            for token in symptom_tokens:
                if len(token) >= 2 and token in error_lower:
                    score += 0.15
                    matched.append(f'symptom:{symptom}')
                    break

        title_lower = issue.title.lower()
        title_tokens = title_lower.split()
        for token in title_tokens:
            if len(token) >= 2 and token in error_lower:
                score += 0.1
                matched.append(f'title:{issue.title}')
                break

        for tag in issue.tags:
            if tag.lower() in error_lower:
                score += 0.1
                matched.append(f'tag:{tag}')

        score = min(score, 1.0)
        return score, matched

    def _compute_keyword_match_score(self, query_text: str,
                                      issue: KnownIssue) -> Tuple[float, List[str]]:
        score = 0.0
        matched = []
        query_lower = query_text.lower()
        query_tokens = set(query_lower.split())

        for pattern in issue.error_patterns:
            pattern_lower = pattern.lower()
            pattern_tokens = set(re.findall(r'[\w\u4e00-\u9fff]+', pattern_lower))
            overlap = query_tokens & pattern_tokens
            if overlap:
                ratio = len(overlap) / max(len(pattern_tokens), 1)
                score += 0.3 * ratio
                matched.append(f'error_pattern_kw:{pattern}')

        for symptom in issue.symptoms:
            symptom_lower = symptom.lower()
            symptom_tokens = set(re.findall(r'[\w\u4e00-\u9fff]+', symptom_lower))
            overlap = query_tokens & symptom_tokens
            if overlap:
                ratio = len(overlap) / max(len(symptom_tokens), 1)
                score += 0.2 * ratio
                matched.append(f'symptom_kw:{symptom}')

        title_tokens = set(re.findall(r'[\w\u4e00-\u9fff]+', issue.title.lower()))
        overlap = query_tokens & title_tokens
        if overlap:
            ratio = len(overlap) / max(len(title_tokens), 1)
            score += 0.25 * ratio
            matched.append(f'title_kw:{issue.title}')

        for tag in issue.tags:
            if tag.lower() in query_lower:
                score += 0.15
                matched.append(f'tag_kw:{tag}')

        if issue.root_cause:
            cause_tokens = set(re.findall(r'[\w\u4e00-\u9fff]+', issue.root_cause.lower()))
            overlap = query_tokens & cause_tokens
            if overlap:
                ratio = len(overlap) / max(len(cause_tokens), 1)
                score += 0.1 * ratio
                matched.append(f'root_cause_kw:{issue.root_cause[:30]}')

        score = min(score, 1.0)
        return score, matched

    def get_issues_for_bank(self, bank_name: str) -> List[KnownIssue]:
        self.load()
        entry = self._entries.get(bank_name)
        if entry is None:
            return []
        return list(entry.known_issues)

    def get_column_descriptions(self, bank_name: str) -> List[ColumnDescription]:
        self.load()
        entry = self._entries.get(bank_name)
        if entry is None:
            return []
        return list(entry.column_descriptions)

    def get_template_screenshots(self, bank_name: str) -> List[TemplateScreenshot]:
        self.load()
        entry = self._entries.get(bank_name)
        if entry is None:
            return []
        return list(entry.template_screenshots)

    def get_general_pitfalls(self, bank_name: str) -> List[str]:
        self.load()
        entry = self._entries.get(bank_name)
        if entry is None:
            return []
        return list(entry.general_pitfalls)

    def get_config_version(self, bank_name: str) -> str:
        self.load()
        entry = self._entries.get(bank_name)
        if entry is None:
            return ''
        return entry.config_version

    def get_knowledge_path(self) -> str:
        return self._knowledge_path

    def get_statistics(self) -> Dict[str, Any]:
        self.load()
        total_banks = len(self._entries)
        total_issues = sum(len(e.known_issues) for e in self._entries.values())
        total_screenshots = sum(len(e.template_screenshots) for e in self._entries.values())
        total_columns = sum(len(e.column_descriptions) for e in self._entries.values())
        return {
            'total_banks': total_banks,
            'total_known_issues': total_issues,
            'total_template_screenshots': total_screenshots,
            'total_column_descriptions': total_columns,
            'knowledge_base_path': self._knowledge_path,
        }

    def to_dict(self) -> Dict[str, Any]:
        self.load()
        return {
            'banks': {name: entry.to_dict() for name, entry in self._entries.items()},
            'statistics': self.get_statistics(),
        }


def get_knowledge_base(knowledge_path=None) -> BankKnowledgeBase:
    return BankKnowledgeBase(knowledge_path)


def search_known_issues(error_message: str,
                        bank_name: Optional[str] = None,
                        min_score: float = 0.3) -> List[SearchResult]:
    kb = get_knowledge_base()
    return kb.search_all(error_message, bank_name, min_score)


def diagnose_from_knowledge_base(log_lines: List[str],
                                  bank_name: Optional[str] = None) -> List[SearchResult]:
    """
    从日志行中提取错误信息，检索知识库匹配已知问题与解法。

    Args:
        log_lines: 日志行列表
        bank_name: 可选，限定搜索到特定银行

    Returns:
        按匹配分数降序排列的 SearchResult 列表
    """
    logger = get_logger()
    kb = get_knowledge_base()

    error_lines = []
    error_patterns = [
        r'\bERROR\b', r'\bCRITICAL\b', r'\bWARNING\b',
        r'失败', r'错误', r'异常', r'无法', r'不匹配',
        r'ImportError', r'FileNotFoundError', r'PermissionError',
        r'BadZipFile', r'InvalidFileException',
    ]

    for line in log_lines:
        for pattern in error_patterns:
            if re.search(pattern, line, re.IGNORECASE):
                error_lines.append(line.strip())
                break

    if not error_lines:
        return []

    error_text = '\n'.join(error_lines[:50])
    results = kb.search_all(error_text, bank_name, min_score=0.3)

    if results:
        logger.info('知识库匹配到 %d 个已知问题', len(results))
        for r in results[:5]:
            logger.info(
                '  [%.0f%%] %s - %s (%s)',
                r.match_score * 100, r.bank_name, r.issue.title,
                ', '.join(r.matched_patterns[:3]),
            )

    return results
