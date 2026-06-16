#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
银行格式知识库模块测试
"""

import os
import sys
import tempfile
import shutil

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'backend'))

from bank_knowledge_base import (
    ColumnDescription,
    KnownIssue,
    TemplateScreenshot,
    BankWikiEntry,
    SearchResult,
    BankKnowledgeBase,
    get_knowledge_base,
    search_known_issues,
    diagnose_from_knowledge_base,
)


@pytest.fixture
def temp_dir():
    d = tempfile.mkdtemp(prefix='test_kb_')
    yield d
    shutil.rmtree(d, ignore_errors=True)


@pytest.fixture
def kb_instance(temp_dir):
    BankKnowledgeBase._instance = None
    BankKnowledgeBase._instance = None
    knowledge_path = os.path.join(temp_dir, 'bank_knowledge.yaml')
    kb = BankKnowledgeBase(knowledge_path=knowledge_path)
    yield kb
    BankKnowledgeBase._instance = None


@pytest.fixture
def sample_entry():
    return BankWikiEntry(
        bank_name='测试银行',
        display_name='测试银行',
        description='测试银行描述',
        config_version='1.0',
        last_verified_date='2026-06-01',
        verified_by='tester',
        template_screenshots=[
            TemplateScreenshot(
                name='标准模板',
                description='测试银行标准导出格式截图',
                file_path='screenshots/test_bank.png',
                upload_date='2026-06-01',
                config_version='1.0',
                notes='',
            ),
        ],
        column_descriptions=[
            ColumnDescription(
                field_key='trade_date',
                display_name='交易日期',
                description='交易发生日期',
                example_values=['2024-01-15'],
                pitfalls=['日期格式可能不统一'],
            ),
            ColumnDescription(
                field_key='payment',
                display_name='支出金额',
                description='资金流出金额',
                example_values=['5000.00'],
                pitfalls=[],
            ),
        ],
        known_issues=[
            KnownIssue(
                issue_id='test_header_shift',
                title='测试银行表头行偏移',
                description='表头不在预期行',
                symptoms=['数据解析异常', '列映射错位'],
                root_cause='网银系统升级后导出模板变化',
                fix_steps=['修改 start_row', '重新运行'],
                error_patterns=[
                    r'表头.*不匹配.*测试银行',
                    r'header.*mismatch.*test',
                ],
                severity='warning',
                config_version='1.0',
                last_seen_date='2026-05-01',
                tags=['表头偏移', '测试银行'],
            ),
        ],
        general_pitfalls=['测试银行需要注意的事项'],
        notes='测试备注',
        tags=['测试'],
    )


class TestColumnDescription:
    def test_create(self):
        cd = ColumnDescription(
            field_key='trade_date',
            display_name='交易日期',
            description='交易日期',
            example_values=['2024-01-15'],
            pitfalls=['格式问题'],
        )
        assert cd.field_key == 'trade_date'
        assert cd.display_name == '交易日期'
        assert len(cd.example_values) == 1
        assert len(cd.pitfalls) == 1

    def test_to_dict(self):
        cd = ColumnDescription(field_key='balance', display_name='余额')
        d = cd.to_dict()
        assert d['field_key'] == 'balance'
        assert d['display_name'] == '余额'
        assert d['example_values'] == []
        assert d['pitfalls'] == []

    def test_from_dict(self):
        data = {
            'field_key': 'receipt',
            'display_name': '收入金额',
            'description': '收入',
            'example_values': ['1000.00'],
            'pitfalls': ['可能为空'],
        }
        cd = ColumnDescription.from_dict(data)
        assert cd.field_key == 'receipt'
        assert cd.display_name == '收入金额'
        assert len(cd.example_values) == 1

    def test_roundtrip(self):
        cd = ColumnDescription(
            field_key='payment',
            display_name='支出',
            description='支出金额',
            example_values=['5000'],
            pitfalls=['注意符号'],
        )
        d = cd.to_dict()
        cd2 = ColumnDescription.from_dict(d)
        assert cd2.field_key == cd.field_key
        assert cd2.display_name == cd.display_name
        assert cd2.example_values == cd.example_values
        assert cd2.pitfalls == cd.pitfalls


class TestKnownIssue:
    def test_create(self):
        issue = KnownIssue(
            issue_id='test_001',
            title='测试问题',
            description='描述',
            symptoms=['症状1'],
            root_cause='原因',
            fix_steps=['步骤1'],
            error_patterns=[r'test.*error'],
            severity='warning',
            config_version='1.0',
            tags=['测试'],
        )
        assert issue.issue_id == 'test_001'
        assert issue.severity == 'warning'

    def test_to_dict(self):
        issue = KnownIssue(issue_id='i1', title='标题')
        d = issue.to_dict()
        assert d['issue_id'] == 'i1'
        assert d['title'] == '标题'
        assert d['symptoms'] == []
        assert d['fix_steps'] == []
        assert d['error_patterns'] == []

    def test_from_dict(self):
        data = {
            'issue_id': 'i2',
            'title': '标题2',
            'description': '描述2',
            'symptoms': ['s1'],
            'root_cause': '原因2',
            'fix_steps': ['f1', 'f2'],
            'error_patterns': [r'pattern'],
            'severity': 'error',
            'config_version': '2.0',
            'last_seen_date': '2026-06-01',
            'tags': ['tag1'],
        }
        issue = KnownIssue.from_dict(data)
        assert issue.issue_id == 'i2'
        assert len(issue.fix_steps) == 2
        assert issue.severity == 'error'

    def test_roundtrip(self):
        issue = KnownIssue(
            issue_id='rt1',
            title='RT测试',
            symptoms=['s1', 's2'],
            error_patterns=[r'e1'],
            tags=['t1'],
        )
        d = issue.to_dict()
        issue2 = KnownIssue.from_dict(d)
        assert issue2.issue_id == issue.issue_id
        assert issue2.symptoms == issue.symptoms
        assert issue2.error_patterns == issue.error_patterns


class TestTemplateScreenshot:
    def test_create(self):
        ss = TemplateScreenshot(
            name='模板1',
            description='描述',
            file_path='screenshots/test.png',
        )
        assert ss.name == '模板1'
        assert ss.file_path == 'screenshots/test.png'

    def test_roundtrip(self):
        ss = TemplateScreenshot(
            name='模板2',
            description='描述2',
            file_path='path/to/img.png',
            upload_date='2026-06-01',
            config_version='1.0',
            notes='备注',
        )
        d = ss.to_dict()
        ss2 = TemplateScreenshot.from_dict(d)
        assert ss2.name == ss.name
        assert ss2.file_path == ss.file_path
        assert ss2.notes == ss.notes


class TestBankWikiEntry:
    def test_create(self):
        entry = BankWikiEntry(bank_name='测试银行')
        assert entry.bank_name == '测试银行'
        assert entry.config_version == '1.0'
        assert entry.known_issues == []
        assert entry.column_descriptions == []

    def test_to_dict_with_nested(self, sample_entry):
        d = sample_entry.to_dict()
        assert d['bank_name'] == '测试银行'
        assert len(d['known_issues']) == 1
        assert len(d['column_descriptions']) == 2
        assert len(d['template_screenshots']) == 1
        assert d['known_issues'][0]['issue_id'] == 'test_header_shift'

    def test_from_dict(self):
        data = {
            'bank_name': '示例银行',
            'display_name': '示例银行',
            'config_version': '2.0',
            'known_issues': [
                {
                    'issue_id': 'i1',
                    'title': '问题1',
                    'symptoms': ['s1'],
                    'error_patterns': [r'pat1'],
                },
            ],
            'column_descriptions': [
                {
                    'field_key': 'trade_date',
                    'display_name': '日期',
                },
            ],
            'general_pitfalls': ['注意1', '注意2'],
        }
        entry = BankWikiEntry.from_dict(data)
        assert entry.bank_name == '示例银行'
        assert len(entry.known_issues) == 1
        assert len(entry.column_descriptions) == 1
        assert len(entry.general_pitfalls) == 2

    def test_roundtrip(self, sample_entry):
        d = sample_entry.to_dict()
        entry2 = BankWikiEntry.from_dict(d)
        assert entry2.bank_name == sample_entry.bank_name
        assert len(entry2.known_issues) == len(sample_entry.known_issues)
        assert len(entry2.column_descriptions) == len(sample_entry.column_descriptions)
        assert entry2.known_issues[0].issue_id == 'test_header_shift'
        assert entry2.column_descriptions[0].field_key == 'trade_date'


class TestSearchResult:
    def test_create(self):
        issue = KnownIssue(issue_id='s1', title='搜索测试')
        sr = SearchResult(
            bank_name='银行A',
            issue=issue,
            match_score=0.85,
            matched_patterns=['error_pattern:test'],
        )
        assert sr.bank_name == '银行A'
        assert sr.match_score == 0.85
        assert len(sr.matched_patterns) == 1

    def test_to_dict(self):
        issue = KnownIssue(issue_id='s2', title='搜索2')
        sr = SearchResult(bank_name='银行B', issue=issue, match_score=0.6)
        d = sr.to_dict()
        assert d['bank_name'] == '银行B'
        assert d['match_score'] == 0.6
        assert d['issue']['issue_id'] == 's2'


class TestBankKnowledgeBase:
    def test_empty_kb(self, kb_instance):
        names = kb_instance.get_all_bank_names()
        assert names == []

    def test_add_entry(self, kb_instance, sample_entry):
        ok = kb_instance.add_or_update_entry(sample_entry)
        assert ok is True

        entry = kb_instance.get_entry('测试银行')
        assert entry is not None
        assert entry.bank_name == '测试银行'
        assert len(entry.known_issues) == 1

    def test_update_entry(self, kb_instance, sample_entry):
        kb_instance.add_or_update_entry(sample_entry)

        sample_entry.description = '更新后的描述'
        kb_instance.add_or_update_entry(sample_entry)

        entry = kb_instance.get_entry('测试银行')
        assert entry.description == '更新后的描述'

    def test_remove_entry(self, kb_instance, sample_entry):
        kb_instance.add_or_update_entry(sample_entry)
        ok = kb_instance.remove_entry('测试银行')
        assert ok is True
        assert kb_instance.get_entry('测试银行') is None

    def test_persistence(self, kb_instance, sample_entry, temp_dir):
        kb_instance.add_or_update_entry(sample_entry)

        knowledge_path = os.path.join(temp_dir, 'bank_knowledge.yaml')
        assert os.path.exists(knowledge_path)

        BankKnowledgeBase._instance = None
        kb2 = BankKnowledgeBase(knowledge_path=knowledge_path)
        entry = kb2.get_entry('测试银行')
        assert entry is not None
        assert entry.bank_name == '测试银行'
        assert len(entry.known_issues) == 1
        BankKnowledgeBase._instance = None

    def test_add_known_issue(self, kb_instance):
        entry = BankWikiEntry(bank_name='新银行')
        kb_instance.add_or_update_entry(entry)

        issue = KnownIssue(
            issue_id='new_issue_1',
            title='新问题',
            symptoms=['症状1'],
            error_patterns=[r'新问题.*error'],
        )
        ok = kb_instance.add_known_issue('新银行', issue)
        assert ok is True

        issues = kb_instance.get_issues_for_bank('新银行')
        assert len(issues) == 1
        assert issues[0].issue_id == 'new_issue_1'

    def test_add_known_issue_updates_existing(self, kb_instance):
        entry = BankWikiEntry(bank_name='银行X')
        issue1 = KnownIssue(issue_id='i1', title='问题1')
        entry.known_issues = [issue1]
        kb_instance.add_or_update_entry(entry)

        issue2 = KnownIssue(issue_id='i1', title='问题1-更新', severity='error')
        kb_instance.add_known_issue('银行X', issue2)

        issues = kb_instance.get_issues_for_bank('银行X')
        assert len(issues) == 1
        assert issues[0].title == '问题1-更新'
        assert issues[0].severity == 'error'

    def test_add_column_description(self, kb_instance):
        entry = BankWikiEntry(bank_name='银行Y')
        kb_instance.add_or_update_entry(entry)

        cd = ColumnDescription(field_key='trade_date', display_name='交易日期')
        ok = kb_instance.add_column_description('银行Y', cd)
        assert ok is True

        cols = kb_instance.get_column_descriptions('银行Y')
        assert len(cols) == 1
        assert cols[0].field_key == 'trade_date'

    def test_add_template_screenshot(self, kb_instance):
        entry = BankWikiEntry(bank_name='银行Z')
        kb_instance.add_or_update_entry(entry)

        ss = TemplateScreenshot(name='模板1', file_path='test.png')
        ok = kb_instance.add_template_screenshot('银行Z', ss)
        assert ok is True

        screenshots = kb_instance.get_template_screenshots('银行Z')
        assert len(screenshots) == 1

    def test_get_general_pitfalls(self, kb_instance, sample_entry):
        kb_instance.add_or_update_entry(sample_entry)
        pitfalls = kb_instance.get_general_pitfalls('测试银行')
        assert len(pitfalls) == 1
        assert '测试银行需要注意的事项' in pitfalls

    def test_get_config_version(self, kb_instance, sample_entry):
        kb_instance.add_or_update_entry(sample_entry)
        version = kb_instance.get_config_version('测试银行')
        assert version == '1.0'

    def test_get_statistics(self, kb_instance, sample_entry):
        kb_instance.add_or_update_entry(sample_entry)
        stats = kb_instance.get_statistics()
        assert stats['total_banks'] == 1
        assert stats['total_known_issues'] == 1
        assert stats['total_template_screenshots'] == 1
        assert stats['total_column_descriptions'] == 2

    def test_nonexistent_bank(self, kb_instance):
        assert kb_instance.get_entry('不存在的银行') is None
        assert kb_instance.get_issues_for_bank('不存在的银行') == []
        assert kb_instance.get_column_descriptions('不存在的银行') == []
        assert kb_instance.get_general_pitfalls('不存在的银行') == []
        assert kb_instance.get_config_version('不存在的银行') == ''


class TestSearchByError:
    def test_match_error_pattern(self, kb_instance, sample_entry):
        kb_instance.add_or_update_entry(sample_entry)

        results = kb_instance.search_issues_by_error(
            '表头不匹配 测试银行',
            bank_name='测试银行',
        )
        assert len(results) >= 1
        assert results[0].bank_name == '测试银行'
        assert results[0].match_score > 0

    def test_match_symptom(self, kb_instance, sample_entry):
        kb_instance.add_or_update_entry(sample_entry)

        results = kb_instance.search_issues_by_error(
            '数据解析异常',
            bank_name='测试银行',
            min_score=0.1,
        )
        assert len(results) >= 1

    def test_no_match(self, kb_instance, sample_entry):
        kb_instance.add_or_update_entry(sample_entry)

        results = kb_instance.search_issues_by_error(
            '完全无关的消息 xyz123',
            bank_name='测试银行',
            min_score=0.5,
        )
        assert len(results) == 0

    def test_filter_by_bank(self, kb_instance, sample_entry):
        kb_instance.add_or_update_entry(sample_entry)

        other_entry = BankWikiEntry(bank_name='其他银行')
        kb_instance.add_or_update_entry(other_entry)

        results = kb_instance.search_issues_by_error(
            '表头不匹配',
            bank_name='测试银行',
        )
        for r in results:
            assert r.bank_name == '测试银行'

    def test_search_all_banks(self, kb_instance, sample_entry):
        kb_instance.add_or_update_entry(sample_entry)

        results = kb_instance.search_issues_by_error('表头不匹配 测试银行')
        assert len(results) >= 1


class TestSearchByKeywords:
    def test_keyword_match(self, kb_instance, sample_entry):
        kb_instance.add_or_update_entry(sample_entry)

        results = kb_instance.search_issues_by_keywords(
            ['测试银行', '表头', '偏移'],
            bank_name='测试银行',
            min_score=0.1,
        )
        assert len(results) >= 1

    def test_no_keyword_match(self, kb_instance, sample_entry):
        kb_instance.add_or_update_entry(sample_entry)

        results = kb_instance.search_issues_by_keywords(
            ['不相关的', '关键词'],
            bank_name='测试银行',
            min_score=0.5,
        )
        assert len(results) == 0


class TestSearchAll:
    def test_combined_search(self, kb_instance, sample_entry):
        kb_instance.add_or_update_entry(sample_entry)

        results = kb_instance.search_all('表头不匹配 测试银行')
        assert len(results) >= 1

    def test_dedup(self, kb_instance, sample_entry):
        kb_instance.add_or_update_entry(sample_entry)

        results = kb_instance.search_all('表头不匹配')
        issue_ids = set()
        for r in results:
            key = (r.bank_name, r.issue.issue_id)
            assert key not in issue_ids, f'重复结果: {key}'
            issue_ids.add(key)


class TestDiagnoseFromKnowledgeBase:
    def test_diagnose_with_error_lines(self, kb_instance, sample_entry):
        kb_instance.add_or_update_entry(sample_entry)

        log_lines = [
            '[INFO] 正常日志',
            '[ERROR] 表头不匹配 测试银行',
            '[WARNING] 数据解析异常',
        ]
        results = diagnose_from_knowledge_base(log_lines, bank_name='测试银行')
        assert len(results) >= 1

    def test_diagnose_no_error_lines(self, kb_instance, sample_entry):
        kb_instance.add_or_update_entry(sample_entry)

        log_lines = [
            '[INFO] 正常日志1',
            '[INFO] 正常日志2',
        ]
        results = diagnose_from_knowledge_base(log_lines)
        assert len(results) == 0


class TestSearchKnownIssuesHelper:
    def test_search_known_issues(self, kb_instance, sample_entry):
        kb_instance.add_or_update_entry(sample_entry)

        results = search_known_issues('表头不匹配 测试银行', bank_name='测试银行')
        assert len(results) >= 1


class TestLoadRealKnowledgeBase:
    def test_load_existing_yaml(self):
        knowledge_path = os.path.join(
            os.path.dirname(__file__), '..', 'backend', 'bank_knowledge.yaml'
        )
        if not os.path.exists(knowledge_path):
            pytest.skip('bank_knowledge.yaml 不存在')

        BankKnowledgeBase._instance = None
        try:
            kb = BankKnowledgeBase(knowledge_path=knowledge_path)
            names = kb.get_all_bank_names()
            assert len(names) > 0

            for name in names:
                entry = kb.get_entry(name)
                assert entry is not None
                assert entry.bank_name == name
        finally:
            BankKnowledgeBase._instance = None
