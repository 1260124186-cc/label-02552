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


class TestPreserveNestedOnUpdate:
    def test_add_or_update_preserves_screenshots(self, kb_instance, sample_entry):
        kb_instance.add_or_update_entry(sample_entry)
        assert len(kb_instance.get_template_screenshots('测试银行')) == 1

        new_entry = BankWikiEntry(
            bank_name='测试银行',
            display_name='测试银行-更新',
            description='新的描述',
        )
        assert len(new_entry.template_screenshots) == 0

        ok = kb_instance.add_or_update_entry(new_entry, preserve_nested=True)
        assert ok is True

        entry = kb_instance.get_entry('测试银行')
        assert entry.display_name == '测试银行-更新'
        assert entry.description == '新的描述'
        assert len(entry.template_screenshots) == 1
        assert entry.template_screenshots[0].name == '标准模板'

    def test_add_or_update_preserves_column_descriptions(self, kb_instance, sample_entry):
        kb_instance.add_or_update_entry(sample_entry)
        assert len(kb_instance.get_column_descriptions('测试银行')) == 2

        new_entry = BankWikiEntry(bank_name='测试银行')
        assert len(new_entry.column_descriptions) == 0

        kb_instance.add_or_update_entry(new_entry, preserve_nested=True)

        entry = kb_instance.get_entry('测试银行')
        assert len(entry.column_descriptions) == 2
        field_keys = [c.field_key for c in entry.column_descriptions]
        assert 'trade_date' in field_keys
        assert 'payment' in field_keys

    def test_add_or_update_preserves_known_issues(self, kb_instance, sample_entry):
        kb_instance.add_or_update_entry(sample_entry)
        assert len(kb_instance.get_issues_for_bank('测试银行')) == 1

        new_entry = BankWikiEntry(
            bank_name='测试银行',
            config_version='2.0',
        )
        assert len(new_entry.known_issues) == 0

        kb_instance.add_or_update_entry(new_entry, preserve_nested=True)

        entry = kb_instance.get_entry('测试银行')
        assert entry.config_version == '2.0'
        assert len(entry.known_issues) == 1
        assert entry.known_issues[0].issue_id == 'test_header_shift'

    def test_add_or_update_preserves_general_pitfalls(self, kb_instance, sample_entry):
        kb_instance.add_or_update_entry(sample_entry)
        assert len(kb_instance.get_general_pitfalls('测试银行')) == 1

        new_entry = BankWikiEntry(bank_name='测试银行')
        assert len(new_entry.general_pitfalls) == 0

        kb_instance.add_or_update_entry(new_entry, preserve_nested=True)

        entry = kb_instance.get_entry('测试银行')
        assert len(entry.general_pitfalls) == 1
        assert '测试银行需要注意的事项' in entry.general_pitfalls

    def test_add_or_update_preserves_all_nested(self, kb_instance, sample_entry):
        kb_instance.add_or_update_entry(sample_entry)
        original_issues = list(sample_entry.known_issues)
        original_cols = list(sample_entry.column_descriptions)
        original_ss = list(sample_entry.template_screenshots)
        original_pitfalls = list(sample_entry.general_pitfalls)

        new_entry = BankWikiEntry(
            bank_name='测试银行',
            display_name='更新后的名称',
            description='更新后的描述',
            config_version='2.0',
            last_verified_date='2026-06-16',
            verified_by='admin',
            notes='更新备注',
            tags=['更新后'],
        )

        kb_instance.add_or_update_entry(new_entry, preserve_nested=True)

        entry = kb_instance.get_entry('测试银行')
        assert entry.display_name == '更新后的名称'
        assert entry.description == '更新后的描述'
        assert entry.config_version == '2.0'
        assert entry.last_verified_date == '2026-06-16'
        assert entry.verified_by == 'admin'
        assert entry.notes == '更新备注'
        assert entry.tags == ['更新后']

        assert len(entry.template_screenshots) == len(original_ss)
        assert len(entry.column_descriptions) == len(original_cols)
        assert len(entry.known_issues) == len(original_issues)
        assert len(entry.general_pitfalls) == len(original_pitfalls)

        assert entry.known_issues[0].issue_id == original_issues[0].issue_id
        assert entry.column_descriptions[0].field_key == original_cols[0].field_key
        assert entry.template_screenshots[0].name == original_ss[0].name
        assert entry.general_pitfalls[0] == original_pitfalls[0]

    def test_add_or_update_with_explicit_nested_overrides(self, kb_instance, sample_entry):
        kb_instance.add_or_update_entry(sample_entry)

        new_issue = KnownIssue(issue_id='new_issue', title='新问题')
        new_entry = BankWikiEntry(
            bank_name='测试银行',
            known_issues=[new_issue],
        )
        assert len(new_entry.known_issues) == 1

        kb_instance.add_or_update_entry(new_entry, preserve_nested=True)

        entry = kb_instance.get_entry('测试银行')
        assert len(entry.known_issues) == 1
        assert entry.known_issues[0].issue_id == 'new_issue'
        assert len(entry.template_screenshots) == 1
        assert len(entry.column_descriptions) == 2

    def test_add_or_update_force_override(self, kb_instance, sample_entry):
        kb_instance.add_or_update_entry(sample_entry)

        new_entry = BankWikiEntry(bank_name='测试银行')
        kb_instance.add_or_update_entry(new_entry, preserve_nested=False)

        entry = kb_instance.get_entry('测试银行')
        assert len(entry.known_issues) == 0
        assert len(entry.template_screenshots) == 0
        assert len(entry.column_descriptions) == 0
        assert len(entry.general_pitfalls) == 0

    def test_add_or_update_default_preserves(self, kb_instance, sample_entry):
        kb_instance.add_or_update_entry(sample_entry)

        new_entry = BankWikiEntry(bank_name='测试银行')
        kb_instance.add_or_update_entry(new_entry)

        entry = kb_instance.get_entry('测试银行')
        assert len(entry.known_issues) == 1
        assert len(entry.template_screenshots) == 1
        assert len(entry.column_descriptions) == 2

    def test_add_new_entry_no_preserve(self, kb_instance):
        entry = BankWikiEntry(bank_name='全新银行')
        ok = kb_instance.add_or_update_entry(entry, preserve_nested=True)
        assert ok is True
        assert kb_instance.get_entry('全新银行') is not None


class TestUpdateBankBasicInfo:
    def test_update_basic_info(self, kb_instance, sample_entry):
        kb_instance.add_or_update_entry(sample_entry)

        ok = kb_instance.update_bank_basic_info(
            '测试银行',
            display_name='新显示名',
            config_version='2.0',
            description='新描述',
        )
        assert ok is True

        entry = kb_instance.get_entry('测试银行')
        assert entry.display_name == '新显示名'
        assert entry.config_version == '2.0'
        assert entry.description == '新描述'
        assert len(entry.known_issues) == 1
        assert len(entry.template_screenshots) == 1
        assert len(entry.column_descriptions) == 2

    def test_update_tags(self, kb_instance, sample_entry):
        kb_instance.add_or_update_entry(sample_entry)

        ok = kb_instance.update_bank_basic_info(
            '测试银行',
            tags=['生产', '已验证'],
        )
        assert ok is True

        entry = kb_instance.get_entry('测试银行')
        assert entry.tags == ['生产', '已验证']
        assert len(entry.known_issues) == 1

    def test_update_last_verified_date(self, kb_instance, sample_entry):
        kb_instance.add_or_update_entry(sample_entry)

        ok = kb_instance.update_bank_basic_info(
            '测试银行',
            last_verified_date='2026-06-16',
            verified_by='operator',
        )
        assert ok is True

        entry = kb_instance.get_entry('测试银行')
        assert entry.last_verified_date == '2026-06-16'
        assert entry.verified_by == 'operator'
        assert len(entry.known_issues) == 1

    def test_update_notes(self, kb_instance, sample_entry):
        kb_instance.add_or_update_entry(sample_entry)

        ok = kb_instance.update_bank_basic_info(
            '测试银行',
            notes='新增备注信息',
        )
        assert ok is True

        entry = kb_instance.get_entry('测试银行')
        assert entry.notes == '新增备注信息'
        assert len(entry.column_descriptions) == 2

    def test_update_nonexistent_bank(self, kb_instance):
        ok = kb_instance.update_bank_basic_info(
            '不存在的银行',
            display_name='测试',
        )
        assert ok is False

    def test_update_ignores_unknown_fields(self, kb_instance, sample_entry):
        kb_instance.add_or_update_entry(sample_entry)

        ok = kb_instance.update_bank_basic_info(
            '测试银行',
            display_name='合法更新',
            known_issues=[],
            template_screenshots=[],
        )
        assert ok is True

        entry = kb_instance.get_entry('测试银行')
        assert entry.display_name == '合法更新'
        assert len(entry.known_issues) == 1
        assert len(entry.template_screenshots) == 1

    def test_update_only_selected_fields(self, kb_instance, sample_entry):
        kb_instance.add_or_update_entry(sample_entry)
        original_desc = sample_entry.description
        original_version = sample_entry.config_version

        ok = kb_instance.update_bank_basic_info(
            '测试银行',
            verified_by='new_operator',
        )
        assert ok is True

        entry = kb_instance.get_entry('测试银行')
        assert entry.verified_by == 'new_operator'
        assert entry.description == original_desc
        assert entry.config_version == original_version
        assert len(entry.known_issues) == 1

    def test_update_preserves_searchability(self, kb_instance, sample_entry):
        kb_instance.add_or_update_entry(sample_entry)

        kb_instance.update_bank_basic_info(
            '测试银行',
            config_version='2.0',
        )

        results = kb_instance.search_issues_by_error(
            '表头不匹配 测试银行',
            bank_name='测试银行',
        )
        assert len(results) >= 1
        assert results[0].issue.config_version == '1.0'
        assert kb_instance.get_config_version('测试银行') == '2.0'


class TestSafeUpdatePersistence:
    def test_update_persists_to_disk(self, kb_instance, sample_entry, temp_dir):
        kb_instance.add_or_update_entry(sample_entry)
        original_ss_count = len(sample_entry.template_screenshots)
        original_issues_count = len(sample_entry.known_issues)

        kb_instance.update_bank_basic_info(
            '测试银行',
            description='更新后的描述',
            config_version='2.0',
        )

        BankKnowledgeBase._instance = None
        knowledge_path = os.path.join(temp_dir, 'bank_knowledge.yaml')
        kb2 = BankKnowledgeBase(knowledge_path=knowledge_path)

        entry = kb2.get_entry('测试银行')
        assert entry.description == '更新后的描述'
        assert entry.config_version == '2.0'
        assert len(entry.template_screenshots) == original_ss_count
        assert len(entry.known_issues) == original_issues_count
        BankKnowledgeBase._instance = None

    def test_add_or_update_preserve_persists(self, kb_instance, sample_entry, temp_dir):
        kb_instance.add_or_update_entry(sample_entry)

        new_entry = BankWikiEntry(
            bank_name='测试银行',
            display_name='更新显示名',
        )
        kb_instance.add_or_update_entry(new_entry, preserve_nested=True)

        BankKnowledgeBase._instance = None
        knowledge_path = os.path.join(temp_dir, 'bank_knowledge.yaml')
        kb2 = BankKnowledgeBase(knowledge_path=knowledge_path)

        entry = kb2.get_entry('测试银行')
        assert entry.display_name == '更新显示名'
        assert len(entry.known_issues) == 1
        assert len(entry.template_screenshots) == 1
        assert len(entry.column_descriptions) == 2
        BankKnowledgeBase._instance = None


class TestKnownIssueUpdateNoSideEffects:
    def test_add_known_issue_preserves_other_issues(self, kb_instance, sample_entry):
        kb_instance.add_or_update_entry(sample_entry)

        new_issue = KnownIssue(
            issue_id='issue_2',
            title='第二个问题',
            error_patterns=[r'error.*2'],
        )
        ok = kb_instance.add_known_issue('测试银行', new_issue)
        assert ok is True

        entry = kb_instance.get_entry('测试银行')
        assert len(entry.known_issues) == 2
        issue_ids = [i.issue_id for i in entry.known_issues]
        assert 'test_header_shift' in issue_ids
        assert 'issue_2' in issue_ids
        assert len(entry.template_screenshots) == 1
        assert len(entry.column_descriptions) == 2

    def test_add_known_issue_update_existing(self, kb_instance, sample_entry):
        kb_instance.add_or_update_entry(sample_entry)

        updated_issue = KnownIssue(
            issue_id='test_header_shift',
            title='更新后的标题',
            description='更新后的描述',
            symptoms=['新症状'],
            root_cause='新根因',
            fix_steps=['新步骤'],
            error_patterns=[r'new.*pattern'],
            severity='error',
            config_version='2.0',
            last_seen_date='2026-06-16',
            tags=['更新后'],
        )
        ok = kb_instance.add_known_issue('测试银行', updated_issue)
        assert ok is True

        entry = kb_instance.get_entry('测试银行')
        assert len(entry.known_issues) == 1
        issue = entry.known_issues[0]
        assert issue.title == '更新后的标题'
        assert issue.severity == 'error'
        assert len(issue.error_patterns) == 1
        assert '新根因' in issue.root_cause
        assert len(entry.template_screenshots) == 1
        assert len(entry.column_descriptions) == 2


class TestTemplateScreenshotFiles:
    def test_screenshots_directory_exists(self):
        script_dir = os.path.join(
            os.path.dirname(__file__), '..', 'backend'
        )
        screenshots_dir = os.path.join(script_dir, 'screenshots')
        assert os.path.isdir(screenshots_dir)

    def test_all_bank_screenshots_exist(self):
        script_dir = os.path.join(
            os.path.dirname(__file__), '..', 'backend'
        )
        expected_files = [
            'beijing_bank_template.png',
            'east_asia_bank_template.png',
            'icbc_template.png',
            'ccb_template.png',
            'cmb_template.png',
        ]
        for filename in expected_files:
            filepath = os.path.join(script_dir, 'screenshots', filename)
            assert os.path.exists(filepath), f'截图文件不存在: {filename}'
            assert os.path.getsize(filepath) > 0, f'截图文件为空: {filename}'

    def test_knowledge_base_screenshot_paths_valid(self):
        knowledge_path = os.path.join(
            os.path.dirname(__file__), '..', 'backend', 'bank_knowledge.yaml'
        )
        if not os.path.exists(knowledge_path):
            pytest.skip('bank_knowledge.yaml 不存在')

        BankKnowledgeBase._instance = None
        try:
            kb = BankKnowledgeBase(knowledge_path=knowledge_path)
            names = kb.get_all_bank_names()
            assert len(names) >= 5

            script_dir = os.path.dirname(knowledge_path)
            for name in names:
                screenshots = kb.get_template_screenshots(name)
                for ss in screenshots:
                    assert ss.file_path, f'{name} 的截图路径为空'
                    full_path = os.path.join(script_dir, ss.file_path)
                    assert os.path.exists(full_path), f'截图文件不存在: {ss.file_path}'
                    assert ss.upload_date, f'{name} 的截图缺少 upload_date'
                    assert ss.name, f'{name} 的截图缺少 name'
        finally:
            BankKnowledgeBase._instance = None

    def test_screenshot_metadata_complete(self):
        knowledge_path = os.path.join(
            os.path.dirname(__file__), '..', 'backend', 'bank_knowledge.yaml'
        )
        if not os.path.exists(knowledge_path):
            pytest.skip('bank_knowledge.yaml 不存在')

        BankKnowledgeBase._instance = None
        try:
            kb = BankKnowledgeBase(knowledge_path=knowledge_path)
            for name in kb.get_all_bank_names():
                screenshots = kb.get_template_screenshots(name)
                for ss in screenshots:
                    assert ss.upload_date, f'{name} - {ss.name} 缺少 upload_date'
                    assert ss.config_version, f'{name} - {ss.name} 缺少 config_version'
                    assert ss.description, f'{name} - {ss.name} 缺少 description'
        finally:
            BankKnowledgeBase._instance = None


class TestFailureLogToKnowledgeBaseResolution:
    def test_beijing_bank_header_mismatch_resolution(self, kb_instance):
        kb_instance.add_or_update_entry(sample_entry_with_issues())

        log_lines = [
            '[INFO] 开始处理北京银行流水文件',
            '[INFO] 识别银行: 北京银行',
            '[ERROR] 表头不匹配 北京银行，预期第4行为表头',
            '[WARNING] 数据解析异常，列映射错位',
            '[ERROR] 处理失败: header mismatch for beijing bank',
            '[ERROR] 退出码: 1',
        ]

        results = diagnose_from_knowledge_base(log_lines, bank_name='北京银行')

        assert len(results) >= 1, '应匹配到至少一个已知问题'
        assert results[0].match_score >= 0.4, f'匹配分数过低: {results[0].match_score}'

        issue = results[0].issue
        assert issue.title == '北京银行表头行偏移'
        assert issue.issue_id == 'bj_header_shift'
        assert len(issue.fix_steps) >= 2
        assert any('start_row' in step for step in issue.fix_steps)
        assert any('修改 bank_rules.yaml' in step for step in issue.fix_steps)

        matched_patterns = results[0].matched_patterns
        assert len(matched_patterns) >= 2
        assert any('error_pattern' in p for p in matched_patterns)

    def test_icbc_column_order_mismatch_resolution(self):
        knowledge_path = os.path.join(
            os.path.dirname(__file__), '..', 'backend', 'bank_knowledge.yaml'
        )
        if not os.path.exists(knowledge_path):
            pytest.skip('bank_knowledge.yaml 不存在')

        BankKnowledgeBase._instance = None
        try:
            kb = BankKnowledgeBase(knowledge_path=knowledge_path)

            log_lines = [
                '[INFO] 处理工商银行流水',
                '[ERROR] 收入金额显示为负数 工商银行',
                '[WARNING] 支出金额全部为0，可能列顺序错误',
                '[ERROR] 余额计算异常',
            ]

            results = kb.search_all('收入 支出 反 工商', bank_name='工商银行')
            assert len(results) >= 1

            issue = results[0].issue
            assert '收入' in issue.title or '支出' in issue.title or '顺序' in issue.title
            assert len(issue.fix_steps) >= 2
            assert any('receipt' in step.lower() or 'payment' in step.lower()
                       for step in issue.fix_steps)

        finally:
            BankKnowledgeBase._instance = None

    def test_east_asia_merged_column_resolution(self):
        knowledge_path = os.path.join(
            os.path.dirname(__file__), '..', 'backend', 'bank_knowledge.yaml'
        )
        if not os.path.exists(knowledge_path):
            pytest.skip('bank_knowledge.yaml 不存在')

        BankKnowledgeBase._instance = None
        try:
            kb = BankKnowledgeBase(knowledge_path=knowledge_path)

            log_lines = [
                '[INFO] 处理东亚银行流水',
                '[WARNING] 摘要列和对方户名列内容相同 东亚银行',
                '[ERROR] 无法单独提取对方户名',
            ]

            results = diagnose_from_knowledge_base(log_lines, bank_name='东亚银行')
            assert len(results) >= 1

            issue = results[0].issue
            assert ('合并' in issue.title or '对方户名' in issue.title)
            assert issue.severity == 'info' or issue.severity == 'warning'

        finally:
            BankKnowledgeBase._instance = None

    def test_search_returns_complete_solution(self):
        knowledge_path = os.path.join(
            os.path.dirname(__file__), '..', 'backend', 'bank_knowledge.yaml'
        )
        if not os.path.exists(knowledge_path):
            pytest.skip('bank_knowledge.yaml 不存在')

        BankKnowledgeBase._instance = None
        try:
            kb = BankKnowledgeBase(knowledge_path=knowledge_path)

            error_msg = '表头不匹配 北京银行 start_row 配置错误'
            results = kb.search_all(error_msg, bank_name='北京银行')

            assert len(results) >= 1
            best_match = results[0]

            assert best_match.issue.root_cause, '应包含根因分析'
            assert len(best_match.issue.fix_steps) >= 2, '应包含修复步骤'
            assert best_match.issue.config_version, '应关联配置版本'
            assert best_match.issue.error_patterns, '应包含错误模式'

            assert best_match.match_score >= 0.3
            assert len(best_match.matched_patterns) >= 1

        finally:
            BankKnowledgeBase._instance = None

    def test_cross_bank_error_search(self):
        knowledge_path = os.path.join(
            os.path.dirname(__file__), '..', 'backend', 'bank_knowledge.yaml'
        )
        if not os.path.exists(knowledge_path):
            pytest.skip('bank_knowledge.yaml 不存在')

        BankKnowledgeBase._instance = None
        try:
            kb = BankKnowledgeBase(knowledge_path=knowledge_path)

            error_msg = '工商银行 收入支出列顺序 金额负数 招商银行 收入 支出 反'
            results = kb.search_all(error_msg, min_score=0.1)

            assert len(results) >= 2, f'应匹配到多家银行的列顺序问题，实际: {len(results)}'

            bank_names = [r.bank_name for r in results]
            assert '工商银行' in bank_names or '招商银行' in bank_names or '建设银行' in bank_names

            results.sort(key=lambda r: r.match_score, reverse=True)
            assert results[0].match_score >= results[-1].match_score

        finally:
            BankKnowledgeBase._instance = None

    def test_troubleshooter_integration_with_knowledge_base(self):
        try:
            from troubleshooter import run_troubleshooting, TroubleshootingReport
        except ImportError:
            pytest.skip('troubleshooter 模块未安装')

        knowledge_path = os.path.join(
            os.path.dirname(__file__), '..', 'backend', 'bank_knowledge.yaml'
        )
        if not os.path.exists(knowledge_path):
            pytest.skip('bank_knowledge.yaml 不存在')

        BankKnowledgeBase._instance = None

        log_content = """[INFO] 2026-06-16 10:00:00 启动银行流水处理
[INFO] 2026-06-16 10:00:01 加载银行规则配置
[INFO] 2026-06-16 10:00:02 识别到银行: 北京银行
[ERROR] 2026-06-16 10:00:03 表头不匹配 北京银行，预期第4行为表头
[WARNING] 2026-06-16 10:00:03 数据解析异常，列映射错位
[ERROR] 2026-06-16 10:00:04 处理失败: header mismatch for beijing bank
[INFO] 2026-06-16 10:00:04 正在尝试自救修复
[ERROR] 2026-06-16 10:00:05 自救修复失败，退出
"""

        import tempfile
        with tempfile.NamedTemporaryFile(mode='w', suffix='.log', delete=False) as f:
            f.write(log_content)
            log_path = f.name

        try:
            class MockResult:
                bank_name = '北京银行'
                success = False
                last_bank_name = '北京银行'

            report = run_troubleshooting(
                log_path=log_path,
                processing_result=MockResult(),
            )

            assert isinstance(report, TroubleshootingReport)
            assert report.total_issues >= 1

            assert hasattr(report, 'knowledge_base_matches')
            assert len(report.knowledge_base_matches) >= 1

            kb_match = report.knowledge_base_matches[0]
            assert 'bank_name' in kb_match
            assert 'issue' in kb_match
            assert 'match_score' in kb_match
            assert kb_match['bank_name'] == '北京银行'
            assert kb_match['match_score'] >= 0.3

            issue_data = kb_match['issue']
            assert 'fix_steps' in issue_data
            assert len(issue_data['fix_steps']) >= 2
            assert 'root_cause' in issue_data
            assert 'title' in issue_data

            found_kb_issue = False
            for issue in report.issues:
                if issue.issue_id == 'knowledge_base_match':
                    found_kb_issue = True
                    assert '知识库匹配到' in issue.title
                    assert len(issue.fix_steps) >= 2
                    assert len(issue.evidence) >= 1
                    assert any('北京银行' in ev for ev in issue.evidence)
                    assert any('匹配度' in ev for ev in issue.evidence)
                    assert any('表头行偏移' in ev or '乱码' in ev for ev in issue.evidence)
                    break
            assert found_kb_issue, '应包含知识库匹配的问题条目'

        finally:
            os.unlink(log_path)
            BankKnowledgeBase._instance = None

    def test_encoding_error_diagnosis(self):
        knowledge_path = os.path.join(
            os.path.dirname(__file__), '..', 'backend', 'bank_knowledge.yaml'
        )
        if not os.path.exists(knowledge_path):
            pytest.skip('bank_knowledge.yaml 不存在')

        BankKnowledgeBase._instance = None
        try:
            kb = BankKnowledgeBase(knowledge_path=knowledge_path)

            log_lines = [
                '[ERROR] UnicodeDecodeError 北京银行 CSV文件',
                '[WARNING] 摘要列显示乱码',
                '[ERROR] 表头识别失败 编码错误',
            ]

            results = diagnose_from_knowledge_base(log_lines, bank_name='北京银行')
            assert len(results) >= 1

            encoding_issue = None
            for r in results:
                if '乱码' in r.issue.title or '编码' in r.issue.title:
                    encoding_issue = r
                    break

            assert encoding_issue is not None, '应匹配到编码相关问题'
            assert 'UTF-8' in ' '.join(encoding_issue.issue.fix_steps) or \
                   'GBK' in ' '.join(encoding_issue.issue.fix_steps)
            assert encoding_issue.issue.severity == 'warning'

        finally:
            BankKnowledgeBase._instance = None


def sample_entry_with_issues():
    return BankWikiEntry(
        bank_name='北京银行',
        display_name='北京银行',
        known_issues=[
            KnownIssue(
                issue_id='bj_header_shift',
                title='北京银行表头行偏移',
                description='表头不在预期行',
                symptoms=['数据解析异常', '列映射错位'],
                root_cause='网银系统升级后导出模板变化',
                fix_steps=[
                    '打开流水文件确认实际表头行位置',
                    '修改 bank_rules.yaml 中北京银行的 start_row 配置',
                    '重新运行流水处理',
                ],
                error_patterns=[
                    r'表头.*不匹配.*北京银行',
                    r'header.*mismatch.*beijing',
                ],
                severity='warning',
                config_version='1.0',
            ),
            KnownIssue(
                issue_id='bj_encoding_garbled',
                title='北京银行导出文件中文乱码',
                description='编码问题导致乱码',
                symptoms=['摘要列显示乱码', '表头识别失败'],
                root_cause='CSV文件使用GBK编码',
                fix_steps=[
                    '将文件另存为 UTF-8 编码',
                    '或另存为 xlsx 格式',
                ],
                error_patterns=[
                    r'乱码',
                    r'UnicodeDecodeError.*北京银行',
                    r'gbk.*utf',
                ],
                severity='warning',
                config_version='1.0',
            ),
        ],
        template_screenshots=[
            TemplateScreenshot(
                name='北京银行标准模板',
                description='标准导出格式',
                file_path='screenshots/beijing_bank_template.png',
                upload_date='2026-06-16',
                config_version='1.0',
            ),
        ],
        column_descriptions=[
            ColumnDescription(field_key='trade_date', display_name='交易日期'),
            ColumnDescription(field_key='payment', display_name='支出金额'),
        ],
    )
