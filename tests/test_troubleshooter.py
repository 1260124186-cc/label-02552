#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
智能排障助手模块测试
"""

import os
import sys
import tempfile
import json

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'backend'))

import troubleshooter as ts
from troubleshooter import (
    RuleEngine,
    TroubleshootingIssue,
    TroubleshootingReport,
    SeverityLevel,
    read_log_file,
    find_latest_log_file,
    run_troubleshooting,
)


class TestSeverityLevel:
    def test_severity_values(self):
        assert SeverityLevel.CRITICAL.value == 'critical'
        assert SeverityLevel.ERROR.value == 'error'
        assert SeverityLevel.WARNING.value == 'warning'
        assert SeverityLevel.INFO.value == 'info'


class TestTroubleshootingIssue:
    def test_create_issue(self):
        issue = TroubleshootingIssue(
            issue_id='test_issue',
            title='测试问题',
            severity=SeverityLevel.WARNING,
            description='这是一个测试问题',
            evidence=['证据1', '证据2'],
            fix_steps=['步骤1', '步骤2'],
            confidence=0.85,
        )
        assert issue.issue_id == 'test_issue'
        assert issue.title == '测试问题'
        assert issue.severity == SeverityLevel.WARNING
        assert issue.description == '这是一个测试问题'
        assert len(issue.evidence) == 2
        assert len(issue.fix_steps) == 2
        assert issue.confidence == 0.85

    def test_to_dict(self):
        issue = TroubleshootingIssue(
            issue_id='test',
            title='测试',
            severity=SeverityLevel.ERROR,
            description='描述',
            confidence=0.75,
        )
        d = issue.to_dict()
        assert d['issue_id'] == 'test'
        assert d['severity'] == 'error'
        assert d['confidence'] == 0.75
        assert isinstance(d['evidence'], list)
        assert isinstance(d['fix_steps'], list)


class TestTroubleshootingReport:
    def test_create_report(self):
        report = TroubleshootingReport(
            timestamp='2024-01-01 12:00:00',
            log_file='/path/to/log',
            total_issues=2,
        )
        assert report.timestamp == '2024-01-01 12:00:00'
        assert report.log_file == '/path/to/log'
        assert report.total_issues == 2

    def test_to_dict(self):
        issue = TroubleshootingIssue(
            issue_id='test',
            title='测试',
            severity=SeverityLevel.WARNING,
            description='描述',
            confidence=0.8,
        )
        report = TroubleshootingReport(
            timestamp='2024-01-01',
            log_file='test.log',
            total_issues=1,
            issues=[issue],
            summary='测试总结',
        )
        d = report.to_dict()
        assert d['total_issues'] == 1
        assert len(d['issues']) == 1
        assert d['summary'] == '测试总结'


class TestReadLogFile:
    def test_read_existing_log(self, tmp_path):
        log_file = tmp_path / 'bankcheck.log'
        log_content = (
            '[2024-01-01 10:00:00] INFO - 开始处理\n'
            '[2024-01-01 10:00:01] WARNING - 未找到主体查找表\n'
            '[2024-01-01 10:00:02] ERROR - 处理失败\n'
        )
        log_file.write_text(log_content, encoding='utf-8')

        lines = read_log_file(str(log_file))
        assert len(lines) == 3
        assert '未找到主体查找表' in lines[1]

    def test_read_nonexistent_log(self):
        lines = read_log_file('/nonexistent/path/bankcheck.log')
        assert lines == []

    def test_read_log_max_lines(self, tmp_path):
        log_file = tmp_path / 'bankcheck.log'
        lines_content = [f'line {i}\n' for i in range(100)]
        log_file.write_text(''.join(lines_content), encoding='utf-8')

        lines = read_log_file(str(log_file), max_lines=10)
        assert len(lines) == 10
        assert 'line 90' in lines[0]


class TestFindLatestLogFile:
    def test_find_log_in_script_dir(self, tmp_path, monkeypatch):
        log_file = tmp_path / 'bankcheck.log'
        log_file.write_text('test log', encoding='utf-8')

        monkeypatch.setattr(ts, 'get_script_dir', lambda: str(tmp_path))

        found = find_latest_log_file(str(tmp_path))
        assert found is not None
        assert 'bankcheck.log' in found

    def test_find_no_log(self, tmp_path):
        found = find_latest_log_file(str(tmp_path))
        assert found is None


class MockProcessingResult:
    def __init__(self, **kwargs):
        self.lookup_missing = kwargs.get('lookup_missing', False)
        self.unprocessed_files = kwargs.get('unprocessed_files', [])
        self.error_files = kwargs.get('error_files', [])
        self.folder_empty = kwargs.get('folder_empty', False)


class TestRuleEngine:
    def test_init_engine(self):
        engine = RuleEngine()
        assert len(engine.rules) > 0

    def test_check_lookup_missing_with_result(self):
        engine = RuleEngine()
        result = MockProcessingResult(lookup_missing=True)
        issues = engine.analyze([], result)

        lookup_issues = [i for i in issues if i.issue_id == 'lookup_missing']
        assert len(lookup_issues) == 1
        assert lookup_issues[0].confidence > 0.5

    def test_check_lookup_missing_with_log(self):
        engine = RuleEngine()
        log_lines = [
            '[2024-01-01 10:00:00] WARNING - 未找到主体查找表，"主体"列将为空',
        ]
        issues = engine.analyze(log_lines, None)

        lookup_issues = [i for i in issues if i.issue_id == 'lookup_missing']
        assert len(lookup_issues) == 1
        assert lookup_issues[0].confidence > 0

    def test_check_unrecognized_files_with_result(self):
        engine = RuleEngine()
        result = MockProcessingResult(
            unprocessed_files=['/path/未知银行_流水.xlsx', '/path/某某银行.xlsx']
        )
        issues = engine.analyze([], result)

        unrec_issues = [i for i in issues if i.issue_id == 'unrecognized_files']
        assert len(unrec_issues) == 1
        assert unrec_issues[0].confidence > 0.3

    def test_check_unrecognized_files_with_log(self):
        engine = RuleEngine()
        log_lines = [
            '[2024-01-01 10:00:00] WARNING - 文件「未知银行.xlsx」无法识别银行类型',
        ]
        issues = engine.analyze(log_lines, None)

        unrec_issues = [i for i in issues if i.issue_id == 'unrecognized_files']
        assert len(unrec_issues) == 1

    def test_check_xls_dependency_with_log(self):
        engine = RuleEngine()
        log_lines = [
            '[2024-01-01 10:00:00] ERROR - 处理 .xls 文件需要 xlrd 库',
            '[2024-01-01 10:00:01] ERROR - 缺少 xlrd 库，无法处理 .xls 文件',
        ]
        issues = engine.analyze(log_lines, None)

        xls_issues = [i for i in issues if i.issue_id == 'xls_missing_dependency']
        assert len(xls_issues) == 1
        assert xls_issues[0].confidence > 0.3

    def test_check_xls_dependency_with_result(self):
        engine = RuleEngine()
        result = MockProcessingResult(
            error_files=[('/path/test.xls', "ImportError: 缺少 xlrd 库")]
        )
        issues = engine.analyze([], result)

        xls_issues = [i for i in issues if i.issue_id == 'xls_missing_dependency']
        assert len(xls_issues) == 1

    def test_check_empty_folder_with_result(self):
        engine = RuleEngine()
        result = MockProcessingResult(folder_empty=True)
        issues = engine.analyze([], result)

        empty_issues = [i for i in issues if i.issue_id == 'empty_folder']
        assert len(empty_issues) == 1

    def test_check_empty_folder_with_log(self):
        engine = RuleEngine()
        log_lines = [
            '[2024-01-01 10:00:00] WARNING - 文件夹为空，未找到 Excel 文件',
        ]
        issues = engine.analyze(log_lines, None)

        empty_issues = [i for i in issues if i.issue_id == 'empty_folder']
        assert len(empty_issues) == 1

    def test_check_file_format_error(self):
        engine = RuleEngine()
        log_lines = [
            '[2024-01-01 10:00:00] ERROR - 文件格式错误',
            '[2024-01-01 10:00:01] ERROR - BadZipFile: File is not a zip file',
        ]
        issues = engine.analyze(log_lines, None)

        format_issues = [i for i in issues if i.issue_id == 'file_format_error']
        assert len(format_issues) == 1

    def test_check_header_mismatch(self):
        engine = RuleEngine()
        log_lines = [
            '[2024-01-01 10:00:00] WARNING - 表头不匹配',
            '[2024-01-01 10:00:01] WARNING - 银行模板表头校验不通过',
        ]
        issues = engine.analyze(log_lines, None)

        header_issues = [i for i in issues if i.issue_id == 'header_mismatch']
        assert len(header_issues) == 1

    def test_check_permission_error(self):
        engine = RuleEngine()
        log_lines = [
            '[2024-01-01 10:00:00] ERROR - PermissionError: 权限不足',
        ]
        issues = engine.analyze(log_lines, None)

        perm_issues = [i for i in issues if i.issue_id == 'permission_error']
        assert len(perm_issues) == 1

    def test_analyze_no_issues(self):
        engine = RuleEngine()
        log_lines = [
            '[2024-01-01 10:00:00] INFO - 处理完成',
            '[2024-01-01 10:00:01] INFO - 成功提取 100 条记录',
        ]
        issues = engine.analyze(log_lines, None)
        assert len(issues) == 0

    def test_analyze_multiple_issues_sorted(self):
        engine = RuleEngine()
        log_lines = [
            '[2024-01-01 10:00:00] ERROR - 缺少 xlrd 库，无法处理 .xls 文件',
            '[2024-01-01 10:00:01] WARNING - 未找到主体查找表',
        ]
        issues = engine.analyze(log_lines, None)

        assert len(issues) >= 2

        severity_order = {'critical': 0, 'error': 1, 'warning': 2, 'info': 3}
        for i in range(len(issues) - 1):
            assert severity_order[issues[i].severity.value] <= severity_order[issues[i + 1].severity.value]

    def test_issue_has_fix_steps(self):
        engine = RuleEngine()
        log_lines = [
            '[2024-01-01 10:00:00] WARNING - 未找到主体查找表',
        ]
        issues = engine.analyze(log_lines, None)

        assert len(issues) > 0
        for issue in issues:
            assert len(issue.fix_steps) > 0

    def test_issue_has_evidence(self):
        engine = RuleEngine()
        log_lines = [
            '[2024-01-01 10:00:00] WARNING - 未找到主体查找表，"主体"列将为空',
        ]
        issues = engine.analyze(log_lines, None)

        assert len(issues) > 0
        for issue in issues:
            assert len(issue.evidence) > 0


class TestRunTroubleshooting:
    def test_run_with_log_and_result(self, tmp_path):
        log_file = tmp_path / 'bankcheck.log'
        log_content = (
            '[2024-01-01 10:00:00] INFO - 开始处理\n'
            '[2024-01-01 10:00:01] WARNING - 未找到主体查找表\n'
            '[2024-01-01 10:00:02] INFO - 处理完成\n'
        )
        log_file.write_text(log_content, encoding='utf-8')

        result = MockProcessingResult(lookup_missing=True)
        report = run_troubleshooting(
            log_path=str(log_file),
            processing_result=result,
            script_dir=str(tmp_path),
        )

        assert isinstance(report, TroubleshootingReport)
        assert report.log_file == str(log_file)
        assert report.total_issues > 0
        assert len(report.issues) == report.total_issues
        assert report.summary != ''

    def test_run_without_log(self, tmp_path):
        report = run_troubleshooting(
            log_path=None,
            processing_result=None,
            script_dir=str(tmp_path),
        )

        assert isinstance(report, TroubleshootingReport)

    def test_run_no_issues(self, tmp_path):
        log_file = tmp_path / 'bankcheck.log'
        log_content = (
            '[2024-01-01 10:00:00] INFO - 开始处理\n'
            '[2024-01-01 10:00:02] INFO - 处理完成\n'
        )
        log_file.write_text(log_content, encoding='utf-8')

        report = run_troubleshooting(
            log_path=str(log_file),
            processing_result=None,
            script_dir=str(tmp_path),
        )

        assert report.total_issues == 0
        assert '正常' in report.summary

    def test_print_report(self, capsys):
        issue = TroubleshootingIssue(
            issue_id='test',
            title='测试问题',
            severity=SeverityLevel.WARNING,
            description='测试描述',
            evidence=['证据1'],
            fix_steps=['步骤1'],
            confidence=0.8,
        )
        report = TroubleshootingReport(
            timestamp='2024-01-01 12:00:00',
            log_file='test.log',
            total_issues=1,
            issues=[issue],
            summary='测试总结',
        )

        ts.print_report(report)

        captured = capsys.readouterr()
        assert '智能排障助手' in captured.out
        assert '诊断报告' in captured.out
        assert '测试问题' in captured.out
        assert '测试总结' in captured.out


class TestIntegrationScenarios:
    def test_lookup_missing_scenario(self):
        log_lines = [
            '[2024-01-01 10:00:00] INFO - 开始处理文件夹',
            '[2024-01-01 10:00:01] WARNING - 未找到主体查找表，"主体"列将为空',
            '[2024-01-01 10:00:02] WARNING - 程序目录下未找到任何 Excel 文件作为主体查找表',
            '[2024-01-01 10:00:03] INFO - 正在处理文件: 北京银行_流水.xlsx',
            '[2024-01-01 10:00:05] INFO - 处理完成，共提取 50 条记录',
        ]
        result = MockProcessingResult(lookup_missing=True)

        engine = RuleEngine()
        issues = engine.analyze(log_lines, result)

        lookup_issues = [i for i in issues if i.issue_id == 'lookup_missing']
        assert len(lookup_issues) == 1
        assert lookup_issues[0].confidence > 0.7
        assert len(lookup_issues[0].fix_steps) > 0

    def test_unrecognized_files_scenario(self):
        log_lines = [
            '[2024-01-01 10:00:00] INFO - 开始处理文件夹',
            '[2024-01-01 10:00:01] INFO - 正在处理文件: 北京银行_流水.xlsx',
            '[2024-01-01 10:00:02] WARNING - 文件「未知银行A.xlsx」无法识别银行类型',
            '[2024-01-01 10:00:03] WARNING - 文件「某某银行流水.xlsx」无法识别银行类型',
            '[2024-01-01 10:00:04] INFO - 处理完成',
        ]
        result = MockProcessingResult(
            unprocessed_files=[
                '/path/未知银行A.xlsx',
                '/path/某某银行流水.xlsx',
            ]
        )

        engine = RuleEngine()
        issues = engine.analyze(log_lines, result)

        unrec_issues = [i for i in issues if i.issue_id == 'unrecognized_files']
        assert len(unrec_issues) == 1
        assert '无法识别' in unrec_issues[0].title

    def test_xls_dependency_scenario(self):
        log_lines = [
            '[2024-01-01 10:00:00] INFO - 开始处理文件夹',
            '[2024-01-01 10:00:01] INFO - 检测到 .xls 文件: 旧流水.xls',
            '[2024-01-01 10:00:02] ERROR - 处理 .xls 文件需要 xlrd 库，请运行: pip install xlrd',
            '[2024-01-01 10:00:03] ERROR - 缺少 xlrd 库，无法处理 .xls 文件。请运行: pip install xlrd',
            '[2024-01-01 10:00:04] WARNING - 文件处理失败: 旧流水.xls',
        ]
        result = MockProcessingResult(
            error_files=[('/path/旧流水.xls', "ImportError: 缺少 xlrd 库")]
        )

        engine = RuleEngine()
        issues = engine.analyze(log_lines, result)

        xls_issues = [i for i in issues if i.issue_id == 'xls_missing_dependency']
        assert len(xls_issues) == 1
        assert xls_issues[0].severity == SeverityLevel.ERROR
        assert any('xlrd' in step for step in xls_issues[0].fix_steps)

    def test_multiple_issues_scenario(self):
        log_lines = [
            '[2024-01-01 10:00:00] INFO - 开始处理文件夹',
            '[2024-01-01 10:00:01] WARNING - 未找到主体查找表',
            '[2024-01-01 10:00:02] WARNING - 文件「未知银行.xlsx」无法识别银行类型',
            '[2024-01-01 10:00:03] ERROR - 缺少 xlrd 库，无法处理 .xls 文件',
            '[2024-01-01 10:00:04] INFO - 处理完成',
        ]
        result = MockProcessingResult(
            lookup_missing=True,
            unprocessed_files=['/path/未知银行.xlsx'],
            error_files=[('/path/旧流水.xls', "ImportError: 缺少 xlrd 库")]
        )

        engine = RuleEngine()
        issues = engine.analyze(log_lines, result)

        assert len(issues) >= 3

        issue_ids = [i.issue_id for i in issues]
        assert 'lookup_missing' in issue_ids
        assert 'unrecognized_files' in issue_ids
        assert 'xls_missing_dependency' in issue_ids


if __name__ == '__main__':
    pytest.main([__file__, '-v'])
