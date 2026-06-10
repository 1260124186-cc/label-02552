import os
import json

import pytest

import bankcheck
from bankcheck import (
    CounterpartyRule,
    CounterpartyRuleConfig,
    get_counterparty_rule_config,
    apply_counterparty_rules,
    export_counterparty_tags,
    add_counterparty_keyword_rule,
    _match_counterparty,
)


def _make_record(uid='1', counterparty='北京XX科技有限公司'):
    return {
        '唯一id': uid,
        '银行': '北京银行',
        '银行账号': '123',
        '主体': '测试',
        '交易日期': '2024-01-01',
        '付款': -100,
        '收款': None,
        '摘要': '转账',
        '对方户名': counterparty,
        '余额': 1000,
        '交易流水号': 'T001',
    }


def _reset_singleton():
    bankcheck._counterparty_rule_config_instance = None


@pytest.fixture(autouse=True)
def reset_counterparty_singleton():
    _reset_singleton()
    yield
    _reset_singleton()


class TestCounterpartyRuleConfig:

    def test_create_rule(self, tmp_dir):
        config = CounterpartyRuleConfig(script_dir=tmp_dir)
        rule = CounterpartyRule(
            rule_id='',
            name='测试规则',
            rule_type='blacklist',
            keywords=['测试'],
            match_mode='contains',
        )
        rule_id = config.add_rule(rule)
        assert rule_id
        assert rule.rule_id == rule_id
        with open(config.config_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        assert len(data['rules']) == 1
        assert data['rules'][0]['name'] == '测试规则'

    def test_get_rules(self, tmp_dir):
        config = CounterpartyRuleConfig(script_dir=tmp_dir)
        config.add_rule(CounterpartyRule(
            rule_id='', name='黑1', rule_type='blacklist', keywords=['a'],
        ))
        config.add_rule(CounterpartyRule(
            rule_id='', name='白1', rule_type='whitelist', keywords=['b'],
        ))
        config.add_rule(CounterpartyRule(
            rule_id='', name='黑2', rule_type='blacklist', keywords=['c'], enabled=False,
        ))
        all_rules = config.get_rules()
        assert len(all_rules) == 3
        blacklist = config.get_rules(rule_type='blacklist')
        assert len(blacklist) == 2
        whitelist = config.get_rules(rule_type='whitelist')
        assert len(whitelist) == 1
        enabled_rules = config.get_rules(enabled=True)
        assert len(enabled_rules) == 2
        enabled_blacklist = config.get_rules(rule_type='blacklist', enabled=True)
        assert len(enabled_blacklist) == 1

    def test_update_rule(self, tmp_dir):
        config = CounterpartyRuleConfig(script_dir=tmp_dir)
        rule_id = config.add_rule(CounterpartyRule(
            rule_id='', name='旧名', rule_type='blacklist', keywords=['旧词'],
        ))
        ok = config.update_rule(rule_id, {'name': '新名', 'keywords': ['新词']})
        assert ok is True
        rules = config.get_rules()
        assert rules[0].name == '新名'
        assert rules[0].keywords == ['新词']
        ok_missing = config.update_rule('nonexistent', {'name': 'x'})
        assert ok_missing is False

    def test_delete_rule(self, tmp_dir):
        config = CounterpartyRuleConfig(script_dir=tmp_dir)
        rule_id = config.add_rule(CounterpartyRule(
            rule_id='', name='删除测试', rule_type='blacklist', keywords=['a'],
        ))
        assert len(config.get_rules()) == 1
        ok = config.delete_rule(rule_id)
        assert ok is True
        assert len(config.get_rules()) == 0
        ok_missing = config.delete_rule('nonexistent')
        assert ok_missing is False

    def test_toggle_rule(self, tmp_dir):
        config = CounterpartyRuleConfig(script_dir=tmp_dir)
        rule_id = config.add_rule(CounterpartyRule(
            rule_id='', name='开关测试', rule_type='blacklist', keywords=['a'],
        ))
        rule = config.get_rules()[0]
        assert rule.enabled is True
        ok = config.toggle_rule(rule_id, False)
        assert ok is True
        assert config.get_rules()[0].enabled is False
        ok = config.toggle_rule(rule_id, True)
        assert ok is True
        assert config.get_rules()[0].enabled is True
        ok_missing = config.toggle_rule('nonexistent', False)
        assert ok_missing is False


class TestCounterpartyMatching:

    def test_match_contains(self):
        rule = CounterpartyRule(
            rule_id='r1', name='包含', rule_type='blacklist',
            keywords=['科技'], match_mode='contains',
        )
        assert _match_counterparty('北京XX科技有限公司', rule) == '科技'
        assert _match_counterparty('北京XX贸易公司', rule) is None

    def test_match_exact(self):
        rule = CounterpartyRule(
            rule_id='r2', name='精确', rule_type='blacklist',
            keywords=['北京XX科技有限公司'], match_mode='exact',
        )
        assert _match_counterparty('北京XX科技有限公司', rule) == '北京XX科技有限公司'
        assert _match_counterparty('北京XX科技有限公司分部', rule) is None

    def test_match_startswith(self):
        rule = CounterpartyRule(
            rule_id='r3', name='前缀', rule_type='blacklist',
            keywords=['北京'], match_mode='startswith',
        )
        assert _match_counterparty('北京XX科技有限公司', rule) == '北京'
        assert _match_counterparty('上海XX科技有限公司', rule) is None

    def test_match_endswith(self):
        rule = CounterpartyRule(
            rule_id='r4', name='后缀', rule_type='blacklist',
            keywords=['有限公司'], match_mode='endswith',
        )
        assert _match_counterparty('北京XX科技有限公司', rule) == '有限公司'
        assert _match_counterparty('北京XX科技合伙企业', rule) is None

    def test_match_regex(self):
        rule = CounterpartyRule(
            rule_id='r5', name='正则', rule_type='blacklist',
            keywords=['北京.*科技'], match_mode='regex',
        )
        assert _match_counterparty('北京XX科技有限公司', rule) == '北京.*科技'
        assert _match_counterparty('上海XX科技有限公司', rule) is None

    def test_no_match(self, tmp_dir):
        config = CounterpartyRuleConfig(script_dir=tmp_dir)
        config.add_rule(CounterpartyRule(
            rule_id='', name='无命中', rule_type='blacklist', keywords=['不存在的关键词'],
        ))
        records = [_make_record()]
        result, summary = apply_counterparty_rules(records, script_dir=tmp_dir)
        assert summary['tagged_count'] == 0
        assert result[0]['黑白名单标签'] == ''

    def test_blacklist_and_whitelist(self, tmp_dir):
        config = CounterpartyRuleConfig(script_dir=tmp_dir)
        config.add_rule(CounterpartyRule(
            rule_id='', name='黑名单规则', rule_type='blacklist', keywords=['科技'],
        ))
        config.add_rule(CounterpartyRule(
            rule_id='', name='白名单规则', rule_type='whitelist', keywords=['北京'],
        ))
        records = [_make_record(counterparty='北京XX科技有限公司')]
        result, summary = apply_counterparty_rules(records, script_dir=tmp_dir)
        assert summary['tagged_count'] == 1
        assert summary['blacklist_hits'] == 1
        assert summary['whitelist_hits'] == 1
        assert '黑名单' in result[0]['黑白名单标签']
        assert '白名单' in result[0]['黑白名单标签']

    def test_disabled_rule_skipped(self, tmp_dir):
        config = CounterpartyRuleConfig(script_dir=tmp_dir)
        config.add_rule(CounterpartyRule(
            rule_id='', name='禁用规则', rule_type='blacklist',
            keywords=['科技'], enabled=False,
        ))
        records = [_make_record(counterparty='北京XX科技有限公司')]
        result, summary = apply_counterparty_rules(records, script_dir=tmp_dir)
        assert summary['tagged_count'] == 0
        assert result[0]['黑白名单标签'] == ''

    def test_empty_counterparty(self, tmp_dir):
        config = CounterpartyRuleConfig(script_dir=tmp_dir)
        config.add_rule(CounterpartyRule(
            rule_id='', name='测试', rule_type='blacklist', keywords=['科技'],
        ))
        records_empty = [_make_record(counterparty='')]
        records_none = [_make_record(counterparty=None)]
        for recs in [records_empty, records_none]:
            result, summary = apply_counterparty_rules(recs, script_dir=tmp_dir)
            assert summary['tagged_count'] == 0
            assert result[0]['黑白名单标签'] == ''

    def test_multiple_rule_hits(self, tmp_dir):
        config = CounterpartyRuleConfig(script_dir=tmp_dir)
        config.add_rule(CounterpartyRule(
            rule_id='', name='规则A', rule_type='blacklist', keywords=['科技'],
        ))
        config.add_rule(CounterpartyRule(
            rule_id='', name='规则B', rule_type='blacklist', keywords=['北京'],
        ))
        config.add_rule(CounterpartyRule(
            rule_id='', name='规则C', rule_type='whitelist', keywords=['有限公司'],
        ))
        records = [_make_record(counterparty='北京XX科技有限公司')]
        result, summary = apply_counterparty_rules(records, script_dir=tmp_dir)
        assert summary['tagged_count'] == 1
        assert summary['blacklist_hits'] == 2
        assert summary['whitelist_hits'] == 1
        tags = result[0]['黑白名单标签']
        assert '规则A' in tags
        assert '规则B' in tags
        assert '规则C' in tags
        names = result[0]['命中规则名称']
        assert '规则A' in names
        assert '规则B' in names
        assert '规则C' in names
        kw = result[0]['命中关键词']
        assert '科技' in kw
        assert '北京' in kw
        assert '有限公司' in kw


class TestApplyCounterpartyRules:

    def test_apply_returns_summary(self, tmp_dir):
        config = CounterpartyRuleConfig(script_dir=tmp_dir)
        config.add_rule(CounterpartyRule(
            rule_id='', name='测试', rule_type='blacklist', keywords=['科技'],
        ))
        records = [_make_record()]
        result, summary = apply_counterparty_rules(records, script_dir=tmp_dir)
        assert 'total_records' in summary
        assert 'tagged_count' in summary
        assert 'blacklist_hits' in summary
        assert 'whitelist_hits' in summary
        assert 'rule_hit_counts' in summary
        assert summary['total_records'] == 1

    def test_apply_adds_tag_fields(self, tmp_dir):
        config = CounterpartyRuleConfig(script_dir=tmp_dir)
        config.add_rule(CounterpartyRule(
            rule_id='', name='黑名单规则', rule_type='blacklist', keywords=['科技'],
        ))
        records = [_make_record()]
        result, _ = apply_counterparty_rules(records, script_dir=tmp_dir)
        assert '黑白名单标签' in result[0]
        assert '命中规则名称' in result[0]
        assert '命中关键词' in result[0]
        assert '黑名单' in result[0]['黑白名单标签']
        assert result[0]['命中规则名称'] == '黑名单规则'
        assert result[0]['命中关键词'] == '科技'


class TestExportCounterpartyTags:

    def test_export_creates_file(self, tmp_dir):
        config = CounterpartyRuleConfig(script_dir=tmp_dir)
        config.add_rule(CounterpartyRule(
            rule_id='', name='导出测试', rule_type='blacklist', keywords=['科技'],
        ))
        records = [_make_record()]
        records, _ = apply_counterparty_rules(records, script_dir=tmp_dir)
        output_path = os.path.join(tmp_dir, 'tags_export.xlsx')
        result = export_counterparty_tags(records, output_path)
        assert result == output_path
        assert os.path.exists(output_path)

    def test_export_only_tagged(self, tmp_dir):
        config = CounterpartyRuleConfig(script_dir=tmp_dir)
        config.add_rule(CounterpartyRule(
            rule_id='', name='有标签', rule_type='blacklist', keywords=['科技'],
        ))
        records = [
            _make_record(uid='1', counterparty='北京XX科技有限公司'),
            _make_record(uid='2', counterparty='上海YY贸易公司'),
        ]
        records, _ = apply_counterparty_rules(records, script_dir=tmp_dir)
        output_path = os.path.join(tmp_dir, 'tags_filtered.xlsx')
        export_counterparty_tags(records, output_path)
        import pandas as pd
        df = pd.read_excel(output_path, engine='openpyxl')
        assert len(df) == 1
        assert '科技' in str(df.iloc[0]['命中关键词'])


class TestConvenienceFunction:

    def test_add_counterparty_keyword_rule(self, tmp_dir):
        rule_id = add_counterparty_keyword_rule(
            name='便捷规则',
            rule_type='blacklist',
            keywords=['便捷'],
            match_mode='contains',
            script_dir=tmp_dir,
        )
        assert rule_id
        config = get_counterparty_rule_config(script_dir=tmp_dir)
        rules = config.get_rules()
        found = [r for r in rules if r.rule_id == rule_id]
        assert len(found) == 1
        assert found[0].name == '便捷规则'
        assert found[0].rule_type == 'blacklist'
        assert found[0].keywords == ['便捷']
        assert found[0].enabled is True
