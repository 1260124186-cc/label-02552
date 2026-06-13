import os
import json

import pytest

import bankcheck
from bankcheck import (
    CashflowClassificationRule,
    CashflowRuleConfig,
    CASHFLOW_CATEGORIES,
    CASHFLOW_SUBCATEGORIES,
    CASHFLOW_CATEGORY_HIERARCHY,
    get_cashflow_rule_config,
    get_default_cashflow_rules,
    apply_cashflow_classification,
    summarize_cashflow_by_category,
    export_cashflow_summary,
    add_cashflow_rule,
    init_default_cashflow_rules,
    get_cashflow_classification,
    _match_cashflow_text,
    _check_amount_direction,
    _reset_cashflow_singleton,
)


def _make_record(uid='1', summary='', counterpart='', payment=None, receipt=None):
    return {
        '唯一id': uid,
        '银行': '北京银行',
        '银行账号': '123',
        '主体': '测试主体',
        '交易日期': '2024-01-01',
        '付款': payment,
        '收款': receipt,
        '摘要': summary,
        '对方户名': counterpart,
        '余额': 1000,
        '交易流水号': f'T{uid}',
    }


def _reset_singleton():
    _reset_cashflow_singleton()


@pytest.fixture(autouse=True)
def reset_cashflow_singleton():
    _reset_singleton()
    yield
    _reset_singleton()


class TestCashflowConstants:

    def test_category_hierarchy_consistency(self):
        for sub_code in CASHFLOW_SUBCATEGORIES:
            assert sub_code in CASHFLOW_CATEGORY_HIERARCHY, f"子类别 {sub_code} 缺少层级映射"
            main_code = CASHFLOW_CATEGORY_HIERARCHY[sub_code]
            assert main_code in CASHFLOW_CATEGORIES, f"主类别 {main_code} 不存在"

    def test_category_names_not_empty(self):
        assert len(CASHFLOW_CATEGORIES) > 0
        assert len(CASHFLOW_SUBCATEGORIES) > 0
        assert len(CASHFLOW_CATEGORY_HIERARCHY) > 0


class TestCashflowClassificationRule:

    def test_rule_contains_match(self):
        rule = CashflowClassificationRule(
            rule_id='r1', name='测试', subcategory='salary',
            summary_keywords=['工资'], match_mode='contains')
        assert _match_cashflow_text('发放工资', rule.summary_keywords, rule.match_mode) == '工资'
        assert _match_cashflow_text('报销费用', rule.summary_keywords, rule.match_mode) is None

    def test_rule_exact_match(self):
        rule = CashflowClassificationRule(
            rule_id='r1', name='测试', subcategory='salary',
            summary_keywords=['工资'], match_mode='exact')
        assert _match_cashflow_text('工资', rule.summary_keywords, rule.match_mode) == '工资'
        assert _match_cashflow_text('发放工资', rule.summary_keywords, rule.match_mode) is None

    def test_rule_startswith_match(self):
        rule = CashflowClassificationRule(
            rule_id='r1', name='测试', subcategory='salary',
            summary_keywords=['工资'], match_mode='startswith')
        assert _match_cashflow_text('工资发放', rule.summary_keywords, rule.match_mode) == '工资'
        assert _match_cashflow_text('发放工资', rule.summary_keywords, rule.match_mode) is None

    def test_rule_endswith_match(self):
        rule = CashflowClassificationRule(
            rule_id='r1', name='测试', subcategory='salary',
            summary_keywords=['工资'], match_mode='endswith')
        assert _match_cashflow_text('发放工资', rule.summary_keywords, rule.match_mode) == '工资'
        assert _match_cashflow_text('工资发放', rule.summary_keywords, rule.match_mode) is None

    def test_rule_regex_match(self):
        rule = CashflowClassificationRule(
            rule_id='r1', name='测试', subcategory='salary',
            summary_keywords=[r'工资\d+'], match_mode='regex')
        assert _match_cashflow_text('工资2024', rule.summary_keywords, rule.match_mode) == r'工资\d+'
        assert _match_cashflow_text('工资发放', rule.summary_keywords, rule.match_mode) is None

    def test_rule_main_category_property(self):
        rule = CashflowClassificationRule(
            rule_id='r1', name='测试', subcategory='salary',
            summary_keywords=['工资'])
        assert rule.main_category == 'operating'

        rule2 = CashflowClassificationRule(
            rule_id='r2', name='测试', subcategory='loan_out',
            summary_keywords=['还款'])
        assert rule2.main_category == 'financing'


class TestAmountDirection:

    def test_any_direction(self):
        record_pay = _make_record(payment=-100)
        record_receive = _make_record(receipt=100)
        assert _check_amount_direction(record_pay, 'any') is True
        assert _check_amount_direction(record_receive, 'any') is True

    def test_payment_direction(self):
        record_pay = _make_record(payment=-100)
        record_receive = _make_record(receipt=100)
        assert _check_amount_direction(record_pay, 'payment') is True
        assert _check_amount_direction(record_receive, 'payment') is False

    def test_receipt_direction(self):
        record_pay = _make_record(payment=-100)
        record_receive = _make_record(receipt=100)
        assert _check_amount_direction(record_pay, 'receipt') is False
        assert _check_amount_direction(record_receive, 'receipt') is True

    def test_zero_amount(self):
        record_zero_pay = _make_record(payment=0)
        record_zero_receive = _make_record(receipt=0)
        assert _check_amount_direction(record_zero_pay, 'payment') is False
        assert _check_amount_direction(record_zero_receive, 'receipt') is False


class TestCashflowRuleConfig:

    def test_create_config(self, tmp_dir):
        config = CashflowRuleConfig(script_dir=tmp_dir)
        assert config.config_path == os.path.join(tmp_dir, 'cashflow_rules.json')
        assert os.path.exists(config.config_path)
        with open(config.config_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        assert 'rules' in data

    def test_default_rules_loaded(self, tmp_dir):
        config = CashflowRuleConfig(script_dir=tmp_dir)
        rules = config.get_rules()
        assert len(rules) > 0

    def test_add_rule(self, tmp_dir):
        config = CashflowRuleConfig(script_dir=tmp_dir)
        initial_count = len(config.get_rules())
        rule = CashflowClassificationRule(
            rule_id='',
            name='自定义规则',
            subcategory='operating_other',
            summary_keywords=['测试关键词'],
        )
        rule_id = config.add_rule(rule)
        assert rule_id
        assert rule.rule_id == rule_id
        assert len(config.get_rules()) == initial_count + 1

    def test_get_rules_by_subcategory(self, tmp_dir):
        config = CashflowRuleConfig(script_dir=tmp_dir)
        salary_rules = config.get_rules(subcategory='salary')
        assert len(salary_rules) > 0
        for r in salary_rules:
            assert r.subcategory == 'salary'

    def test_get_rules_by_main_category(self, tmp_dir):
        config = CashflowRuleConfig(script_dir=tmp_dir)
        operating_rules = config.get_rules(main_category='operating')
        assert len(operating_rules) > 0
        for r in operating_rules:
            assert CASHFLOW_CATEGORY_HIERARCHY.get(r.subcategory) == 'operating'

    def test_get_enabled_rules(self, tmp_dir):
        config = CashflowRuleConfig(script_dir=tmp_dir)
        all_rules = config.get_rules()
        enabled_rules = config.get_rules(enabled=True)
        assert len(enabled_rules) == len(all_rules)

    def test_update_rule(self, tmp_dir):
        config = CashflowRuleConfig(script_dir=tmp_dir)
        rule_id = config.add_rule(CashflowClassificationRule(
            rule_id='', name='旧名', subcategory='salary', summary_keywords=['旧词']))
        ok = config.update_rule(rule_id, {'name': '新名', 'summary_keywords': ['新词']})
        assert ok is True
        rules = config.get_rules()
        found = [r for r in rules if r.rule_id == rule_id][0]
        assert found.name == '新名'
        assert found.summary_keywords == ['新词']

    def test_delete_rule(self, tmp_dir):
        config = CashflowRuleConfig(script_dir=tmp_dir)
        initial_count = len(config.get_rules())
        rule_id = config.add_rule(CashflowClassificationRule(
            rule_id='', name='删除测试', subcategory='salary', summary_keywords=['a']))
        assert len(config.get_rules()) == initial_count + 1
        ok = config.delete_rule(rule_id)
        assert ok is True
        assert len(config.get_rules()) == initial_count

    def test_toggle_rule(self, tmp_dir):
        config = CashflowRuleConfig(script_dir=tmp_dir)
        rule_id = config.add_rule(CashflowClassificationRule(
            rule_id='', name='开关测试', subcategory='salary', summary_keywords=['a']))
        rule = [r for r in config.get_rules() if r.rule_id == rule_id][0]
        assert rule.enabled is True
        ok = config.toggle_rule(rule_id, False)
        assert ok is True
        rule = [r for r in config.get_rules() if r.rule_id == rule_id][0]
        assert rule.enabled is False

    def test_rules_sorted_by_priority(self, tmp_dir):
        config = CashflowRuleConfig(script_dir=tmp_dir)
        config.add_rule(CashflowClassificationRule(
            rule_id='', name='低优先级', subcategory='salary',
            summary_keywords=['a'], priority=200))
        config.add_rule(CashflowClassificationRule(
            rule_id='', name='高优先级', subcategory='salary',
            summary_keywords=['b'], priority=10))
        rules = config.get_rules()
        priorities = [r.priority for r in rules]
        assert priorities == sorted(priorities)

    def test_singleton(self, tmp_dir):
        config1 = get_cashflow_rule_config(script_dir=tmp_dir)
        config2 = get_cashflow_rule_config(script_dir=tmp_dir)
        assert config1 is config2


class TestDefaultRules:

    def test_default_rules_count(self):
        rules = get_default_cashflow_rules()
        assert len(rules) >= 16

    def test_default_rules_have_required_fields(self):
        rules = get_default_cashflow_rules()
        for rule in rules:
            assert rule.name
            assert rule.subcategory
            assert rule.enabled is True
            assert rule.created_by == 'system'

    def test_salary_rule(self):
        rules = get_default_cashflow_rules()
        salary_rules = [r for r in rules if r.subcategory == 'salary']
        assert len(salary_rules) > 0
        salary_rule = salary_rules[0]
        assert '工资' in salary_rule.summary_keywords
        assert salary_rule.amount_direction == 'payment'
        assert salary_rule.priority == 10

    def test_tax_rule(self):
        rules = get_default_cashflow_rules()
        tax_rules = [r for r in rules if r.subcategory == 'tax']
        assert len(tax_rules) > 0
        tax_rule = tax_rules[0]
        assert '税' in tax_rule.summary_keywords
        assert '税务局' in tax_rule.counterpart_keywords
        assert tax_rule.amount_direction == 'payment'

    def test_loan_rules(self):
        rules = get_default_cashflow_rules()
        loan_in_rules = [r for r in rules if r.subcategory == 'loan_in']
        loan_out_rules = [r for r in rules if r.subcategory == 'loan_out']
        assert len(loan_in_rules) > 0
        assert len(loan_out_rules) > 0
        assert loan_in_rules[0].amount_direction == 'receipt'
        assert loan_out_rules[0].amount_direction == 'payment'

    def test_investment_rules(self):
        rules = get_default_cashflow_rules()
        inv_out = [r for r in rules if r.subcategory == 'investment_out']
        inv_in = [r for r in rules if r.subcategory == 'investment_in']
        assert len(inv_out) > 0
        assert len(inv_in) > 0
        assert inv_out[0].amount_direction == 'payment'
        assert inv_in[0].amount_direction == 'receipt'


class TestClassification:

    def test_classify_salary(self, tmp_dir):
        records = [_make_record(summary='发放1月工资', payment=-50000)]
        result, summary = apply_cashflow_classification(records, script_dir=tmp_dir)
        assert summary['classified_count'] == 1
        assert result[0]['现金流子类别'] == '工资薪金'
        assert result[0]['现金流主类别'] == '经营活动现金流'
        assert result[0]['现金流分类匹配关键词'] == '工资'
        assert result[0]['现金流分类匹配来源'] == 'summary'

    def test_classify_tax_by_counterpart(self, tmp_dir):
        records = [_make_record(summary='往来款', counterpart='北京市海淀区税务局', payment=-10000)]
        result, summary = apply_cashflow_classification(records, script_dir=tmp_dir)
        assert summary['classified_count'] == 1
        assert result[0]['现金流子类别'] == '税费缴纳'
        assert result[0]['现金流分类匹配来源'] == 'counterpart'

    def test_classify_tax_by_summary(self, tmp_dir):
        records = [_make_record(summary='缴纳增值税', payment=-5000)]
        result, summary = apply_cashflow_classification(records, script_dir=tmp_dir)
        assert result[0]['现金流子类别'] == '税费缴纳'

    def test_classify_loan_receipt(self, tmp_dir):
        records = [_make_record(summary='收到银行贷款', receipt=1000000)]
        result, summary = apply_cashflow_classification(records, script_dir=tmp_dir)
        assert result[0]['现金流子类别'] == '取得借款'
        assert result[0]['现金流主类别'] == '筹资活动现金流'

    def test_classify_loan_repayment(self, tmp_dir):
        records = [_make_record(summary='偿还银行贷款本金', payment=-500000)]
        result, summary = apply_cashflow_classification(records, script_dir=tmp_dir)
        assert result[0]['现金流子类别'] == '偿还借款'

    def test_classify_investment_purchase(self, tmp_dir):
        records = [_make_record(summary='购买理财产品', payment=-100000)]
        result, summary = apply_cashflow_classification(records, script_dir=tmp_dir)
        assert result[0]['现金流子类别'] == '投资支付'
        assert result[0]['现金流主类别'] == '投资活动现金流'

    def test_classify_investment_income(self, tmp_dir):
        records = [_make_record(summary='收到理财利息', receipt=5000)]
        result, summary = apply_cashflow_classification(records, script_dir=tmp_dir)
        assert result[0]['现金流子类别'] == '投资收益'

    def test_classify_transfer(self, tmp_dir):
        records = [_make_record(summary='账户转账', payment=-10000)]
        result, summary = apply_cashflow_classification(records, script_dir=tmp_dir)
        assert result[0]['现金流子类别'] == '内部转账'
        assert result[0]['现金流主类别'] == '未分类'

    def test_classify_customer_receipt(self, tmp_dir):
        records = [_make_record(summary='收到货款', receipt=200000)]
        result, summary = apply_cashflow_classification(records, script_dir=tmp_dir)
        assert result[0]['现金流子类别'] == '客户收款'

    def test_classify_supplier_payment(self, tmp_dir):
        records = [_make_record(summary='支付原材料款', payment=-80000)]
        result, summary = apply_cashflow_classification(records, script_dir=tmp_dir)
        assert result[0]['现金流子类别'] == '供应商付款'

    def test_classify_fixed_asset(self, tmp_dir):
        records = [_make_record(summary='购买办公设备', payment=-50000)]
        result, summary = apply_cashflow_classification(records, script_dir=tmp_dir)
        assert result[0]['现金流子类别'] == '固定资产'

    def test_classify_dividend(self, tmp_dir):
        records = [_make_record(summary='分配股东分红', payment=-100000)]
        result, summary = apply_cashflow_classification(records, script_dir=tmp_dir)
        assert result[0]['现金流子类别'] == '分配股利'
        assert result[0]['现金流主类别'] == '筹资活动现金流'

    def test_classify_capital_in(self, tmp_dir):
        records = [_make_record(summary='股东投资款', receipt=500000)]
        result, summary = apply_cashflow_classification(records, script_dir=tmp_dir)
        assert result[0]['现金流子类别'] == '实收资本'

    def test_classify_interest(self, tmp_dir):
        records = [_make_record(summary='支付季度利息', payment=-15000, counterpart='中国工商银行北京分行')]
        result, summary = apply_cashflow_classification(records, script_dir=tmp_dir)
        assert result[0]['现金流子类别'] == '利息支出'

    def test_unclassified(self, tmp_dir):
        records = [_make_record(summary='模糊摘要无匹配关键词', payment=-1000)]
        result, summary = apply_cashflow_classification(records, script_dir=tmp_dir)
        assert summary['unclassified_count'] == 1
        assert result[0]['现金流子类别'] == '未分类'

    def test_priority_order(self, tmp_dir):
        records = [_make_record(summary='工资转账', payment=-50000)]
        result, summary = apply_cashflow_classification(records, script_dir=tmp_dir)
        assert result[0]['现金流子类别'] == '工资薪金'

    def test_amount_direction_filter(self, tmp_dir):
        records = [_make_record(summary='工资', receipt=50000)]
        result, summary = apply_cashflow_classification(records, script_dir=tmp_dir)
        assert result[0]['现金流子类别'] != '工资薪金'

    def test_multiple_records(self, tmp_dir):
        records = [
            _make_record(uid='1', summary='发放工资', payment=-50000),
            _make_record(uid='2', summary='缴纳税费', payment=-10000),
            _make_record(uid='3', summary='收到货款', receipt=200000),
            _make_record(uid='4', summary='购买设备', payment=-80000),
            _make_record(uid='5', summary='收到贷款', receipt=500000),
            _make_record(uid='6', summary='无法识别', payment=-5000),
        ]
        result, summary = apply_cashflow_classification(records, script_dir=tmp_dir)
        assert summary['total_records'] == 6
        assert summary['classified_count'] == 5
        assert summary['unclassified_count'] == 1
        assert summary['classification_rate'] == pytest.approx(83.33, 0.01)

    def test_adds_classification_fields(self, tmp_dir):
        records = [_make_record(summary='工资', payment=-1000)]
        result, _ = apply_cashflow_classification(records, script_dir=tmp_dir)
        assert '现金流主类别' in result[0]
        assert '现金流子类别' in result[0]
        assert '现金流分类主类别编码' in result[0]
        assert '现金流分类子类别编码' in result[0]
        assert '现金流分类规则名称' in result[0]
        assert '现金流分类匹配关键词' in result[0]
        assert '现金流分类匹配来源' in result[0]

    def test_disabled_rule_skipped(self, tmp_dir):
        config = CashflowRuleConfig(script_dir=tmp_dir)
        salary_rule = [r for r in config.get_rules() if r.subcategory == 'salary'][0]
        config.toggle_rule(salary_rule.rule_id, False)
        records = [_make_record(summary='发放工资', payment=-50000)]
        result, summary = apply_cashflow_classification(records, script_dir=tmp_dir)
        assert result[0]['现金流子类别'] != '工资薪金'

    def test_summary_returns_summary_dict(self, tmp_dir):
        records = [_make_record(summary='工资', payment=-1000)]
        _, summary = apply_cashflow_classification(records, script_dir=tmp_dir)
        assert 'total_records' in summary
        assert 'classified_count' in summary
        assert 'unclassified_count' in summary
        assert 'classification_rate' in summary
        assert 'category_counts' in summary
        assert 'subcategory_counts' in summary
        assert 'rule_hit_counts' in summary


class TestSummary:

    def test_summarize_by_subcategory(self, tmp_dir):
        records = [
            _make_record(uid='1', summary='工资', payment=-50000),
            _make_record(uid='2', summary='税费', payment=-10000),
            _make_record(uid='3', summary='货款', receipt=200000),
        ]
        classified, _ = apply_cashflow_classification(records, script_dir=tmp_dir)
        summary = summarize_cashflow_by_category(classified, group_by='sub')
        assert len(summary) >= 3
        salary = [s for s in summary if s['子类别编码'] == 'salary'][0]
        assert salary['交易笔数'] == 1
        assert salary['流出金额'] == 50000
        assert salary['流入金额'] == 0
        assert salary['净额'] == -50000

    def test_summarize_by_main_category(self, tmp_dir):
        records = [
            _make_record(uid='1', summary='工资', payment=-50000),
            _make_record(uid='2', summary='货款', receipt=200000),
            _make_record(uid='3', summary='贷款', receipt=100000),
        ]
        classified, _ = apply_cashflow_classification(records, script_dir=tmp_dir)
        summary = summarize_cashflow_by_category(classified, group_by='main')
        operating = [s for s in summary if s['主类别编码'] == 'operating'][0]
        assert operating['交易笔数'] == 2
        assert operating['流入金额'] == 200000
        assert operating['流出金额'] == 50000
        assert operating['净额'] == 150000

    def test_summarize_all(self, tmp_dir):
        records = [
            _make_record(uid='1', summary='工资', payment=-50000),
            _make_record(uid='2', summary='贷款', receipt=100000),
        ]
        classified, _ = apply_cashflow_classification(records, script_dir=tmp_dir)
        summary = summarize_cashflow_by_category(classified, group_by='all')
        main_items = [s for s in summary if s['汇总维度'] == '主类别']
        sub_items = [s for s in summary if s['汇总维度'] == '子类别']
        assert len(main_items) >= 2
        assert len(sub_items) >= 2

    def test_summary_amounts_rounded(self, tmp_dir):
        records = [_make_record(summary='工资', payment=-1000.123)]
        classified, _ = apply_cashflow_classification(records, script_dir=tmp_dir)
        summary = summarize_cashflow_by_category(classified, group_by='sub')
        salary = [s for s in summary if s['子类别编码'] == 'salary'][0]
        assert salary['流出金额'] == pytest.approx(1000.12, 0.01)

    def test_summary_sorted_by_category(self, tmp_dir):
        records = [
            _make_record(uid='1', summary='贷款', receipt=100000),
            _make_record(uid='2', summary='工资', payment=-50000),
        ]
        classified, _ = apply_cashflow_classification(records, script_dir=tmp_dir)
        summary = summarize_cashflow_by_category(classified, group_by='sub')
        categories = [s['主类别编码'] for s in summary]
        assert categories == sorted(categories)


class TestExport:

    def test_export_creates_file(self, tmp_dir):
        records = [
            _make_record(uid='1', summary='工资', payment=-50000),
            _make_record(uid='2', summary='税费', payment=-10000),
            _make_record(uid='3', summary='货款', receipt=200000),
        ]
        classified, _ = apply_cashflow_classification(records, script_dir=tmp_dir)
        output_path = os.path.join(tmp_dir, 'cashflow_summary.xlsx')
        result = export_cashflow_summary(classified, output_path)
        assert result == output_path
        assert os.path.exists(output_path)

    def test_export_has_sheets(self, tmp_dir):
        import pandas as pd
        records = [_make_record(summary='工资', payment=-50000)]
        classified, _ = apply_cashflow_classification(records, script_dir=tmp_dir)
        output_path = os.path.join(tmp_dir, 'cashflow_summary.xlsx')
        export_cashflow_summary(classified, output_path)
        xl = pd.ExcelFile(output_path, engine='openpyxl')
        assert '主类别汇总' in xl.sheet_names
        assert '子类别汇总' in xl.sheet_names
        assert '交易明细' in xl.sheet_names

    def test_export_without_details(self, tmp_dir):
        import pandas as pd
        records = [_make_record(summary='工资', payment=-50000)]
        classified, _ = apply_cashflow_classification(records, script_dir=tmp_dir)
        output_path = os.path.join(tmp_dir, 'cashflow_no_details.xlsx')
        export_cashflow_summary(classified, output_path, include_details=False)
        xl = pd.ExcelFile(output_path, engine='openpyxl')
        assert '主类别汇总' in xl.sheet_names
        assert '子类别汇总' in xl.sheet_names
        assert '交易明细' not in xl.sheet_names

    def test_export_main_summary_content(self, tmp_dir):
        import pandas as pd
        records = [
            _make_record(uid='1', summary='工资', payment=-50000),
            _make_record(uid='2', summary='货款', receipt=200000),
        ]
        classified, _ = apply_cashflow_classification(records, script_dir=tmp_dir)
        output_path = os.path.join(tmp_dir, 'cashflow_content.xlsx')
        export_cashflow_summary(classified, output_path)
        df = pd.read_excel(output_path, sheet_name='主类别汇总', engine='openpyxl')
        assert len(df) > 0
        assert '主类别编码' in df.columns
        assert '流入金额(元)' in df.columns


class TestConvenienceFunctions:

    def test_add_cashflow_rule(self, tmp_dir):
        rule_id = add_cashflow_rule(
            name='自定义规则',
            subcategory='operating_other',
            summary_keywords=['自定义关键词'],
            script_dir=tmp_dir,
        )
        assert rule_id
        config = get_cashflow_rule_config(script_dir=tmp_dir)
        rules = config.get_rules()
        found = [r for r in rules if r.rule_id == rule_id]
        assert len(found) == 1
        assert found[0].name == '自定义规则'

    def test_init_default_rules(self, tmp_dir):
        rule_ids = init_default_cashflow_rules(script_dir=tmp_dir)
        assert len(rule_ids) == len(get_default_cashflow_rules())

    def test_get_cashflow_classification(self, tmp_dir):
        records = [
            _make_record(uid='1', summary='工资', payment=-50000),
            _make_record(uid='2', summary='货款', receipt=200000),
        ]
        classified, class_summary, summary = get_cashflow_classification(
            records, script_dir=tmp_dir)
        assert len(classified) == 2
        assert class_summary['classified_count'] == 2
        assert len(summary) > 0


class TestIntegration:

    def test_full_workflow(self, tmp_dir):
        records = [
            _make_record(uid='1', summary='发放1月工资', payment=-50000),
            _make_record(uid='2', summary='缴纳增值税', payment=-15000, counterpart='北京市税务局'),
            _make_record(uid='3', summary='收到客户货款', receipt=300000),
            _make_record(uid='4', summary='支付供应商货款', payment=-150000),
            _make_record(uid='5', summary='购买办公设备', payment=-80000),
            _make_record(uid='6', summary='收到银行贷款', receipt=500000),
            _make_record(uid='7', summary='偿还贷款利息', payment=-15000, counterpart='中国工商银行'),
            _make_record(uid='8', summary='购买理财产品', payment=-200000),
            _make_record(uid='9', summary='收到理财分红', receipt=10000),
            _make_record(uid='10', summary='账户间转账', payment=-100000),
        ]

        classified, class_summary, summary = get_cashflow_classification(
            records, script_dir=tmp_dir)

        assert class_summary['total_records'] == 10
        assert class_summary['classified_count'] == 10
        assert class_summary['classification_rate'] == 100.0

        output_path = os.path.join(tmp_dir, 'full_workflow.xlsx')
        export_cashflow_summary(classified, output_path)

        assert os.path.exists(output_path)

        sub_summary = [s for s in summary if s['汇总维度'] == '子类别']
        category_codes = [s['子类别编码'] for s in sub_summary]
        assert 'salary' in category_codes
        assert 'tax' in category_codes
        assert 'customer' in category_codes
        assert 'supplier' in category_codes
        assert 'fixed_asset' in category_codes
        assert 'loan_in' in category_codes
        assert 'interest' in category_codes
        assert 'investment_out' in category_codes
        assert 'investment_income' in category_codes
        assert 'transfer' in category_codes

        salary_item = [s for s in sub_summary if s['子类别编码'] == 'salary'][0]
        assert salary_item['流出金额'] == 50000
        assert salary_item['净额'] == -50000

        loan_in_item = [s for s in sub_summary if s['子类别编码'] == 'loan_in'][0]
        assert loan_in_item['流入金额'] == 500000
        assert loan_in_item['净额'] == 500000
