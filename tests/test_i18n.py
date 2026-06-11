# -*- coding: utf-8 -*-
"""
多语言国际化模块单元测试
测试覆盖：单例模式、语言加载、翻译、变量插值、嵌套键、语言切换、回退机制
"""

import os
import sys
import json
import tempfile
import shutil

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'backend'))

import i18n
from i18n import (
    I18nManager,
    get_i18n,
    init_i18n,
    t,
    set_language,
    get_language,
    get_available_languages,
    get_i18n_dir,
    get_locales_dir,
)


class TestI18nDirectories:
    """测试目录路径获取"""

    def test_get_i18n_dir(self):
        """测试获取i18n目录"""
        dir_path = get_i18n_dir()
        assert os.path.isdir(dir_path)
        assert os.path.basename(dir_path) == 'i18n'

    def test_get_locales_dir(self):
        """测试获取语言包目录"""
        dir_path = get_locales_dir()
        assert os.path.isdir(dir_path)
        assert os.path.basename(dir_path) == 'locales'


class TestI18nManagerSingleton:
    """测试单例模式"""

    def test_singleton_instance(self):
        """测试多次实例化返回同一对象"""
        instance1 = I18nManager()
        instance2 = I18nManager()
        assert instance1 is instance2

    def test_get_i18n_returns_same_instance(self):
        """测试get_i18n返回同一实例"""
        instance1 = get_i18n()
        instance2 = get_i18n()
        assert instance1 is instance2

    def test_init_i18n_creates_new_instance(self):
        """测试init_i18n重置并创建新实例"""
        old_instance = get_i18n()
        new_instance = init_i18n()
        assert old_instance is not new_instance


class TestLanguagePacks:
    """测试语言包加载"""

    def test_zh_cn_language_pack_exists(self):
        """测试简体中文语言包存在且有效"""
        zh_file = os.path.join(get_locales_dir(), 'zh_CN.json')
        assert os.path.exists(zh_file)

        with open(zh_file, 'r', encoding='utf-8') as f:
            data = json.load(f)

        assert isinstance(data, dict)
        assert 'metadata' in data
        assert data['metadata']['code'] == 'zh_CN'
        assert data['metadata']['name'] == '简体中文'

    def test_en_us_language_pack_exists(self):
        """测试英文语言包存在且有效"""
        en_file = os.path.join(get_locales_dir(), 'en_US.json')
        assert os.path.exists(en_file)

        with open(en_file, 'r', encoding='utf-8') as f:
            data = json.load(f)

        assert isinstance(data, dict)
        assert 'metadata' in data
        assert data['metadata']['code'] == 'en_US'
        assert data['metadata']['name'] == 'English'

    def test_language_packs_structure_match(self):
        """测试两个语言包的键结构一致"""
        zh_file = os.path.join(get_locales_dir(), 'zh_CN.json')
        en_file = os.path.join(get_locales_dir(), 'en_US.json')

        with open(zh_file, 'r', encoding='utf-8') as f:
            zh_data = json.load(f)
        with open(en_file, 'r', encoding='utf-8') as f:
            en_data = json.load(f)

        def get_all_keys(d, prefix=''):
            keys = set()
            for k, v in d.items():
                full_key = f"{prefix}.{k}" if prefix else k
                if isinstance(v, dict):
                    keys.update(get_all_keys(v, full_key))
                else:
                    keys.add(full_key)
            return keys

        zh_keys = get_all_keys(zh_data)
        en_keys = get_all_keys(en_data)

        missing_in_en = zh_keys - en_keys
        missing_in_zh = en_keys - zh_keys

        assert len(missing_in_en) == 0, f"英文语言包缺少以下键: {missing_in_en}"
        assert len(missing_in_zh) == 0, f"中文语言包缺少以下键: {missing_in_zh}"

    def test_language_packs_have_expected_categories(self):
        """测试语言包包含预期的分类"""
        expected_categories = [
            'metadata', 'gui', 'cli', 'common', 'errors', 'success',
            'warnings', 'info', 'result_messages', 'report', 'api',
            'modes', 'diff', 'web_ui'
        ]

        zh_file = os.path.join(get_locales_dir(), 'zh_CN.json')
        with open(zh_file, 'r', encoding='utf-8') as f:
            zh_data = json.load(f)

        for cat in expected_categories:
            assert cat in zh_data, f"中文语言包缺少分类: {cat}"


class TestLanguageSwitching:
    """测试语言切换功能"""

    def test_default_language_is_zh_cn(self):
        """测试默认语言是简体中文"""
        instance = init_i18n(language='zh_CN')
        assert instance.get_language() == 'zh_CN'

    def test_set_language_to_english(self):
        """测试切换到英文"""
        instance = init_i18n(language='zh_CN')
        assert instance.get_language() == 'zh_CN'

        result = instance.set_language('en_US')
        assert result is True
        assert instance.get_language() == 'en_US'

    def test_set_language_to_unsupported(self):
        """测试设置不支持的语言"""
        instance = init_i18n(language='zh_CN')
        result = instance.set_language('ja_JP')
        assert result is False
        assert instance.get_language() == 'zh_CN'

    def test_convenience_set_language(self):
        """测试便捷函数set_language"""
        init_i18n(language='zh_CN')
        assert get_language() == 'zh_CN'

        result = set_language('en_US')
        assert result is True
        assert get_language() == 'en_US'

    def test_get_language_name(self):
        """测试获取语言显示名称"""
        instance = init_i18n(language='zh_CN')
        assert instance.get_language_name() == '简体中文'

        instance.set_language('en_US')
        assert instance.get_language_name() == 'English'

    def test_get_available_languages(self):
        """测试获取可用语言列表"""
        instance = init_i18n()
        langs = instance.get_available_languages()
        assert isinstance(langs, dict)
        assert 'zh_CN' in langs
        assert 'en_US' in langs
        assert langs['zh_CN'] == '简体中文'
        assert langs['en_US'] == 'English'

        conv_langs = get_available_languages()
        assert conv_langs == langs


class TestTranslation:
    """测试翻译功能"""

    def test_translate_simple_key_zh(self):
        """测试简体中文简单键翻译"""
        instance = init_i18n(language='zh_CN')
        instance.set_language('zh_CN')

        result = instance.translate('gui.select_folder')
        assert result == '请选择银行流水文件夹'

    def test_translate_simple_key_en(self):
        """测试英文简单键翻译"""
        instance = init_i18n(language='zh_CN')
        instance.set_language('en_US')

        result = instance.translate('gui.select_folder')
        assert result == 'Please select bank statement folder'

    def test_convenience_t_function(self):
        """测试便捷翻译函数t"""
        init_i18n(language='zh_CN')
        set_language('zh_CN')

        assert t('gui.select_folder') == '请选择银行流水文件夹'

        set_language('en_US')
        assert t('gui.select_folder') == 'Please select bank statement folder'

    def test_translate_with_variable_interpolation(self):
        """测试带变量插值的翻译"""
        instance = init_i18n(language='zh_CN')

        result = instance.translate('errors.file_not_found', file='test.xlsx')
        assert result == '文件不存在: test.xlsx'

        instance.set_language('en_US')
        result = instance.translate('errors.file_not_found', file='test.xlsx')
        assert result == 'File not found: test.xlsx'

    def test_translate_nested_key(self):
        """测试嵌套键访问"""
        instance = init_i18n(language='zh_CN')

        result = instance.translate('report.title')
        assert result == '银行流水检验报告'

    def test_translate_multiple_variables(self):
        """测试多个变量的插值"""
        instance = init_i18n(language='zh_CN')
        instance.set_language('zh_CN')

        result = instance.translate(
            'diff.summary',
            added=5,
            removed=3,
            modified=2
        )
        assert '新增 5 条' in result
        assert '删除 3 条' in result
        assert '变更 2 条' in result

        instance.set_language('en_US')
        result = instance.translate(
            'diff.summary',
            added=5,
            removed=3,
            modified=2
        )
        assert '5 added' in result
        assert '3 removed' in result
        assert '2 modified' in result

    def test_translate_missing_key_returns_key(self):
        """测试翻译不存在的键返回键本身"""
        instance = init_i18n(language='zh_CN')
        result = instance.translate('this.key.does.not.exist')
        assert result == 'this.key.does.not.exist'

    def test_translate_empty_key(self):
        """测试翻译空键"""
        instance = init_i18n(language='zh_CN')
        assert instance.translate('') == ''
        assert instance.translate(None) is None

    def test_fallback_to_default_language(self):
        """测试回退到默认语言"""
        instance = init_i18n(language='en_US')
        instance.set_language('en_US')

        assert instance.has_key('gui.select_folder')
        assert instance.translate('gui.select_folder') == 'Please select bank statement folder'


class TestTranslationDict:
    """测试获取翻译字典"""

    def test_get_report_translation_dict(self):
        """测试获取report分类的翻译字典"""
        instance = init_i18n(language='zh_CN')
        report_dict = instance.get_translation_dict('report')

        assert isinstance(report_dict, dict)
        assert 'title' in report_dict
        assert 'batch_id' in report_dict

    def test_get_web_ui_translation_dict(self):
        """测试获取web_ui分类的翻译字典"""
        instance = init_i18n(language='en_US')
        instance.set_language('en_US')
        web_dict = instance.get_translation_dict('web_ui')

        assert isinstance(web_dict, dict)

    def test_get_nonexistent_dict_returns_none(self):
        """测试获取不存在的分类返回None"""
        instance = init_i18n(language='zh_CN')
        result = instance.get_translation_dict('nonexistent.category')
        assert result is None


class TestHasKey:
    """测试键存在性检查"""

    def test_has_existing_key(self):
        """测试检查存在的键"""
        instance = init_i18n(language='zh_CN')
        assert instance.has_key('gui.select_folder') is True
        assert instance.has_key('errors.file_not_found') is True

    def test_has_nonexistent_key(self):
        """测试检查不存在的键"""
        instance = init_i18n(language='zh_CN')
        assert instance.has_key('nonexistent.key') is False

    def test_has_key_in_other_language(self):
        """测试检查其他语言的键"""
        instance = init_i18n(language='zh_CN')
        instance.set_language('en_US')
        assert instance.has_key('gui.select_folder') is True


class TestVariableInterpolation:
    """测试变量插值的各种情况"""

    def test_missing_variable_in_kwargs(self):
        """测试缺少变量时返回原模板"""
        instance = init_i18n(language='zh_CN')
        result = instance.translate('errors.file_not_found')
        assert result == '文件不存在: {file}'

    def test_extra_variables_ignored(self):
        """测试额外变量被忽略"""
        instance = init_i18n(language='zh_CN')
        result = instance.translate(
            'errors.file_not_found',
            file='test.xlsx',
            extra_var='should be ignored'
        )
        assert result == '文件不存在: test.xlsx'

    def test_numeric_variables(self):
        """测试数字类型变量"""
        instance = init_i18n(language='zh_CN')
        instance.set_language('zh_CN')
        result = instance.translate(
            'diff.summary',
            added=100,
            removed=50,
            modified=25
        )
        assert '100' in result
        assert '50' in result
        assert '25' in result


class TestReportTranslations:
    """测试报告相关翻译"""

    def test_report_title_translations(self):
        """测试报告标题翻译"""
        instance = init_i18n(language='zh_CN')
        instance.set_language('zh_CN')
        assert instance.translate('report.title') == '银行流水检验报告'

        instance.set_language('en_US')
        assert instance.translate('report.title') == 'Bank Statement Inspection Report'

    def test_report_section_translations(self):
        """测试报告章节标题翻译"""
        instance = init_i18n(language='zh_CN')
        instance.set_language('zh_CN')

        assert '处理统计' in instance.translate('report.section_processing_stats')
        assert '文件清单' in instance.translate('report.section_file_list')

        instance.set_language('en_US')
        assert 'Processing Statistics' in instance.translate('report.section_processing_stats')
        assert 'File List' in instance.translate('report.section_file_list')

    def test_report_metric_translations(self):
        """测试报告指标翻译"""
        instance = init_i18n(language='zh_CN')
        instance.set_language('zh_CN')

        assert instance.translate('report.processed_files') == '处理文件数'
        assert instance.translate('report.total_records') == '总记录数'

        instance.set_language('en_US')
        assert instance.translate('report.processed_files') == 'Processed Files'
        assert instance.translate('report.total_records') == 'Total Records'


class TestModeTranslations:
    """测试运行模式翻译"""

    def test_mode_names_zh(self):
        """测试中文模式名称"""
        instance = init_i18n(language='zh_CN')

        assert instance.translate('modes.pipeline_name') == '主流程'
        assert instance.translate('modes.diff_name') == '变更对比'
        assert instance.translate('modes.export_name') == '财务导出'

    def test_mode_names_en(self):
        """测试英文模式名称"""
        instance = init_i18n(language='en_US')
        instance.set_language('en_US')

        assert instance.translate('modes.pipeline_name') == 'Main Pipeline'
        assert instance.translate('modes.diff_name') == 'Diff Comparison'
        assert instance.translate('modes.export_name') == 'Financial Export'

    def test_mode_descriptions(self):
        """测试模式描述翻译"""
        instance = init_i18n(language='zh_CN')
        assert '处理银行流水文件夹' in instance.translate('modes.pipeline_desc')

        instance.set_language('en_US')
        assert 'Process bank statement folder' in instance.translate('modes.pipeline_desc')


class TestErrorTranslations:
    """测试错误信息翻译"""

    def test_common_errors_zh(self):
        """测试中文错误信息"""
        instance = init_i18n(language='zh_CN')

        assert '文件不存在' in instance.translate('errors.file_not_found', file='test.xlsx')
        assert '文件夹不存在' in instance.translate('errors.folder_not_found', folder='test_dir')
        assert '读取失败' in instance.translate('errors.read_failed', file='test.xlsx')

    def test_common_errors_en(self):
        """测试英文错误信息"""
        instance = init_i18n(language='en_US')
        instance.set_language('en_US')

        assert 'File not found' in instance.translate('errors.file_not_found', file='test.xlsx')
        assert 'Folder not found' in instance.translate('errors.folder_not_found', folder='test_dir')
        assert 'Failed to read' in instance.translate('errors.read_failed', file='test.xlsx')


class TestApiTranslations:
    """测试API响应翻译"""

    def test_api_messages_zh(self):
        """测试中文API消息"""
        instance = init_i18n(language='zh_CN')

        assert '添加成功' in instance.translate('api.add_success')
        assert '更新成功' in instance.translate('api.update_success')
        assert '删除成功' in instance.translate('api.delete_success')

    def test_api_messages_en(self):
        """测试英文API消息"""
        instance = init_i18n(language='en_US')
        instance.set_language('en_US')

        assert 'successfully' in instance.translate('api.add_success')
        assert 'successfully' in instance.translate('api.update_success')
        assert 'successfully' in instance.translate('api.delete_success')


class TestWebUITranslations:
    """测试Web界面翻译"""

    def test_web_ui_zh(self):
        """测试中文Web界面"""
        instance = init_i18n(language='zh_CN')

        assert '主体查找表' in instance.translate('web_ui.title')
        assert '搜索' in instance.translate('web_ui.search')
        assert '新增' in instance.translate('web_ui.add')

    def test_web_ui_en(self):
        """测试英文Web界面"""
        instance = init_i18n(language='en_US')
        instance.set_language('en_US')

        assert 'Lookup Table' in instance.translate('web_ui.title')
        assert 'Search' in instance.translate('web_ui.search')
        assert 'Add' in instance.translate('web_ui.add')


class TestCLITranslations:
    """测试CLI界面翻译"""

    def test_cli_zh(self):
        """测试中文CLI"""
        instance = init_i18n(language='zh_CN')

        assert '请选择运行模式' in instance.translate('cli.select_mode')
        assert '请输入文件夹路径' in instance.translate('cli.enter_folder_path')

    def test_cli_en(self):
        """测试英文CLI"""
        instance = init_i18n(language='en_US')
        instance.set_language('en_US')

        assert 'Please select' in instance.translate('cli.select_mode')
        assert 'folder path' in instance.translate('cli.enter_folder_path')


class TestFormattingEdgeCases:
    """测试格式化边缘情况"""

    def test_special_characters_in_variables(self):
        """测试变量中包含特殊字符"""
        instance = init_i18n(language='zh_CN')
        result = instance.translate(
            'errors.file_not_found',
            file='path/to/file with spaces & special.xlsx'
        )
        assert 'path/to/file with spaces & special.xlsx' in result

    def test_unicode_variables(self):
        """测试Unicode变量"""
        instance = init_i18n(language='zh_CN')
        result = instance.translate(
            'errors.file_not_found',
            file='测试文件.xlsx'
        )
        assert '测试文件.xlsx' in result


class TestModuleIntegration:
    """测试模块集成"""

    def test_import_convenience_functions(self):
        """测试便捷函数可以正确导入"""
        from i18n import t, set_language, get_language
        assert callable(t)
        assert callable(set_language)
        assert callable(get_language)

    def test_translation_consistency_across_calls(self):
        """测试多次调用翻译结果一致"""
        init_i18n(language='zh_CN')
        set_language('zh_CN')
        result1 = t('gui.select_folder')
        result2 = t('gui.select_folder')
        assert result1 == result2

    def test_language_switch_affects_subsequent_calls(self):
        """测试语言切换影响后续调用"""
        init_i18n(language='zh_CN')
        set_language('zh_CN')
        zh_result = t('gui.select_folder')

        set_language('en_US')
        en_result = t('gui.select_folder')

        assert zh_result != en_result
        assert zh_result == '请选择银行流水文件夹'
        assert en_result == 'Please select bank statement folder'


if __name__ == '__main__':
    pytest.main([__file__, '-v'])
