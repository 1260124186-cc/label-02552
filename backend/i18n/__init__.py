# -*- coding: utf-8 -*-
"""
多语言国际化模块 (i18n)
提供语言包加载、切换、翻译功能，支持简体中文与英文
"""

import os
import json
import logging
from typing import Dict, Any, Optional, List
from string import Formatter


def get_i18n_dir() -> str:
    """获取 i18n 模块所在目录"""
    return os.path.dirname(os.path.abspath(__file__))


def get_locales_dir() -> str:
    """获取语言包目录"""
    return os.path.join(get_i18n_dir(), 'locales')


class I18nManager:
    """
    多语言管理器 - 单例模式
    
    功能：
    1. 加载 JSON 格式的语言包
    2. 支持运行时语言切换
    3. 提供翻译函数，支持变量替换
    4. 支持嵌套键访问（如 'gui.select_folder'）
    """
    
    _instance = None
    _initialized = False
    
    DEFAULT_LANGUAGE = 'zh_CN'
    SUPPORTED_LANGUAGES = {
        'zh_CN': '简体中文',
        'en_US': 'English',
    }
    
    def __new__(cls, *args, **kwargs):
        if cls._instance is None:
            cls._instance = super(I18nManager, cls).__new__(cls)
        return cls._instance
    
    def __init__(self, language: Optional[str] = None, auto_detect: bool = True):
        if self._initialized:
            return
        
        self._initialized = True
        self._current_language = self.DEFAULT_LANGUAGE
        self._translations: Dict[str, Dict[str, Any]] = {}
        self._logger = logging.getLogger('bankcheck.i18n')
        
        self._load_all_languages()
        
        if language and language in self.SUPPORTED_LANGUAGES:
            self.set_language(language)
        elif auto_detect:
            self._auto_detect_language()
    
    def _load_all_languages(self) -> None:
        """加载所有支持的语言包"""
        locales_dir = get_locales_dir()
        
        for lang_code in self.SUPPORTED_LANGUAGES:
            lang_file = os.path.join(locales_dir, f'{lang_code}.json')
            if os.path.exists(lang_file):
                try:
                    with open(lang_file, 'r', encoding='utf-8') as f:
                        self._translations[lang_code] = json.load(f)
                    self._logger.info('Loaded language pack: %s', lang_code)
                except (json.JSONDecodeError, IOError) as e:
                    self._logger.error('Failed to load language pack %s: %s', lang_code, e)
                    self._translations[lang_code] = {}
            else:
                self._logger.warning('Language pack not found: %s', lang_file)
                self._translations[lang_code] = {}
    
    def _auto_detect_language(self) -> None:
        """自动检测系统语言"""
        try:
            import locale
            try:
                system_lang = locale.getlocale()[0]
            except (AttributeError, TypeError):
                system_lang = locale.getdefaultlocale()[0]
            if system_lang:
                if system_lang.startswith('zh'):
                    self.set_language('zh_CN')
                else:
                    self.set_language('en_US')
        except Exception:
            self.set_language(self.DEFAULT_LANGUAGE)
    
    def set_language(self, language: str) -> bool:
        """
        设置当前语言
        
        Args:
            language: 语言代码，如 'zh_CN' 或 'en_US'
            
        Returns:
            bool: 设置是否成功
        """
        if language not in self.SUPPORTED_LANGUAGES:
            self._logger.warning('Unsupported language: %s', language)
            return False
        
        if language not in self._translations or not self._translations[language]:
            self._logger.warning('Language pack not loaded: %s', language)
            return False
        
        self._current_language = language
        self._logger.info('Language set to: %s (%s)', 
                         language, self.SUPPORTED_LANGUAGES[language])
        return True
    
    def get_language(self) -> str:
        """获取当前语言代码"""
        return self._current_language
    
    def get_language_name(self) -> str:
        """获取当前语言显示名称"""
        return self.SUPPORTED_LANGUAGES.get(self._current_language, self._current_language)
    
    def get_available_languages(self) -> Dict[str, str]:
        """获取所有支持的语言"""
        return dict(self.SUPPORTED_LANGUAGES)
    
    def _get_nested_value(self, data: Dict[str, Any], key: str) -> Optional[Any]:
        """
        嵌套获取字典值
        
        Args:
            data: 字典数据
            key: 嵌套键，如 'gui.select_folder'
            
        Returns:
            找到的值，未找到返回 None
        """
        keys = key.split('.')
        value = data
        
        for k in keys:
            if isinstance(value, dict) and k in value:
                value = value[k]
            else:
                return None
        
        return value
    
    def _format_string(self, template: str, **kwargs) -> str:
        """
        格式化字符串，支持变量替换
        
        Args:
            template: 模板字符串，如 '文件不存在: {file}'
            **kwargs: 变量键值对
            
        Returns:
            格式化后的字符串
        """
        try:
            return template.format(**kwargs)
        except (KeyError, ValueError):
            return template
    
    def translate(self, key: str, **kwargs) -> str:
        """
        翻译函数
        
        Args:
            key: 翻译键，如 'gui.select_folder' 或 'errors.file_not_found'
            **kwargs: 变量替换参数
            
        Returns:
            翻译后的字符串。如果键不存在，返回键本身。
        """
        if not key:
            return key
        
        translations = self._translations.get(self._current_language, {})
        value = self._get_nested_value(translations, key)
        
        if value is None:
            fallback_lang = self.DEFAULT_LANGUAGE
            if self._current_language != fallback_lang:
                fallback_translations = self._translations.get(fallback_lang, {})
                value = self._get_nested_value(fallback_translations, key)
            
            if value is None:
                self._logger.debug('Translation key not found: %s', key)
                return key
        
        if isinstance(value, str) and kwargs:
            return self._format_string(value, **kwargs)
        
        return str(value) if value is not None else key
    
    def get_translation_dict(self, key: str) -> Optional[Dict[str, Any]]:
        """
        获取翻译字典（用于获取一组相关翻译）
        
        Args:
            key: 翻译键，如 'report' 或 'web_ui'
            
        Returns:
            翻译字典，未找到返回 None
        """
        translations = self._translations.get(self._current_language, {})
        value = self._get_nested_value(translations, key)
        
        if value is None and self._current_language != self.DEFAULT_LANGUAGE:
            fallback_translations = self._translations.get(self.DEFAULT_LANGUAGE, {})
            value = self._get_nested_value(fallback_translations, key)
        
        if isinstance(value, dict):
            return value
        
        return None
    
    def has_key(self, key: str) -> bool:
        """检查翻译键是否存在"""
        translations = self._translations.get(self._current_language, {})
        if self._get_nested_value(translations, key) is not None:
            return True
        
        if self._current_language != self.DEFAULT_LANGUAGE:
            fallback_translations = self._translations.get(self.DEFAULT_LANGUAGE, {})
            return self._get_nested_value(fallback_translations, key) is not None
        
        return False
    
    def reload_language(self, language: Optional[str] = None) -> bool:
        """
        重新加载指定语言包（用于热更新）
        
        Args:
            language: 语言代码，不指定则重载当前语言
            
        Returns:
            bool: 重载是否成功
        """
        target_lang = language or self._current_language
        if target_lang not in self.SUPPORTED_LANGUAGES:
            return False
        
        lang_file = os.path.join(get_locales_dir(), f'{target_lang}.json')
        if not os.path.exists(lang_file):
            return False
        
        try:
            with open(lang_file, 'r', encoding='utf-8') as f:
                self._translations[target_lang] = json.load(f)
            self._logger.info('Reloaded language pack: %s', target_lang)
            return True
        except (json.JSONDecodeError, IOError) as e:
            self._logger.error('Failed to reload language pack %s: %s', target_lang, e)
            return False
    
    def reload_all(self) -> None:
        """重新加载所有语言包"""
        for lang_code in self.SUPPORTED_LANGUAGES:
            self.reload_language(lang_code)


_i18n_instance: Optional[I18nManager] = None


def get_i18n() -> I18nManager:
    """获取 I18nManager 单例实例"""
    global _i18n_instance
    if _i18n_instance is None:
        _i18n_instance = I18nManager()
    return _i18n_instance


def init_i18n(language: Optional[str] = None) -> I18nManager:
    """
    初始化多语言模块（重置单例）
    
    Args:
        language: 初始语言代码，如 'zh_CN' 或 'en_US'
        
    Returns:
        I18nManager 实例
    """
    global _i18n_instance
    I18nManager._instance = None
    I18nManager._initialized = False
    _i18n_instance = I18nManager(language=language, auto_detect=(language is None))
    return _i18n_instance


def t(key: str, **kwargs) -> str:
    """
    便捷翻译函数
    
    Args:
        key: 翻译键
        **kwargs: 变量参数
        
    Returns:
        翻译后的字符串
    """
    return get_i18n().translate(key, **kwargs)


def set_language(language: str) -> bool:
    """
    便捷函数：设置当前语言
    
    Args:
        language: 语言代码
        
    Returns:
        bool: 设置是否成功
    """
    return get_i18n().set_language(language)


def get_language() -> str:
    """便捷函数：获取当前语言代码"""
    return get_i18n().get_language()


def get_available_languages() -> Dict[str, str]:
    """便捷函数：获取所有支持的语言"""
    return get_i18n().get_available_languages()
