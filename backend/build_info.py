# -*- coding: utf-8 -*-
"""
构建信息模块
统一管理应用版本号、构建时间与构建平台，打包时由构建脚本自动生成。

使用方式：
    from build_info import get_version, get_build_time, get_build_info, get_build_platform

开发环境下未生成该文件时，动态回退到默认值；
打包时由 build.bat / build.sh 覆写本文件，嵌入真实构建时间与平台。
"""

import os
import sys
import platform
from datetime import datetime

__all__ = ['get_version', 'get_build_time', 'get_build_info', 'get_build_platform',
           'APP_VERSION', 'BUILD_TIME', 'BUILD_PLATFORM']

APP_VERSION = "1.0.0"
BUILD_TIME = None
BUILD_PLATFORM = None


def _is_frozen() -> bool:
    """判断是否为 PyInstaller 打包后的环境"""
    return getattr(sys, 'frozen', False)


def _detect_build_time() -> str:
    """检测构建时间：优先使用本文件的修改时间，开发环境取当前时间"""
    if BUILD_TIME:
        return BUILD_TIME
    try:
        if _is_frozen():
            exe_path = sys.executable
            ts = os.path.getmtime(exe_path)
        else:
            here = os.path.abspath(os.path.dirname(__file__))
            ts = os.path.getmtime(os.path.join(here, 'build_info.py'))
        return datetime.fromtimestamp(ts).strftime('%Y-%m-%d %H:%M:%S')
    except Exception:
        return datetime.now().strftime('%Y-%m-%d %H:%M:%S')


def _detect_build_platform() -> str:
    """检测构建平台：优先使用打包时嵌入的值，否则运行时推断"""
    if BUILD_PLATFORM:
        return BUILD_PLATFORM
    return platform.system()


def get_version() -> str:
    """获取应用版本号"""
    return APP_VERSION


def get_build_time() -> str:
    """获取构建时间字符串"""
    return _detect_build_time()


def get_build_platform() -> str:
    """获取构建平台字符串 (Windows / Darwin / Linux)"""
    return _detect_build_platform()


def get_build_info() -> dict:
    """获取完整的构建信息字典"""
    return {
        'version': get_version(),
        'build_time': get_build_time(),
        'platform': get_build_platform(),
    }


def format_version_banner(app_name: str = "银行流水检验工具") -> str:
    """格式化版本横幅，用于启动时打印"""
    v = get_version()
    bt = get_build_time()
    bp = get_build_platform()
    platform_display = {"Windows": "Windows", "Darwin": "macOS", "Linux": "Linux"}.get(bp, bp)
    line = "=" * 48
    return (
        f"{line}\n"
        f"  {app_name}\n"
        f"  版本: v{v}\n"
        f"  构建时间: {bt}\n"
        f"  构建平台: {platform_display}\n"
        f"{line}"
    )


if __name__ == '__main__':
    print(format_version_banner())
