@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

set APP_NAME=bankcheck
set APP_NAME_CN=银行流水检验工具
set VERSION=1.0.0
set ENTRY_SCRIPT=bankcheck.py
set BUILD_INFO_FILE=build_info.py

echo ========================================
echo   %APP_NAME_CN% - 跨平台打包脚本
echo   版本: v%VERSION%
echo ========================================
echo.

REM ── 解析命令行参数 ──
set BUILD_MODE=onefile
set WITH_CONSOLE=--windowed
set CLEAN_BUILD=--clean
set TARGET_PLATFORM=auto
set ICON_FILE=

:parse_args
if "%~1"=="" goto end_parse
if /i "%~1"=="--onedir" set BUILD_MODE=onedir
if /i "%~1"=="--console" set WITH_CONSOLE=
if /i "%~1"=="--no-clean" set CLEAN_BUILD=
if /i "%~1"=="--win" set TARGET_PLATFORM=win
if /i "%~1"=="--mac" set TARGET_PLATFORM=mac
if /i "%~1"=="--macos" set TARGET_PLATFORM=mac
if /i "%~1"=="--all" set TARGET_PLATFORM=all
if /i "%~1"=="--icon" (
    set ICON_FILE=%~2
    shift
)
shift
goto parse_args
:end_parse

REM ── 检测当前运行平台 ──
set CURRENT_OS=windows
if defined OSTYPE (
    echo %OSTYPE% | findstr /i "darwin linux" >nul && set CURRENT_OS=unix
)

REM ── 自动选择目标平台 ──
if "%TARGET_PLATFORM%"=="auto" (
    if "%CURRENT_OS%"=="windows" (
        set TARGET_PLATFORM=win
    ) else (
        set TARGET_PLATFORM=mac
    )
)

echo [信息] 打包模式: %BUILD_MODE%
echo [信息] 目标平台: %TARGET_PLATFORM%
if "%WITH_CONSOLE%"=="" (
    echo [信息] 显示控制台窗口
) else (
    echo [信息] 隐藏控制台窗口（GUI 模式）
)
echo.

REM ── 安装依赖 ──
echo [1/5] 检查并安装依赖...
if exist requirements-lock.txt (
    echo       使用锁文件 (requirements-lock.txt) 安装依赖以确保可复现...
    pip install -r requirements-lock.txt
) else (
    echo       未找到锁文件，使用 requirements.txt 安装依赖...
    pip install -r requirements.txt
)
if %errorlevel% neq 0 (
    echo [错误] 依赖安装失败，请检查 Python 环境。
    pause
    exit /b 1
)
echo       依赖安装完成。
echo.

REM ── 生成构建信息 ──
echo [2/5] 生成构建信息...
for /f "tokens=2 delims==" %%I in ('wmic os get localdatetime /value') do set dt=%%I
set BUILD_YEAR=%dt:~0,4%
set BUILD_MONTH=%dt:~4,2%
set BUILD_DAY=%dt:~6,2%
set BUILD_HOUR=%dt:~8,2%
set BUILD_MINUTE=%dt:~10,2%
set BUILD_SECOND=%dt:~12,2%
set BUILD_TIME_STR=%BUILD_YEAR%-%BUILD_MONTH%-%BUILD_DAY% %BUILD_HOUR%:%BUILD_MINUTE%:%BUILD_SECOND%

set BUILD_PLATFORM_WIN=Windows
set BUILD_PLATFORM_MAC=Darwin

call :generate_build_info "%BUILD_PLATFORM_WIN%"
echo       版本: v%VERSION%
echo       构建时间: %BUILD_TIME_STR%
echo       构建平台: %BUILD_PLATFORM_WIN%
echo       构建信息已写入 %BUILD_INFO_FILE%
echo.

REM ── Windows 打包 ──
if "%TARGET_PLATFORM%"=="win" goto :build_windows
if "%TARGET_PLATFORM%"=="all" goto :build_windows
goto :skip_windows

:build_windows
echo [3/5] 准备 Windows 打包参数...

set PYI_ARGS_WIN=--name %APP_NAME% %WITH_CONSOLE% %CLEAN_BUILD%

if "%BUILD_MODE%"=="onefile" (
    set PYI_ARGS_WIN=%PYI_ARGS_WIN% --onefile
) else (
    set PYI_ARGS_WIN=%PYI_ARGS_WIN% --onedir
)

if not "%ICON_FILE%"=="" (
    if exist "%ICON_FILE%" (
        set PYI_ARGS_WIN=%PYI_ARGS_WIN% --icon %ICON_FILE%
    )
)

set PYI_ARGS_WIN=%PYI_ARGS_WIN% --add-data "i18n;i18n"
set PYI_ARGS_WIN=%PYI_ARGS_WIN% --add-data "static;static"
set PYI_ARGS_WIN=%PYI_ARGS_WIN% --add-data "templates;templates"
set PYI_ARGS_WIN=%PYI_ARGS_WIN% --add-data "bank_rules.yaml;."
set PYI_ARGS_WIN=%PYI_ARGS_WIN% --add-data "bank_directories.yaml;."
set PYI_ARGS_WIN=%PYI_ARGS_WIN% --add-data "task_presets.json;."
set PYI_ARGS_WIN=%PYI_ARGS_WIN% --add-data "task_queue_config.json;."
set PYI_ARGS_WIN=%PYI_ARGS_WIN% --add-data "scheduler_config.json;."
set PYI_ARGS_WIN=%PYI_ARGS_WIN% --add-data "requirements.txt;."

set PYI_ARGS_WIN=%PYI_ARGS_WIN% --collect-all openpyxl
set PYI_ARGS_WIN=%PYI_ARGS_WIN% --collect-all pandas
set PYI_ARGS_WIN=%PYI_ARGS_WIN% --collect-all yaml

echo       模式: %BUILD_MODE%
echo       包含资源: i18n, static, templates, 配置文件
echo.

echo [4/5] 正在打包 Windows 版本，请稍候...
echo.

pyinstaller %PYI_ARGS_WIN% %ENTRY_SCRIPT%

if %errorlevel% neq 0 (
    echo.
    echo [错误] Windows 打包失败，请检查上方日志。
    pause
    exit /b 1
)

echo.
echo   Windows 打包完成！
if "%BUILD_MODE%"=="onefile" (
    echo   输出文件: dist\%APP_NAME%.exe
) else (
    echo   输出目录: dist\%APP_NAME%\
    echo   主程序: dist\%APP_NAME%\%APP_NAME%.exe
)
echo.
goto :after_windows

:skip_windows
echo [3/5] 跳过 Windows 打包（当前非 Windows 环境或指定了 --mac）
echo.

:after_windows

REM ── macOS 打包 ──
if "%TARGET_PLATFORM%"=="mac" goto :build_mac
if "%TARGET_PLATFORM%"=="all" goto :build_mac
goto :skip_mac

:build_mac
echo [4/5] 准备 macOS 打包参数...

call :generate_build_info "%BUILD_PLATFORM_MAC%"

set PYI_ARGS_MAC=--name %APP_NAME% %WITH_CONSOLE% %CLEAN_BUILD%

if "%BUILD_MODE%"=="onefile" (
    set PYI_ARGS_MAC=%PYI_ARGS_MAC% --onefile
) else (
    set PYI_ARGS_MAC=%PYI_ARGS_MAC% --onedir
)

if not "%ICON_FILE%"=="" (
    if exist "%ICON_FILE%" (
        set PYI_ARGS_MAC=%PYI_ARGS_MAC% --icon %ICON_FILE%
    )
)

set PYI_ARGS_MAC=%PYI_ARGS_MAC% --add-data "i18n:i18n"
set PYI_ARGS_MAC=%PYI_ARGS_MAC% --add-data "static:static"
set PYI_ARGS_MAC=%PYI_ARGS_MAC% --add-data "templates:templates"
set PYI_ARGS_MAC=%PYI_ARGS_MAC% --add-data "bank_rules.yaml:."
set PYI_ARGS_MAC=%PYI_ARGS_MAC% --add-data "bank_directories.yaml:."
set PYI_ARGS_MAC=%PYI_ARGS_MAC% --add-data "task_presets.json:."
set PYI_ARGS_MAC=%PYI_ARGS_MAC% --add-data "task_queue_config.json:."
set PYI_ARGS_MAC=%PYI_ARGS_MAC% --add-data "scheduler_config.json:."
set PYI_ARGS_MAC=%PYI_ARGS_MAC% --add-data "requirements.txt:."

set PYI_ARGS_MAC=%PYI_ARGS_MAC% --collect-all openpyxl
set PYI_ARGS_MAC=%PYI_ARGS_MAC% --collect-all pandas
set PYI_ARGS_MAC=%PYI_ARGS_MAC% --collect-all yaml

echo       模式: %BUILD_MODE%
echo       包含资源: i18n, static, templates, 配置文件
echo.

echo [5/5] 正在打包 macOS 版本，请稍候...
echo.

pyinstaller %PYI_ARGS_MAC% %ENTRY_SCRIPT%

if %errorlevel% neq 0 (
    echo.
    echo [错误] macOS 打包失败，请检查上方日志。
    pause
    exit /b 1
)

echo.
echo   macOS 打包完成！
if "%BUILD_MODE%"=="onefile" (
    echo   输出文件: dist\%APP_NAME%
) else (
    echo   输出目录: dist\%APP_NAME%.app
    echo   主程序: dist\%APP_NAME%.app\Contents\MacOS\%APP_NAME%
)
echo.
goto :after_mac

:skip_mac
echo [5/5] 跳过 macOS 打包（指定了 --win）
echo.

:after_mac

REM ── 最终汇总 ──
echo ========================================
echo   打包流程完成！
echo   版本: v%VERSION%
echo   构建时间: %BUILD_TIME_STR%
echo.
if "%TARGET_PLATFORM%"=="win" (
    if "%BUILD_MODE%"=="onefile" (
        echo   Windows: dist\%APP_NAME%.exe
    ) else (
        echo   Windows: dist\%APP_NAME%\
    )
)
if "%TARGET_PLATFORM%"=="mac" (
    if "%BUILD_MODE%"=="onefile" (
        echo   macOS:   dist\%APP_NAME%
    ) else (
        echo   macOS:   dist\%APP_NAME%.app
    )
)
if "%TARGET_PLATFORM%"=="all" (
    if "%BUILD_MODE%"=="onefile" (
        echo   Windows: dist\%APP_NAME%.exe
        echo   macOS:   dist\%APP_NAME%
    ) else (
        echo   Windows: dist\%APP_NAME%\
        echo   macOS:   dist\%APP_NAME%.app
    )
)
echo ========================================
echo.
echo 提示:
echo   --onedir    打包为目录模式（启动更快，体积稍大）
echo   --console   显示控制台窗口（便于调试）
echo   --no-clean  不清理上次构建缓存
echo   --win       仅打包 Windows 版本
echo   --mac       仅打包 macOS 版本
echo   --all       同时打包 Windows 和 macOS 版本
echo   --icon      指定图标文件（Windows: .ico, macOS: .icns）
echo.
pause
exit /b 0

REM ── 子程序：生成 build_info.py ──
:generate_build_info
set BUILD_PLAT=%~1
(
echo # -*- coding: utf-8 -*-
echo """
echo 构建信息模块
echo 本文件由打包脚本自动生成，请勿手动修改。
echo """
echo.
echo import os
echo import sys
echo import platform
echo from datetime import datetime
echo.
echo __all__ = ['get_version', 'get_build_time', 'get_build_info', 'get_build_platform',
echo            'APP_VERSION', 'BUILD_TIME', 'BUILD_PLATFORM']
echo.
echo APP_VERSION = "%VERSION%"
echo BUILD_TIME = "%BUILD_TIME_STR%"
echo BUILD_PLATFORM = "%BUILD_PLAT%"
echo.
echo.
echo def _is_frozen():
echo     return getattr(sys, 'frozen', False)
echo.
echo.
echo def _detect_build_time():
echo     if BUILD_TIME:
echo         return BUILD_TIME
echo     try:
echo         if _is_frozen():
echo             exe_path = sys.executable
echo             ts = os.path.getmtime(exe_path)
echo         else:
echo             here = os.path.abspath(os.path.dirname(__file__))
echo             ts = os.path.getmtime(os.path.join(here, 'build_info.py'))
echo         return datetime.fromtimestamp(ts).strftime('%%Y-%%m-%%d %%H:%%M:%%S')
echo     except Exception:
echo         return datetime.now().strftime('%%Y-%%m-%%d %%H:%%M:%%S')
echo.
echo.
echo def _detect_build_platform():
echo     if BUILD_PLATFORM:
echo         return BUILD_PLATFORM
echo     return platform.system()
echo.
echo.
echo def get_version():
echo     return APP_VERSION
echo.
echo.
echo def get_build_time():
echo     return _detect_build_time()
echo.
echo.
echo def get_build_platform():
echo     return _detect_build_platform()
echo.
echo.
echo def get_build_info():
echo     return {
echo         'version': get_version(),
echo         'build_time': get_build_time(),
echo         'platform': get_build_platform(),
echo     }
echo.
echo.
echo def format_version_banner(app_name="%APP_NAME_CN%"):
echo     v = get_version()
echo     bt = get_build_time()
echo     bp = get_build_platform()
echo     platform_display = {"Windows": "Windows", "Darwin": "macOS", "Linux": "Linux"}.get(bp, bp)
echo     line = "=" * 48
echo     return (
echo         f"{{line}}\n"
echo         f"  {{app_name}}\n"
echo         f"  版本: v{{v}}\n"
echo         f"  构建时间: {{bt}}\n"
echo         f"  构建平台: {{platform_display}}\n"
echo         f"{{line}}"
echo     )
echo.
echo.
echo if __name__ == '__main__':
echo     print(format_version_banner())
) > %BUILD_INFO_FILE%
goto :eof
