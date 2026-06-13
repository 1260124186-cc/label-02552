#!/usr/bin/env bash
# ========================================
#   银行流水检验工具 - 跨平台打包脚本
#   支持生成 Windows exe / macOS app 包
# ========================================

set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

APP_NAME="bankcheck"
APP_NAME_CN="银行流水检验工具"
VERSION="1.0.0"
ENTRY_SCRIPT="bankcheck.py"
BUILD_INFO_FILE="build_info.py"
ICON_FILE=""

BUILD_MODE="onefile"
WITH_CONSOLE="--windowed"
CLEAN_BUILD="--clean"
CREATE_DMG=false
TARGET_PLATFORM="auto"

usage() {
    echo "用法: $0 [选项]"
    echo ""
    echo "选项:"
    echo "  --onedir      打包为目录模式（启动更快，体积稍大）"
    echo "  --console     显示控制台窗口（便于调试）"
    echo "  --no-clean    不清理上次构建缓存"
    echo "  --dmg         打包后生成 .dmg 磁盘镜像（仅 macOS onedir 模式）"
    echo "  --icon <文件> 指定应用图标 (.icns / .ico)"
    echo "  --win         仅打包 Windows 版本"
    echo "  --mac         仅打包 macOS 版本"
    echo "  --all         同时打包 Windows 和 macOS 版本"
    echo "  -h, --help    显示帮助"
    exit 0
}

while [[ $# -gt 0 ]]; do
    case "$1" in
        --onedir)
            BUILD_MODE="onedir"
            shift
            ;;
        --console)
            WITH_CONSOLE=""
            shift
            ;;
        --no-clean)
            CLEAN_BUILD=""
            shift
            ;;
        --dmg)
            CREATE_DMG=true
            shift
            ;;
        --icon)
            ICON_FILE="$2"
            shift 2
            ;;
        --win)
            TARGET_PLATFORM="win"
            shift
            ;;
        --mac|--macos)
            TARGET_PLATFORM="mac"
            shift
            ;;
        --all)
            TARGET_PLATFORM="all"
            shift
            ;;
        -h|--help)
            usage
            ;;
        *)
            echo "[警告] 未知参数: $1"
            shift
            ;;
    esac
done

CURRENT_OS="$(uname -s)"

if [ "$TARGET_PLATFORM" = "auto" ]; then
    case "$CURRENT_OS" in
        Darwin*)  TARGET_PLATFORM="mac" ;;
        MINGW*|MSYS*|CYGWIN*|Windows_NT) TARGET_PLATFORM="win" ;;
        Linux*)   TARGET_PLATFORM="mac" ;;
        *)        TARGET_PLATFORM="mac" ;;
    esac
fi

echo "========================================"
echo "  ${APP_NAME_CN} - 跨平台打包脚本"
echo "  版本: v${VERSION}"
echo "========================================"
echo ""
echo "[信息] 打包模式: ${BUILD_MODE}"
echo "[信息] 目标平台: ${TARGET_PLATFORM}"
echo "[信息] 当前系统: ${CURRENT_OS}"
if [ -z "$WITH_CONSOLE" ]; then
    echo "[信息] 显示控制台窗口"
else
    echo "[信息] 隐藏控制台窗口（GUI 模式）"
fi
echo ""

# ── 安装依赖 ──
echo "[1/5] 检查并安装依赖..."

if command -v python3 &>/dev/null; then
    PYTHON_CMD="python3"
elif command -v python &>/dev/null; then
    PYTHON_CMD="python"
else
    echo "[错误] 未找到 Python，请先安装 Python 3.9 及以上版本。"
    exit 1
fi

if [ -f "requirements-lock.txt" ]; then
    echo "      使用锁文件 (requirements-lock.txt) 安装依赖以确保可复现..."
    $PYTHON_CMD -m pip install -r requirements-lock.txt
else
    echo "      未找到锁文件，使用 requirements.txt 安装依赖..."
    $PYTHON_CMD -m pip install -r requirements.txt
fi

echo "      依赖安装完成。"
echo ""

# ── 生成构建信息 ──
echo "[2/5] 生成构建信息..."

BUILD_TIME_STR=$(date '+%Y-%m-%d %H:%M:%S')

generate_build_info() {
    local build_plat="$1"
    cat > "$BUILD_INFO_FILE" << EOF
# -*- coding: utf-8 -*-
"""
构建信息模块
本文件由打包脚本自动生成，请勿手动修改。
"""

import os
import sys
import platform
from datetime import datetime

__all__ = ['get_version', 'get_build_time', 'get_build_info', 'get_build_platform',
           'APP_VERSION', 'BUILD_TIME', 'BUILD_PLATFORM']

APP_VERSION = "${VERSION}"
BUILD_TIME = "${BUILD_TIME_STR}"
BUILD_PLATFORM = "${build_plat}"


def _is_frozen():
    return getattr(sys, 'frozen', False)


def _detect_build_time():
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


def _detect_build_platform():
    if BUILD_PLATFORM:
        return BUILD_PLATFORM
    return platform.system()


def get_version():
    return APP_VERSION


def get_build_time():
    return _detect_build_time()


def get_build_platform():
    return _detect_build_platform()


def get_build_info():
    return {
        'version': get_version(),
        'build_time': get_build_time(),
        'platform': get_build_platform(),
    }


def format_version_banner(app_name="${APP_NAME_CN}"):
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
EOF
    echo "      构建信息已写入 ${BUILD_INFO_FILE} (平台: ${build_plat})"
}

generate_build_info "Darwin"
echo "      版本: v${VERSION}"
echo "      构建时间: ${BUILD_TIME_STR}"
echo ""

# ── 公共 PyInstaller 资源参数 ──
build_pyinstaller_args() {
    local separator="$1"
    local args="--name ${APP_NAME} ${WITH_CONSOLE} ${CLEAN_BUILD}"

    if [ "$BUILD_MODE" = "onefile" ]; then
        args="$args --onefile"
    else
        args="$args --onedir"
    fi

    if [ -n "$ICON_FILE" ] && [ -f "$ICON_FILE" ]; then
        args="$args --icon ${ICON_FILE}"
        echo "      应用图标: ${ICON_FILE}"
    fi

    args="$args --add-data 'i18n${separator}i18n'"
    args="$args --add-data 'static${separator}static'"
    args="$args --add-data 'templates${separator}templates'"
    args="$args --add-data 'bank_rules.yaml${separator}.'"
    args="$args --add-data 'bank_directories.yaml${separator}.'"
    args="$args --add-data 'task_presets.json${separator}.'"
    args="$args --add-data 'task_queue_config.json${separator}.'"
    args="$args --add-data 'scheduler_config.json${separator}.'"
    args="$args --add-data 'requirements.txt${separator}.'"

    args="$args --collect-all openpyxl"
    args="$args --collect-all pandas"
    args="$args --collect-all yaml"

    echo "$args"
}

# ── macOS 打包 ──
if [ "$TARGET_PLATFORM" = "mac" ] || [ "$TARGET_PLATFORM" = "all" ]; then
    echo "[3/5] 准备 macOS 打包参数..."

    generate_build_info "Darwin"

    PYI_ARGS_MAC=$(build_pyinstaller_args ":")

    echo "      模式: ${BUILD_MODE}"
    echo "      包含资源: i18n, static, templates, 配置文件"
    echo ""

    echo "[4/5] 正在打包 macOS 版本，请稍候..."
    echo ""

    $PYTHON_CMD -m PyInstaller $PYI_ARGS_MAC "$ENTRY_SCRIPT"

    echo ""
    echo "  macOS 打包完成！"
    if [ "$BUILD_MODE" = "onefile" ]; then
        echo "  输出文件: dist/${APP_NAME}"
    else
        echo "  输出目录: dist/${APP_NAME}.app"
        echo "  主程序: dist/${APP_NAME}.app/Contents/MacOS/${APP_NAME}"
    fi
    echo ""

    if [ "$CREATE_DMG" = true ] && [ -d "dist/${APP_NAME}.app" ]; then
        echo ""
        echo "[附加] 正在生成 .dmg 磁盘镜像..."

        DMG_NAME="${APP_NAME}_v${VERSION}.dmg"
        FINAL_DMG="dist/${DMG_NAME}"

        if command -v hdiutil &>/dev/null; then
            hdiutil create -volname "${APP_NAME_CN} v${VERSION}" \
                -srcfolder "dist/${APP_NAME}.app" \
                -ov -format UDZO \
                "$FINAL_DMG"

            echo "  DMG 镜像: ${FINAL_DMG}"
        else
            echo "[警告] 未找到 hdiutil，无法生成 DMG。"
        fi
    fi
else
    echo "[3/5] 跳过 macOS 打包（指定了 --win）"
    echo ""
fi

# ── Windows 打包 ──
if [ "$TARGET_PLATFORM" = "win" ] || [ "$TARGET_PLATFORM" = "all" ]; then
    echo "[4/5] 准备 Windows 打包参数..."

    generate_build_info "Windows"

    PYI_ARGS_WIN=$(build_pyinstaller_args ";")

    echo "      模式: ${BUILD_MODE}"
    echo "      包含资源: i18n, static, templates, 配置文件"
    echo ""

    echo "[5/5] 正在打包 Windows 版本，请稍候..."
    echo ""

    $PYTHON_CMD -m PyInstaller $PYI_ARGS_WIN "$ENTRY_SCRIPT"

    echo ""
    echo "  Windows 打包完成！"
    if [ "$BUILD_MODE" = "onefile" ]; then
        echo "  输出文件: dist/${APP_NAME}.exe"
    else
        echo "  输出目录: dist/${APP_NAME}/"
        echo "  主程序: dist/${APP_NAME}/${APP_NAME}.exe"
    fi
    echo ""
else
    if [ "$TARGET_PLATFORM" = "mac" ]; then
        echo "[5/5] 跳过 Windows 打包（指定了 --mac）"
        echo ""
    fi
fi

# ── 最终汇总 ──
echo "========================================"
echo "  打包流程完成！"
echo "  版本: v${VERSION}"
echo "  构建时间: ${BUILD_TIME_STR}"
echo ""

case "$TARGET_PLATFORM" in
    win)
        if [ "$BUILD_MODE" = "onefile" ]; then
            echo "  Windows: dist/${APP_NAME}.exe"
        else
            echo "  Windows: dist/${APP_NAME}/"
        fi
        ;;
    mac)
        if [ "$BUILD_MODE" = "onefile" ]; then
            echo "  macOS:   dist/${APP_NAME}"
        else
            echo "  macOS:   dist/${APP_NAME}.app"
        fi
        ;;
    all)
        if [ "$BUILD_MODE" = "onefile" ]; then
            echo "  Windows: dist/${APP_NAME}.exe"
            echo "  macOS:   dist/${APP_NAME}"
        else
            echo "  Windows: dist/${APP_NAME}/"
            echo "  macOS:   dist/${APP_NAME}.app"
        fi
        ;;
esac
echo "========================================"

echo ""
echo "提示:"
echo "  --onedir    打包为 .app 目录模式（启动更快，体积稍大）"
echo "  --console   显示控制台窗口（便于调试）"
echo "  --no-clean  不清理上次构建缓存"
echo "  --dmg       打包后生成 .dmg 磁盘镜像（仅 macOS onedir 模式）"
echo "  --icon      指定图标文件（Windows: .ico, macOS: .icns）"
echo "  --win       仅打包 Windows 版本"
echo "  --mac       仅打包 macOS 版本"
echo "  --all       同时打包 Windows 和 macOS 版本"
echo ""
