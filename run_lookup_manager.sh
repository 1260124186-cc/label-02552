#!/usr/bin/env bash
# ========================================
#   主体查找表管理系统 - macOS / Linux 启动脚本
# ========================================

set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
BACKEND_DIR="$SCRIPT_DIR/backend"
REQ_FILE="$BACKEND_DIR/requirements.txt"
LOCK_FILE="$BACKEND_DIR/requirements-lock.txt"

echo "========================================"
echo "  主体查找表管理系统 - 启动脚本"
echo "========================================"
echo ""

detect_python() {
    if command -v python3 &>/dev/null; then
        PYTHON_CMD="python3"
    elif command -v python &>/dev/null; then
        PYTHON_CMD="python"
    else
        PYTHON_CMD=""
    fi
}

detect_python

if [ -z "$PYTHON_CMD" ]; then
    echo "[提示] 未检测到 Python，正在尝试自动安装..."

    if [ "$(uname)" = "Darwin" ]; then
        if command -v brew &>/dev/null; then
            echo "[信息] 通过 Homebrew 安装 Python..."
            brew install python3
        else
            echo "[信息] 未找到 Homebrew，正在先安装 Homebrew..."
            /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
            if [ -f /opt/homebrew/bin/brew ]; then
                eval "$(/opt/homebrew/bin/brew shellenv)"
            elif [ -f /usr/local/bin/brew ]; then
                eval "$(/usr/local/bin/brew shellenv)"
            fi
            brew install python3
        fi
    elif command -v apt-get &>/dev/null; then
        echo "[信息] 通过 apt 安装 Python..."
        if command -v sudo &>/dev/null; then
            sudo apt-get update && sudo apt-get install -y python3 python3-pip
        else
            apt-get update && apt-get install -y python3 python3-pip
        fi
    elif command -v dnf &>/dev/null; then
        echo "[信息] 通过 dnf 安装 Python..."
        if command -v sudo &>/dev/null; then
            sudo dnf install -y python3 python3-pip
        else
            dnf install -y python3 python3-pip
        fi
    elif command -v yum &>/dev/null; then
        echo "[信息] 通过 yum 安装 Python..."
        if command -v sudo &>/dev/null; then
            sudo yum install -y python3 python3-pip
        else
            yum install -y python3 python3-pip
        fi
    else
        echo "[错误] 无法自动安装 Python，请手动安装 Python 3.9 及以上版本。"
        echo "       下载地址: https://www.python.org/downloads/"
        exit 1
    fi

    detect_python
    if [ -z "$PYTHON_CMD" ]; then
        echo "[错误] Python 安装后仍未检测到，请检查安装是否成功并将其加入 PATH。"
        exit 1
    fi
fi

echo "[信息] 检测到 Python: $($PYTHON_CMD --version)"
echo ""

if ! $PYTHON_CMD -m pip --version &>/dev/null; then
    echo "[提示] 未检测到 pip，正在安装..."
    if $PYTHON_CMD -m ensurepip --upgrade 2>/dev/null; then
        echo "[信息] 通过 ensurepip 安装 pip 成功。"
    else
        echo "[提示] ensurepip 不可用，通过 get-pip.py 安装..."
        curl -sS https://bootstrap.pypa.io/get-pip.py | $PYTHON_CMD
    fi
fi

echo "[信息] 正在检查并安装依赖..."
if [ -f "$LOCK_FILE" ]; then
    echo "[信息] 使用锁文件 (requirements-lock.txt) 安装依赖以确保可复现..."
    $PYTHON_CMD -m pip install -r "$LOCK_FILE"
else
    echo "[信息] 未找到锁文件，使用 requirements.txt 安装依赖..."
    $PYTHON_CMD -m pip install -r "$REQ_FILE"
fi
if [ $? -ne 0 ]; then
    echo "[错误] 依赖安装失败，请检查网络连接或 Python 环境。"
    exit 1
fi

echo ""
echo "[信息] 依赖已就绪，正在启动主体查找表管理系统..."
echo ""
echo "========================================"
echo "  启动完成后，请在浏览器中访问:"
echo "  http://127.0.0.1:5000"
echo "========================================"
echo ""
$PYTHON_CMD "$BACKEND_DIR/app.py"
