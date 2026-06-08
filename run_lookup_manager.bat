@echo off
chcp 65001 >nul
echo ========================================
echo   主体查找表管理系统 - Windows 启动脚本
echo ========================================
echo.

where python >nul 2>nul
if %errorlevel% neq 0 (
    echo [提示] 未检测到 Python，正在尝试通过 winget 自动安装...
    where winget >nul 2>nul
    if %errorlevel% neq 0 (
        echo [错误] 未找到 winget 包管理器，无法自动安装 Python。
        echo         请手动前往 https://www.python.org/downloads/ 下载安装 Python 3.9 及以上版本。
        echo         安装时请勾选 "Add Python to PATH"。
        pause
        exit /b 1
    )
    winget install Python.Python.3.12 --accept-package-agreements --accept-source-agreements
    if %errorlevel% neq 0 (
        echo [错误] 自动安装 Python 失败。
        echo         请手动前往 https://www.python.org/downloads/ 下载安装 Python 3.9 及以上版本。
        echo         安装时请勾选 "Add Python to PATH"。
        pause
        exit /b 1
    )
    echo [提示] Python 安装完成，请关闭此窗口并重新运行本脚本以使环境变量生效。
    pause
    exit /b 0
)

echo [信息] 检测到 Python:
python --version
echo.

python -m pip --version >nul 2>nul
if %errorlevel% neq 0 (
    echo [提示] 未检测到 pip，正在通过 ensurepip 安装...
    python -m ensurepip --upgrade >nul 2>nul
    if %errorlevel% neq 0 (
        echo [错误] pip 安装失败，请手动安装 pip。
        echo         参考: https://pip.pypa.io/en/stable/installation/
        pause
        exit /b 1
    )
)

echo [信息] 正在检查并安装依赖...
if exist backend\requirements-lock.txt (
    echo [信息] 使用锁文件 (requirements-lock.txt) 安装依赖以确保可复现...
    python -m pip install -r backend\requirements-lock.txt
) else (
    echo [信息] 未找到锁文件，使用 requirements.txt 安装依赖...
    python -m pip install -r backend\requirements.txt
)
if %errorlevel% neq 0 (
    echo [错误] 依赖安装失败，请检查网络连接或 Python 环境。
    pause
    exit /b 1
)

echo.
echo [信息] 依赖已就绪，正在启动主体查找表管理系统...
echo.
echo ========================================
echo   启动完成后，请在浏览器中访问:
echo   http://127.0.0.1:5000
echo ========================================
echo.
python backend\app.py

if %errorlevel% neq 0 (
    echo.
    echo [错误] 程序运行异常，请查看上方日志或 backend\lookup_manager.log 获取详情。
    pause
    exit /b 1
)

pause
