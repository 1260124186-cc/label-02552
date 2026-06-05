@echo off
chcp 65001 >nul
echo ========================================
echo   银行流水检验工具 - Windows 启动脚本
echo ========================================
echo.

REM 检测 Python 是否已安装
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

REM 检测 pip 是否可用
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

REM 检测并安装依赖
echo [信息] 正在检查并安装依赖...
python -m pip install -r backend\requirements.txt
if %errorlevel% neq 0 (
    echo [错误] 依赖安装失败，请检查网络连接或 Python 环境。
    pause
    exit /b 1
)

echo.
echo [信息] 依赖已就绪，正在启动程序...
echo.
python backend\bankcheck.py

if %errorlevel% neq 0 (
    echo.
    echo [错误] 程序运行异常，请查看上方日志或 backend\bankcheck.log 获取详情。
    pause
    exit /b 1
)

pause
