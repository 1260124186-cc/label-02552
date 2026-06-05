@echo off
chcp 65001 >nul
echo ========================================
echo   银行流水检验工具 - 打包脚本
echo ========================================
echo.

pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo [错误] 依赖安装失败，请检查 Python 环境。
    pause
    exit /b 1
)

echo.
echo 正在打包为 exe，请稍候...
pyinstaller --onefile --windowed --name bankcheck bankcheck.py

if %errorlevel% neq 0 (
    echo [错误] 打包失败，请检查日志。
    pause
    exit /b 1
)

echo.
echo 打包完成！exe 文件位于 dist\bankcheck.exe
pause
