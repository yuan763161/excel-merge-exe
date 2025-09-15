@echo off
chcp 65001 >nul
echo ====================================
echo    Excel合并工具 - 一键打包脚本
echo ====================================
echo.

REM 检查Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未检测到Python，请先安装Python 3.8或更高版本
    echo 下载地址: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo [1/4] 安装必要的包...
pip install pandas openpyxl pyinstaller -i https://pypi.org/simple

echo.
echo [2/4] 清理旧文件...
if exist build rd /s /q build
if exist dist rd /s /q dist
if exist *.spec del *.spec

echo.
echo [3/4] 开始打包为单个exe文件...
pyinstaller --onefile --windowed --noconfirm ^
    --name "Excel合并工具" ^
    --icon=NONE ^
    --add-data "*.py;." ^
    --hidden-import pandas ^
    --hidden-import openpyxl ^
    excel_merger_simple.py

echo.
echo [4/4] 打包完成！
echo.

if exist "dist\Excel合并工具.exe" (
    echo ✓ 成功生成exe文件！
    echo.
    echo 文件位置: %cd%\dist\Excel合并工具.exe
    echo 文件大小:
    dir "dist\Excel合并工具.exe" | find "Excel"
    echo.
    echo 可以将 dist\Excel合并工具.exe 复制到任何地方使用
    echo.
    explorer dist
) else (
    echo × 打包失败，请检查错误信息
)

echo.
pause