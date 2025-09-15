@echo off
echo ========================================
echo Excel合并工具 - PyInstaller打包脚本 (改进版)
echo ========================================
echo.

REM 检查Python是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误: 未找到Python，请确保Python已正确安装并添加到PATH环境变量
    pause
    exit /b 1
)

REM 安装依赖
echo 正在安装依赖包...
pip install -r requirements.txt
if errorlevel 1 (
    echo 错误: 依赖包安装失败
    pause
    exit /b 1
)
echo.

REM 清理之前的构建文件
echo 正在清理旧的构建文件...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist "Excel合并工具.spec" del "Excel合并工具.spec"
echo.

REM 使用PyInstaller打包 - 添加更多选项以提高兼容性
echo 正在打包程序为exe...
pyinstaller ^
    --onefile ^
    --windowed ^
    --name "Excel合并工具" ^
    --add-data="requirements.txt;." ^
    --hidden-import=tkinter ^
    --hidden-import=tkinter.ttk ^
    --hidden-import=tkinter.filedialog ^
    --hidden-import=tkinter.messagebox ^
    --hidden-import=pandas ^
    --hidden-import=openpyxl ^
    --exclude-module=test ^
    --exclude-module=unittest ^
    --exclude-module=email ^
    --exclude-module=html ^
    --exclude-module=http ^
    --exclude-module=urllib ^
    --exclude-module=xml ^
    excel_merger.py

if errorlevel 1 (
    echo 错误: 打包失败
    pause
    exit /b 1
)

echo.
echo ========================================
echo 打包完成！
echo ========================================
echo.
echo 生成的exe文件位于: dist\Excel合并工具.exe
echo 文件大小:
dir "dist\Excel合并工具.exe" | find "Excel合并工具.exe"
echo.
echo 这是一个独立的可执行文件，可以在任何Windows机器上运行
echo.
pause