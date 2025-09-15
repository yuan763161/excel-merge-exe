@echo off
echo ========================================
echo Excel合并工具 - PyInstaller Spec文件打包脚本
echo ========================================
echo.

REM 检查Python是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误: 未找到Python，请确保Python已正确安装并添加到PATH环境变量
    pause
    exit /b 1
)

echo 正在安装依赖包...
pip install -r requirements.txt
if errorlevel 1 (
    echo 错误: 依赖包安装失败
    pause
    exit /b 1
)

echo.
echo 正在清理旧的构建文件...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo.
echo 正在使用PyInstaller spec文件打包程序...
pyinstaller pyinstaller_spec.spec
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
echo.
pause