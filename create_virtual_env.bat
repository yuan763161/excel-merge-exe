@echo off
echo ========================================
echo 创建Python虚拟环境用于打包
echo ========================================
echo.

REM 检查Python是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误: 未找到Python，请确保Python已正确安装并添加到PATH环境变量
    pause
    exit /b 1
)

echo 正在创建虚拟环境...
python -m venv venv_packaging
if errorlevel 1 (
    echo 错误: 虚拟环境创建失败
    pause
    exit /b 1
)

echo.
echo 激活虚拟环境...
call venv_packaging\Scripts\activate.bat

echo.
echo 升级pip...
python -m pip install --upgrade pip

echo.
echo 安装基础依赖...
pip install -r requirements.txt

echo.
echo 安装所有打包工具...
pip install pyinstaller cx_Freeze
pip install py2exe

echo.
echo ========================================
echo 虚拟环境创建完成！
echo ========================================
echo.
echo 使用方法:
echo 1. 运行: venv_packaging\Scripts\activate.bat
echo 2. 然后运行任意打包脚本
echo 3. 完成后运行: deactivate
echo.
echo 或者直接运行: build_in_venv.bat
echo.
pause