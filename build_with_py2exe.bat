@echo off
echo ========================================
echo Excel合并工具 - py2exe打包脚本
echo ========================================
echo.

REM 检查Python是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误: 未找到Python，请确保Python已正确安装并添加到PATH环境变量
    pause
    exit /b 1
)

echo 正在检查并安装必要的包...
echo.

REM 安装py2exe
echo 安装py2exe...
pip install py2exe
if errorlevel 1 (
    echo 错误: py2exe安装失败
    pause
    exit /b 1
)

REM 安装项目依赖
echo 安装项目依赖...
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
echo 正在使用py2exe打包程序...
python setup_py2exe.py py2exe
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
echo 生成的可执行文件位于: dist\ 目录下
echo 主文件: Excel合并工具.exe
echo.
echo 注意: 需要将整个dist目录复制到目标机器才能运行
echo.
pause