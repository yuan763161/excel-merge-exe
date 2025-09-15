@echo off
echo ========================================
echo Excel合并工具 - cx_Freeze打包脚本
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

REM 安装cx_Freeze
echo 安装cx_Freeze...
pip install cx_Freeze
if errorlevel 1 (
    echo 错误: cx_Freeze安装失败
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
echo 正在使用cx_Freeze打包程序...
python setup_cx_freeze.py build
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
echo 生成的可执行文件位于: build\exe.win-amd64-3.x\ 目录下
echo 文件名: Excel合并工具.exe
echo.
echo 将整个 build\exe.win-amd64-3.x\ 目录复制到目标机器即可运行
echo.
pause