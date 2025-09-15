@echo off
echo ========================================
echo 运行应用程序测试
echo ========================================
echo.

REM 检查Python是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误: 未找到Python，请确保Python已正确安装并添加到PATH环境变量
    pause
    exit /b 1
)

echo 正在运行依赖测试...
echo.

python test_application.py

echo.
pause