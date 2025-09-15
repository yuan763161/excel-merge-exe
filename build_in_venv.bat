@echo off
echo ========================================
echo 在虚拟环境中打包Excel合并工具
echo ========================================
echo.

REM 检查虚拟环境是否存在
if not exist "venv_packaging" (
    echo 虚拟环境不存在，正在创建...
    call create_virtual_env.bat
    if errorlevel 1 (
        echo 虚拟环境创建失败
        pause
        exit /b 1
    )
)

echo 激活虚拟环境...
call venv_packaging\Scripts\activate.bat

echo.
echo 确保依赖已安装...
pip install -r requirements.txt
pip install pyinstaller

echo.
echo 清理旧的构建文件...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo.
echo 在虚拟环境中打包...
pyinstaller ^
    --onefile ^
    --windowed ^
    --name "Excel合并工具" ^
    --hidden-import=tkinter ^
    --hidden-import=tkinter.ttk ^
    --hidden-import=tkinter.filedialog ^
    --hidden-import=tkinter.messagebox ^
    --hidden-import=pandas ^
    --hidden-import=openpyxl ^
    --exclude-module=test ^
    --exclude-module=unittest ^
    excel_merger.py

if errorlevel 1 (
    echo 打包失败
    pause
    exit /b 1
)

echo.
echo 停用虚拟环境...
call deactivate

echo.
echo ========================================
echo 打包完成！
echo ========================================
echo.
echo 生成的exe文件位于: dist\Excel合并工具.exe
echo.
echo 虚拟环境位于: venv_packaging\
echo 可以安全删除虚拟环境以节省空间
echo.
pause