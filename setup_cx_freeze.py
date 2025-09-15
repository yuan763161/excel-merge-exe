"""
cx_Freeze setup script for Excel Merger Application
Alternative packaging method that works well for tkinter applications
"""
import sys
from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but it might need fine tuning.
build_options = {
    'packages': ['tkinter', 'pandas', 'openpyxl', 'pathlib', 'threading', 'os'],
    'excludes': ['test', 'unittest', 'email', 'html', 'http', 'urllib', 'xml'],
    'include_files': [],
    'optimize': 1,
    'zip_include_packages': ['encodings', 'importlib']
}

base = 'Win32GUI' if sys.platform == 'win32' else None

executables = [
    Executable(
        'excel_merger.py',
        base=base,
        target_name='Excel合并工具.exe',
        icon=None  # You can add an .ico file path here if you have one
    )
]

setup(
    name='Excel合并工具',
    version='1.0',
    description='Excel文件合并工具',
    author='Your Name',
    options={'build_exe': build_options},
    executables=executables
)