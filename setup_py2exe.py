"""
py2exe setup script for Excel Merger Application
Windows-specific packaging solution
"""
from distutils.core import setup
import py2exe
import sys

# py2exe setup
setup(
    windows=[{
        'script': 'excel_merger.py',
        'dest_base': 'Excel合并工具',
        'icon_resources': [],  # You can add icon here: [(1, 'icon.ico')]
    }],
    options={
        'py2exe': {
            'packages': ['tkinter', 'pandas', 'openpyxl'],
            'includes': [
                'tkinter.ttk',
                'tkinter.filedialog',
                'tkinter.messagebox',
                'pathlib',
                'threading',
                'os'
            ],
            'excludes': [
                'test',
                'unittest',
                'email',
                'html',
                'http',
                'urllib',
                'xml',
                'pydoc',
                'doctest'
            ],
            'dll_excludes': [
                'w9xpopen.exe',
                'MSVCP90.dll',
                'mswsock.dll',
                'powrprof.dll'
            ],
            'optimize': 2,
            'compressed': True,
            'bundle_files': 2,  # 1 = everything in one file, 2 = DLLs separate, 3 = everything separate
        }
    },
    zipfile=None,  # Create library.zip
    data_files=[],
)