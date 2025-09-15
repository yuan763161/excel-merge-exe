#!/usr/bin/env python3
"""
Test script to verify that all dependencies are working correctly
Run this before packaging to ensure everything works
"""

import sys
import traceback

def test_imports():
    """Test all required imports"""
    print("Testing imports...")

    try:
        import tkinter as tk
        print("✓ tkinter imported successfully")
    except ImportError as e:
        print(f"✗ tkinter import failed: {e}")
        return False

    try:
        from tkinter import ttk, filedialog, messagebox
        print("✓ tkinter submodules imported successfully")
    except ImportError as e:
        print(f"✗ tkinter submodules import failed: {e}")
        return False

    try:
        import pandas as pd
        print(f"✓ pandas {pd.__version__} imported successfully")
    except ImportError as e:
        print(f"✗ pandas import failed: {e}")
        return False

    try:
        import openpyxl
        print(f"✓ openpyxl {openpyxl.__version__} imported successfully")
    except ImportError as e:
        print(f"✗ openpyxl import failed: {e}")
        return False

    try:
        from pathlib import Path
        import threading
        import os
        print("✓ Standard library modules imported successfully")
    except ImportError as e:
        print(f"✗ Standard library import failed: {e}")
        return False

    return True

def test_tkinter_functionality():
    """Test basic tkinter functionality"""
    print("\nTesting tkinter functionality...")

    try:
        root = tk.Tk()
        root.withdraw()  # Hide the window

        # Test creating widgets
        frame = tk.Frame(root)
        label = tk.Label(frame, text="Test")
        button = tk.Button(frame, text="Test Button")
        listbox = tk.Listbox(frame)
        progress = ttk.Progressbar(frame)

        print("✓ Basic tkinter widgets created successfully")

        # Clean up
        root.destroy()
        return True

    except Exception as e:
        print(f"✗ tkinter functionality test failed: {e}")
        traceback.print_exc()
        return False

def test_pandas_functionality():
    """Test basic pandas functionality"""
    print("\nTesting pandas functionality...")

    try:
        # Create test dataframe
        df = pd.DataFrame({
            'Column1': [1, 2, 3],
            'Column2': ['A', 'B', 'C']
        })

        # Test operations that our app uses
        columns = list(df.columns)
        concat_df = pd.concat([df, df], ignore_index=True)

        print("✓ Basic pandas operations work correctly")
        return True

    except Exception as e:
        print(f"✗ pandas functionality test failed: {e}")
        traceback.print_exc()
        return False

def test_openpyxl_functionality():
    """Test basic openpyxl functionality"""
    print("\nTesting openpyxl functionality...")

    try:
        import tempfile
        import os

        # Create a test Excel file
        df = pd.DataFrame({
            'Test1': [1, 2, 3],
            'Test2': ['X', 'Y', 'Z']
        })

        # Test writing Excel file
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
            temp_path = tmp.name

        df.to_excel(temp_path, index=False)

        # Test reading Excel file
        read_df = pd.read_excel(temp_path)

        # Clean up
        os.unlink(temp_path)

        print("✓ Excel read/write operations work correctly")
        return True

    except Exception as e:
        print(f"✗ openpyxl functionality test failed: {e}")
        traceback.print_exc()
        return False

def main():
    """Main test function"""
    print("=" * 50)
    print("Excel合并工具 - 依赖测试")
    print("=" * 50)
    print(f"Python version: {sys.version}")
    print(f"Platform: {sys.platform}")
    print()

    tests = [
        test_imports,
        test_tkinter_functionality,
        test_pandas_functionality,
        test_openpyxl_functionality
    ]

    results = []
    for test in tests:
        try:
            result = test()
            results.append(result)
        except Exception as e:
            print(f"✗ Test {test.__name__} crashed: {e}")
            traceback.print_exc()
            results.append(False)

    print("\n" + "=" * 50)
    print("测试结果汇总")
    print("=" * 50)

    if all(results):
        print("✓ 所有测试通过！应用程序可以正常打包")
        return 0
    else:
        print("✗ 部分测试失败，请检查依赖安装")
        failed_tests = sum(1 for r in results if not r)
        print(f"失败测试数量: {failed_tests}/{len(results)}")
        return 1

if __name__ == "__main__":
    exit_code = main()

    print("\n按任意键退出...")
    try:
        input()
    except:
        pass

    sys.exit(exit_code)