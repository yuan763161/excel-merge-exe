# Excel文件合并工具

一个简单易用的Windows应用程序，用于合并多个具有相同列名的Excel文件。

## 功能特点

- 支持批量选择多个Excel文件（.xlsx, .xls）
- 自动检查所有文件的列名是否一致
- 将所有数据合并到一个新的Excel文件中
- 美观的图形界面，操作简单直观
- 实时显示处理进度

## 使用方法

### 方法一：直接运行exe文件（推荐）

1. 双击 `dist` 文件夹中的 `Excel合并工具.exe`
2. 点击"选择Excel文件"按钮，选择要合并的多个Excel文件
3. 确认列表中显示了所有要合并的文件
4. 点击"开始合并"按钮
5. 选择保存位置和文件名
6. 等待合并完成

### 方法二：从源码运行

1. 安装Python 3.8或更高版本
2. 安装依赖包：
   ```bash
   pip install -r requirements.txt
   ```
3. 运行程序：
   ```bash
   python excel_merger.py
   ```

## 打包为exe

如果需要重新打包为exe文件，请在Windows系统上执行：

```bash
# 方法1：使用批处理文件
build_exe.bat

# 方法2：手动执行
pip install -r requirements.txt
pyinstaller --onefile --windowed --name "Excel合并工具" excel_merger.py
```

打包完成后，exe文件会生成在 `dist` 文件夹中。

## 注意事项

- 所有要合并的Excel文件必须具有相同的列名（列的顺序和名称都要相同）
- 合并时会按照文件选择的顺序依次添加数据
- 合并后的文件会保留所有原始数据，不会进行去重
- 建议在合并前备份原始文件

## 系统要求

- Windows 7/8/10/11
- 至少100MB可用磁盘空间
- 建议4GB以上内存（处理大文件时）

## 常见问题

**Q: 程序提示列名不一致怎么办？**
A: 请检查所有Excel文件的列名是否完全相同，包括列的顺序。可以先手动调整列名和顺序后再进行合并。

**Q: 可以合并多少个文件？**
A: 理论上没有限制，但受限于计算机内存。建议单次合并不超过100个文件。

**Q: 支持哪些Excel格式？**
A: 支持 .xlsx 和 .xls 格式的Excel文件。

## 技术栈

- Python 3.8+
- tkinter (GUI界面)
- pandas (数据处理)
- openpyxl (Excel读写)
- PyInstaller (exe打包)