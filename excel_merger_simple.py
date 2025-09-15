#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel合并工具 - 简化版
适合使用auto-py-to-exe或在线服务打包
"""

import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os
from pathlib import Path

class ExcelMerger:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Excel合并工具")
        self.root.geometry("600x500")
        self.root.resizable(False, False)

        # 设置窗口居中
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - 300
        y = (self.root.winfo_screenheight() // 2) - 250
        self.root.geometry(f"600x500+{x}+{y}")

        self.files = []
        self.setup_ui()

    def setup_ui(self):
        # 标题
        tk.Label(
            self.root,
            text="Excel 文件合并工具",
            font=("Arial", 18, "bold"),
            bg="#2196F3",
            fg="white",
            pady=15
        ).pack(fill="x")

        # 说明
        tk.Label(
            self.root,
            text="选择多个列名相同的Excel文件进行合并",
            font=("Arial", 10),
            pady=10
        ).pack()

        # 按钮框架
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=10)

        tk.Button(
            btn_frame,
            text="选择文件",
            command=self.select_files,
            bg="#4CAF50",
            fg="white",
            font=("Arial", 10),
            padx=20,
            pady=5
        ).pack(side="left", padx=5)

        tk.Button(
            btn_frame,
            text="清空列表",
            command=self.clear_files,
            bg="#f44336",
            fg="white",
            font=("Arial", 10),
            padx=20,
            pady=5
        ).pack(side="left", padx=5)

        # 文件列表
        list_frame = tk.Frame(self.root)
        list_frame.pack(fill="both", expand=True, padx=20, pady=10)

        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side="right", fill="y")

        self.listbox = tk.Listbox(
            list_frame,
            yscrollcommand=scrollbar.set,
            font=("Arial", 10),
            height=12
        )
        self.listbox.pack(fill="both", expand=True)
        scrollbar.config(command=self.listbox.yview)

        # 状态标签
        self.status = tk.Label(
            self.root,
            text="请选择要合并的Excel文件",
            font=("Arial", 10),
            fg="#666"
        )
        self.status.pack(pady=5)

        # 合并按钮
        self.merge_btn = tk.Button(
            self.root,
            text="开始合并",
            command=self.merge_files,
            bg="#2196F3",
            fg="white",
            font=("Arial", 12, "bold"),
            padx=30,
            pady=10,
            state="disabled"
        )
        self.merge_btn.pack(pady=15)

    def select_files(self):
        files = filedialog.askopenfilenames(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls")]
        )

        for file in files:
            if file not in self.files:
                self.files.append(file)
                self.listbox.insert(tk.END, os.path.basename(file))

        self.update_status()

    def clear_files(self):
        self.files = []
        self.listbox.delete(0, tk.END)
        self.update_status()

    def update_status(self):
        count = len(self.files)
        if count == 0:
            self.status.config(text="请选择要合并的Excel文件")
            self.merge_btn.config(state="disabled")
        elif count == 1:
            self.status.config(text="至少需要选择2个文件")
            self.merge_btn.config(state="disabled")
        else:
            self.status.config(text=f"已选择 {count} 个文件")
            self.merge_btn.config(state="normal")

    def merge_files(self):
        if len(self.files) < 2:
            messagebox.showwarning("警告", "请选择至少2个文件")
            return

        try:
            # 读取所有文件
            self.status.config(text="正在读取文件...")
            self.root.update()

            dfs = []
            first_cols = None

            for i, file in enumerate(self.files):
                self.status.config(text=f"读取 {i+1}/{len(self.files)}: {os.path.basename(file)}")
                self.root.update()

                df = pd.read_excel(file)

                if first_cols is None:
                    first_cols = list(df.columns)
                elif list(df.columns) != first_cols:
                    messagebox.showerror(
                        "错误",
                        f"文件列名不一致:\n{os.path.basename(file)}"
                    )
                    self.status.config(text="合并失败：列名不一致")
                    return

                dfs.append(df)

            # 合并数据
            self.status.config(text="正在合并数据...")
            self.root.update()

            merged = pd.concat(dfs, ignore_index=True)

            # 保存文件
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel文件", "*.xlsx")]
            )

            if save_path:
                merged.to_excel(save_path, index=False)
                messagebox.showinfo(
                    "成功",
                    f"合并完成！\n共 {len(merged)} 行数据\n保存至: {save_path}"
                )
                self.status.config(text=f"合并成功：{len(merged)} 行")
            else:
                self.status.config(text="已取消保存")

        except Exception as e:
            messagebox.showerror("错误", str(e))
            self.status.config(text="合并失败")

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = ExcelMerger()
    app.run()