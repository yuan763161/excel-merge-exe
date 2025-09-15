import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from pathlib import Path
import threading
import os

class ExcelMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel文件合并工具")
        self.root.geometry("800x600")

        # 设置现代化的颜色主题
        self.bg_color = "#f0f0f0"
        self.primary_color = "#2196F3"
        self.secondary_color = "#4CAF50"
        self.danger_color = "#f44336"
        self.text_color = "#333333"

        self.root.configure(bg=self.bg_color)

        # 存储选中的文件
        self.selected_files = []

        self.setup_ui()

    def setup_ui(self):
        # 标题
        title_frame = tk.Frame(self.root, bg=self.primary_color)
        title_frame.pack(fill="x", pady=(0, 20))

        title_label = tk.Label(
            title_frame,
            text="Excel 文件合并工具",
            font=("Microsoft YaHei", 20, "bold"),
            bg=self.primary_color,
            fg="white",
            pady=20
        )
        title_label.pack()

        # 说明文字
        desc_label = tk.Label(
            self.root,
            text="选择多个列名相同的Excel文件，将它们合并为一个文件",
            font=("Microsoft YaHei", 11),
            bg=self.bg_color,
            fg=self.text_color
        )
        desc_label.pack(pady=(0, 20))

        # 文件选择区域
        file_frame = tk.Frame(self.root, bg=self.bg_color)
        file_frame.pack(pady=10, padx=20, fill="both", expand=True)

        # 选择文件按钮
        btn_frame = tk.Frame(file_frame, bg=self.bg_color)
        btn_frame.pack(fill="x", pady=(0, 10))

        self.select_btn = tk.Button(
            btn_frame,
            text="选择Excel文件",
            command=self.select_files,
            font=("Microsoft YaHei", 11),
            bg=self.primary_color,
            fg="white",
            padx=20,
            pady=10,
            relief="flat",
            cursor="hand2"
        )
        self.select_btn.pack(side="left", padx=(0, 10))

        self.clear_btn = tk.Button(
            btn_frame,
            text="清空列表",
            command=self.clear_files,
            font=("Microsoft YaHei", 11),
            bg=self.danger_color,
            fg="white",
            padx=20,
            pady=10,
            relief="flat",
            cursor="hand2"
        )
        self.clear_btn.pack(side="left")

        # 文件列表
        list_frame = tk.Frame(file_frame, bg="white", relief="ridge", bd=1)
        list_frame.pack(fill="both", expand=True)

        # 列表标题
        list_title = tk.Label(
            list_frame,
            text="已选择的文件：",
            font=("Microsoft YaHei", 10),
            bg="white",
            fg=self.text_color,
            anchor="w"
        )
        list_title.pack(fill="x", padx=10, pady=(10, 5))

        # 文件列表框
        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side="right", fill="y")

        self.file_listbox = tk.Listbox(
            list_frame,
            yscrollcommand=scrollbar.set,
            font=("Microsoft YaHei", 10),
            bg="white",
            fg=self.text_color,
            selectbackground=self.primary_color,
            selectforeground="white",
            relief="flat",
            bd=0
        )
        self.file_listbox.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        scrollbar.config(command=self.file_listbox.yview)

        # 状态标签
        self.status_label = tk.Label(
            self.root,
            text="请选择要合并的Excel文件",
            font=("Microsoft YaHei", 10),
            bg=self.bg_color,
            fg=self.text_color
        )
        self.status_label.pack(pady=10)

        # 进度条
        self.progress = ttk.Progressbar(
            self.root,
            mode='indeterminate',
            length=760
        )
        self.progress.pack(pady=(0, 10), padx=20)

        # 合并按钮
        self.merge_btn = tk.Button(
            self.root,
            text="开始合并",
            command=self.merge_files,
            font=("Microsoft YaHei", 12, "bold"),
            bg=self.secondary_color,
            fg="white",
            padx=30,
            pady=12,
            relief="flat",
            cursor="hand2",
            state="disabled"
        )
        self.merge_btn.pack(pady=(10, 20))

        # 添加鼠标悬停效果
        self.add_hover_effects()

    def add_hover_effects(self):
        def on_enter(e, btn, color):
            btn.config(bg=self.darken_color(color))

        def on_leave(e, btn, color):
            btn.config(bg=color)

        buttons = [
            (self.select_btn, self.primary_color),
            (self.clear_btn, self.danger_color),
            (self.merge_btn, self.secondary_color)
        ]

        for btn, color in buttons:
            btn.bind("<Enter>", lambda e, b=btn, c=color: on_enter(e, b, c))
            btn.bind("<Leave>", lambda e, b=btn, c=color: on_leave(e, b, c))

    def darken_color(self, color):
        # 简单的颜色变暗函数
        color_map = {
            "#2196F3": "#1976D2",
            "#4CAF50": "#388E3C",
            "#f44336": "#d32f2f"
        }
        return color_map.get(color, color)

    def select_files(self):
        files = filedialog.askopenfilenames(
            title="选择Excel文件",
            filetypes=[
                ("Excel文件", "*.xlsx *.xls"),
                ("所有文件", "*.*")
            ]
        )

        if files:
            for file in files:
                if file not in self.selected_files:
                    self.selected_files.append(file)
                    self.file_listbox.insert(tk.END, os.path.basename(file))

            self.update_status(f"已选择 {len(self.selected_files)} 个文件")

            if len(self.selected_files) >= 2:
                self.merge_btn.config(state="normal")

    def clear_files(self):
        self.selected_files = []
        self.file_listbox.delete(0, tk.END)
        self.update_status("请选择要合并的Excel文件")
        self.merge_btn.config(state="disabled")

    def update_status(self, message):
        self.status_label.config(text=message)
        self.root.update()

    def merge_files(self):
        if len(self.selected_files) < 2:
            messagebox.showwarning("警告", "请选择至少2个Excel文件进行合并")
            return

        # 在新线程中执行合并操作
        thread = threading.Thread(target=self.perform_merge)
        thread.daemon = True
        thread.start()

    def perform_merge(self):
        try:
            # 禁用按钮
            self.merge_btn.config(state="disabled")
            self.select_btn.config(state="disabled")
            self.clear_btn.config(state="disabled")

            # 开始进度条
            self.progress.start(10)
            self.update_status("正在读取文件...")

            # 读取所有Excel文件
            dataframes = []
            file_columns = {}

            for i, file in enumerate(self.selected_files):
                self.update_status(f"正在读取文件 {i+1}/{len(self.selected_files)}: {os.path.basename(file)}")
                df = pd.read_excel(file)
                dataframes.append(df)
                file_columns[os.path.basename(file)] = list(df.columns)

            # 检查列名是否一致
            self.update_status("检查列名一致性...")
            first_columns = list(dataframes[0].columns)

            for i, df in enumerate(dataframes[1:], 1):
                if list(df.columns) != first_columns:
                    self.progress.stop()
                    messagebox.showerror(
                        "错误",
                        f"文件 '{os.path.basename(self.selected_files[i])}' 的列名与第一个文件不一致\n\n"
                        f"第一个文件列名: {first_columns}\n"
                        f"当前文件列名: {list(df.columns)}"
                    )
                    return

            # 合并数据
            self.update_status("正在合并数据...")
            merged_df = pd.concat(dataframes, ignore_index=True)

            # 保存文件
            save_path = filedialog.asksaveasfilename(
                title="保存合并后的文件",
                defaultextension=".xlsx",
                filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
            )

            if save_path:
                self.update_status("正在保存文件...")
                merged_df.to_excel(save_path, index=False)

                self.progress.stop()
                self.update_status(f"合并完成！总共 {len(merged_df)} 行数据")

                messagebox.showinfo(
                    "成功",
                    f"文件合并成功！\n\n"
                    f"合并了 {len(self.selected_files)} 个文件\n"
                    f"总共 {len(merged_df)} 行数据\n"
                    f"保存位置: {save_path}"
                )
            else:
                self.progress.stop()
                self.update_status("已取消保存")

        except Exception as e:
            self.progress.stop()
            messagebox.showerror("错误", f"合并文件时发生错误：\n{str(e)}")
            self.update_status("合并失败")

        finally:
            # 恢复按钮状态
            self.merge_btn.config(state="normal" if len(self.selected_files) >= 2 else "disabled")
            self.select_btn.config(state="normal")
            self.clear_btn.config(state="normal")
            self.progress.stop()

def main():
    root = tk.Tk()
    app = ExcelMergerApp(root)

    # 设置窗口居中
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f"{width}x{height}+{x}+{y}")

    root.mainloop()

if __name__ == "__main__":
    main()