"""
Tab4 - 插入缺失行功能
"""
import tkinter as tk
from tkinter import ttk, messagebox
import threading
import os

from excel_toolkit.insert_rows import process_insert_rows


class Tab4InsertMixin:
    """Tab4 插入缺失行 Mixin"""
    
    def create_tab4_insert(self, tab):
        """创建Tab4界面"""
        # 变量已在 _initialize_all_variables 中创建，这里不再重复创建
        if not hasattr(self, 'file_x_var'):
            self.file_x_var = tk.StringVar(value="未选择文件X")
            self.file_y_var = tk.StringVar(value="未选择文件Y")
            self.sheet_x_var = tk.StringVar()
            self.sheet_y_var = tk.StringVar()
            self._trace_persist(self.file_x_var)
            self._trace_persist(self.file_y_var)
            self._trace_persist(self.sheet_x_var)
            self._trace_persist(self.sheet_y_var)

        # 表格X
        f1 = ttk.Frame(tab)
        f1.pack(fill='x', pady=5)
        ttk.Button(f1, text="选择表格X (源)", 
                  command=lambda: self.select_file_and_sheets(
                      self.file_x_var, self.sheet_x_var, self.combo_x, 
                      "选择源表格X")).pack(side='left', padx=5)
        ttk.Label(f1, textvariable=self.file_x_var).pack(side='left', padx=5)

        f2 = ttk.Frame(tab)
        f2.pack(fill='x', pady=5)
        ttk.Label(f2, text="X子表:").pack(side='left', padx=5)
        self.combo_x = ttk.Combobox(f2, textvariable=self.sheet_x_var, 
                                    state="readonly", width=22)
        self.combo_x.pack(side='left', padx=5)

        # 表格Y
        f3 = ttk.Frame(tab)
        f3.pack(fill='x', pady=5)
        ttk.Button(f3, text="选择表格Y (目标)", 
                  command=lambda: self.select_file_and_sheets(
                      self.file_y_var, self.sheet_y_var, self.combo_y, 
                      "选择目标表格Y")).pack(side='left', padx=5)
        ttk.Label(f3, textvariable=self.file_y_var).pack(side='left', padx=5)

        f4 = ttk.Frame(tab)
        f4.pack(fill='x', pady=5)
        ttk.Label(f4, text="Y子表:").pack(side='left', padx=5)
        self.combo_y = ttk.Combobox(f4, textvariable=self.sheet_y_var, 
                                    state="readonly", width=22)
        self.combo_y.pack(side='left', padx=5)

        # 执行按钮
        f5 = ttk.Frame(tab)
        f5.pack(fill='x', pady=10)
        ttk.Button(f5, text="[4] 开始插入缺失行", command=self.run_tool4, 
                  style='Accent.TButton').pack(side='left', padx=5)
        self.logger4, clear_log4 = self.create_log_widget(tab)
        ttk.Button(f5, text="清空日志", command=clear_log4, 
                  style='Secondary.TButton').pack(side='left', padx=5)

    def run_tool4(self):
        """执行插入缺失行"""
        file_x = self.file_x_var.get()
        file_y = self.file_y_var.get()
        sheet_x = self.sheet_x_var.get()
        sheet_y = self.sheet_y_var.get()
        
        if not file_x or file_x == "未选择文件X":
            messagebox.showwarning("⚠️ 警告", "请先选择表格X。")
            return
        if not file_y or file_y == "未选择文件Y":
            messagebox.showwarning("⚠️ 警告", "请先选择表格Y。")
            return
        if not sheet_x:
            messagebox.showwarning("⚠️ 警告", "请选择表格X的子表。")
            return
        if not sheet_y:
            messagebox.showwarning("⚠️ 警告", "请选择表格Y的子表。")
            return

        self.logger4("=" * 50)
        self.logger4(f"▶️ 开始运行 [4] 插入缺失行...")
        
        self._update_status("正在处理...", icon="⏳", show_progress=True)
        self.master.config(cursor="watch")
        
        def thread_target():
            try:
                def safe_logger(msg):
                    self.master.after(0, lambda m=msg: self.logger4(m))
                
                stats = process_insert_rows(file_x, sheet_x, file_y, sheet_y, safe_logger)
                
                def on_success():
                    self.master.config(cursor="")
                    self._update_status("就绪", icon="✅", show_progress=False)
                    
                    msg = (
                        f"插入完成！\n\n"
                        f"缺失行数: {stats['missing_count']}\n"
                        f"已插入行数: {stats['inserted_rows']}\n"
                        f"文件已保存: {os.path.basename(file_y)}"
                    )
                    messagebox.showinfo("✅ 完成", msg)
                    self.logger4(msg.replace("\n\n", "\n"))
                
                self.master.after(0, on_success)
                
            except Exception as e:
                error_msg = str(e)
                def on_error(msg=error_msg):
                    self.master.config(cursor="")
                    self._update_status("错误", icon="❌", show_progress=False)
                    messagebox.showerror("❌ 错误", msg)
                    self.logger4(f"❌ 发生错误: {msg}")
                self.master.after(0, on_error)

        threading.Thread(target=thread_target, daemon=True).start()





























