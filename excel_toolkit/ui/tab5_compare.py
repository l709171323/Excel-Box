"""
Tab5 - 对比列数据功能
"""
import tkinter as tk
from tkinter import ttk, messagebox
import threading

from excel_toolkit.compare import process_compare_columns


class Tab5CompareMixin:
    """Tab5 对比列数据 Mixin"""
    
    def create_tab5_compare(self, tab):
        """创建Tab5界面"""
        # 检查变量是否已经在_initialize_all_variables中创建
        if not hasattr(self, 'file5_x_var'):
            self.file5_x_var = tk.StringVar(value="未选择文件X")
            self.file5_y_var = tk.StringVar(value="未选择文件Y")
            self.sheet5_y_var = tk.StringVar()
            self.col5_x_var = tk.StringVar(value="A")
            self.col5_y_var = tk.StringVar(value="A")
            self.ignore_dups_var = tk.BooleanVar(value=True)
            
            self._trace_persist(self.file5_x_var)
            self._trace_persist(self.file5_y_var)
            self._trace_persist(self.sheet5_y_var)
            self._trace_persist(self.col5_x_var)
            self._trace_persist(self.col5_y_var)
            self._trace_persist(self.ignore_dups_var)

        # 表格X（可多选子表）
        f1 = ttk.Frame(tab)
        f1.pack(fill='x', pady=5)
        ttk.Button(f1, text="选择表格X", 
                  command=lambda: self.select_file_and_listbox(
                      self.file5_x_var, self.listbox5_x, 
                      "选择要对比的表格X")).pack(side='left', padx=5)
        ttk.Label(f1, textvariable=self.file5_x_var).pack(side='left', padx=5)

        f2 = ttk.Frame(tab)
        f2.pack(fill='x', pady=5)
        ttk.Label(f2, text="X子表(可多选):").pack(side='left', padx=5)
        
        listbox_frame = ttk.Frame(f2)
        listbox_frame.pack(side='left', padx=5)
        self.listbox5_x = tk.Listbox(listbox_frame, selectmode='extended', 
                                     width=25, height=4, exportselection=False)
        self.listbox5_x.pack(side='left')
        scrollbar = ttk.Scrollbar(listbox_frame, orient='vertical', 
                                  command=self.listbox5_x.yview)
        scrollbar.pack(side='left', fill='y')
        self.listbox5_x.config(yscrollcommand=scrollbar.set)
        
        ttk.Label(f2, text="X列号:").pack(side='left', padx=(20, 5))
        ttk.Entry(f2, textvariable=self.col5_x_var, width=6).pack(side='left', padx=5)

        # 表格Y
        f3 = ttk.Frame(tab)
        f3.pack(fill='x', pady=5)
        ttk.Button(f3, text="选择表格Y", 
                  command=lambda: self.select_file_and_sheets(
                      self.file5_y_var, self.sheet5_y_var, self.combo5_y, 
                      "选择要对比的表格Y")).pack(side='left', padx=5)
        ttk.Label(f3, textvariable=self.file5_y_var).pack(side='left', padx=5)

        f4 = ttk.Frame(tab)
        f4.pack(fill='x', pady=5)
        ttk.Label(f4, text="Y子表:").pack(side='left', padx=5)
        self.combo5_y = ttk.Combobox(f4, textvariable=self.sheet5_y_var, 
                                     state="readonly", width=22)
        self.combo5_y.pack(side='left', padx=5)
        ttk.Label(f4, text="Y列号:").pack(side='left', padx=(20, 5))
        ttk.Entry(f4, textvariable=self.col5_y_var, width=6).pack(side='left', padx=5)

        # 选项
        f5 = ttk.Frame(tab)
        f5.pack(fill='x', pady=5)
        ttk.Checkbutton(f5, text="忽略重复值（集合比较）", 
                       variable=self.ignore_dups_var).pack(side='left', padx=5)

        # 执行按钮
        f6 = ttk.Frame(tab)
        f6.pack(fill='x', pady=10)
        ttk.Button(f6, text="[5] 开始对比", command=self.run_tool5, 
                  style='Accent.TButton').pack(side='left', padx=5)
        self.logger5, clear_log5 = self.create_log_widget(tab)
        ttk.Button(f6, text="清空日志", command=clear_log5, 
                  style='Secondary.TButton').pack(side='left', padx=5)

    def run_tool5(self):
        """执行对比列数据"""
        file_x = self.file5_x_var.get()
        file_y = self.file5_y_var.get()
        sheet_y = self.sheet5_y_var.get()
        col_x = self.col5_x_var.get()
        col_y = self.col5_y_var.get()
        ignore_dups = self.ignore_dups_var.get()
        
        # 获取选中的X子表
        selected_indices = self.listbox5_x.curselection()
        sheets_x = [self.listbox5_x.get(i) for i in selected_indices]
        
        if not file_x or file_x == "未选择文件X":
            messagebox.showwarning("⚠️ 警告", "请先选择表格X。")
            return
        if not file_y or file_y == "未选择文件Y":
            messagebox.showwarning("⚠️ 警告", "请先选择表格Y。")
            return
        if not sheets_x:
            messagebox.showwarning("⚠️ 警告", "请至少选择一个X子表。")
            return
        if not sheet_y:
            messagebox.showwarning("⚠️ 警告", "请选择Y子表。")
            return

        self.logger5("=" * 50)
        self.logger5(f"▶️ 开始运行 [5] 对比列数据...")
        self.logger5(f"X子表: {sheets_x}")
        self.logger5(f"Y子表: {sheet_y}")
        self.logger5(f"忽略重复: {ignore_dups}")
        
        self._update_status("正在对比...", icon="⏳", show_progress=True)
        self.master.config(cursor="watch")
        
        def thread_target():
            try:
                def safe_logger(msg):
                    self.master.after(0, lambda m=msg: self.logger5(m))
                
                result = process_compare_columns(
                    file_x, sheets_x, col_x,
                    file_y, sheet_y, col_y,
                    safe_logger, ignore_dups
                )
                
                def on_success():
                    self.master.config(cursor="")
                    self._update_status("就绪", icon="✅", show_progress=False)
                    messagebox.showinfo("✅ 完成", result)
                    self.logger5(result)
                
                self.master.after(0, on_success)
                
            except Exception as e:
                error_msg = str(e)
                def on_error(msg=error_msg):
                    self.master.config(cursor="")
                    self._update_status("错误", icon="❌", show_progress=False)
                    messagebox.showerror("❌ 错误", msg)
                    self.logger5(f"❌ 发生错误: {msg}")
                self.master.after(0, on_error)

        threading.Thread(target=thread_target, daemon=True).start()





























