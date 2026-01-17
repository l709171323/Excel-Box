"""
Tab7 - 前缀填充承运商功能
"""
import tkinter as tk
from tkinter import ttk, messagebox
import threading

from excel_toolkit.prefix_fill import process_prefix_fill


class Tab7PrefixMixin:
    """Tab7 前缀填充 Mixin"""
    
    def create_tab7_prefix(self, tab):
        """创建Tab7界面"""
        # 检查变量是否已经在_initialize_all_variables中创建
        if not hasattr(self, 'file7_var'):
            self.file7_var = tk.StringVar(value="未选择文件")
            self.src7_var = tk.StringVar(value="A")
            self.dst7_var = tk.StringVar(value="B")
            self._trace_persist(self.file7_var)
            self._trace_persist(self.src7_var)
            self._trace_persist(self.dst7_var)

        f1 = ttk.Frame(tab)
        f1.pack(fill='x', pady=5)
        ttk.Button(f1, text="选择文件", 
                  command=lambda: self.select_file_and_sheets(
                      self.file7_var, None, None, 
                      "选择要处理的文件")).pack(side='left', padx=5)
        ttk.Label(f1, textvariable=self.file7_var).pack(side='left', padx=5)

        f2 = ttk.Frame(tab)
        f2.pack(fill='x', pady=5)
        ttk.Label(f2, text="源列号（含单号）:").pack(side='left', padx=5)
        ttk.Entry(f2, textvariable=self.src7_var, width=8).pack(side='left', padx=5)
        ttk.Label(f2, text="目标列号（填充承运商）:").pack(side='left', padx=(20, 5))
        ttk.Entry(f2, textvariable=self.dst7_var, width=8).pack(side='left', padx=5)

        # 规则说明
        f_info = ttk.LabelFrame(tab, text="填充规则", style="Section.TLabelframe")
        f_info.pack(fill='x', pady=5, padx=5)
        ttk.Label(f_info, text="• 首字符 '9' → 填充 'usps'").pack(anchor='w', padx=10)
        ttk.Label(f_info, text="• 首字符 'G' → 填充 'GOFO'").pack(anchor='w', padx=10)
        ttk.Label(f_info, text="• 首字符 'U' → 填充 'UniUni'").pack(anchor='w', padx=10)

        f3 = ttk.Frame(tab)
        f3.pack(fill='x', pady=10)
        ttk.Button(f3, text="[7] 开始前缀填充", command=self.run_tool7, 
                  style='Accent.TButton').pack(side='left', padx=5)
        self.logger7, clear_log7 = self.create_log_widget(tab)
        ttk.Button(f3, text="清空日志", command=clear_log7, 
                  style='Secondary.TButton').pack(side='left', padx=5)

    def run_tool7(self):
        """执行前缀填充"""
        file = self.file7_var.get()
        src = self.src7_var.get()
        dst = self.dst7_var.get()
        
        if not file or file == "未选择文件":
            messagebox.showwarning("⚠️ 警告", "请先选择文件。")
            return
        if not src:
            messagebox.showwarning("⚠️ 警告", "请输入源列号。")
            return
        if not dst:
            messagebox.showwarning("⚠️ 警告", "请输入目标列号。")
            return

        self.logger7("=" * 50)
        self.logger7(f"▶️ 开始运行 [7] 前缀填充承运商...")
        
        self._update_status("正在处理...", icon="⏳", show_progress=True)
        self.master.config(cursor="watch")
        
        def thread_target():
            try:
                def safe_logger(msg):
                    self.master.after(0, lambda m=msg: self.logger7(m))
                
                result = process_prefix_fill(file, src, dst, safe_logger)
                
                def on_success():
                    self.master.config(cursor="")
                    self._update_status("就绪", icon="✅", show_progress=False)
                    messagebox.showinfo("✅ 完成", result)
                    self.logger7(result)
                
                self.master.after(0, on_success)
                
            except Exception as e:
                error_msg = str(e)
                def on_error(msg=error_msg):
                    self.master.config(cursor="")
                    self._update_status("错误", icon="❌", show_progress=False)
                    messagebox.showerror("❌ 错误", msg)
                    self.logger7(f"❌ 发生错误: {msg}")
                self.master.after(0, on_error)

        threading.Thread(target=thread_target, daemon=True).start()





























