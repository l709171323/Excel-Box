"""
Tab1 - 州名转换功能
"""
import tkinter as tk
from tkinter import ttk, messagebox
from tkinter.scrolledtext import ScrolledText
import threading
import os

from excel_toolkit.states import process_states
from excel_toolkit.exceptions import ExcelToolkitError
from excel_toolkit.error_handler import get_user_friendly_error, log_error
from excel_toolkit.tooltip import create_tooltip


class Tab1StatesMixin:
    """Tab1 州名转换 Mixin"""
    
    def create_tab1_states(self, tab):
        """创建Tab1界面(优化版)"""
        # 变量已在 _initialize_all_variables 中创建，这里不再重复创建
        if not hasattr(self, 'file1_var'):
            self.file1_var = tk.StringVar(value="未选择文件")
            self.sheet1_var = tk.StringVar()
            self.col1_var = tk.StringVar(value="G")
            self._trace_persist(self.file1_var)
            self._trace_persist(self.sheet1_var)
            self._trace_persist(self.col1_var)

        # ===== 卡片1: 文件选择 =====
        file_card = ttk.LabelFrame(tab, text="📊 文件选择", padding=12)
        file_card.pack(fill='x', padx=15, pady=(15, 8))
        
        file_row = ttk.Frame(file_card)
        file_row.pack(fill='x', pady=4)
        
        select_btn = ttk.Button(
            file_row, 
            text="📂 选择文件",
            width=12,
            command=lambda: self.select_file_and_sheets(
                self.file1_var, self.sheet1_var, self.combo1, 
                "选择要转换州名的Excel文件")
        )
        select_btn.pack(side='left', padx=(0, 8))
        create_tooltip(select_btn, "选择包含州名数据的Excel文件")
        
        ttk.Label(file_row, textvariable=self.file1_var, foreground='#6B7280').pack(side='left', fill='x', expand=True)

        # ===== 卡片2: 参数配置 =====
        param_card = ttk.LabelFrame(tab, text="⚙️ 参数配置", padding=12)
        param_card.pack(fill='x', padx=15, pady=8)
        
        param_row = ttk.Frame(param_card)
        param_row.pack(fill='x', pady=4)
        
        ttk.Label(param_row, text="工作表:", width=8).pack(side='left', padx=(0, 8))
        self.combo1 = ttk.Combobox(param_row, textvariable=self.sheet1_var, state="readonly", width=20)
        self.combo1.pack(side='left', padx=(0, 16))
        create_tooltip(self.combo1, "选择包含州名数据的工作表")
        
        ttk.Label(param_row, text="目标列:", width=8).pack(side='left', padx=(0, 8))
        col_entry = ttk.Entry(param_row, textvariable=self.col1_var, width=6)
        col_entry.pack(side='left')
        create_tooltip(col_entry, "输入列号(如A、B、G等)")
        
        # 提示信息
        hint_row = ttk.Frame(param_card)
        hint_row.pack(fill='x', pady=(8, 4))
        ttk.Label(
            hint_row,
            text="ℹ️ 提示: 程序会将选中列的州全名转换为两字母缩写(如 California → CA)",
            foreground='#6B7280',
            font=("Microsoft YaHei UI", 9)
        ).pack(side='left')

        # ===== 操作按钮 =====
        action_frame = ttk.Frame(tab)
        action_frame.pack(fill='x', padx=15, pady=15)
        
        run_btn = ttk.Button(
            action_frame,
            text="▶️ 开始转换",
            command=self.run_tool1,
            style='Accent.TButton',
            width=16
        )
        run_btn.pack(side='left', padx=(0, 8))
        create_tooltip(run_btn, "开始执行州名转换(快捷键: Ctrl+R)")
        
        # ===== 日志区域 =====
        log_card = ttk.LabelFrame(tab, text="📝 执行日志", padding=12)
        log_card.pack(fill='both', expand=True, padx=15, pady=(0, 15))
        
        log_widget = ScrolledText(
            log_card,
            height=10,
            state="disabled",
            font=("Consolas", 9),
            wrap='word'
        )
        log_widget.pack(fill='both', expand=True)
        
        try:
            log_widget.configure(
                bg="#F9FAFB",
                fg="#111827",
                insertbackground="#111827",
                relief='flat',
                borderwidth=1
            )
        except Exception:
            pass
        
        if hasattr(self, '_text_widgets'):
            self._text_widgets.append(log_widget)
        
        def logger(text):
            log_widget.config(state="normal")
            log_widget.insert("end", str(text) + "\n")
            log_widget.see("end")
            log_widget.config(state="disabled")
        
        def clear_log():
            log_widget.config(state="normal")
            log_widget.delete("1.0", "end")
            logger("✅ 日志已清空")
        
        self.logger1 = logger
        clear_log1 = clear_log
        
        ttk.Button(
            action_frame,
            text="🧹 清空日志",
            command=clear_log1,
            style='Secondary.TButton',
            width=12
        ).pack(side='left')
        create_tooltip(action_frame.winfo_children()[-1], "清空下方的日志记录")

    def run_tool1(self):
        """执行州名转换"""
        file = self.file1_var.get()
        sheet = self.sheet1_var.get()
        col = self.col1_var.get()
        
        if not file or file == "未选择文件":
            messagebox.showwarning("⚠️ 警告", "请先选择一个文件。")
            return
        if not sheet:
            messagebox.showwarning("⚠️ 警告", "请选择一个子表。")
            return
        if not col:
            messagebox.showwarning("⚠️ 警告", "请输入一个列号。")
            return

        self.logger1("=" * 60)
        self.logger1(f"▶️ 开始执行州名转换...")
        self.logger1(f"  文件: {os.path.basename(file)}")
        self.logger1(f"  工作表: {sheet}")
        self.logger1(f"  目标列: {col}")
        self.logger1("=" * 60)
        
        self._update_status("正在处理州名转换...", icon="⏳", show_progress=True)
        self.master.config(cursor="watch")
        
        def thread_target():
            try:
                def safe_logger(msg):
                    self.master.after(0, lambda m=msg: self.logger1(m))
                
                stats = process_states(file, sheet, col, safe_logger)
                
                def on_success():
                    self.master.config(cursor="")
                    self._update_status("就绪", icon="✅", show_progress=False)
                    
                    msg = (
                        f"✅ 州名转换完成！\n\n"
                        f"总共处理: {stats['total']} 行\n"
                        f"成功转换: {stats['success']} 行\n"
                        f"未找到/保持原值: {stats['failed']} 行\n\n"
                        f"文件已保存: {os.path.basename(file)}"
                    )
                    
                    messagebox.showinfo("✅ 完成", msg)
                    self.logger1("\n" + "=" * 60)
                    self.logger1(f"✅ 转换完成")
                    self.logger1(f"  总计: {stats['total']} 行")
                    self.logger1(f"  成功: {stats['success']} 行")
                    self.logger1(f"  跳过: {stats['failed']} 行")
                    self.logger1("=" * 60)

                self.master.after(0, on_success)

            except ExcelToolkitError as e:
                # 自定义异常(包含友好信息)
                user_msg = e.get_user_message()
                err_msg = e.message
                err_solution = e.solution
                def on_custom_error(umsg=user_msg, msg=err_msg, sol=err_solution):
                    self.master.config(cursor="")
                    self._update_status("错误", icon="❌", show_progress=False)
                    messagebox.showerror("❌ 错误", umsg)
                    self.logger1(f"\n❌ {msg}")
                    if sol:
                        self.logger1(f"💡 解决方案: {sol}")
                
                self.master.after(0, on_custom_error)
            
            except Exception as e:
                # 未预期的异常
                log_error(e, "州名转换")
                error_msg = get_user_friendly_error(e)
                error_str = str(e)
                def on_error(msg=error_msg, err=error_str):
                    self.master.config(cursor="")
                    self._update_status("错误", icon="❌", show_progress=False)
                    messagebox.showerror("❌ 错误", msg)
                    self.logger1(f"\n❌ 发生错误: {err}")
                    self.logger1(f"💡 请查看日志文件获取详细信息")
                
                self.master.after(0, on_error)

        threading.Thread(target=thread_target, daemon=True).start()





























