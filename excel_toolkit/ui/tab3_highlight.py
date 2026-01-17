"""
Tab3 - 高亮重复项功能(优化布局版)
"""
import tkinter as tk
from tkinter import ttk, messagebox
from tkinter.scrolledtext import ScrolledText
import threading
import os

from excel_toolkit.highlight import highlight_duplicates
from excel_toolkit.tooltip import create_tooltip


class Tab3HighlightMixin:
    """Tab3 高亮重复项 Mixin"""
    
    def create_tab3_highlight(self, tab):
        """创建Tab3界面(优化版)"""
        # 检查变量是否已经在_initialize_all_variables中创建
        if not hasattr(self, 'file3_var'):
            self.file3_var = tk.StringVar(value="未选择文件")
            self.col3_var = tk.StringVar(value="A")
            self.sheet3_var = tk.StringVar()
            self._trace_persist(self.file3_var)
            self._trace_persist(self.col3_var)
            self._trace_persist(self.sheet3_var)

        # ===== 卡片1: 文件选择 =====
        file_card = ttk.LabelFrame(tab, text="📊 文件选择", padding=12)
        file_card.pack(fill='x', padx=15, pady=(15, 8))
        
        file_row = ttk.Frame(file_card)
        file_row.pack(fill='x', pady=4)
        
        select_btn = ttk.Button(
            file_row, 
            text="📂 选择文件",
            width=12,
            command=self._select_file3
        )
        select_btn.pack(side='left', padx=(0, 8))
        create_tooltip(select_btn, "选择包含重复数据的Excel文件")
        
        ttk.Label(file_row, textvariable=self.file3_var, foreground='#6B7280').pack(side='left', fill='x', expand=True)
        
        # 工作表选择
        sheet_row = ttk.Frame(file_card)
        sheet_row.pack(fill='x', pady=(8, 4))
        
        ttk.Label(sheet_row, text="工作表:", width=8).pack(side='left', padx=(0, 8))
        self.sheet3_combo = ttk.Combobox(sheet_row, textvariable=self.sheet3_var,
                                         state="readonly", width=25)
        self.sheet3_combo.pack(side='left', padx=(0, 8))
        create_tooltip(self.sheet3_combo, "选择要处理的工作表，不选择则处理所有工作表")
        
        ttk.Label(sheet_row, text="（不选择=处理所有工作表）", 
                 foreground='#6B7280', font=("Microsoft YaHei UI", 9)).pack(side='left')

        # ===== 卡片2: 参数配置 =====
        param_card = ttk.LabelFrame(tab, text="⚙️ 参数配置", padding=12)
        param_card.pack(fill='x', padx=15, pady=8)
        
        param_row = ttk.Frame(param_card)
        param_row.pack(fill='x', pady=4)
        
        ttk.Label(param_row, text="目标列:", width=8).pack(side='left', padx=(0, 8))
        col_entry = ttk.Entry(param_row, textvariable=self.col3_var, width=6)
        col_entry.pack(side='left')
        create_tooltip(col_entry, "输入要检查重复的列号(如A、B、C等)")
        
        # 提示信息
        hint_row = ttk.Frame(param_card)
        hint_row.pack(fill='x', pady=(8, 4))
        ttk.Label(
            hint_row,
            text="ℹ️ 提示: 程序会自动检测所有工作表中指定列的重复值,并用不同颜色高亮标记",
            foreground='#6B7280',
            font=("Microsoft YaHei UI", 9)
        ).pack(side='left')

        # ===== 操作按钮 =====
        action_frame = ttk.Frame(tab)
        action_frame.pack(fill='x', padx=15, pady=15)
        
        run_btn = ttk.Button(
            action_frame,
            text="▶️ 开始高亮",
            command=self.run_tool3,
            style='Accent.TButton',
            width=16
        )
        run_btn.pack(side='left', padx=(0, 8))
        create_tooltip(run_btn, "开始检测并高亮重复项")
        
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
        
        self.logger3 = logger
        clear_log3 = clear_log
        
        ttk.Button(
            action_frame,
            text="🧹 清空日志",
            command=clear_log3,
            style='Secondary.TButton',
            width=12
        ).pack(side='left')
        create_tooltip(action_frame.winfo_children()[-1], "清空下方的日志记录")

    def _select_file3(self):
        """选择文件"""
        from excel_toolkit.ui.mixins import get_sheet_names
        from tkinter import filedialog
        path = filedialog.askopenfilename(
            title="选择要高亮重复项的Excel文件",
            filetypes=[("表格文件", "*.xlsx;*.xlsm;*.xls"), ("所有文件", "*.*")]
        )
        if path:
            self.file3_var.set(path)
            self.logger3(f"已选择文件: {path}")
            names = get_sheet_names(path)
            if names:
                self._update_combobox_options(self.sheet3_combo, self.sheet3_var, names)
                self.logger3(f"  工作表: {', '.join(names)}")
    
    def run_tool3(self):
        """执行高亮重复项"""
        file = self.file3_var.get()
        col = self.col3_var.get().strip()
        sheet = self.sheet3_var.get().strip() if hasattr(self, 'sheet3_var') else None
        
        if not file or file == "未选择文件":
            messagebox.showwarning("⚠️ 警告", "请先选择要处理的Excel文件。")
            return
        if not col:
            messagebox.showwarning("⚠️ 警告", "请输入要检查的列号。")
            return

        self.logger3("=" * 60)
        self.logger3(f"▶️ 开始执行高亮重复项...")
        self.logger3(f"  文件: {os.path.basename(file)}")
        if sheet:
            self.logger3(f"  工作表: {sheet}")
        else:
            self.logger3(f"  工作表: 全部")
        self.logger3(f"  目标列: {col}")
        self.logger3("=" * 60)
        
        self._update_status("正在高亮重复项...", icon="⏳", show_progress=True)
        self.master.config(cursor="watch")
        
        def thread_target():
            try:
                def safe_logger(msg):
                    self.master.after(0, lambda m=msg: self.logger3(m))
                
                stats = highlight_duplicates(file, col, safe_logger, sheet)
                
                def on_success():
                    self.master.config(cursor="")
                    self._update_status("就绪", icon="✅", show_progress=False)
                    
                    msg = (
                        f"✅ 高亮完成！\n\n"
                        f"处理工作表数: {stats['sheets_processed']}\n"
                        f"高亮单元格数: {stats['cells_highlighted']}\n\n"
                        f"文件已保存: {os.path.basename(file)}"
                    )
                    messagebox.showinfo("✅ 完成", msg)
                    self.logger3("\n" + "=" * 60)
                    self.logger3(f"✅ 高亮完成")
                    self.logger3(f"  处理工作表: {stats['sheets_processed']} 个")
                    self.logger3(f"  高亮单元格: {stats['cells_highlighted']} 个")
                    self.logger3("=" * 60)
                
                self.master.after(0, on_success)
                
            except Exception as e:
                error_msg = str(e)
                def on_error(msg=error_msg):
                    self.master.config(cursor="")
                    self._update_status("错误", icon="❌", show_progress=False)
                    messagebox.showerror("❌ 错误", msg)
                    self.logger3(f"❌ 发生错误: {msg}")
                self.master.after(0, on_error)

        threading.Thread(target=thread_target, daemon=True).start()





























