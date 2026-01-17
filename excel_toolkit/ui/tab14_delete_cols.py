"""
Tab14 - 批量删除列功能（含模板管理）
"""
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from tkinter.scrolledtext import ScrolledText
import threading
import os
import json

from excel_toolkit.delete_cols import delete_columns, parse_column_input
from excel_toolkit.tooltip import create_tooltip


class Tab14DeleteColsMixin:
    """Tab14 批量删除列 Mixin"""
    
    def _get_templates14_path(self):
        """获取模板配置文件路径"""
        config_dir = self._config_dir()
        return os.path.join(config_dir, "delete_cols_templates.json")
    
    def _load_templates14(self):
        """加载模板配置"""
        path = self._get_templates14_path()
        if os.path.exists(path):
            try:
                with open(path, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception:
                return {}
        return {}
    
    def _save_templates14(self, templates):
        """保存模板配置"""
        path = self._get_templates14_path()
        try:
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(templates, f, ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showerror("错误", f"保存模板失败: {e}")
    
    def _refresh_template14_combo(self):
        """刷新模板下拉框"""
        templates = self._load_templates14()
        names = list(templates.keys())
        self.template14_combo['values'] = ["（选择模板）"] + names
        if not self.template14_var.get() or self.template14_var.get() not in names:
            self.template14_var.set("（选择模板）")
    
    def create_tab14_delete_cols(self, tab):
        """创建Tab14界面"""
        # 检查变量是否已经在_initialize_all_variables中创建
        if not hasattr(self, 'file14_var'):
            self.file14_var = tk.StringVar(value="未选择文件")
            self.sheet14_var = tk.StringVar()
            self.cols14_var = tk.StringVar(value="")
            self._trace_persist(self.file14_var)
            self._trace_persist(self.sheet14_var)
            self._trace_persist(self.cols14_var)
        
        # 模板选择变量
        if not hasattr(self, 'template14_var'):
            self.template14_var = tk.StringVar(value="（选择模板）")

        # ===== 卡片1: 文件选择 =====
        file_card = ttk.LabelFrame(tab, text="📊 文件选择", padding=12)
        file_card.pack(fill='x', padx=15, pady=(15, 8))
        
        file_row = ttk.Frame(file_card)
        file_row.pack(fill='x', pady=4)
        
        select_btn = ttk.Button(
            file_row, 
            text="📂 选择文件",
            width=12,
            command=self._select_file14
        )
        select_btn.pack(side='left', padx=(0, 8))
        create_tooltip(select_btn, "选择要删除列的Excel文件")
        
        ttk.Label(file_row, textvariable=self.file14_var, foreground='#6B7280').pack(side='left', fill='x', expand=True)
        
        # 工作表选择
        sheet_row = ttk.Frame(file_card)
        sheet_row.pack(fill='x', pady=(8, 4))
        
        ttk.Label(sheet_row, text="工作表:", width=8).pack(side='left', padx=(0, 8))
        self.sheet14_combo = ttk.Combobox(sheet_row, textvariable=self.sheet14_var,
                                         state="readonly", width=25)
        self.sheet14_combo.pack(side='left', padx=(0, 8))
        create_tooltip(self.sheet14_combo, "选择要处理的工作表，不选择则处理所有工作表")
        
        ttk.Label(sheet_row, text="（不选择=处理所有工作表）", 
                 foreground='#6B7280', font=("Microsoft YaHei UI", 9)).pack(side='left')

        # ===== 卡片2: 列配置 =====
        col_card = ttk.LabelFrame(tab, text="🗑️ 要删除的列", padding=12)
        col_card.pack(fill='x', padx=15, pady=8)
        
        # 模板选择行
        template_row = ttk.Frame(col_card)
        template_row.pack(fill='x', pady=(0, 8))
        
        ttk.Label(template_row, text="模板:", width=8).pack(side='left', padx=(0, 8))
        self.template14_combo = ttk.Combobox(template_row, textvariable=self.template14_var,
                                              state="readonly", width=20)
        self.template14_combo.pack(side='left', padx=(0, 8))
        self.template14_combo.bind('<<ComboboxSelected>>', self._on_template14_selected)
        create_tooltip(self.template14_combo, "选择已保存的模板，快速填充列配置")
        
        # 模板操作按钮
        save_tpl_btn = ttk.Button(template_row, text="💾 保存", width=8,
                                  command=self._save_template14)
        save_tpl_btn.pack(side='left', padx=(0, 4))
        create_tooltip(save_tpl_btn, "将当前列配置保存为新模板")
        
        rename_tpl_btn = ttk.Button(template_row, text="✏️ 重命名", width=8,
                                    command=self._rename_template14)
        rename_tpl_btn.pack(side='left', padx=(0, 4))
        create_tooltip(rename_tpl_btn, "重命名当前选中的模板")
        
        delete_tpl_btn = ttk.Button(template_row, text="🗑️ 删除", width=8,
                                    command=self._delete_template14)
        delete_tpl_btn.pack(side='left')
        create_tooltip(delete_tpl_btn, "删除当前选中的模板")
        
        # 列输入行
        col_row = ttk.Frame(col_card)
        col_row.pack(fill='x', pady=4)
        
        ttk.Label(col_row, text="列标识:", width=8).pack(side='left', padx=(0, 8))
        col_entry = ttk.Entry(col_row, textvariable=self.cols14_var, width=30)
        col_entry.pack(side='left', padx=(0, 8))
        create_tooltip(col_entry, "输入要删除的列，如: D,E 或 D-F 或 A C E")
        
        # 提示信息
        hint_row = ttk.Frame(col_card)
        hint_row.pack(fill='x', pady=(8, 4))
        
        hint_text = ttk.Label(
            hint_row,
            text="ℹ️ 支持格式: \"D,E\" 或 \"D-F\" (范围) 或 \"A C E\" (空格分隔)",
            foreground='#6B7280',
            font=("Microsoft YaHei UI", 9)
        )
        hint_text.pack(side='left')
        
        # 警告提示
        warning_row = ttk.Frame(col_card)
        warning_row.pack(fill='x', pady=(4, 4))
        
        warning_text = ttk.Label(
            warning_row,
            text="⚠️ 注意: 删除操作会直接修改原文件，建议提前备份！",
            foreground='#DC2626',
            font=("Microsoft YaHei UI", 9, "bold")
        )
        warning_text.pack(side='left')

        # ===== 操作按钮 =====
        action_frame = ttk.Frame(tab)
        action_frame.pack(fill='x', padx=15, pady=15)
        
        run_btn = ttk.Button(
            action_frame,
            text="▶️ 开始删除",
            command=self.run_tool14,
            style='Accent.TButton',
            width=16
        )
        run_btn.pack(side='left', padx=(0, 8))
        create_tooltip(run_btn, "开始删除指定的列")
        
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
        
        self.logger14 = logger
        clear_log14 = clear_log
        
        ttk.Button(
            action_frame,
            text="🧹 清空日志",
            command=clear_log14,
            style='Secondary.TButton',
            width=12
        ).pack(side='left')
        create_tooltip(action_frame.winfo_children()[-1], "清空下方的日志记录")
        
        # 初始化模板列表
        self._refresh_template14_combo()

    def _on_template14_selected(self, event=None):
        """模板选择事件"""
        tpl_name = self.template14_var.get()
        if tpl_name and tpl_name != "（选择模板）":
            templates = self._load_templates14()
            if tpl_name in templates:
                self.cols14_var.set(templates[tpl_name])
                self.logger14(f"已加载模板 [{tpl_name}]: {templates[tpl_name]}")
    
    def _save_template14(self):
        """保存当前配置为模板"""
        cols = self.cols14_var.get().strip()
        if not cols:
            messagebox.showwarning("警告", "请先输入要删除的列标识")
            return
        
        # 弹出输入框获取模板名称
        name = simpledialog.askstring("保存模板", "请输入模板名称:", parent=self.master)
        if not name:
            return
        name = name.strip()
        if not name:
            messagebox.showwarning("警告", "模板名称不能为空")
            return
        
        templates = self._load_templates14()
        if name in templates:
            if not messagebox.askyesno("确认", f"模板 [{name}] 已存在，是否覆盖？"):
                return
        
        templates[name] = cols
        self._save_templates14(templates)
        self._refresh_template14_combo()
        self.template14_var.set(name)
        self.logger14(f"✅ 模板 [{name}] 已保存: {cols}")
        messagebox.showinfo("成功", f"模板 [{name}] 已保存")
    
    def _rename_template14(self):
        """重命名当前模板"""
        old_name = self.template14_var.get()
        if not old_name or old_name == "（选择模板）":
            messagebox.showwarning("警告", "请先选择要重命名的模板")
            return
        
        templates = self._load_templates14()
        if old_name not in templates:
            messagebox.showwarning("警告", f"模板 [{old_name}] 不存在")
            return
        
        new_name = simpledialog.askstring("重命名模板", 
                                          f"请输入新名称 (当前: {old_name}):",
                                          parent=self.master,
                                          initialvalue=old_name)
        if not new_name:
            return
        new_name = new_name.strip()
        if not new_name:
            messagebox.showwarning("警告", "模板名称不能为空")
            return
        if new_name == old_name:
            return
        
        if new_name in templates:
            messagebox.showwarning("警告", f"模板 [{new_name}] 已存在")
            return
        
        # 重命名
        templates[new_name] = templates.pop(old_name)
        self._save_templates14(templates)
        self._refresh_template14_combo()
        self.template14_var.set(new_name)
        self.logger14(f"✅ 模板已重命名: [{old_name}] → [{new_name}]")
        messagebox.showinfo("成功", f"模板已重命名为 [{new_name}]")
    
    def _delete_template14(self):
        """删除当前模板"""
        tpl_name = self.template14_var.get()
        if not tpl_name or tpl_name == "（选择模板）":
            messagebox.showwarning("警告", "请先选择要删除的模板")
            return
        
        if not messagebox.askyesno("确认删除", f"确定要删除模板 [{tpl_name}] 吗？"):
            return
        
        templates = self._load_templates14()
        if tpl_name in templates:
            del templates[tpl_name]
            self._save_templates14(templates)
            self._refresh_template14_combo()
            self.logger14(f"✅ 模板 [{tpl_name}] 已删除")
            messagebox.showinfo("成功", f"模板 [{tpl_name}] 已删除")

    def _select_file14(self):
        """选择文件"""
        from excel_toolkit.ui.mixins import get_sheet_names
        from tkinter import filedialog
        path = filedialog.askopenfilename(
            title="选择要删除列的Excel文件",
            filetypes=[("表格文件", "*.xlsx;*.xlsm;*.xls"), ("所有文件", "*.*")]
        )
        if path:
            self.file14_var.set(path)
            self.logger14(f"已选择文件: {path}")
            names = get_sheet_names(path)
            if names:
                self._update_combobox_options(self.sheet14_combo, self.sheet14_var, names)
                # 清空选择，让用户可以选择处理所有工作表
                self.sheet14_var.set("")
                self.logger14(f"  工作表: {', '.join(names)}")
    
    def run_tool14(self):
        """执行批量删除列"""
        file = self.file14_var.get()
        cols_str = self.cols14_var.get().strip()
        sheet = self.sheet14_var.get().strip() if hasattr(self, 'sheet14_var') else None
        
        if not file or file == "未选择文件":
            messagebox.showwarning("⚠️ 警告", "请先选择要处理的Excel文件。")
            return
        if not cols_str:
            messagebox.showwarning("⚠️ 警告", "请输入要删除的列标识。")
            return
        
        # 解析列标识
        columns = parse_column_input(cols_str)
        if not columns:
            messagebox.showwarning("⚠️ 警告", "无法解析列标识，请检查输入格式。")
            return
        
        # 确认删除操作
        cols_display = ", ".join(sorted(columns, key=lambda x: ord(x[0]) if len(x) == 1 else ord(x[0])*26 + ord(x[1])))
        confirm_msg = f"确定要删除以下列吗？\n\n列: {cols_display}\n\n⚠️ 此操作会直接修改原文件！"
        if not messagebox.askyesno("确认删除", confirm_msg):
            self.logger14("❌ 用户取消操作")
            return

        self.logger14("=" * 60)
        self.logger14(f"▶️ 开始执行批量删除列...")
        self.logger14(f"  文件: {os.path.basename(file)}")
        if sheet:
            self.logger14(f"  工作表: {sheet}")
        else:
            self.logger14(f"  工作表: 全部")
        self.logger14(f"  删除列: {cols_display}")
        self.logger14("=" * 60)
        
        self._update_status("正在删除列...", icon="⏳", show_progress=True)
        self.master.config(cursor="watch")
        
        def thread_target():
            try:
                def safe_logger(msg):
                    self.master.after(0, lambda m=msg: self.logger14(m))
                
                stats = delete_columns(file, columns, safe_logger, sheet if sheet else None)
                
                def on_success():
                    self.master.config(cursor="")
                    self._update_status("就绪", icon="✅", show_progress=False)
                    
                    msg = (
                        f"✅ 删除完成！\n\n"
                        f"处理工作表数: {stats['sheets_processed']}\n"
                        f"删除列数: {stats['columns_deleted']}\n\n"
                        f"文件已保存: {os.path.basename(file)}"
                    )
                    messagebox.showinfo("✅ 完成", msg)
                    self.logger14("\n" + "=" * 60)
                    self.logger14(f"✅ 删除完成")
                    self.logger14(f"  处理工作表: {stats['sheets_processed']} 个")
                    self.logger14(f"  删除列: {stats['columns_deleted']} 个")
                    self.logger14("=" * 60)
                
                self.master.after(0, on_success)
                
            except Exception as e:
                error_msg = str(e)
                def on_error(msg=error_msg):
                    self.master.config(cursor="")
                    self._update_status("错误", icon="❌", show_progress=False)
                    messagebox.showerror("❌ 错误", msg)
                    self.logger14(f"❌ 发生错误: {msg}")
                self.master.after(0, on_error)

        threading.Thread(target=thread_target, daemon=True).start()
