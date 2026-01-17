"""
Tab11 - 发货模板填充功能
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os

from excel_toolkit.shipping_fill import (
    process_shipping_fill, 
    get_warehouses_from_config, 
    get_mapping_choices_from_config,
    check_template_has_data
)


class Tab11ShippingMixin:
    """Tab11 发货模板填充 Mixin"""
    
    def create_tab11_shipping(self, tab):
        """创建Tab11界面 - 发货模板填充"""
        # 变量已在 _initialize_all_variables 中创建，这里不再重复创建
        if not hasattr(self, 'config11_var'):
            self.config11_var = tk.StringVar(value="未选择配置文件")
            self.order11_file_var = tk.StringVar(value="未选择订单文件")
            self.order11_sheet_var = tk.StringVar()
            self.template11_file_var = tk.StringVar(value="未选择模板文件")
            self.template11_sheet_var = tk.StringVar()
            self.mapping11_choice_var = tk.StringVar(value="映射1")
            
            # 持久化追踪
            self._trace_persist(self.config11_var)
            self._trace_persist(self.order11_file_var)
            self._trace_persist(self.order11_sheet_var)
            self._trace_persist(self.template11_file_var)
            self._trace_persist(self.template11_sheet_var)
            self._trace_persist(self.mapping11_choice_var)
        
        # 初始化仓库复选框字典（每次创建UI时重新初始化）
        self.warehouses11_checks = {}
        
        # ===== 配置文件 =====
        f_config = ttk.LabelFrame(tab, text="配置文件（列映射 + 物流渠道）", style="Section.TLabelframe")
        f_config.pack(fill='x', pady=6, padx=5)
        
        f_config_inner = ttk.Frame(f_config)
        f_config_inner.pack(fill='x', padx=6, pady=6)
        
        ttk.Button(f_config_inner, text="选择配置文件", 
                  command=self._select_config11).pack(side='left', padx=5)
        ttk.Label(f_config_inner, textvariable=self.config11_var, 
                 wraplength=500).pack(side='left', padx=5, fill='x', expand=True)
        
        f_config_choice = ttk.Frame(f_config)
        f_config_choice.pack(fill='x', padx=6, pady=(0, 6))
        
        ttk.Label(f_config_choice, text="选择映射关系:").pack(side='left', padx=5)
        self.combo11_mapping = ttk.Combobox(f_config_choice, textvariable=self.mapping11_choice_var,
                                            state="readonly", width=15, values=["映射1", "映射2", "映射3"])
        self.combo11_mapping.pack(side='left', padx=5)
        ttk.Label(f_config_choice, text="（子表1=映射1，子表2=映射2，子表3=映射3）", 
                 font=("Segoe UI", 9)).pack(side='left', padx=5)
        
        # ===== 订单文件 =====
        f_order = ttk.LabelFrame(tab, text="订单信息文件", style="Section.TLabelframe")
        f_order.pack(fill='x', pady=6, padx=5)
        
        f_order1 = ttk.Frame(f_order)
        f_order1.pack(fill='x', padx=6, pady=6)
        
        ttk.Button(f_order1, text="选择订单文件", 
                  command=self._select_order11).pack(side='left', padx=5)
        ttk.Label(f_order1, textvariable=self.order11_file_var,
                 wraplength=500).pack(side='left', padx=5, fill='x', expand=True)
        
        f_order2 = ttk.Frame(f_order)
        f_order2.pack(fill='x', padx=6, pady=(0, 6))
        
        ttk.Label(f_order2, text="工作表:").pack(side='left', padx=5)
        self.combo11_order = ttk.Combobox(f_order2, textvariable=self.order11_sheet_var,
                                          state="readonly", width=25)
        self.combo11_order.pack(side='left', padx=5)
        
        # ===== 模板文件 =====
        f_template = ttk.LabelFrame(tab, text="发货模板文件（将直接填充此文件）", style="Section.TLabelframe")
        f_template.pack(fill='x', pady=6, padx=5)
        
        f_tpl1 = ttk.Frame(f_template)
        f_tpl1.pack(fill='x', padx=6, pady=6)
        
        ttk.Button(f_tpl1, text="选择模板文件", 
                  command=self._select_template11).pack(side='left', padx=5)
        ttk.Label(f_tpl1, textvariable=self.template11_file_var,
                 wraplength=500).pack(side='left', padx=5, fill='x', expand=True)
        
        f_tpl2 = ttk.Frame(f_template)
        f_tpl2.pack(fill='x', padx=6, pady=(0, 6))
        
        ttk.Label(f_tpl2, text="工作表:").pack(side='left', padx=5)
        self.combo11_template = ttk.Combobox(f_tpl2, textvariable=self.template11_sheet_var,
                                             state="readonly", width=25)
        self.combo11_template.pack(side='left', padx=5)
        
        # ===== 仓库筛选（多选） =====
        f_wh_filter = ttk.LabelFrame(tab, text="筛选仓库（勾选要填充的仓库，不勾选=全部）", 
                                     style="Section.TLabelframe")
        f_wh_filter.pack(fill='x', pady=6, padx=5)
        
        wh_wrap = ttk.Frame(f_wh_filter)
        wh_wrap.pack(fill='both', expand=True, padx=6, pady=6)
        
        self.wh11_canvas = tk.Canvas(wh_wrap, height=80)
        self.wh11_canvas.pack(side='left', fill='both', expand=True)
        
        sb_wh = ttk.Scrollbar(wh_wrap, orient='vertical', command=self.wh11_canvas.yview)
        sb_wh.pack(side='right', fill='y')
        self.wh11_canvas.configure(yscrollcommand=sb_wh.set)
        
        self.wh11_inner = ttk.Frame(self.wh11_canvas)
        self.wh11_canvas_window = self.wh11_canvas.create_window((0, 0), window=self.wh11_inner, anchor='nw')
        self.wh11_inner.bind('<Configure>', 
                            lambda e: self.wh11_canvas.configure(scrollregion=self.wh11_canvas.bbox('all')))
        
        ctrl_wh = ttk.Frame(f_wh_filter)
        ctrl_wh.pack(fill='x', padx=6, pady=(0, 6))
        
        ttk.Button(ctrl_wh, text="刷新仓库列表", 
                  command=self._refresh_warehouses11).pack(side='left', padx=4)
        ttk.Button(ctrl_wh, text="全选", 
                  command=lambda: self._select_all_warehouses11(True)).pack(side='left', padx=4)
        ttk.Button(ctrl_wh, text="清空", 
                  command=lambda: self._select_all_warehouses11(False)).pack(side='left', padx=4)
        ttk.Label(ctrl_wh, text="提示: 不勾选任何仓库 = 填充全部订单", 
                 font=("Segoe UI", 9)).pack(side='left', padx=10)
        
        # ===== 执行按钮和日志 =====
        f_run = ttk.Frame(tab)
        f_run.pack(fill='x', pady=10, padx=5)
        
        ttk.Button(f_run, text="[11] 开始填充发货模板", 
                  command=self.run_tool11, 
                  style='Accent.TButton').pack(side='left', padx=5)
        
        self.logger11, clear_log11 = self.create_log_widget(tab)
        
        ttk.Button(f_run, text="清空日志", 
                  command=clear_log11, 
                  style='Secondary.TButton').pack(side='left', padx=5)
        
        ttk.Button(f_run, text="查看配置映射", 
                  command=self._show_config11).pack(side='left', padx=5)
        
        # Tab11创建完成后，检查是否需要自动加载数据
        self.master.after(100, self._auto_load_tab11_data)
    
    def _auto_load_tab11_data(self):
        """Tab11创建完成后自动加载数据"""
        try:
            config_path = self.config11_var.get()
            if config_path and config_path != "未选择配置文件" and os.path.exists(config_path):
                # 刷新映射关系选项
                self._refresh_mapping_choices11()
                # 刷新仓库列表（这是用户看到的筛选仓库显示框）
                self._refresh_warehouses11()
                self.logger11(f"✅ 自动加载配置文件: {os.path.basename(config_path)}")
                self.logger11(f"✅ 仓库列表已自动刷新")
        except Exception as e:
            self.logger11(f"⚠️ 自动加载配置失败: {e}")

    def _select_config11(self):
        """选择配置文件"""
        path = filedialog.askopenfilename(
            title="选择配置文件（列映射+物流渠道）",
            filetypes=[("Excel文件", "*.xlsx;*.xlsm;*.xls"), ("所有文件", "*.*")]
        )
        if path:
            self.config11_var.set(path)
            self.logger11(f"已选择配置文件: {path}")
            self._refresh_mapping_choices11()
            self._refresh_warehouses11()
    
    def _select_order11(self):
        """选择订单文件"""
        from excel_toolkit.ui.mixins import get_sheet_names
        path = filedialog.askopenfilename(
            title="选择订单信息文件",
            filetypes=[("Excel文件", "*.xlsx;*.xlsm;*.xls"), ("所有文件", "*.*")]
        )
        if path:
            self.order11_file_var.set(path)
            self.logger11(f"已选择订单文件: {path}")
            names = get_sheet_names(path)
            if names:
                self._update_combobox_options(self.combo11_order, self.order11_sheet_var, names)
                self.logger11(f"  工作表: {', '.join(names)}")
    
    def _select_template11(self):
        """选择模板文件"""
        from excel_toolkit.ui.mixins import get_sheet_names
        path = filedialog.askopenfilename(
            title="选择发货模板文件",
            filetypes=[("Excel文件", "*.xlsx;*.xlsm;*.xls"), ("所有文件", "*.*")]
        )
        if path:
            self.template11_file_var.set(path)
            self.logger11(f"已选择模板文件: {path}")
            names = get_sheet_names(path)
            if names:
                self._update_combobox_options(self.combo11_template, self.template11_sheet_var, names)
                self.logger11(f"  工作表: {', '.join(names)}")
    
    def _refresh_mapping_choices11(self):
        """刷新映射关系选项"""
        config_path = self.config11_var.get()
        if not config_path or config_path == "未选择配置文件":
            return

        if not os.path.exists(config_path):
            return

        try:
            choices = get_mapping_choices_from_config(config_path)
            self.combo11_mapping['values'] = choices

            current = self.mapping11_choice_var.get()
            if current not in choices:
                self.mapping11_choice_var.set("映射1")

            if hasattr(self, 'logger11'):
                if len(choices) > 1:
                    self.logger11(f"已检测到 {len(choices)} 套映射关系: {', '.join(choices)}")
        except Exception as e:
            if hasattr(self, 'logger11'):
                self.logger11(f"[WARNING] 刷新映射关系选项失败: {e}")
            else:
                print(f"[WARNING] 刷新映射关系选项失败: {e}")
    
    def _refresh_warehouses11(self):
        """刷新仓库列表"""
        config_path = self.config11_var.get()
        if not config_path or config_path == "未选择配置文件":
            return

        if not os.path.exists(config_path):
            if hasattr(self, 'logger11'):
                self.logger11(f"[WARNING] 配置文件不存在: {config_path}")
            return

        try:
            warehouses = get_warehouses_from_config(config_path)

            for widget in self.wh11_inner.winfo_children():
                widget.destroy()
            self.warehouses11_checks.clear()

            for wh in warehouses:
                var = tk.BooleanVar(value=False)
                cb = ttk.Checkbutton(self.wh11_inner, text=wh, variable=var)
                cb.pack(side='left', padx=8, pady=2)
                self.warehouses11_checks[wh] = var

            if hasattr(self, 'logger11'):
                self.logger11(f"已加载 {len(warehouses)} 个仓库: {', '.join(warehouses)}")

        except Exception as e:
            if hasattr(self, 'logger11'):
                self.logger11(f"[ERROR] 加载仓库列表失败: {e}")
            else:
                print(f"[ERROR] 加载仓库列表失败: {e}")
    
    def _select_all_warehouses11(self, select: bool):
        """全选/清空仓库"""
        for var in self.warehouses11_checks.values():
            var.set(select)
    
    def _show_fill_mode_dialog(self, existing_rows: int):
        """
        显示填充模式选择对话框
        
        Args:
            existing_rows: 模板文件中现有的数据行数
        
        Returns:
            "overwrite" 或 "append"，如果用户取消则返回 None
        """
        dialog = tk.Toplevel(self.master)
        dialog.title("⚠️ 检测到已有数据")
        dialog.geometry("450x220")
        dialog.transient(self.master)
        dialog.grab_set()
        
        # 居中显示
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        
        result = {"choice": None}
        
        # 提示信息
        msg_frame = ttk.Frame(dialog, padding=20)
        msg_frame.pack(fill='both', expand=True)
        
        ttk.Label(msg_frame, text="⚠️ 模板文件中检测到已有数据", 
                 font=("Segoe UI", 11, "bold")).pack(pady=(0, 10))
        ttk.Label(msg_frame, text=f"现有数据行数: {existing_rows} 行", 
                 font=("Segoe UI", 10)).pack(pady=5)
        ttk.Label(msg_frame, text="请选择填充模式:", 
                 font=("Segoe UI", 10)).pack(pady=(10, 5))
        
        # 按钮框架
        btn_frame = ttk.Frame(dialog, padding=(20, 0, 20, 20))
        btn_frame.pack(fill='x')
        
        def on_overwrite():
            result["choice"] = "overwrite"
            dialog.destroy()
        
        def on_append():
            result["choice"] = "append"
            dialog.destroy()
        
        def on_cancel():
            result["choice"] = None
            dialog.destroy()
        
        ttk.Button(btn_frame, text="覆盖模式 (从第2行开始，覆盖现有数据)", 
                  command=on_overwrite, width=40).pack(pady=5)
        ttk.Label(btn_frame, text="⚠️ 警告: 现有数据将被覆盖", 
                 foreground="red", font=("Segoe UI", 9)).pack()
        
        ttk.Button(btn_frame, text="追加模式 (在现有数据后追加新数据)", 
                  command=on_append, width=40).pack(pady=(15, 5))
        ttk.Label(btn_frame, text="✓ 保留现有数据，追加到末尾", 
                 foreground="green", font=("Segoe UI", 9)).pack()
        
        ttk.Button(btn_frame, text="取消", command=on_cancel, 
                  style='Secondary.TButton').pack(pady=(15, 0))
        
        dialog.wait_window()
        return result["choice"]
    
    def _show_config11(self):
        """显示配置映射"""
        config_path = self.config11_var.get()
        if not config_path or config_path == "未选择配置文件":
            messagebox.showwarning("⚠️ 警告", "请先选择配置文件。")
            return
        
        if not os.path.exists(config_path):
            messagebox.showerror("❌ 错误", f"配置文件不存在: {config_path}")
            return
        
        try:
            from excel_toolkit.shipping_fill import load_config_mapping, get_mapping_choices_from_config
            
            # 先获取可用的映射选项
            available_mappings = get_mapping_choices_from_config(config_path)
            
            config1 = load_config_mapping(config_path, "映射1", logger=lambda x: None)
            config2 = load_config_mapping(config_path, "映射2", logger=lambda x: None) if "映射2" in available_mappings else None
            config3 = load_config_mapping(config_path, "映射3", logger=lambda x: None) if "映射3" in available_mappings else None
            
            win = tk.Toplevel(self.master)
            win.title("配置映射预览")
            win.geometry("700x500")
            win.transient(self.master)
            
            text_frame = ttk.Frame(win, padding=10)
            text_frame.pack(fill='both', expand=True)
            
            scrollbar = ttk.Scrollbar(text_frame)
            scrollbar.pack(side='right', fill='y')
            
            text = tk.Text(text_frame, wrap='word', yscrollcommand=scrollbar.set,
                          font=("Consolas", 10), padx=10, pady=10)
            text.pack(side='left', fill='both', expand=True)
            scrollbar.config(command=text.yview)
            
            content = "📋 映射1（子表1）列映射关系:\n"
            content += "-" * 50 + "\n"
            for order_col, template_col in config1["column_mapping_1"].items():
                content += f"  {order_col}  →  {template_col}\n"
            
            if config2 and config2.get("column_mapping_2"):
                content += "\n📋 映射2（子表2）列映射关系:\n"
                content += "-" * 50 + "\n"
                for order_col, template_col in config2["column_mapping_2"].items():
                    content += f"  {order_col}  →  {template_col}\n"
            
            if config3 and config3.get("column_mapping_3"):
                content += "\n📋 映射3（子表3）列映射关系:\n"
                content += "-" * 50 + "\n"
                for order_col, template_col in config3["column_mapping_3"].items():
                    content += f"  {order_col}  →  {template_col}\n"
            
            content += "\n📦 仓库物流渠道配置:\n"
            content += "-" * 50 + "\n"
            for wh, carriers in config1["shipping_map"].items():
                content += f"\n【{wh}】\n"
                for carrier, service in carriers.items():
                    content += f"  {carrier}  →  {service}\n"
            
            text.insert('1.0', content)
            text.config(state='disabled')
            
            ttk.Button(win, text="关闭", command=win.destroy).pack(pady=10)
            
        except Exception as e:
            messagebox.showerror("❌ 错误", f"读取配置失败: {e}")
    
    def run_tool11(self):
        """执行发货模板填充"""
        config_file = self.config11_var.get()
        order_file = self.order11_file_var.get()
        order_sheet = self.order11_sheet_var.get()
        template_file = self.template11_file_var.get()
        template_sheet = self.template11_sheet_var.get()
        
        if not config_file or config_file == "未选择配置文件":
            messagebox.showwarning("⚠️ 警告", "请先选择配置文件。")
            return
        if not order_file or order_file == "未选择订单文件":
            messagebox.showwarning("⚠️ 警告", "请先选择订单文件。")
            return
        if not order_sheet:
            messagebox.showwarning("⚠️ 警告", "请选择订单文件的工作表。")
            return
        if not template_file or template_file == "未选择模板文件":
            messagebox.showwarning("⚠️ 警告", "请先选择模板文件。")
            return
        if not template_sheet:
            messagebox.showwarning("⚠️ 警告", "请选择模板文件的工作表。")
            return
        
        # 检查模板文件是否已有数据
        fill_mode = "overwrite"  # 默认覆盖模式
        try:
            data_check = check_template_has_data(template_file, template_sheet)
            if data_check["has_data"]:
                # 弹出对话框让用户选择
                choice = self._show_fill_mode_dialog(data_check["data_rows"])
                if choice is None:  # 用户取消
                    self.logger11("❌ 用户取消了操作")
                    return
                fill_mode = choice
        except Exception as e:
            self.logger11(f"⚠️ 检测模板数据失败: {e}，使用默认覆盖模式")
        
        selected_warehouses = [name for name, var in self.warehouses11_checks.items() if var.get()]
        warehouse_filter = selected_warehouses if selected_warehouses else None
        
        self.logger11("=" * 50)
        self.logger11("▶️ 开始运行 [11] 发货模板填充...")
        if warehouse_filter:
            self.logger11(f"   筛选仓库: {', '.join(warehouse_filter)}")
        else:
            self.logger11("   筛选仓库: 全部")
        
        mode_text = "覆盖模式" if fill_mode == "overwrite" else "追加模式"
        self.logger11(f"   填充模式: {mode_text}")
        
        self._update_status("正在填充...", icon="⏳", show_progress=True)
        self.master.config(cursor="watch")
        
        def thread_target():
            try:
                def safe_logger(msg):
                    self.master.after(0, lambda m=msg: self.logger11(m))
                
                mapping_choice = self.mapping11_choice_var.get()
                result = process_shipping_fill(
                    order_file=order_file,
                    order_sheet_name=order_sheet,
                    template_file=template_file,
                    template_sheet_name=template_sheet,
                    config_file=config_file,
                    logger=safe_logger,
                    warehouse_filter=warehouse_filter,
                    mapping_choice=mapping_choice,
                    fill_mode=fill_mode
                )
                
                def on_success():
                    self.master.config(cursor="")
                    self._update_status("就绪", icon="✅", show_progress=False)
                    messagebox.showinfo("✅ 完成", result)
                    self.logger11(f"✅ {result}")
                
                self.master.after(0, on_success)
                
            except Exception as e:
                import traceback
                error_msg = str(e)
                trace = traceback.format_exc()
                
                # 生成友好的错误提示
                friendly_msg = error_msg
                if "Permission denied" in error_msg or "PermissionError" in trace:
                    friendly_msg = "文件被占用，请关闭Excel中打开的模板文件后重试"
                elif "not subscriptable" in error_msg:
                    friendly_msg = "Excel文件格式读取错误，请确保文件格式正确"
                elif "FileNotFoundError" in trace or "不存在" in error_msg:
                    friendly_msg = "文件不存在，请检查文件路径是否正确"
                elif "没有找到列映射" in error_msg:
                    friendly_msg = "配置文件中没有找到列映射关系，请检查配置文件格式"
                elif "工作表" in error_msg and "不存在" in error_msg:
                    friendly_msg = error_msg  # 已经是友好提示
                
                def on_error():
                    self.master.config(cursor="")
                    self._update_status("错误", icon="❌", show_progress=False)
                    messagebox.showerror("操作失败", friendly_msg)
                    self.logger11(f"❌ {friendly_msg}")
                    if friendly_msg != error_msg:
                        self.logger11(f"   详细信息: {error_msg}")
                
                self.master.after(0, on_error)
        
        threading.Thread(target=thread_target, daemon=True).start()
