"""
Tab9 - 建议发货仓库功能
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading

from excel_toolkit.warehouse_router import process_warehouse_routing, read_inventory


class Tab9RouterMixin:
    """Tab9 仓库路由 Mixin"""
    
    def create_tab9_router(self, tab):
        """创建Tab9界面"""
        # 检查变量是否已经在_initialize_all_variables中创建
        if not hasattr(self, 'file9_var'):
            self.file9_var = tk.StringVar(value="未选择文件")
            self.sheet9_var = tk.StringVar()
            self.inv9_var = tk.StringVar(value="未选择发货信息表格")
            self.sku9_var = tk.StringVar(value="A")
            self.state9_var = tk.StringVar(value="B")
            self.dst9_var = tk.StringVar(value="C")
            self.block9_var = tk.BooleanVar(value=False)
            
            self._trace_persist(self.file9_var)
            self._trace_persist(self.sheet9_var)
            self._trace_persist(self.inv9_var)
            self._trace_persist(self.sku9_var)
            self._trace_persist(self.state9_var)
            self._trace_persist(self.dst9_var)
            self._trace_persist(self.block9_var)

        # 收件信息文件
        f1 = ttk.Frame(tab)
        f1.pack(fill='x', pady=5)
        ttk.Button(f1, text="选择收件信息表格", 
                  command=lambda: self.select_file_and_sheets(
                      self.file9_var, self.sheet9_var, self.combo9, 
                      "选择收件信息表格")).pack(side='left', padx=5)
        ttk.Label(f1, textvariable=self.file9_var).pack(side='left', padx=5)

        f2 = ttk.Frame(tab)
        f2.pack(fill='x', pady=5)
        ttk.Label(f2, text="子表:").pack(side='left', padx=5)
        self.combo9 = ttk.Combobox(f2, textvariable=self.sheet9_var, 
                                   state="readonly", width=22)
        self.combo9.pack(side='left', padx=5)

        # 库存文件
        f3 = ttk.Frame(tab)
        f3.pack(fill='x', pady=5)
        ttk.Button(f3, text="选择发货信息表格（库存）", 
                  command=self._select_inventory9).pack(side='left', padx=5)
        ttk.Label(f3, textvariable=self.inv9_var).pack(side='left', padx=5)

        # 列配置
        f4 = ttk.LabelFrame(tab, text="列配置", style="Section.TLabelframe")
        f4.pack(fill='x', pady=5, padx=5)
        
        ttk.Label(f4, text="SKU列:").pack(side='left', padx=5)
        ttk.Entry(f4, textvariable=self.sku9_var, width=6).pack(side='left', padx=5)
        ttk.Label(f4, text="州列:").pack(side='left', padx=(20, 5))
        ttk.Entry(f4, textvariable=self.state9_var, width=6).pack(side='left', padx=5)
        ttk.Label(f4, text="输出列:").pack(side='left', padx=(20, 5))
        ttk.Entry(f4, textvariable=self.dst9_var, width=6).pack(side='left', padx=5)

        # 屏蔽选项
        f5 = ttk.Frame(tab)
        f5.pack(fill='x', pady=5)
        ttk.Checkbutton(f5, text="屏蔽科技单(州)：GA/TX", 
                       variable=self.block9_var).pack(side='left', padx=5)
        
        # 仓库屏蔽列表
        self.block9_frame = ttk.LabelFrame(tab, text="按名称屏蔽仓库", 
                                           style="Section.TLabelframe")
        self.block9_frame.pack(fill='x', pady=5, padx=5)
        self.block9_inner = ttk.Frame(self.block9_frame)
        self.block9_inner.pack(fill='x', padx=5, pady=5)
        self.block9_checks = {}

        # 执行按钮
        f6 = ttk.Frame(tab)
        f6.pack(fill='x', pady=10)
        ttk.Button(f6, text="[9] 开始计算发货仓库", command=self.run_tool9, 
                  style='Accent.TButton').pack(side='left', padx=5)
        self.logger9, clear_log9 = self.create_log_widget(tab)
        ttk.Button(f6, text="清空日志", command=clear_log9, 
                  style='Secondary.TButton').pack(side='left', padx=5)

    def _select_inventory9(self):
        """选择库存文件"""
        path = filedialog.askopenfilename(
            title="选择发货信息表格",
            filetypes=[("Excel文件", "*.xlsx;*.xlsm;*.xls"), ("所有文件", "*.*")]
        )
        if path:
            self.inv9_var.set(path)
            self.logger9(f"已选择库存文件: {path}")
            self._refresh_block9_from_inventory()

    def _refresh_block9_from_inventory(self):
        """刷新仓库屏蔽列表"""
        inv_path = self.inv9_var.get()
        
        # 如果是从数据库加载，直接使用数据库数据
        if inv_path == "[数据库]":
            try:
                from excel_toolkit.db_operations import load_warehouse_inventory
                data = load_warehouse_inventory()
                if data:
                    warehouse_data, sku_data = data
                    
                    # 清除旧的复选框
                    for widget in self.block9_inner.winfo_children():
                        widget.destroy()
                    self.block9_checks.clear()
                    
                    # 创建新的复选框
                    all_wh = sorted(warehouse_data.keys())
                    for wh in all_wh:
                        var = tk.BooleanVar(value=False)
                        cb = ttk.Checkbutton(self.block9_inner, text=wh, variable=var)
                        cb.pack(side='left', padx=5)
                        self.block9_checks[wh] = var
                    
                    self.logger9(f"已从数据库加载 {len(all_wh)} 个仓库")
                    return
            except Exception as e:
                self.logger9(f"从数据库加载失败: {e}")
                return
        
        # 原有逻辑：从Excel文件加载
        if not inv_path or inv_path == "未选择发货信息表格":
            return
        
        try:
            sku_by_wh, wh_state = read_inventory(inv_path, logger=lambda x: None)
            
            # 清除旧的复选框
            for widget in self.block9_inner.winfo_children():
                widget.destroy()
            self.block9_checks.clear()
            
            # 创建新的复选框
            all_wh = sorted(set(list(sku_by_wh.keys()) + list(wh_state.keys())))
            for wh in all_wh:
                var = tk.BooleanVar(value=False)
                cb = ttk.Checkbutton(self.block9_inner, text=wh, variable=var)
                cb.pack(side='left', padx=5)
                self.block9_checks[wh] = var
            
            self.logger9(f"已加载 {len(all_wh)} 个仓库")
        except Exception as e:
            self.logger9(f"加载库存文件失败: {e}")

    def run_tool9(self):
        """执行仓库路由计算"""
        file = self.file9_var.get()
        sheet = self.sheet9_var.get()
        inv = self.inv9_var.get()
        sku_col = self.sku9_var.get()
        state_col = self.state9_var.get()
        dst_col = self.dst9_var.get()
        
        if not file or file == "未选择文件":
            messagebox.showwarning("⚠️ 警告", "请先选择收件信息表格。")
            return
        if not sheet:
            messagebox.showwarning("⚠️ 警告", "请选择子表。")
            return
        if not inv or inv == "未选择发货信息表格":
            messagebox.showwarning("⚠️ 警告", "请先选择发货信息表格。")
            return

        # 获取屏蔽的仓库
        blocked_wh = [name for name, var in self.block9_checks.items() if var.get()]

        self.logger9("=" * 50)
        self.logger9(f"▶️ 开始运行 [9] 建议发货仓库...")
        
        self._update_status("正在计算...", icon="⏳", show_progress=True)
        self.master.config(cursor="watch")
        
        def thread_target():
            try:
                def safe_logger(msg):
                    self.master.after(0, lambda m=msg: self.logger9(m))
                
                result = process_warehouse_routing(
                    file, sheet, sku_col, state_col, dst_col, inv,
                    safe_logger, self.block9_var.get(), blocked_wh
                )
                
                def on_success():
                    self.master.config(cursor="")
                    self._update_status("就绪", icon="✅", show_progress=False)
                    messagebox.showinfo("✅ 完成", result)
                    self.logger9(result)
                
                self.master.after(0, on_success)
                
            except Exception as e:
                error_msg = str(e)
                def on_error(msg=error_msg):
                    self.master.config(cursor="")
                    self._update_status("错误", icon="❌", show_progress=False)
                    messagebox.showerror("❌ 错误", msg)
                    self.logger9(f"❌ 发生错误: {msg}")
                self.master.after(0, on_error)

        threading.Thread(target=thread_target, daemon=True).start()


