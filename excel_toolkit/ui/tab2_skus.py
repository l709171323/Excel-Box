"""
Tab2 - SKU智能填充功能
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
import json

from excel_toolkit.sku_fill import process_skus, identify_header_mapping


class Tab2SkusMixin:
    """Tab2 SKU填充 Mixin"""
    
    def create_tab2_skus(self, tab):
        """创建Tab2界面"""
        # 变量已在 _initialize_all_variables 中创建，这里只需添加trace和加载配置
        # 为订单文件和SKU数据库使用独立的持久化保存
        if not hasattr(self, '_tab2_traces_added'):
            self.file2_var.trace_add("write", lambda *a: self._save_tab2_order_file())
            self.sku_db2_var.trace_add("write", lambda *a: self._save_tab2_sku_db_file())
            self.sku_db2_sheet_var.trace_add("write", lambda *a: self._save_tab2_sku_db_file())
            self._trace_persist(self.order2_sheet_var)
            # 加载订单文件路径（SKU数据库文件将在UI创建后加载）
            self._load_tab2_order_file()
            self._tab2_traces_added = True  # 标记已添加trace，避免重复

        # 文件选择区
        f1 = ttk.Frame(tab)
        f1.pack(fill='x', pady=5)
        ttk.Button(f1, text="选择订单表格", 
                  command=self._select_order2).pack(side='left', padx=5)
        ttk.Label(f1, textvariable=self.file2_var).pack(side='left', padx=5)
        
        # 订单工作表选择
        f1_sheet = ttk.Frame(tab)
        f1_sheet.pack(fill='x', pady=5, padx=5)
        ttk.Label(f1_sheet, text="订单工作表:").pack(side='left', padx=5)
        self.order2_sheet_combo = ttk.Combobox(f1_sheet, textvariable=self.order2_sheet_var,
                                                state="readonly", width=25)
        self.order2_sheet_combo.pack(side='left', padx=5)
        ttk.Label(f1_sheet, text="（不选择=处理所有工作表）", 
                 font=("Segoe UI", 9)).pack(side='left', padx=5)

        f2 = ttk.Frame(tab)
        f2.pack(fill='x', pady=5, padx=5)
        ttk.Button(f2, text="选择SKU数据库", 
                  command=self._select_sku_database2).pack(side='left', padx=5)
        ttk.Label(f2, textvariable=self.sku_db2_var).pack(side='left', padx=5)
        
        # SKU数据库工作表选择
        f2_sheet = ttk.Frame(tab)
        f2_sheet.pack(fill='x', pady=5, padx=5)
        ttk.Label(f2_sheet, text="SKU数据库工作表:").pack(side='left', padx=5)
        
        self.sku_db2_sheet_combo = ttk.Combobox(f2_sheet, textvariable=self.sku_db2_sheet_var,
                                                state="readonly", width=25)
        self.sku_db2_sheet_combo.pack(side='left', padx=5)
        self.sku_db2_sheet_combo.bind('<<ComboboxSelected>>', self._on_sku_sheet_changed2)

        # SKU数据库列映射区
        f_db_map = ttk.LabelFrame(tab, text="SKU数据库列映射（自动识别，可手动调整）", 
                                  style="Section.TLabelframe")
        f_db_map.pack(fill='x', pady=5, padx=5)
        
        # 变量已在预初始化中创建
        if not hasattr(self, 'db_sku_col'):
            self.db_sku_col = tk.StringVar()
            self.db_l_col = tk.StringVar()
            self.db_w_col = tk.StringVar()
            self.db_h_col = tk.StringVar()
            self.db_wt_col = tk.StringVar()
            
            self._trace_persist(self.db_sku_col)
            self._trace_persist(self.db_l_col)
            self._trace_persist(self.db_w_col)
            self._trace_persist(self.db_h_col)
            self._trace_persist(self.db_wt_col)
        
        row1 = ttk.Frame(f_db_map)
        row1.pack(fill='x', padx=8, pady=4)
        ttk.Label(row1, text="SKU列:", width=10).pack(side='left')
        self.db_sku_combo = ttk.Combobox(row1, textvariable=self.db_sku_col, 
                                         state="readonly", width=20)
        self.db_sku_combo.pack(side='left', padx=5)
        
        ttk.Label(row1, text="长列:", width=10).pack(side='left', padx=(20, 0))
        self.db_l_combo = ttk.Combobox(row1, textvariable=self.db_l_col, 
                                       state="readonly", width=20)
        self.db_l_combo.pack(side='left', padx=5)
        
        row2 = ttk.Frame(f_db_map)
        row2.pack(fill='x', padx=8, pady=4)
        ttk.Label(row2, text="宽列:", width=10).pack(side='left')
        self.db_w_combo = ttk.Combobox(row2, textvariable=self.db_w_col, 
                                       state="readonly", width=20)
        self.db_w_combo.pack(side='left', padx=5)
        
        ttk.Label(row2, text="高列:", width=10).pack(side='left', padx=(20, 0))
        self.db_h_combo = ttk.Combobox(row2, textvariable=self.db_h_col, 
                                       state="readonly", width=20)
        self.db_h_combo.pack(side='left', padx=5)
        
        row3 = ttk.Frame(f_db_map)
        row3.pack(fill='x', padx=8, pady=4)
        ttk.Label(row3, text="重量列:", width=10).pack(side='left')
        self.db_wt_combo = ttk.Combobox(row3, textvariable=self.db_wt_col, 
                                        state="readonly", width=20)
        self.db_wt_combo.pack(side='left', padx=5)

        # 目标表格列配置区
        f_target_map = ttk.LabelFrame(tab, text="目标表格列配置（输入列号如A/B/C或列名）", 
                                      style="Section.TLabelframe")
        f_target_map.pack(fill='x', pady=5, padx=5)
        
        # 模板管理区
        template_mgmt = ttk.Frame(f_target_map)
        template_mgmt.pack(fill='x', padx=8, pady=4)
        ttk.Label(template_mgmt, text="配置模板:").pack(side='left', padx=5)
        
        # 变量已在预初始化中创建
        if not hasattr(self, 'template2_var'):
            self.template2_var = tk.StringVar(value="默认")
            self._trace_persist(self.template2_var)
        
        self.template2_combo = ttk.Combobox(template_mgmt, textvariable=self.template2_var,
                                            state="readonly", width=20)
        self.template2_combo.pack(side='left', padx=5)
        self.template2_combo.bind('<<ComboboxSelected>>', self._load_target_template2)
        
        ttk.Button(template_mgmt, text="保存当前配置", 
                  command=self._save_target_template2).pack(side='left', padx=5)
        ttk.Button(template_mgmt, text="删除模板", 
                  command=self._delete_target_template2).pack(side='left', padx=5)
        
        # 初始化模板列表
        self._refresh_template_list2()
        
        # 变量已在预初始化中创建
        if not hasattr(self, 'target_sku_col'):
            self.target_sku_col = tk.StringVar(value="A")
            self.target_qty_col = tk.StringVar(value="B")
            self.target_l_col = tk.StringVar(value="C")
            self.target_w_col = tk.StringVar(value="D")
            self.target_h_col = tk.StringVar(value="E")
            self.target_wt_col = tk.StringVar(value="F")
            
            self._trace_persist(self.target_sku_col)
            self._trace_persist(self.target_qty_col)
            self._trace_persist(self.target_l_col)
            self._trace_persist(self.target_w_col)
            self._trace_persist(self.target_h_col)
        self._trace_persist(self.target_wt_col)
        
        trow1 = ttk.Frame(f_target_map)
        trow1.pack(fill='x', padx=8, pady=4)
        ttk.Label(trow1, text="SKU列:", width=10).pack(side='left')
        ttk.Entry(trow1, textvariable=self.target_sku_col, width=8).pack(side='left', padx=5)
        ttk.Label(trow1, text="数量列:", width=10).pack(side='left', padx=(20, 0))
        ttk.Entry(trow1, textvariable=self.target_qty_col, width=8).pack(side='left', padx=5)
        
        trow2 = ttk.Frame(f_target_map)
        trow2.pack(fill='x', padx=8, pady=4)
        ttk.Label(trow2, text="长填充到:", width=10).pack(side='left')
        ttk.Entry(trow2, textvariable=self.target_l_col, width=8).pack(side='left', padx=5)
        ttk.Label(trow2, text="宽填充到:", width=10).pack(side='left', padx=(20, 0))
        ttk.Entry(trow2, textvariable=self.target_w_col, width=8).pack(side='left', padx=5)
        
        trow3 = ttk.Frame(f_target_map)
        trow3.pack(fill='x', padx=8, pady=4)
        ttk.Label(trow3, text="高填充到:", width=10).pack(side='left')
        ttk.Entry(trow3, textvariable=self.target_h_col, width=8).pack(side='left', padx=5)
        ttk.Label(trow3, text="重量填充到:", width=10).pack(side='left', padx=(20, 0))
        ttk.Entry(trow3, textvariable=self.target_wt_col, width=8).pack(side='left', padx=5)

        # 忽略数量选项
        trow4 = ttk.Frame(f_target_map)
        trow4.pack(fill='x', padx=8, pady=4)
        self.ignore_qty_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(trow4, text="计算单个SKU数据（忽略数量列）", 
                        variable=self.ignore_qty_var).pack(side='left')

        # 执行按钮区
        f3 = ttk.Frame(tab)
        f3.pack(fill='x', pady=10)
        ttk.Button(f3, text="[2] 开始智能填充SKU", command=self.run_tool2, 
                  style='Accent.TButton').pack(side='left', padx=5)
        self.logger2, clear_log2 = self.create_log_widget(tab)
        ttk.Button(f3, text="清空日志", command=clear_log2, 
                  style='Secondary.TButton').pack(side='left', padx=5)
        
        # UI创建完毕后，加载SKU数据库文件路径和工作表配置
        # 必须在 sku_db2_sheet_combo 创建之后调用，否则无法更新下拉框
        self._load_tab2_sku_db_file()

    def _read_sku_db_sheets2(self, file_path):
        """读取SKU数据库文件的工作表列表（公共方法）"""
        file_ext = os.path.splitext(file_path)[1].lower()
        sheet_names = []
        
        if file_ext in ['.xlsx', '.xlsm']:
            # 使用openpyxl读取新版Excel
            from excel_toolkit.excel_lite import ExcelReader
            wb = ExcelReader(file_path, read_only=True, data_only=True)
            sheet_names = wb.sheetnames
            wb.close()
        elif file_ext == '.xls':
            # 使用xlrd读取旧版Excel
            import xlrd
            wb = xlrd.open_workbook(file_path)
            sheet_names = wb.sheetnames
        else:
            raise ValueError(f"不支持的文件格式: {file_ext}")
        
        return sheet_names
    
    def _select_sku_database2(self):
        """选择外部SKU数据库文件并读取工作表列表"""
        path = filedialog.askopenfilename(
            title="选择SKU数据库表格",
            filetypes=[("Excel文件", "*.xlsx;*.xlsm;*.xls"), ("所有文件", "*.*")]
        )
        if not path:
            return
            
        self.sku_db2_var.set(path)
        self.logger2(f"已选择SKU数据库: {os.path.basename(path)}")
        
        try:
            # 使用公共方法读取工作表列表
            sheet_names = self._read_sku_db_sheets2(path)
            
            if not sheet_names:
                self.logger2("  警告：未找到工作表")
                return
            
            # 更新工作表下拉框
            self.sku_db2_sheet_combo['values'] = sheet_names
            
            # 智能选择默认工作表
            default_sheet = None
            if "商品资料" in sheet_names:
                default_sheet = "商品资料"
            elif "SKU" in sheet_names:
                default_sheet = "SKU"
            elif any("sku" in s.lower() for s in sheet_names):
                default_sheet = next(s for s in sheet_names if "sku" in s.lower())
            else:
                default_sheet = sheet_names[0]
            
            self.sku_db2_sheet_var.set(default_sheet)
            self.logger2(f"  已选择工作表: {default_sheet}")
            
            # 读取选中工作表的表头
            self._load_sku_sheet_headers2()
            
        except Exception as e:
            self.logger2(f"  读取文件失败: {e}")
            messagebox.showwarning("警告", f"无法读取文件。\n错误: {e}")
    
    def _on_sku_sheet_changed2(self, event=None):
        """工作表选择变化时重新读取表头"""
        self._load_sku_sheet_headers2()
    
    def _load_sku_sheet_headers2(self):
        """读取选中工作表的表头并自动识别列映射"""
        sku_db_path = self.sku_db2_var.get()
        sheet_name = self.sku_db2_sheet_var.get()
        
        if not sku_db_path or sku_db_path == "未选择SKU数据库":
            return
        
        if not sheet_name:
            return
        
        try:
            # 根据文件扩展名选择合适的读取方式
            file_ext = os.path.splitext(sku_db_path)[1].lower()
            headers = []
            
            if file_ext in ['.xlsx', '.xlsm']:
                # 使用openpyxl读取新版Excel
                from excel_toolkit.excel_lite import ExcelReader, ExcelWriter
                wb = ExcelReader(sku_db_path, read_only=True, data_only=True)
                
                if sheet_name not in wb.sheetnames:
                    self.logger2(f"  警告：工作表 '{sheet_name}' 不存在")
                    wb.close()
                    return
                
                ws = wb[sheet_name]
                headers = [cell.value for cell in list(ws.rows)[0] if cell.value]
                wb.close()
                
            elif file_ext == '.xls':
                # 使用xlrd读取旧版Excel
                import xlrd
                wb = xlrd.open_workbook(sku_db_path)
                
                if sheet_name not in wb.sheetnames:
                    self.logger2(f"  警告：工作表 '{sheet_name}' 不存在")
                    return
                
                ws = wb.sheet_by_name(sheet_name)
                if ws.nrows > 0:
                    headers = [cell.value for cell in ws.row(0) if cell.value]
            
            if not headers:
                self.logger2("  警告：未找到表头，请手动配置列映射")
                return
            
            # 更新下拉框选项
            self.db_sku_combo['values'] = headers
            self.db_l_combo['values'] = headers
            self.db_w_combo['values'] = headers
            self.db_h_combo['values'] = headers
            self.db_wt_combo['values'] = headers
            
            # 智能识别并自动选择
            mapping = identify_header_mapping(headers)
            if mapping['sku']:
                self.db_sku_col.set(mapping['sku'])
            if mapping['l']:
                self.db_l_col.set(mapping['l'])
            if mapping['w']:
                self.db_w_col.set(mapping['w'])
            if mapping['h']:
                self.db_h_col.set(mapping['h'])
            if mapping['wt']:
                self.db_wt_col.set(mapping['wt'])
            
            self.logger2(f"  已自动识别列映射: SKU={mapping['sku']}, 长={mapping['l']}, "
                        f"宽={mapping['w']}, 高={mapping['h']}, 重量={mapping['wt']}")
            self.logger2("  如有误，请手动调整下拉框")
            
        except Exception as e:
            self.logger2(f"  读取表头失败: {e}")
            messagebox.showwarning("警告", f"无法读取工作表表头，请手动配置列映射。\n错误: {e}")
    
    def _get_templates_file2(self):
        """获取模板配置文件路径"""
        config_dir = self._config_dir()
        return os.path.join(config_dir, "tab2_target_templates.json")
    
    def _get_order_file_path2(self):
        """获取订单文件路径保存文件"""
        config_dir = self._config_dir()
        return os.path.join(config_dir, "tab2_order_file.json")
    
    def _get_sku_db_file_path2(self):
        """获取SKU数据库文件路径保存文件"""
        config_dir = self._config_dir()
        return os.path.join(config_dir, "tab2_sku_db_file.json")
    
    def _save_tab2_order_file(self):
        """保存订单文件路径到独立文件"""
        file_path = self._get_order_file_path2()
        data = {"order_file": self.file2_var.get()}
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存订单文件路径失败: {e}")
    
    def _load_tab2_order_file(self):
        """从独立文件加载订单文件路径"""
        file_path = self._get_order_file_path2()
        if os.path.exists(file_path):
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    order_file = data.get("order_file", "")
                    if order_file and order_file != "未选择文件":
                        self.file2_var.set(order_file)
            except Exception as e:
                print(f"加载订单文件路径失败: {e}")
    
    def _save_tab2_sku_db_file(self):
        """保存SKU数据库文件路径和选中的工作表到独立文件"""
        file_path = self._get_sku_db_file_path2()
        
        # 获取当前值
        sku_db_file = self.sku_db2_var.get()
        sku_db_sheet = self.sku_db2_sheet_var.get() if hasattr(self, 'sku_db2_sheet_var') else ""
        
        # 如果文件路径有效但工作表为空，尝试从现有配置中保留工作表
        if sku_db_file and sku_db_file != "未选择SKU数据库" and not sku_db_sheet:
            if os.path.exists(file_path):
                try:
                    with open(file_path, 'r', encoding='utf-8') as f:
                        existing_data = json.load(f)
                        # 如果是同一个文件，保留原来的工作表设置
                        if existing_data.get("sku_db_file") == sku_db_file:
                            sku_db_sheet = existing_data.get("sku_db_sheet", "")
                except Exception:
                    pass
        
        data = {
            "sku_db_file": sku_db_file,
            "sku_db_sheet": sku_db_sheet
        }
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存SKU数据库文件路径失败: {e}")
    
    def _load_tab2_sku_db_file(self):
        """从独立文件加载SKU数据库文件路径和工作表"""
        file_path = self._get_sku_db_file_path2()
        if os.path.exists(file_path):
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    sku_db_file = data.get("sku_db_file", "")
                    sku_db_sheet = data.get("sku_db_sheet", "")
                    
                    if sku_db_file and sku_db_file != "未选择SKU数据库" and os.path.exists(sku_db_file):
                        # 设置文件路径（不触发保存，因为是加载过程）
                        self.sku_db2_var.set(sku_db_file)
                        
                        # 读取工作表列表
                        try:
                            sheet_names = self._read_sku_db_sheets2(sku_db_file)
                            
                            if sheet_names and hasattr(self, 'sku_db2_sheet_combo'):
                                # 更新工作表下拉框
                                self.sku_db2_sheet_combo['values'] = sheet_names
                                
                                # 恢复之前选中的工作表，如果存在的话
                                if sku_db_sheet and sku_db_sheet in sheet_names:
                                    self.sku_db2_sheet_var.set(sku_db_sheet)
                                    print(f"[自动加载] SKU数据库: {os.path.basename(sku_db_file)}, 工作表: {sku_db_sheet}")
                                else:
                                    # 否则智能选择默认工作表
                                    default_sheet = None
                                    if "商品资料" in sheet_names:
                                        default_sheet = "商品资料"
                                    elif "SKU" in sheet_names:
                                        default_sheet = "SKU"
                                    elif any("sku" in s.lower() for s in sheet_names):
                                        default_sheet = next(s for s in sheet_names if "sku" in s.lower())
                                    else:
                                        default_sheet = sheet_names[0]
                                    
                                    if default_sheet:
                                        self.sku_db2_sheet_var.set(default_sheet)
                                        print(f"[自动加载] SKU数据库: {os.path.basename(sku_db_file)}, 工作表: {default_sheet}（智能选择）")
                                
                                # 读取表头
                                self._load_sku_sheet_headers2()
                        except Exception as e:
                            print(f"[自动加载] 读取SKU数据库工作表失败: {e}")
            except Exception as e:
                print(f"加载SKU数据库文件路径失败: {e}")
    
    def _refresh_template_list2(self):
        """刷新模板列表"""
        templates_file = self._get_templates_file2()
        templates = []
        
        if os.path.exists(templates_file):
            try:
                with open(templates_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    templates = list(data.keys())
            except Exception:
                pass
        
        if not templates:
            templates = ["默认"]
        
        # 保留当前选中的模板（即使它不在文件中）
        current = self.template2_var.get()
        if current and current not in templates:
            templates.append(current)
        
        self.template2_combo['values'] = templates
        
        # 如果没有当前选择，选择第一个
        if not current:
            self.template2_var.set(templates[0])
    
    def _save_target_template2(self):
        """保存当前目标列配置为模板"""
        # 创建自定义对话框，继承置顶状态
        dialog = tk.Toplevel(self.master)
        dialog.title("保存配置模板")
        dialog.transient(self.master)
        dialog.grab_set()
        
        # 如果主窗口是置顶的，对话框也置顶
        if self.master.attributes('-topmost'):
            dialog.attributes('-topmost', True)
        
        # 居中显示
        dialog.geometry("400x150")
        x = self.master.winfo_x() + (self.master.winfo_width() - 400) // 2
        y = self.master.winfo_y() + (self.master.winfo_height() - 150) // 2
        dialog.geometry(f"+{x}+{y}")
        
        # 提示文本
        ttk.Label(dialog, text="请输入模板名称:", 
                 font=("Segoe UI", 10)).pack(pady=(20, 10))
        
        # 输入框
        template_name_var = tk.StringVar(value=self.template2_var.get())
        entry = ttk.Entry(dialog, textvariable=template_name_var, width=30, font=("Segoe UI", 10))
        entry.pack(pady=10)
        entry.focus_set()
        entry.select_range(0, tk.END)
        
        result_name = None
        
        def on_ok():
            nonlocal result_name
            result_name = template_name_var.get()
            dialog.destroy()
        
        def on_cancel():
            dialog.destroy()
        
        # 按钮区
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="确定", command=on_ok, width=10).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="取消", command=on_cancel, width=10).pack(side='left', padx=5)
        
        # 绑定回车键
        entry.bind('<Return>', lambda e: on_ok())
        entry.bind('<Escape>', lambda e: on_cancel())
        
        # 等待对话框关闭
        dialog.wait_window()
        
        template_name = result_name
        
        if not template_name or not template_name.strip():
            return
        
        template_name = template_name.strip()
        
        # 收集当前配置
        config = {
            'sku': self.target_sku_col.get(),
            'qty': self.target_qty_col.get(),
            'l': self.target_l_col.get(),
            'w': self.target_w_col.get(),
            'h': self.target_h_col.get(),
            'wt': self.target_wt_col.get()
        }
        
        # 读取现有模板
        templates_file = self._get_templates_file2()
        templates = {}
        
        if os.path.exists(templates_file):
            try:
                with open(templates_file, 'r', encoding='utf-8') as f:
                    templates = json.load(f)
            except Exception:
                pass
        
        # 保存模板
        templates[template_name] = config
        
        try:
            with open(templates_file, 'w', encoding='utf-8') as f:
                json.dump(templates, f, ensure_ascii=False, indent=2)
            
            self.logger2(f"✅ 配置模板 '{template_name}' 已保存")
            messagebox.showinfo("✅ 成功", f"配置模板 '{template_name}' 已保存！")
            
            # 刷新模板列表并选中新模板
            self._refresh_template_list2()
            self.template2_var.set(template_name)
            
        except Exception as e:
            messagebox.showerror("❌ 错误", f"保存模板失败: {e}")
    
    def _load_target_template2(self, event=None):
        """加载选中的目标列配置模板"""
        template_name = self.template2_var.get()
        
        if not template_name or template_name == "默认":
            return
        
        templates_file = self._get_templates_file2()
        
        if not os.path.exists(templates_file):
            return
        
        try:
            with open(templates_file, 'r', encoding='utf-8') as f:
                templates = json.load(f)
            
            if template_name not in templates:
                self.logger2(f"⚠️ 模板 '{template_name}' 不存在")
                return
            
            config = templates[template_name]
            
            # 填充配置
            self.target_sku_col.set(config.get('sku', 'A'))
            self.target_qty_col.set(config.get('qty', 'B'))
            self.target_l_col.set(config.get('l', 'C'))
            self.target_w_col.set(config.get('w', 'D'))
            self.target_h_col.set(config.get('h', 'E'))
            self.target_wt_col.set(config.get('wt', 'F'))
            
            self.logger2(f"✅ 已加载模板 '{template_name}': SKU={config.get('sku')}, 数量={config.get('qty')}, "
                        f"长={config.get('l')}, 宽={config.get('w')}, 高={config.get('h')}, 重量={config.get('wt')}")
            
        except Exception as e:
            messagebox.showerror("❌ 错误", f"加载模板失败: {e}")
    
    def _delete_target_template2(self):
        """删除选中的配置模板"""
        template_name = self.template2_var.get()
        
        if not template_name or template_name == "默认":
            messagebox.showwarning("⚠️ 警告", "不能删除默认模板")
            return
        
        # 确认删除
        if not messagebox.askyesno("确认删除", f"确定要删除模板 '{template_name}' 吗？"):
            return
        
        templates_file = self._get_templates_file2()
        
        if not os.path.exists(templates_file):
            return
        
        try:
            with open(templates_file, 'r', encoding='utf-8') as f:
                templates = json.load(f)
            
            if template_name in templates:
                del templates[template_name]
                
                with open(templates_file, 'w', encoding='utf-8') as f:
                    json.dump(templates, f, ensure_ascii=False, indent=2)
                
                self.logger2(f"✅ 模板 '{template_name}' 已删除")
                messagebox.showinfo("✅ 成功", f"模板 '{template_name}' 已删除！")
                
                # 刷新模板列表
                self._refresh_template_list2()
            
        except Exception as e:
            messagebox.showerror("❌ 错误", f"删除模板失败: {e}")

    def _select_order2(self):
        """选择订单文件"""
        from excel_toolkit.ui.mixins import get_sheet_names
        from tkinter import filedialog
        path = filedialog.askopenfilename(
            title="选择订单表格",
            filetypes=[("表格文件", "*.xlsx;*.xlsm;*.xls"), ("所有文件", "*.*")]
        )
        if path:
            self.file2_var.set(path)
            self.logger2(f"已选择订单文件: {path}")
            names = get_sheet_names(path)
            if names:
                self._update_combobox_options(self.order2_sheet_combo, self.order2_sheet_var, names)
                self.logger2(f"  工作表: {', '.join(names)}")
    
    def run_tool2(self):
        """执行SKU填充"""
        file = self.file2_var.get()
        sku_db = self.sku_db2_var.get()
        sku_db_sheet = self.sku_db2_sheet_var.get()
        order_sheet = self.order2_sheet_var.get().strip() if hasattr(self, 'order2_sheet_var') else None
        
        if not file or file == "未选择文件":
            messagebox.showwarning("⚠️ 警告", "请先选择订单表格。")
            return
        if not sku_db or sku_db == "未选择SKU数据库":
            messagebox.showwarning("⚠️ 警告", "请先选择SKU数据库表格。")
            return
        if not sku_db_sheet:
            messagebox.showwarning("⚠️ 警告", "请选择SKU数据库工作表。")
            return
        
        db_col_map = {
            'sku': self.db_sku_col.get(),
            'l': self.db_l_col.get(),
            'w': self.db_w_col.get(),
            'h': self.db_h_col.get(),
            'wt': self.db_wt_col.get()
        }
        
        missing_db = [k for k, v in db_col_map.items() if not v]
        if missing_db:
            messagebox.showwarning("⚠️ 警告", f"请配置SKU数据库的列映射: {', '.join(missing_db)}")
            return
        
        target_col_map = {
            'sku': self.target_sku_col.get(),
            'qty': self.target_qty_col.get(),
            'l': self.target_l_col.get(),
            'w': self.target_w_col.get(),
            'h': self.target_h_col.get(),
            'wt': self.target_wt_col.get()
        }
        
        missing_target = [k for k, v in target_col_map.items() if not v]
        if missing_target:
            messagebox.showwarning("⚠️ 警告", f"请配置目标表格的列: {', '.join(missing_target)}")
            return

        self.logger2("=" * 50)
        self.logger2(f"▶️ 开始运行 [2] 填充SKU信息...")
        self.logger2(f"SKU数据库工作表: {sku_db_sheet}")
        self.logger2(f"数据库映射: {db_col_map}")
        self.logger2(f"目标列配置: {target_col_map}")
        
        self._update_status("正在填充SKU信息...", icon="⏳", show_progress=True)
        self.master.config(cursor="watch")
        
        def thread_target():
            try:
                def safe_logger(msg):
                    self.master.after(0, lambda m=msg: self.logger2(m))
                    
                ignore_qty = self.ignore_qty_var.get()
                stats = process_skus(file, sku_db, db_col_map, target_col_map, safe_logger, sku_db_sheet, ignore_qty, order_sheet)
                
                def on_success():
                    self.master.config(cursor="")
                    self._update_status("就绪", icon="✅", show_progress=False)
                    
                    if stats['sheets_processed'] > 0:
                        msg = (
                            f"SKU填充完成！\n\n"
                            f"处理工作表数: {stats['sheets_processed']}\n"
                            f"填充行数: {stats['rows_filled']}\n"
                            f"文件已保存: {os.path.basename(file)}"
                        )
                    else:
                        msg = "SKU填充完成。\n未处理任何数据（可能未找到匹配列或数据为空）。"
                        
                    messagebox.showinfo("✅ 完成", msg)
                    self.logger2(msg.replace("\n\n", "\n"))
                
                self.master.after(0, on_success)
                
            except Exception as e:
                error_msg = str(e)  # 在闭包外捕获错误信息
                def on_error(msg=error_msg):
                    self.master.config(cursor="")
                    self._update_status("错误", icon="❌", show_progress=False)
                    messagebox.showerror("❌ 错误", msg)
                    self.logger2(f"❌ 发生错误: {msg}")
                self.master.after(0, on_error)

        threading.Thread(target=thread_target, daemon=True).start()





























