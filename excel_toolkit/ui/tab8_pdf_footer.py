import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import threading

from excel_toolkit.pdf_ocr import add_pdf_footer_for_shipping_label, load_sku_full_name_map_from_excel

# 常用字体列表
COMMON_FONTS = [
    ("STSong-Light", "宋体 (STSong)"),
    ("STHeiti-Regular", "黑体 (STHeiti)"),
    ("msyh", "微软雅黑"),
    ("Helvetica", "Helvetica"),
    ("Helvetica-Bold", "Helvetica 粗体"),
    ("Times-Roman", "Times Roman"),
    ("Times-Bold", "Times 粗体"),
    ("Courier", "Courier"),
    ("Courier-Bold", "Courier 粗体"),
]

# 常用字号列表
COMMON_FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24]


class Tab8PdfFooterMixin:
    def create_tab8_pdf_footer(self, tab):
        # 检查变量是否已经在_initialize_all_variables中创建
        if not hasattr(self, 'pdf8_input_var'):
            self.pdf8_input_var = tk.StringVar(value="未选择PDF")
            self.pdf8_output_var = tk.StringVar(value="未选择输出目录")
            self.pdf8_map_excel_var = tk.StringVar(value="未选择SKU映射Excel")
            self.pdf8_map_sheet_var = tk.StringVar(value="")
            self.pdf8_short_col_var = tk.StringVar(value="")
            self.pdf8_full_col_var = tk.StringVar(value="")
            self.pdf8_overwrite_var = tk.BooleanVar(value=False)
            self.pdf8_font_var = tk.StringVar(value="STSong-Light")
            self.pdf8_fontsize_var = tk.StringVar(value="10")

            for v in [
                self.pdf8_input_var,
                self.pdf8_output_var,
                self.pdf8_map_excel_var,
                self.pdf8_map_sheet_var,
                self.pdf8_short_col_var,
                self.pdf8_full_col_var,
                self.pdf8_overwrite_var,
                self.pdf8_font_var,
                self.pdf8_fontsize_var,
            ]:
                self._trace_persist(v)
        
        # 这些不需要持久化的变量每次都重新初始化
        self.pdf8_sheet_names = []  # 存储sheet名称列表
        self.pdf8_selected_files = []  # 存储选中的PDF文件列表

        f_pdf = ttk.LabelFrame(tab, text="PDF 文件", style="Section.TLabelframe")
        f_pdf.pack(fill='x', pady=6, padx=5)

        f_pdf1 = ttk.Frame(f_pdf)
        f_pdf1.pack(fill='x', padx=6, pady=6)
        ttk.Button(f_pdf1, text="选择PDF文件(可多选)", command=self._select_pdf8_input).pack(side='left', padx=5)
        ttk.Label(f_pdf1, textvariable=self.pdf8_input_var, wraplength=400).pack(side='left', padx=5, fill='x', expand=True)
        ttk.Button(f_pdf1, text="清空列表", command=self._clear_pdf8_list).pack(side='left', padx=5)

        f_pdf2 = ttk.Frame(f_pdf)
        f_pdf2.pack(fill='x', padx=6, pady=(0, 6))
        ttk.Button(f_pdf2, text="选择输出目录", command=self._select_pdf8_output).pack(side='left', padx=5)
        ttk.Label(f_pdf2, textvariable=self.pdf8_output_var, wraplength=400).pack(side='left', padx=5, fill='x', expand=True)
        ttk.Checkbutton(f_pdf2, text="覆盖模式（直接修改原PDF）", variable=self.pdf8_overwrite_var, command=self._toggle_overwrite8).pack(side='left', padx=(20, 5))

        f_map = ttk.LabelFrame(tab, text="SKU 映射（简称→全称）", style="Section.TLabelframe")
        f_map.pack(fill='x', pady=6, padx=5)

        f_map1 = ttk.Frame(f_map)
        f_map1.pack(fill='x', padx=6, pady=6)
        ttk.Button(f_map1, text="选择映射Excel", command=self._select_pdf8_map_excel).pack(side='left', padx=5)
        ttk.Label(f_map1, textvariable=self.pdf8_map_excel_var, wraplength=520).pack(side='left', padx=5, fill='x', expand=True)

        f_map2 = ttk.Frame(f_map)
        f_map2.pack(fill='x', padx=6, pady=(0, 6))
        ttk.Label(f_map2, text="Sheet:").pack(side='left', padx=5)
        self.pdf8_sheet_combo = ttk.Combobox(f_map2, textvariable=self.pdf8_map_sheet_var, width=20, state="readonly")
        self.pdf8_sheet_combo.pack(side='left', padx=5)
        ttk.Label(f_map2, text="简称列(可选):").pack(side='left', padx=(20, 5))
        ttk.Entry(f_map2, textvariable=self.pdf8_short_col_var, width=10).pack(side='left', padx=5)
        ttk.Label(f_map2, text="全称列(可选):").pack(side='left', padx=(20, 5))
        ttk.Entry(f_map2, textvariable=self.pdf8_full_col_var, width=10).pack(side='left', padx=5)
        ttk.Button(f_map2, text="查看映射关系", command=self._show_sku_mapping8).pack(side='left', padx=(20, 5))

        # 字体设置区域
        f_font = ttk.LabelFrame(tab, text="页脚样式", style="Section.TLabelframe")
        f_font.pack(fill='x', pady=6, padx=5)
        
        f_font1 = ttk.Frame(f_font)
        f_font1.pack(fill='x', padx=6, pady=6)
        
        ttk.Label(f_font1, text="字体:").pack(side='left', padx=5)
        font_display_names = [f[1] for f in COMMON_FONTS]
        self.pdf8_font_combo = ttk.Combobox(f_font1, textvariable=self.pdf8_font_var, width=18, state="readonly")
        self.pdf8_font_combo['values'] = [f[0] for f in COMMON_FONTS]
        self.pdf8_font_combo.pack(side='left', padx=5)
        
        ttk.Label(f_font1, text="字号:").pack(side='left', padx=(20, 5))
        self.pdf8_fontsize_combo = ttk.Combobox(f_font1, textvariable=self.pdf8_fontsize_var, width=6, state="readonly")
        self.pdf8_fontsize_combo['values'] = COMMON_FONT_SIZES
        self.pdf8_fontsize_combo.pack(side='left', padx=5)
        
        ttk.Label(f_font1, text="(宋体支持中文，其他字体仅支持英文)", font=("Segoe UI", 8), foreground="gray").pack(side='left', padx=(20, 5))

        f_note = ttk.LabelFrame(tab, text="说明", style="Section.TLabelframe")
        f_note.pack(fill='x', pady=6, padx=5)
        ttk.Label(f_note, text="程序将自动从PDF文件名中解析面单规格（格式：SKU简称-X单Y个），支持批量处理多个PDF文件", font=("Segoe UI", 9)).pack(padx=6, pady=4)

        f_run = ttk.Frame(tab)
        f_run.pack(fill='x', pady=10, padx=5)

        ttk.Button(f_run, text="[8] 写入页脚并导出PDF", command=self.run_tool8, style='Accent.TButton').pack(side='left', padx=5)
        self.logger8, clear_log8 = self.create_log_widget(tab)
        ttk.Button(f_run, text="清空日志", command=clear_log8, style='Secondary.TButton').pack(side='left', padx=5)

    def _select_pdf8_input(self):
        paths = filedialog.askopenfilenames(
            title="选择面单PDF文件(可多选)",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")],
        )
        if paths:
            self.pdf8_selected_files = list(paths)
            if len(paths) == 1:
                self.pdf8_input_var.set(f"已选择 1 个PDF文件")
            else:
                self.pdf8_input_var.set(f"已选择 {len(paths)} 个PDF文件")
            
            # 如果没有选择输出目录，且不是覆盖模式，自动设置输出目录
            if not self.pdf8_overwrite_var.get() and (self.pdf8_output_var.get() == "未选择输出目录"):
                first_dir = os.path.dirname(paths[0])
                self.pdf8_output_var.set(first_dir)

    def _select_pdf8_output(self):
        path = filedialog.askdirectory(
            title="选择输出目录",
        )
        if path:
            self.pdf8_output_var.set(path)

    def _clear_pdf8_list(self):
        self.pdf8_selected_files = []
        self.pdf8_input_var.set("未选择PDF")

    def _select_pdf8_map_excel(self):
        path = filedialog.askopenfilename(
            title="选择SKU映射Excel",
            filetypes=[("Excel文件", "*.xlsx;*.xlsm;*.xls"), ("All Files", "*.*")],
        )
        if path:
            self.pdf8_map_excel_var.set(path)
            # 自动加载sheet名称
            self._load_sheet_names8(path)

    def _load_sheet_names8(self, excel_path):
        """加载Excel文件的sheet名称到下拉框"""
        try:
            from excel_toolkit.excel_lite import ExcelReader, ExcelWriter
            wb = ExcelReader(excel_path, read_only=True)
            sheet_names = wb.sheetnames
            wb.close()
            
            self.pdf8_sheet_names = sheet_names
            self.pdf8_sheet_combo['values'] = sheet_names
            
            # 如果当前选择的sheet不在列表中，清空选择
            if self.pdf8_map_sheet_var.get() not in sheet_names:
                self.pdf8_map_sheet_var.set("")
                # 如果有sheet，默认选择第一个
                if sheet_names:
                    self.pdf8_map_sheet_var.set(sheet_names[0])
                    
        except Exception as e:
            self.pdf8_sheet_names = []
            self.pdf8_sheet_combo['values'] = []
            self.pdf8_map_sheet_var.set("")
            
    def _show_sku_mapping8(self):
        mapping_excel = self.pdf8_map_excel_var.get()
        if not mapping_excel or mapping_excel == "未选择SKU映射Excel":
            messagebox.showwarning("⚠️ 警告", "请先选择SKU映射Excel文件。")
            return

        sheet_name = self.pdf8_map_sheet_var.get().strip() or None
        sku_short_col = self.pdf8_short_col_var.get().strip() or None
        sku_full_col = self.pdf8_full_col_var.get().strip() or None

        try:
            mapping = load_sku_full_name_map_from_excel(
                excel_path=mapping_excel,
                sheet_name=sheet_name,
                sku_short_col=sku_short_col,
                sku_full_col=sku_full_col,
            )
            
            if not mapping:
                messagebox.showinfo("ℹ️ 信息", "未找到任何SKU映射关系。")
                return

            # 创建新窗口显示映射关系
            win = tk.Toplevel(self.master)
            win.title("SKU 映射关系")
            win.geometry("600x400")
            
            # 创建表格
            tree = ttk.Treeview(win, columns=("short", "full"), show="headings")
            tree.heading("short", text="SKU简称")
            tree.heading("full", text="SKU全称")
            tree.column("short", width=150)
            tree.column("full", width=400)
            
            # 添加滚动条
            scrollbar = ttk.Scrollbar(win, orient="vertical", command=tree.yview)
            tree.configure(yscrollcommand=scrollbar.set)
            
            tree.pack(side="left", fill="both", expand=True, padx=(10, 0), pady=10)
            scrollbar.pack(side="right", fill="y", padx=(0, 10), pady=10)
            
            # 填充数据
            for short, full in sorted(mapping.items()):
                tree.insert("", "end", values=(short, full))
            
            # 状态栏
            status_label = ttk.Label(win, text=f"共 {len(mapping)} 条映射关系", font=("Segoe UI", 9))
            status_label.pack(side="bottom", pady=5)
            
        except Exception as e:
            messagebox.showerror("❌ 错误", f"读取SKU映射失败: {e}")

    def _toggle_overwrite8(self):
        if self.pdf8_overwrite_var.get():
            self.pdf8_output_var.set("（覆盖模式，将直接修改原PDF）")
        else:
            self.pdf8_output_var.set("未选择输出目录")

    def run_tool8(self):
        if not self.pdf8_selected_files:
            messagebox.showwarning("⚠️ 警告", "请先选择PDF文件。")
            return
            
        output_dir = self.pdf8_output_var.get()
        mapping_excel = self.pdf8_map_excel_var.get()
        overwrite = self.pdf8_overwrite_var.get()

        if not overwrite and (not output_dir or output_dir == "未选择输出目录"):
            messagebox.showwarning("⚠️ 警告", "请先选择输出目录或开启覆盖模式。")
            return
        if not mapping_excel or mapping_excel == "未选择SKU映射Excel":
            messagebox.showwarning("⚠️ 警告", "请先选择SKU映射Excel。")
            return

        sheet_name = self.pdf8_map_sheet_var.get().strip() or None
        sku_short_col = self.pdf8_short_col_var.get().strip() or None
        sku_full_col = self.pdf8_full_col_var.get().strip() or None
        
        # 获取字体和字号设置
        font_name = self.pdf8_font_var.get() or "STSong-Light"
        try:
            font_size = int(self.pdf8_fontsize_var.get() or "10")
        except ValueError:
            font_size = 10

        self.logger8("=" * 50)
        self.logger8("▶️ 开始运行 [8] 面单PDF页脚...")
        self.logger8(f"输入文件数量: {len(self.pdf8_selected_files)}")
        if overwrite:
            self.logger8("模式: 覆盖原文件")
        else:
            self.logger8(f"输出目录: {output_dir}")
        self.logger8(f"映射Excel: {mapping_excel}")
        self.logger8(f"字体: {font_name}, 字号: {font_size}")

        self._update_status("正在写入页脚...", icon="⏳", show_progress=True)
        self.master.config(cursor="watch")

        def thread_target():
            try:
                def safe_logger(msg):
                    self.master.after(0, lambda m=msg: self.logger8(m))
                safe_logger("线程已启动，开始处理...")
                
                # 加载SKU映射（一次性加载，提高效率）
                safe_logger("正在加载SKU映射表...")
                sku_mapping = load_sku_full_name_map_from_excel(
                    excel_path=mapping_excel,
                    sheet_name=sheet_name,
                    sku_short_col=sku_short_col,
                    sku_full_col=sku_full_col,
                )
                safe_logger(f"已加载 {len(sku_mapping)} 条SKU映射")
                
                success_count = 0
                error_count = 0
                
                for idx, input_pdf in enumerate(self.pdf8_selected_files, 1):
                    try:
                        filename = os.path.basename(input_pdf)
                        name_without_ext = os.path.splitext(filename)[0]
                        
                        # 解析面单规格
                        import re
                        m = re.match(r"^(.+?)-(\d+)\s*单\s*(\d+)\s*个$", name_without_ext) or re.match(r"^(.+?)-(\d+)单(\d+)个$", name_without_ext)
                        if not m:
                            safe_logger(f"⚠️ 跳过文件: {filename} (无法解析面单规格)")
                            error_count += 1
                            continue
                        
                        sku_short = m.group(1).strip()
                        x = int(m.group(2))
                        y = int(m.group(3))
                        label_spec = f"{sku_short}-{x}单{y}个"
                        
                        # 设置输出路径
                        if overwrite:
                            output_pdf = input_pdf
                        else:
                            base, ext = os.path.splitext(filename)
                            output_pdf = os.path.join(output_dir, f"{base}_footer{ext}")
                        
                        safe_logger(f"[{idx}/{len(self.pdf8_selected_files)}] 处理: {filename}")
                        safe_logger(f"  解析面单规格: {label_spec}")
                        
                        # 处理PDF
                        result = add_pdf_footer_for_shipping_label(
                            input_pdf=input_pdf,
                            output_pdf=output_pdf,
                            label_spec=label_spec,
                            sku_full_name_map=sku_mapping,
                            hide_multiplier_if_one=True,
                            font_name=font_name,
                            font_size=font_size,
                            logger_func=safe_logger,
                        )
                        
                        safe_logger(f"✅ 完成: {os.path.basename(output_pdf)}")
                        success_count += 1
                        
                    except Exception as e:
                        safe_logger(f"❌ 处理失败: {filename} - {e}")
                        error_count += 1
                
                summary = f"批量处理完成！成功: {success_count}, 失败: {error_count}"
                safe_logger("=" * 50)
                safe_logger(summary)

                def on_success():
                    self.master.config(cursor="")
                    self._update_status("就绪", icon="✅", show_progress=False)
                    messagebox.showinfo("✅ 批量处理完成", summary)

                self.master.after(0, on_success)

            except Exception as e:
                error_msg = str(e)
                def on_error():
                    self.master.config(cursor="")
                    self._update_status("错误", icon="❌", show_progress=False)
                    messagebox.showerror("❌ 错误", error_msg)
                    self.logger8(f"❌ 发生错误: {error_msg}")

                self.master.after(0, on_error)

        threading.Thread(target=thread_target, daemon=True).start()
