"""
Tab6 - 拆分订单PDF功能

注意：此模块包含基本的UI创建和run_tool6函数。
复杂的辅助函数（如open_bbox_selector、test_ocr_regions等）
暂时保留在主app.py中，将在后续版本中迁移。
"""
import json
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os

from excel_toolkit.tooltip import create_tooltip
from excel_toolkit.pdf_ocr import split_pdf_with_ocr, render_page_to_image, ocr_order_number
from excel_toolkit.db_config import get_db_manager
from excel_toolkit.db_operations import save_ocr_template, load_ocr_template
from PIL import ImageTk


class Tab6PdfMixin:
    """Tab6 PDF拆分OCR Mixin
    
    注意：此Mixin需要以下方法在主类中定义：
    - _select_pdf_file()
    - _select_outdir()
    - _on_template_select()
    - _save_region_template()
    - _load_region_template()
    - _auto_load_region_templates()
    - test_ocr_regions()
    - open_bbox_selector()
    """
    
    def create_tab6_pdf(self, tab):
        """创建Tab6界面"""
        # 检查变量是否已经在_initialize_all_variables中创建
        if not hasattr(self, 'pdf_input_var'):
            self.pdf_input_var = tk.StringVar(value="未选择PDF")
            self.pdf_outdir_var = tk.StringVar(value="未选择输出目录")
            self.pdf_bbox_x = tk.StringVar(value="100")
            self.pdf_bbox_y = tk.StringVar(value="200")
            self.pdf_bbox_w = tk.StringVar(value="800")
            self.pdf_bbox_h = tk.StringVar(value="200")
            
            # UniUni 模式：增加第二区域
            self.uniuni_mode_var = tk.BooleanVar(value=False)
            self.pdf_bbox2_x = tk.StringVar(value="120")
            self.pdf_bbox2_y = tk.StringVar(value="220")
            self.pdf_bbox2_w = tk.StringVar(value="800")
            self.pdf_bbox2_h = tk.StringVar(value="200")
            
            # 三区域模式：添加第三区域(GOFO)
            self.three_region_mode_var = tk.BooleanVar(value=False)
            self.pdf_bbox3_x = tk.StringVar(value="100")
            self.pdf_bbox3_y = tk.StringVar(value="300")
            self.pdf_bbox3_w = tk.StringVar(value="800")
            self.pdf_bbox3_h = tk.StringVar(value="200")
            
            self.pdf_dpi_var = tk.StringVar(value="300")
            self.poppler_var = tk.StringVar(value="")
            self.tesseract_var = tk.StringVar(value="")
            self.regex_var = tk.StringVar(value="[A-Za-z0-9#-]{6,32}")
            self.prefix_var = tk.StringVar(value="")
            self.ocr_engine_var = tk.StringVar(value="tesseract")

            # 持久化变量
            for v in [self.pdf_input_var, self.pdf_outdir_var, 
                      self.pdf_bbox_x, self.pdf_bbox_y, self.pdf_bbox_w, self.pdf_bbox_h,
                      self.uniuni_mode_var, 
                      self.pdf_bbox2_x, self.pdf_bbox2_y, self.pdf_bbox2_w, self.pdf_bbox2_h,
                      self.three_region_mode_var,
                      self.pdf_bbox3_x, self.pdf_bbox3_y, self.pdf_bbox3_w, self.pdf_bbox3_h,
                      self.pdf_dpi_var, self.poppler_var, self.tesseract_var, 
                      self.regex_var, self.prefix_var, self.ocr_engine_var]:
                if v:
                    self._trace_persist(v)
        
        # template_choice_var需要单独处理，因为它在后面才创建
        if not hasattr(self, 'template_choice_var'):
            self.template_choice_var = tk.StringVar(value="请选择")
            self._trace_persist(self.template_choice_var)
        
        # 自动加载上次使用的模板
        if hasattr(self, '_auto_load_region_templates'):
            self._auto_load_region_templates()

        # 文件选择
        f1 = ttk.Frame(tab)
        f1.pack(fill='x', pady=5)
        ttk.Button(f1, text="选择合并PDF", 
                  command=self._select_pdf_file).pack(side='left', padx=5)
        ttk.Label(f1, textvariable=self.pdf_input_var).pack(side='left', padx=5)

        # 模板快速选择
        f_template = ttk.Frame(tab)
        f_template.pack(fill='x', pady=5)
        ttk.Label(f_template, text="📋 快速套用模板:").pack(side='left', padx=5)
        # template_choice_var已经在上面创建了，这里不再重复创建
        template_combo = ttk.Combobox(f_template, textvariable=self.template_choice_var, 
                                     state="readonly", width=20,
                                     values=["请选择", "USPS模板", "GOFO模板", "Uni模板", "三区域模式"])
        template_combo.pack(side='left', padx=5)
        template_combo.bind("<<ComboboxSelected>>", self._on_template_select)
        create_tooltip(template_combo, "快速套用预设的面单模板坐标")

        # 输出目录
        f2 = ttk.Frame(tab)
        f2.pack(fill='x', pady=5)
        ttk.Button(f2, text="选择输出目录", 
                  command=self._select_outdir).pack(side='left', padx=5)
        ttk.Label(f2, textvariable=self.pdf_outdir_var).pack(side='left', padx=5)

        # 第一区域 - USPS
        f_bbox = ttk.LabelFrame(tab, text="第一区域 - USPS区域 (像素，左上角为原点)", 
                               style="Section.TLabelframe")
        f_bbox.pack(fill='x', pady=5)
        ttk.Label(f_bbox, text="x").pack(side='left', padx=(8, 2))
        ttk.Entry(f_bbox, textvariable=self.pdf_bbox_x, width=8).pack(side='left')
        ttk.Label(f_bbox, text="y").pack(side='left', padx=(8, 2))
        ttk.Entry(f_bbox, textvariable=self.pdf_bbox_y, width=8).pack(side='left')
        ttk.Label(f_bbox, text="width").pack(side='left', padx=(8, 2))
        ttk.Entry(f_bbox, textvariable=self.pdf_bbox_w, width=8).pack(side='left')
        ttk.Label(f_bbox, text="height").pack(side='left', padx=(8, 2))
        ttk.Entry(f_bbox, textvariable=self.pdf_bbox_h, width=8).pack(side='left')
        ttk.Button(f_bbox, text="💾 保存", 
                  command=lambda: self._save_region_template(1)).pack(side='left', padx=12)
        ttk.Button(f_bbox, text="📂 加载", 
                  command=lambda: self._load_region_template(1)).pack(side='left', padx=5)

        # 第二区域 - Uni
        f_bbox2 = ttk.LabelFrame(tab, text="第二区域 - Uni区域 (三区域模式必填)", 
                                style="Section.TLabelframe")
        f_bbox2.pack(fill='x', pady=5)
        ttk.Label(f_bbox2, text="x").pack(side='left', padx=(8, 2))
        ttk.Entry(f_bbox2, textvariable=self.pdf_bbox2_x, width=8).pack(side='left')
        ttk.Label(f_bbox2, text="y").pack(side='left', padx=(8, 2))
        ttk.Entry(f_bbox2, textvariable=self.pdf_bbox2_y, width=8).pack(side='left')
        ttk.Label(f_bbox2, text="width").pack(side='left', padx=(8, 2))
        ttk.Entry(f_bbox2, textvariable=self.pdf_bbox2_w, width=8).pack(side='left')
        ttk.Label(f_bbox2, text="height").pack(side='left', padx=(8, 2))
        ttk.Entry(f_bbox2, textvariable=self.pdf_bbox2_h, width=8).pack(side='left')
        ttk.Button(f_bbox2, text="💾 保存", 
                  command=lambda: self._save_region_template(2)).pack(side='left', padx=12)
        ttk.Button(f_bbox2, text="📂 加载", 
                  command=lambda: self._load_region_template(2)).pack(side='left', padx=5)

        # 第三区域 - GOFO
        f_bbox3 = ttk.LabelFrame(tab, text="第三区域 - GOFO区域 (三区域模式必填)", 
                                style="Section.TLabelframe")
        f_bbox3.pack(fill='x', pady=5)
        ttk.Label(f_bbox3, text="x").pack(side='left', padx=(8, 2))
        ttk.Entry(f_bbox3, textvariable=self.pdf_bbox3_x, width=8).pack(side='left')
        ttk.Label(f_bbox3, text="y").pack(side='left', padx=(8, 2))
        ttk.Entry(f_bbox3, textvariable=self.pdf_bbox3_y, width=8).pack(side='left')
        ttk.Label(f_bbox3, text="width").pack(side='left', padx=(8, 2))
        ttk.Entry(f_bbox3, textvariable=self.pdf_bbox3_w, width=8).pack(side='left')
        ttk.Label(f_bbox3, text="height").pack(side='left', padx=(8, 2))
        ttk.Entry(f_bbox3, textvariable=self.pdf_bbox3_h, width=8).pack(side='left')
        ttk.Button(f_bbox3, text="💾 保存", 
                  command=lambda: self._save_region_template(3)).pack(side='left', padx=12)
        ttk.Button(f_bbox3, text="📂 加载", 
                  command=lambda: self._load_region_template(3)).pack(side='left', padx=5)

        # OCR设置
        f_opts = ttk.LabelFrame(tab, text="渲染与OCR设置", style="Section.TLabelframe")
        f_opts.pack(fill='x', pady=5)
        ttk.Label(f_opts, text="DPI").pack(side='left', padx=(8, 2))
        ttk.Entry(f_opts, textvariable=self.pdf_dpi_var, width=8).pack(side='left')
        ttk.Label(f_opts, text="OCR引擎").pack(side='left', padx=(12, 2))
        engine_combo = ttk.Combobox(f_opts, textvariable=self.ocr_engine_var, 
                                   values=["tesseract", "umi", "paddle", "rapid"], 
                                   state="readonly", width=10)
        engine_combo.pack(side='left')
        create_tooltip(engine_combo, "rapid: 轻量ONNX推理(~70MB)\\npaddle: PaddlePaddle框架(~600MB)\\numi: 调用Umi-OCR服务\\ntesseract: 传统OCR")
        
        ttk.Label(f_opts, text="提取正则").pack(side='left', padx=(12, 2))
        ttk.Entry(f_opts, textvariable=self.regex_var, width=22).pack(side='left')
        
        def on_three_region_toggle():
            if self.three_region_mode_var.get() and hasattr(self, '_auto_load_region_templates'):
                self._auto_load_region_templates()
        
        ttk.Checkbutton(f_opts, text="启用 UniUni 模式", 
                       variable=self.uniuni_mode_var).pack(side='left', padx=(12, 2))
        ttk.Checkbutton(f_opts, text="✨ 三区域智能识别", 
                       variable=self.three_region_mode_var, 
                       command=on_three_region_toggle).pack(side='left', padx=(12, 2))
        ttk.Label(f_opts, text="文件前缀").pack(side='left', padx=(12, 2))
        ttk.Entry(f_opts, textvariable=self.prefix_var, width=16).pack(side='left')

        # 执行按钮
        f_run = ttk.Frame(tab)
        f_run.pack(fill='x', pady=10)
        ttk.Button(f_run, text="[6] 开始拆分并命名", command=self.run_tool6, 
                  style='Accent.TButton').pack(side='left', padx=5)
        ttk.Button(f_run, text="🔍 测试OCR三区域", command=self.test_ocr_regions, 
                  style='Secondary.TButton').pack(side='left', padx=5)
        self.logger6, clear_log6 = self.create_log_widget(tab)
        ttk.Button(f_run, text="清空日志", command=clear_log6, 
                  style='Secondary.TButton').pack(side='left', padx=5)
        ttk.Button(f_run, text="预览并选择区域", 
                  command=self.open_bbox_selector).pack(side='left', padx=12)

    def run_tool6(self):
        """执行PDF拆分和OCR命名"""
        input_pdf = self.pdf_input_var.get()
        outdir = self.pdf_outdir_var.get()
        
        try:
            x = int(self.pdf_bbox_x.get())
            y = int(self.pdf_bbox_y.get())
            w = int(self.pdf_bbox_w.get())
            h = int(self.pdf_bbox_h.get())
            dpi = int(self.pdf_dpi_var.get())
        except Exception:
            messagebox.showwarning("警告", "请填写正确的 bbox 坐标与 DPI（整数）。")
            return
        
        # 解析第二区域
        bbox2 = None
        bbox3 = None
        if self.uniuni_mode_var.get() or self.three_region_mode_var.get():
            try:
                x2 = int(self.pdf_bbox2_x.get())
                y2 = int(self.pdf_bbox2_y.get())
                w2 = int(self.pdf_bbox2_w.get())
                h2 = int(self.pdf_bbox2_h.get())
                if w2 <= 0 or h2 <= 0:
                    raise ValueError("width/height must be > 0")
                bbox2 = (x2, y2, w2, h2)
            except Exception:
                mode_name = "三区域模式" if self.three_region_mode_var.get() else "UniUni 模式"
                messagebox.showwarning("警告", f"已启用 {mode_name}，但第二区域坐标无效。")
                return
        
        # 解析第三区域
        if self.three_region_mode_var.get():
            try:
                x3 = int(self.pdf_bbox3_x.get())
                y3 = int(self.pdf_bbox3_y.get())
                w3 = int(self.pdf_bbox3_w.get())
                h3 = int(self.pdf_bbox3_h.get())
                if w3 <= 0 or h3 <= 0:
                    raise ValueError("width/height must be > 0")
                bbox3 = (x3, y3, w3, h3)
            except Exception:
                messagebox.showwarning("警告", "已启用三区域模式，但第三区域坐标无效。")
                return
                
        if not input_pdf or input_pdf == "未选择PDF":
            messagebox.showwarning("警告", "请先选择合并的订单PDF文件。")
            return
        if not os.path.exists(input_pdf):
            path = filedialog.askopenfilename(
                title="选择订单PDF", 
                filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")]
            )
            if not path:
                messagebox.showerror("读取错误", f"文件不存在：{input_pdf}")
                return
            self.pdf_input_var.set(path)
            input_pdf = path
        if not outdir or outdir == "未选择输出目录":
            messagebox.showwarning("警告", "请先选择输出目录。")
            return

        self.logger6("----------------------------------")
        if self.three_region_mode_var.get():
            self.logger6("开始运行 [6] 三区域智能识别模式...")
        else:
            self.logger6("开始运行 [6] 拆分订单PDF并OCR命名...")
        
        # 创建进度窗口
        progress_win = self._create_progress_window()
        
        import threading
        def run_in_thread():
            try:
                def progress_callback(current, total, status_text):
                    """进度回调函数"""
                    self.master.after(0, lambda: self._update_progress(
                        progress_win, current, total, status_text
                    ))
                
                msg = split_pdf_with_ocr(
                    input_pdf=input_pdf,
                    out_dir=outdir,
                    bbox=(x, y, w, h),
                    bbox2=bbox2,
                    bbox3=bbox3,
                    uniuni_mode=self.uniuni_mode_var.get(),
                    three_region_mode=self.three_region_mode_var.get(),
                    dpi=dpi,
                    poppler_path=self.poppler_var.get() or None,
                    tesseract_cmd=self.tesseract_var.get() or None,
                    regex=self.regex_var.get() or None,
                    prefix=self.prefix_var.get() or "",
                    logger_func=self.logger6,
                    ocr_engine=self.ocr_engine_var.get(),
                    progress_callback=progress_callback
                )
                
                def on_success():
                    self._close_progress_window(progress_win)
                    self.master.config(cursor="")
                    self.status_var.set("就绪")
                    messagebox.showinfo("完成", msg)
                    self.logger6(msg)
                
                self.master.after(0, on_success)
                
            except Exception as e:
                error_msg = str(e)
                def on_error(msg=error_msg):
                    self._close_progress_window(progress_win)
                    self.master.config(cursor="")
                    self.status_var.set("就绪")
                    messagebox.showerror("发生错误", msg)
                    self.logger6(f"发生错误: {msg}")
                
                self.master.after(0, on_error)
        
        self.status_var.set("处理中...")
        self.master.config(cursor="watch")
        threading.Thread(target=run_in_thread, daemon=True).start()
    
    def _create_progress_window(self):
        """创建进度窗口"""
        win = tk.Toplevel(self.master)
        win.title("PDF拆分进度")
        win.geometry("500x250")
        win.transient(self.master)
        win.grab_set()
        
        # 主框架
        main_frame = ttk.Frame(win, padding=20)
        main_frame.pack(fill='both', expand=True)
        
        # 标题
        title_label = ttk.Label(
            main_frame, 
            text="⚡ 正在拆分PDF文件...",
            font=("Microsoft YaHei UI", 14, "bold")
        )
        title_label.pack(pady=(0, 15))
        
        # 状态文本
        status_label = ttk.Label(
            main_frame,
            text="正在初始化...",
            font=("Microsoft YaHei UI", 10)
        )
        status_label.pack(pady=(0, 10))
        
        # 进度条
        progress = ttk.Progressbar(
            main_frame,
            mode='determinate',
            length=450
        )
        progress.pack(pady=(0, 10))
        
        # 进度文本
        progress_label = ttk.Label(
            main_frame,
            text="0%",
            font=("Microsoft YaHei UI", 12, "bold")
        )
        progress_label.pack(pady=(0, 15))
        
        # 详细信息
        detail_label = ttk.Label(
            main_frame,
            text="",
            font=("Consolas", 9),
            foreground="#6B7280"
        )
        detail_label.pack()
        
        # 保存组件引用
        win.status_label = status_label
        win.progress = progress
        win.progress_label = progress_label
        win.detail_label = detail_label
        win.title_label = title_label
        
        # 居中显示
        win.update_idletasks()
        x = (win.winfo_screenwidth() // 2) - (500 // 2)
        y = (win.winfo_screenheight() // 2) - (250 // 2)
        win.geometry(f"500x250+{x}+{y}")
        
        return win
    
    def _update_progress(self, win, current, total, status_text):
        """更新进度窗口"""
        try:
            if not win or not win.winfo_exists():
                return
            
            # 更新进度条
            percentage = int((current / total) * 100) if total > 0 else 0
            win.progress['value'] = percentage
            
            # 更新进度文本
            win.progress_label.configure(text=f"{percentage}%")
            
            # 更新状态文本
            win.status_label.configure(text=status_text)
            
            # 更新详细信息
            win.detail_label.configure(text=f"已处理: {current}/{total} 页")
            
            # 更新标题动画
            dots = ['⚡', '⚡⚡', '⚡⚡⚡', '⚡⚡⚡⚡']
            dot_index = (current % len(dots))
            win.title_label.configure(text=f"{dots[dot_index]} 正在拆分PDF文件...")
            
            win.update_idletasks()
        except Exception:
            pass
    
    def _close_progress_window(self, win):
        """关闭进度窗口"""
        try:
            if win and win.winfo_exists():
                win.grab_release()
                win.destroy()
        except Exception:
            pass

    # ========== Tab6 辅助方法 ==========
    
    def _select_pdf_file(self):
        """选择合并PDF文件"""
        initial_dir = getattr(self, 't6_last_pdf_dir', None) or os.path.dirname(self.pdf_input_var.get()) if self.pdf_input_var.get() != "未选择PDF" else None
        path = filedialog.askopenfilename(
            title="选择合并订单PDF", 
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")],
            initialdir=initial_dir
        )
        if path:
            self.pdf_input_var.set(path)
            self.t6_last_pdf_dir = os.path.dirname(path)
            self._persist_config()

    def _select_outdir(self):
        """选择输出目录"""
        initial_dir = getattr(self, 't6_last_outdir', None) or (self.pdf_outdir_var.get() if self.pdf_outdir_var.get() != "未选择输出目录" else None)
        path = filedialog.askdirectory(
            title="选择输出目录",
            initialdir=initial_dir
        )
        if path:
            self.pdf_outdir_var.set(path)
            self.t6_last_outdir = path
            self._persist_config()

    def _on_template_select(self, event=None):
        """处理模板选择事件"""
        choice = self.template_choice_var.get()
        
        if choice == "请选择":
            return
        
        if choice == "三区域模式":
            self.three_region_mode_var.set(True)
            self.uniuni_mode_var.set(False)
            
            loaded_any = False
            for region_num in [1, 2, 3]:
                last_tpl = getattr(self, f't6_region{region_num}_template', None)
                if last_tpl and os.path.exists(last_tpl):
                    try:
                        with open(last_tpl, 'r', encoding='utf-8') as f:
                            template_data = json.load(f)
                        bbox = template_data.get("bbox", {})
                        if bbox:
                            if region_num == 1:
                                self.pdf_bbox_x.set(str(bbox.get("x", "100")))
                                self.pdf_bbox_y.set(str(bbox.get("y", "200")))
                                self.pdf_bbox_w.set(str(bbox.get("width", "800")))
                                self.pdf_bbox_h.set(str(bbox.get("height", "100")))
                            elif region_num == 2:
                                self.pdf_bbox2_x.set(str(bbox.get("x", "120")))
                                self.pdf_bbox2_y.set(str(bbox.get("y", "220")))
                                self.pdf_bbox2_w.set(str(bbox.get("width", "800")))
                                self.pdf_bbox2_h.set(str(bbox.get("height", "100")))
                            elif region_num == 3:
                                self.pdf_bbox3_x.set(str(bbox.get("x", "100")))
                                self.pdf_bbox3_y.set(str(bbox.get("y", "300")))
                                self.pdf_bbox3_w.set(str(bbox.get("width", "800")))
                                self.pdf_bbox3_h.set(str(bbox.get("height", "100")))
                            loaded_any = True
                            self.logger6(f"✓ 自动加载区域{region_num}模板: {os.path.basename(last_tpl)}")
                    except Exception:
                        pass
            
            if loaded_any:
                self.logger6("✓ 已启用三区域智能识别模式，并加载上次保存的模板")
            else:
                self.logger6("✓ 已启用三区域智能识别模式（使用默认坐标，请手动配置或加载模板）")
            return
        
        templates = {
            "USPS模板": {"bbox1": {"x": 100, "y": 200, "w": 800, "h": 100}, "uniuni": False, "three_region": False, "name": "USPS"},
            "GOFO模板": {"bbox1": {"x": 100, "y": 300, "w": 800, "h": 100}, "uniuni": False, "three_region": False, "name": "GOFO"},
            "Uni模板": {"bbox1": {"x": 120, "y": 220, "w": 800, "h": 100}, "uniuni": False, "three_region": False, "name": "Uni"},
        }
        
        if choice not in templates:
            return
        
        template = templates[choice]
        self.pdf_bbox_x.set(str(template["bbox1"]["x"]))
        self.pdf_bbox_y.set(str(template["bbox1"]["y"]))
        self.pdf_bbox_w.set(str(template["bbox1"]["w"]))
        self.pdf_bbox_h.set(str(template["bbox1"]["h"]))
        self.uniuni_mode_var.set(template["uniuni"])
        self.three_region_mode_var.set(template["three_region"])
        self.logger6(f"✓ 已套用【{template['name']}】模板坐标（默认值，请使用'预览并选择区域'调整）")

    def _save_region_template(self, region_num):
        """保存单个区域的模板（支持数据库和文件）"""
        region_names = {1: "USPS区域", 2: "Uni区域", 3: "GOFO区域"}
        region_name = region_names.get(region_num, f"区域{region_num}")
        
        try:
            if region_num == 1:
                bbox = {
                    "x": int(self.pdf_bbox_x.get()),
                    "y": int(self.pdf_bbox_y.get()),
                    "width": int(self.pdf_bbox_w.get()),
                    "height": int(self.pdf_bbox_h.get())
                }
            elif region_num == 2:
                bbox = {
                    "x": int(self.pdf_bbox2_x.get()),
                    "y": int(self.pdf_bbox2_y.get()),
                    "width": int(self.pdf_bbox2_w.get()),
                    "height": int(self.pdf_bbox2_h.get())
                }
            elif region_num == 3:
                bbox = {
                    "x": int(self.pdf_bbox3_x.get()),
                    "y": int(self.pdf_bbox3_y.get()),
                    "width": int(self.pdf_bbox3_w.get()),
                    "height": int(self.pdf_bbox3_h.get())
                }
            
            db_manager = get_db_manager()
            use_db = db_manager.config.is_enabled()
            
            if use_db:
                try:
                    success, msg = save_ocr_template(
                        name=region_name,
                        region=region_num,
                        bbox=bbox,
                        description=f"{region_name} OCR区域模板"
                    )
                    if success:
                        messagebox.showinfo("成功", f"{region_name}模板已保存到数据库")
                        self.logger6(f"✓ {region_name}模板已保存到数据库")
                        return
                    else:
                        if messagebox.askyesno("数据库保存失败", 
                                              f"保存到数据库失败:\n{msg}\n\n是否保存到本地文件？"):
                            use_db = False
                        else:
                            return
                except Exception as e:
                    if messagebox.askyesno("数据库错误", 
                                          f"数据库操作出错:\n{e}\n\n是否保存到本地文件？"):
                        use_db = False
                    else:
                        return
            
            if not use_db:
                initial_dir = getattr(self, 't6_last_template_dir', None)
                save_path = filedialog.asksaveasfilename(
                    title=f"保存{region_name}模板",
                    defaultextension=".json",
                    filetypes=[("JSON文件", "*.json"), ("所有文件", "*.*")],
                    initialdir=initial_dir,
                    initialfile=f"{region_name}.json"
                )
                
                if not save_path:
                    return
                
                template_data = {
                    "region": region_num,
                    "name": region_name,
                    "bbox": bbox
                }
                
                with open(save_path, 'w', encoding='utf-8') as f:
                    json.dump(template_data, f, indent=2, ensure_ascii=False)
                
                self.t6_last_template_dir = os.path.dirname(save_path)
                self._persist_config()
                
                messagebox.showinfo("成功", f"{region_name}模板已保存到:\n{save_path}")
                self.logger6(f"✓ {region_name}模板已保存: {os.path.basename(save_path)}")
            
        except ValueError as e:
            messagebox.showerror("输入错误", f"坐标必须是整数:\n{e}")
            self.logger6(f"✗ 保存{region_name}模板失败: 坐标格式错误")
        except Exception as e:
            messagebox.showerror("保存失败", f"保存{region_name}模板时出错:\n{e}")
            self.logger6(f"✗ 保存{region_name}模板失败: {e}")

    def _load_region_template(self, region_num):
        """加载单个区域的模板（支持数据库和文件）"""
        region_names = {1: "USPS区域", 2: "Uni区域", 3: "GOFO区域"}
        region_name = region_names.get(region_num, f"区域{region_num}")
        
        try:
            template_data = None
            
            db_manager = get_db_manager()
            if db_manager.config.is_enabled():
                try:
                    template_data = load_ocr_template(name=region_name, region=region_num)
                    if template_data:
                        self.logger6(f"✓ 从数据库加载{region_name}模板")
                except Exception as e:
                    self.logger6(f"⚠️ 从数据库加载失败: {e}，将尝试从文件加载")
            
            if not template_data:
                last_template_attr = f't6_region{region_num}_template'
                last_template = getattr(self, last_template_attr, None)
                
                if last_template and os.path.exists(last_template):
                    initial_file = last_template
                    initial_dir = os.path.dirname(last_template)
                else:
                    initial_file = None
                    initial_dir = getattr(self, 't6_last_template_dir', None)
                
                load_path = filedialog.askopenfilename(
                    title=f"加载{region_name}模板",
                    filetypes=[("JSON文件", "*.json"), ("所有文件", "*.*")],
                    initialdir=initial_dir,
                    initialfile=os.path.basename(initial_file) if initial_file else None
                )
                
                if not load_path:
                    return
                
                self.t6_last_template_dir = os.path.dirname(load_path)
                self._persist_config()
                
                with open(load_path, 'r', encoding='utf-8') as f:
                    template_data = json.load(f)
                    
                setattr(self, f't6_region{region_num}_template', load_path)
                self._persist_config()
            
            if not template_data:
                return
            
            bbox = None
            
            if "bbox" in template_data and "region" in template_data:
                bbox = template_data.get("bbox", {})
                if template_data.get("region") != region_num:
                    response = messagebox.askyesno(
                        "模板不匹配", 
                        f"此模板是为{template_data.get('name', '其他区域')}保存的，"
                        f"您正在尝试加载到{region_name}。\n\n是否继续？"
                    )
                    if not response:
                        return
            elif f"bbox{region_num}" in template_data:
                bbox = template_data.get(f"bbox{region_num}", {})
                self.logger6(f"  检测到旧格式模板，正在提取区域{region_num}的数据...")
            elif any(f"bbox{i}" in template_data for i in [1, 2, 3]):
                messagebox.showwarning(
                    "区域数据缺失", 
                    f"此模板中没有{region_name}的数据。\n\n"
                    f"提示：这可能是旧版本的模板，只包含部分区域。\n"
                    f"您可以手动配置{region_name}的坐标，然后点击'💾 保存'创建新模板。"
                )
                self.logger6(f"✗ 模板中缺少{region_name}数据")
                return
            else:
                messagebox.showerror("格式错误", f"无法识别的模板格式")
                self.logger6(f"✗ 模板格式错误")
                return
            
            if not bbox:
                messagebox.showerror("错误", f"模板中没有找到{region_name}的坐标数据")
                return
            
            self.logger6(f"  读取到的坐标: x={bbox.get('x')}, y={bbox.get('y')}, width={bbox.get('width')}, height={bbox.get('height')}")
            
            if region_num == 1:
                self.pdf_bbox_x.set(str(bbox.get("x", "100")))
                self.pdf_bbox_y.set(str(bbox.get("y", "200")))
                self.pdf_bbox_w.set(str(bbox.get("width", "800")))
                self.pdf_bbox_h.set(str(bbox.get("height", "100")))
                self.logger6(f"  已应用到第一区域: {self.pdf_bbox_x.get()}, {self.pdf_bbox_y.get()}, {self.pdf_bbox_w.get()}, {self.pdf_bbox_h.get()}")
            elif region_num == 2:
                self.pdf_bbox2_x.set(str(bbox.get("x", "120")))
                self.pdf_bbox2_y.set(str(bbox.get("y", "220")))
                self.pdf_bbox2_w.set(str(bbox.get("width", "800")))
                self.pdf_bbox2_h.set(str(bbox.get("height", "100")))
                self.logger6(f"  已应用到第二区域: {self.pdf_bbox2_x.get()}, {self.pdf_bbox2_y.get()}, {self.pdf_bbox2_w.get()}, {self.pdf_bbox2_h.get()}")
            elif region_num == 3:
                self.pdf_bbox3_x.set(str(bbox.get("x", "100")))
                self.pdf_bbox3_y.set(str(bbox.get("y", "300")))
                self.pdf_bbox3_w.set(str(bbox.get("width", "800")))
                self.pdf_bbox3_h.set(str(bbox.get("height", "100")))
                self.logger6(f"  已应用到第三区域: {self.pdf_bbox3_x.get()}, {self.pdf_bbox3_y.get()}, {self.pdf_bbox3_w.get()}, {self.pdf_bbox3_h.get()}")
            
            from_db = db_manager.config.is_enabled() and 'bbox' in template_data
            if from_db:
                messagebox.showinfo("成功", f"{region_name}模板已从数据库加载")
                self.logger6(f"✓ {region_name}模板已从数据库加载")
            else:
                messagebox.showinfo("成功", f"{region_name}模板已加载")
                self.logger6(f"✓ {region_name}模板已加载")
            
        except Exception as e:
            messagebox.showerror("加载失败", f"加载{region_name}模板时出错:\n{e}")
            self.logger6(f"✗ 加载{region_name}模板失败: {e}")

    def _auto_load_region_templates(self):
        """程序启动时自动加载上次使用的三个区域模板"""
        for region_num in [1, 2, 3]:
            region_names = {1: "USPS区域", 2: "Uni区域", 3: "GOFO区域"}
            region_name = region_names.get(region_num, f"区域{region_num}")
            
            last_template_attr = f't6_region{region_num}_template'
            last_template = getattr(self, last_template_attr, None)
            
            if not last_template or not os.path.exists(last_template):
                continue
            
            try:
                with open(last_template, 'r', encoding='utf-8') as f:
                    template_data = json.load(f)
                
                bbox = None
                if "bbox" in template_data and "region" in template_data:
                    bbox = template_data.get("bbox", {})
                elif f"bbox{region_num}" in template_data:
                    bbox = template_data.get(f"bbox{region_num}", {})
                else:
                    continue
                
                if not bbox:
                    continue
                
                if region_num == 1:
                    self.pdf_bbox_x.set(str(bbox.get("x", "100")))
                    self.pdf_bbox_y.set(str(bbox.get("y", "200")))
                    self.pdf_bbox_w.set(str(bbox.get("width", "800")))
                    self.pdf_bbox_h.set(str(bbox.get("height", "100")))
                elif region_num == 2:
                    self.pdf_bbox2_x.set(str(bbox.get("x", "120")))
                    self.pdf_bbox2_y.set(str(bbox.get("y", "220")))
                    self.pdf_bbox2_w.set(str(bbox.get("width", "800")))
                    self.pdf_bbox2_h.set(str(bbox.get("height", "100")))
                elif region_num == 3:
                    self.pdf_bbox3_x.set(str(bbox.get("x", "100")))
                    self.pdf_bbox3_y.set(str(bbox.get("y", "300")))
                    self.pdf_bbox3_w.set(str(bbox.get("width", "800")))
                    self.pdf_bbox3_h.set(str(bbox.get("height", "100")))
                
                if hasattr(self, 'logger6'):
                    self.logger6(f"✓ 自动加载{region_name}模板: {os.path.basename(last_template)}")
                
            except Exception:
                pass

    def test_ocr_regions(self):
        """测试三个OCR区域，输出识别结果到日志"""
        input_pdf = self.pdf_input_var.get()
        
        if not input_pdf or input_pdf == "未选择PDF":
            messagebox.showwarning("警告", "请先选择PDF文件进行测试。")
            return
        
        if not os.path.exists(input_pdf):
            messagebox.showerror("错误", f"文件不存在：{input_pdf}")
            return
        
        try:
            from pypdf import PdfReader
            reader = PdfReader(input_pdf)
            total_pages = len(reader.pages)
        except Exception as e:
            messagebox.showerror("错误", f"无法读取PDF：{e}")
            return
        
        page_dialog = tk.Toplevel(self.master)
        page_dialog.title("选择测试页面")
        page_dialog.geometry("400x180")
        page_dialog.transient(self.master)
        page_dialog.grab_set()
        
        selected_page = tk.IntVar(value=1)
        
        ttk.Label(page_dialog, text=f"PDF共有 {total_pages} 页，请选择要测试的页面：", 
                 font=('', 10)).pack(pady=15)
        
        page_frame = ttk.Frame(page_dialog)
        page_frame.pack(pady=10)
        ttk.Label(page_frame, text="页码：").pack(side='left', padx=5)
        page_spinbox = ttk.Spinbox(page_frame, from_=1, to=total_pages, 
                                   textvariable=selected_page, width=10)
        page_spinbox.pack(side='left', padx=5)
        ttk.Label(page_frame, text=f"（1-{total_pages}）").pack(side='left')
        
        btn_frame = ttk.Frame(page_dialog)
        btn_frame.pack(pady=15)
        
        def on_confirm():
            page_dialog.destroy()
            self._do_test_ocr_regions(input_pdf, selected_page.get())
        
        def on_cancel():
            page_dialog.destroy()
        
        ttk.Button(btn_frame, text="开始测试", command=on_confirm, 
                  style='Accent.TButton').pack(side='left', padx=10)
        ttk.Button(btn_frame, text="取消", command=on_cancel).pack(side='left', padx=10)
        
        page_dialog.protocol("WM_DELETE_WINDOW", on_cancel)

    def _do_test_ocr_regions(self, input_pdf, page_num):
        """执行OCR区域测试"""
        try:
            x1 = int(self.pdf_bbox_x.get()); y1 = int(self.pdf_bbox_y.get())
            w1 = int(self.pdf_bbox_w.get()); h1 = int(self.pdf_bbox_h.get())
            dpi = int(self.pdf_dpi_var.get())
        except Exception:
            messagebox.showwarning("警告", "请填写正确的区域坐标与 DPI（整数）。")
            return
        
        self.logger6("=" * 60)
        self.logger6("🔍 开始测试OCR三区域...")
        self.logger6(f"PDF文件: {os.path.basename(input_pdf)}")
        self.logger6(f"测试页面: 第{page_num}页")
        self.logger6(f"DPI: {dpi}")
        engine = self.ocr_engine_var.get()
        self.logger6(f"OCR引擎: {engine}")
        self.logger6("-" * 60)
        
        try:
            self.logger6(f"正在渲染PDF第{page_num}页...")
            img = render_page_to_image(
                input_pdf, 
                page_num - 1,
                dpi=dpi, 
                poppler_path=self.poppler_var.get() or None
            )
            
            if self.pdf_bbox_w.get() and int(self.pdf_bbox_w.get()) > 0:
                x1 = int(self.pdf_bbox_x.get()); y1 = int(self.pdf_bbox_y.get())
                w1 = int(self.pdf_bbox_w.get()); h1 = int(self.pdf_bbox_h.get())
                
                self.logger6(f"\n【第一区域 - USPS区域】")
                self.logger6(f"  坐标: x={x1}, y={y1}, width={w1}, height={h1}")
                
                cropped1 = img.crop((x1, y1, x1 + w1, y1 + h1))
                result1 = ocr_order_number(
                    cropped1,
                    tesseract_cmd=self.tesseract_var.get() or None,
                    enable_preprocessing=True,
                    engine=engine
                )
                self.logger6(f"  ✓ OCR结果: '{result1}'")
                if result1 and result1[0] == '9':
                    self.logger6("  ✓ 识别为USPS订单（以9开头）")
            
            if self.pdf_bbox2_w.get() and int(self.pdf_bbox2_w.get()) > 0:
                x2 = int(self.pdf_bbox2_x.get()); y2 = int(self.pdf_bbox2_y.get())
                w2 = int(self.pdf_bbox2_w.get()); h2 = int(self.pdf_bbox2_h.get())
                
                self.logger6(f"\n【第二区域 - Uni区域】")
                self.logger6(f"  坐标: x={x2}, y={y2}, width={w2}, height={h2}")
                
                cropped2 = img.crop((x2, y2, x2 + w2, y2 + h2))
                result2 = ocr_order_number(
                    cropped2,
                    tesseract_cmd=self.tesseract_var.get() or None,
                    enable_preprocessing=True,
                    engine=engine
                )
                self.logger6(f"  ✓ OCR结果: '{result2}'")
                if result2 and result2[0] == 'U':
                    self.logger6("  ✓ 识别为UniUni订单（以U开头）")
            
            if self.pdf_bbox3_w.get() and int(self.pdf_bbox3_w.get()) > 0:
                x3 = int(self.pdf_bbox3_x.get()); y3 = int(self.pdf_bbox3_y.get())
                w3 = int(self.pdf_bbox3_w.get()); h3 = int(self.pdf_bbox3_h.get())
                
                self.logger6(f"\n【第三区域 - GOFO区域】")
                self.logger6(f"  坐标: x={x3}, y={y3}, width={w3}, height={h3}")
                
                cropped3 = img.crop((x3, y3, x3 + w3, y3 + h3))
                result3 = ocr_order_number(
                    cropped3,
                    tesseract_cmd=self.tesseract_var.get() or None,
                    enable_preprocessing=True,
                    engine=engine
                )
                if result3 and result3.upper().startswith("GFUS"):
                     result3 = result3.replace('O', '0').replace('o', '0')
                
                self.logger6(f"  ✓ OCR结果: '{result3}'")
                if result3 and result3[0] == 'G':
                    self.logger6(f"  ✓ 识别为GOFO订单（以G开头）")
            
            self.logger6("-" * 60)
            self.logger6("✓ 测试完成！")
            self.logger6("=" * 60)
            
            messagebox.showinfo("测试完成", "OCR测试已完成，请查看日志窗口。")
            
        except Exception as e:
            self.logger6(f"\n✗ 测试失败: {e}")
            messagebox.showerror("测试失败", f"OCR测试时出错：\n{e}")

    def open_bbox_selector(self):
        """打开OCR区域选择器"""
        input_pdf = self.pdf_input_var.get()
        if not input_pdf or input_pdf == "未选择PDF":
            messagebox.showwarning("警告", "请先选择合并的订单PDF文件。")
            return
        if not os.path.exists(input_pdf):
            path = filedialog.askopenfilename(title="选择订单PDF", filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")])
            if not path:
                messagebox.showerror("读取错误", f"文件不存在：{input_pdf}")
                return
            self.pdf_input_var.set(path)
            input_pdf = path
        try:
            from pypdf import PdfReader
        except Exception:
            messagebox.showerror("缺少依赖", "未安装 pypdf。请先执行: pip install pypdf")
            return
        try:
            reader = PdfReader(input_pdf)
            total_pages = len(reader.pages)
        except Exception as e:
            messagebox.showerror("读取错误", f"无法读取PDF: {e}")
            return

        try:
            dpi = int(self.pdf_dpi_var.get())
        except Exception:
            messagebox.showwarning("警告", "DPI 必须为整数。")
            return

        poppler_path = self.poppler_var.get() or None

        win = tk.Toplevel(self.master)
        win.title("选择 OCR 区域 (左上角为原点)")
        win.geometry("900x720")

        ctrl = ttk.Frame(win); ctrl.pack(fill='x', padx=10, pady=6)
        ttk.Label(ctrl, text=f"总页数: {total_pages}").pack(side='left')
        ttk.Label(ctrl, text="  跳转到页:").pack(side='left', padx=(12, 4))
        page_var = tk.IntVar(value=1)
        page_spin = ttk.Spinbox(ctrl, from_=1, to=total_pages, textvariable=page_var, width=6)
        page_spin.pack(side='left')
        apply_btn1 = ttk.Button(ctrl, text="应用选择为第一区域")
        apply_btn1.pack(side='right', padx=6)
        apply_btn2 = ttk.Button(ctrl, text="应用选择为第二区域")
        apply_btn2.pack(side='right', padx=6)
        apply_btn3 = ttk.Button(ctrl, text="应用选择为第三区域")
        apply_btn3.pack(side='right', padx=6)

        tip = ttk.Label(win, text="提示：在图片上拖拽选择矩形区域；松开鼠标后可更新选择。" )
        tip.pack(fill='x', padx=10)

        canvas = tk.Canvas(win, bg="#f5f5f5")
        canvas.pack(fill='both', expand=True, padx=10, pady=10)

        state = {
            "img": None,
            "photo": None,
            "scale": 1.0,
            "base_scale": 1.0,
            "manual_scale": 1.0,
            "start": None,
            "rect": None,
            "bbox_display": None,
            "prev_bbox1": None,
            "prev_bbox2": None,
            "prev_bbox3": None,
            "rect_prev1": None,
            "rect_prev2": None,
            "rect_prev3": None,
            "origin_x": 10,
            "origin_y": 10,
            "pan_start": None,
            "img_item": None,
            "cache": {},
            "rendering": False,
            "pending_idx": None,
        }

        def render_current_page():
            import threading
            idx = page_var.get() - 1
            if idx in state["cache"]:
                img = state["cache"][idx]
                canvas.update_idletasks()
                max_w = max(600, canvas.winfo_width() - 20)
                max_h = max(400, canvas.winfo_height() - 20)
                iw, ih = img.size
                state["img"] = img
                state["base_scale"] = min(max_w / iw, max_h / ih, 1.0)
                state["manual_scale"] = 1.0
                _refresh_display()
                return
            if state["rendering"]:
                state["pending_idx"] = idx
                return
            def _work():
                try:
                    img_local = render_page_to_image(input_pdf, idx, dpi, poppler_path)
                except Exception as e:
                    def _err():
                        messagebox.showerror("渲染错误", f"无法渲染PDF页面: {e}")
                        state["rendering"] = False
                        state["pending_idx"] = None
                    win.after(0, _err)
                    return
                def _done():
                    state["cache"][idx] = img_local
                    state["rendering"] = False
                    if state["pending_idx"] is not None and state["pending_idx"] != idx:
                        pending = state["pending_idx"]
                        state["pending_idx"] = None
                        render_current_page()
                        return
                    canvas.update_idletasks()
                    max_w = max(600, canvas.winfo_width() - 20)
                    max_h = max(400, canvas.winfo_height() - 20)
                    iw, ih = img_local.size
                    state["img"] = img_local
                    state["base_scale"] = min(max_w / iw, max_h / ih, 1.0)
                    state["manual_scale"] = 1.0
                    _refresh_display()
                def _restore_tip():
                    try:
                        tip.configure(text="提示：拖拽选择；滚轮缩放；右键拖拽平移；松开鼠标后可更新选择。")
                    except Exception:
                        pass
                win.after(0, _done)
                win.after(0, _restore_tip)
            state["rendering"] = True
            try:
                tip.configure(text="正在渲染当前页…")
            except Exception:
                pass
            threading.Thread(target=_work, daemon=True).start()

        def _refresh_display():
            if state["img"] is None:
                return
            iw, ih = state["img"].size
            state["manual_scale"] = max(0.25, min(state["manual_scale"], 4.0))
            scale = state["base_scale"] * state["manual_scale"]
            disp_w, disp_h = int(iw * scale), int(ih * scale)
            img_disp = state["img"].resize((disp_w, disp_h))
            state["photo"] = ImageTk.PhotoImage(img_disp)
            state["scale"] = scale
            canvas.delete('all')
            state["img_item"] = canvas.create_image(state["origin_x"], state["origin_y"], anchor='nw', image=state["photo"])
            state["start"] = None
            if state["rect"] is not None:
                canvas.delete(state["rect"])
            state["rect"] = None
            state["bbox_display"] = None
            _draw_prev_bboxes()

        def _draw_prev_bboxes():
            if state["img"] is None:
                return
            if state["rect_prev1"] is not None:
                canvas.delete(state["rect_prev1"]); state["rect_prev1"] = None
            if state["rect_prev2"] is not None:
                canvas.delete(state["rect_prev2"]); state["rect_prev2"] = None
            if state["rect_prev3"] is not None:
                canvas.delete(state["rect_prev3"]); state["rect_prev3"] = None
            scale = state["scale"] or 1.0
            def draw_box(bbox, color):
                x, y, w, h = bbox
                dx = int(x * scale) + state["origin_x"]
                dy = int(y * scale) + state["origin_y"]
                dw, dh = int(w * scale), int(h * scale)
                return canvas.create_rectangle(dx, dy, dx + dw, dy + dh, outline=color, width=2, dash=(6, 4))
            if state["prev_bbox1"]:
                state["rect_prev1"] = draw_box(state["prev_bbox1"], "#0078D7")
            if state["prev_bbox2"]:
                state["rect_prev2"] = draw_box(state["prev_bbox2"], "#D78A00")
            if state["prev_bbox3"]:
                state["rect_prev3"] = draw_box(state["prev_bbox3"], "#00A000")

        def on_press(event):
            state["start"] = (event.x, event.y)
            if state["rect"] is not None:
                canvas.delete(state["rect"]); state["rect"] = None

        def on_drag(event):
            if state["start"] is None:
                return
            x0, y0 = state["start"]
            x1, y1 = event.x, event.y
            left = state["origin_x"]; top = state["origin_y"]
            right = left + int(state["photo"].width()); bottom = top + int(state["photo"].height())
            x0 = max(left, min(x0, right))
            x1 = max(left, min(x1, right))
            y0 = max(top, min(y0, bottom))
            y1 = max(top, min(y1, bottom))
            if state["rect"] is not None:
                canvas.delete(state["rect"]); state["rect"] = None
            state["rect"] = canvas.create_rectangle(x0, y0, x1, y1, outline="#0078D7", width=2)
            x_min, y_min = min(x0, x1) - state["origin_x"], min(y0, y1) - state["origin_y"]
            x_max, y_max = max(x0, x1) - state["origin_x"], max(y0, y1) - state["origin_y"]
            state["bbox_display"] = (x_min, y_min, x_max - x_min, y_max - y_min)

        def on_release(event):
            pass

        def on_pan_press(event):
            state["pan_start"] = (event.x, event.y)

        def on_pan_drag(event):
            if not state["pan_start"]:
                return
            px, py = state["pan_start"]
            dx, dy = event.x - px, event.y - py
            state["origin_x"] += dx
            state["origin_y"] += dy
            state["pan_start"] = (event.x, event.y)
            if state.get("img_item") is not None:
                canvas.move(state["img_item"], dx, dy)
            if state.get("rect") is not None:
                canvas.move(state["rect"], dx, dy)
            if state.get("rect_prev1") is not None:
                canvas.move(state["rect_prev1"], dx, dy)
            if state.get("rect_prev2") is not None:
                canvas.move(state["rect_prev2"], dx, dy)
            if state.get("rect_prev3") is not None:
                canvas.move(state["rect_prev3"], dx, dy)

        def on_pan_release(event):
            state["pan_start"] = None

        def apply_selection_to(target: int):
            if not state["bbox_display"]:
                messagebox.showwarning("提示", "请先拖拽选择一个区域。")
                return
            x, y, w, h = state["bbox_display"]
            scale = state["scale"] or 1.0
            ox = int(round(x / scale))
            oy = int(round(y / scale))
            ow = int(round(w / scale))
            oh = int(round(h / scale))
            if target == 1:
                self.pdf_bbox_x.set(str(ox))
                self.pdf_bbox_y.set(str(oy))
                self.pdf_bbox_w.set(str(ow))
                self.pdf_bbox_h.set(str(oh))
                state["prev_bbox1"] = (ox, oy, ow, oh)
                msg = "已将选择区域应用到第一区域"
            elif target == 2:
                self.pdf_bbox2_x.set(str(ox))
                self.pdf_bbox2_y.set(str(oy))
                self.pdf_bbox2_w.set(str(ow))
                self.pdf_bbox2_h.set(str(oh))
                state["prev_bbox2"] = (ox, oy, ow, oh)
                msg = "已将选择区域应用到第二区域"
            else:
                self.pdf_bbox3_x.set(str(ox))
                self.pdf_bbox3_y.set(str(oy))
                self.pdf_bbox3_w.set(str(ow))
                self.pdf_bbox3_h.set(str(oh))
                state["prev_bbox3"] = (ox, oy, ow, oh)
                msg = "已将选择区域应用到第三区域"
            _draw_prev_bboxes()
            messagebox.showinfo("已应用", f"{msg}: x={ox}, y={oy}, w={ow}, h={oh}")

        def on_zoom(event):
            delta = event.delta if hasattr(event, 'delta') else 0
            if delta > 0:
                state["manual_scale"] *= 1.1
            elif delta < 0:
                state["manual_scale"] *= 0.9
            _refresh_display()

        canvas.bind('<ButtonPress-1>', on_press)
        canvas.bind('<B1-Motion>', on_drag)
        canvas.bind('<ButtonRelease-1>', on_release)
        canvas.bind('<MouseWheel>', on_zoom)
        canvas.bind('<ButtonPress-3>', on_pan_press)
        canvas.bind('<B3-Motion>', on_pan_drag)
        canvas.bind('<ButtonRelease-3>', on_pan_release)
        apply_btn1.configure(command=lambda: apply_selection_to(1))
        apply_btn2.configure(command=lambda: apply_selection_to(2))
        apply_btn3.configure(command=lambda: apply_selection_to(3))

        def on_page_change(*_):
            render_current_page()

        page_var.trace_add('write', on_page_change)
        render_current_page()

        try:
            tip.configure(text="提示：拖拽选择；滚轮缩放；右键拖拽平移；松开鼠标后可更新选择。")
        except Exception:
            pass

        def _load_prev_from_form():
            try:
                ox = int(self.pdf_bbox_x.get()); oy = int(self.pdf_bbox_y.get()); ow = int(self.pdf_bbox_w.get()); oh = int(self.pdf_bbox_h.get())
                if ow > 0 and oh > 0:
                    state["prev_bbox1"] = (ox, oy, ow, oh)
            except Exception:
                state["prev_bbox1"] = None
            try:
                ox2 = int(self.pdf_bbox2_x.get()); oy2 = int(self.pdf_bbox2_y.get()); ow2 = int(self.pdf_bbox2_w.get()); oh2 = int(self.pdf_bbox2_h.get())
                if ow2 > 0 and oh2 > 0:
                    state["prev_bbox2"] = (ox2, oy2, ow2, oh2)
            except Exception:
                state["prev_bbox2"] = None
            try:
                ox3 = int(self.pdf_bbox3_x.get()); oy3 = int(self.pdf_bbox3_y.get()); ow3 = int(self.pdf_bbox3_w.get()); oh3 = int(self.pdf_bbox3_h.get())
                if ow3 > 0 and oh3 > 0:
                    state["prev_bbox3"] = (ox3, oy3, ow3, oh3)
            except Exception:
                state["prev_bbox3"] = None
            _draw_prev_bboxes()

        _load_prev_from_form()


