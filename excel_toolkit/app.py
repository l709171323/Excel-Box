"""
Excel 工具箱 - 重构版应用主类

使用 Mixin 继承模式，将各个 Tab 的 UI 逻辑拆分到独立模块中。
代码量从 4000+ 行减少到约 800 行。

架构说明：
- ToolkitAppRefactored 继承所有 Mixin 类
- 各 Tab 的创建逻辑在 ui/ 目录下的对应模块中
- 本文件保留：初始化、主题、配置管理、通用辅助函数
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
import tkinter.font as tkfont
from PIL import ImageTk
import os
import json

# 导入所有 Mixin 类
from excel_toolkit.ui import (
    LoggerMixin,
    FileSelectMixin,
    get_sheet_names,
    Tab1StatesMixin,
    Tab2SkusMixin,
    Tab3HighlightMixin,
    Tab4InsertMixin,
    Tab5CompareMixin,
    Tab6PdfMixin,
    Tab7PrefixMixin,
    Tab8PdfFooterMixin,
    Tab9RouterMixin,
    Tab10EntryMixin,
    Tab11ShippingMixin,
    Tab12PptMixin,
    Tab13ImageCompressMixin,
    Tab14DeleteColsMixin,
)

# 导入数据库配置
from excel_toolkit.db_config import get_db_manager

# 导入业务模块
from excel_toolkit.tooltip import create_tooltip
from excel_toolkit.warehouse_router import read_inventory


class ToolkitAppRefactored(
    LoggerMixin,
    FileSelectMixin,
    Tab1StatesMixin,
    Tab2SkusMixin,
    Tab3HighlightMixin,
    Tab4InsertMixin,
    Tab5CompareMixin,
    Tab6PdfMixin,
    Tab7PrefixMixin,
    Tab8PdfFooterMixin,
    Tab9RouterMixin,
    Tab10EntryMixin,
    Tab11ShippingMixin,
    Tab12PptMixin,
    Tab13ImageCompressMixin,
    Tab14DeleteColsMixin,
):
    """
    Excel 工具箱主应用类 (重构版)
    
    通过继承 Mixin 类实现模块化，每个 Tab 的逻辑在对应的 Mixin 中实现。
    """
    
    VERSION = "2.3"
    AUTHOR = "果汁梨"
    
    def __init__(self, master):
        self.master = master
        master.title(f"Excel 工具箱 V{self.VERSION} - By {self.AUTHOR}")
        
        # 设置窗口大小和位置
        self._setup_window_geometry()
        
        # 设置主题
        self._setup_style()
        
        # 初始化变量
        self._text_widgets = []
        self.theme_mode_var = tk.StringVar(value="系统")
        self.topmost_var = tk.BooleanVar(value=True)
        self._accent = "#3b82f6"
        
        # 应用初始主题
        self._apply_theme("系统")
        try:
            self.master.attributes("-topmost", self.topmost_var.get())
        except Exception:
            pass
        
        # 创建UI
        self._create_header()
        self._create_tabs()
        self._create_status_bar()
        self._bind_shortcuts()
        
        # 加载配置
        try:
            self._load_config()
        except FileNotFoundError:
            pass
        except json.JSONDecodeError as e:
            print(f"警告：配置文件格式错误，将使用默认配置。错误: {e}")
        except Exception as e:
            print(f"警告：加载配置时发生错误，将使用默认配置。错误: {e}")
    
    # ==================== 窗口设置 ====================
    
    def _setup_window_geometry(self):
        """设置窗口大小和位置"""
        try:
            sw = self.master.winfo_screenwidth()
            sh = self.master.winfo_screenheight()
            w = min(1080, max(960, sw - 80))
            h = min(780, max(720, sh - 120))
            x = max(0, (sw - w) // 2)
            y = max(0, (sh - h) // 2)
            self.master.geometry(f"{w}x{h}+{x}+{y}")
        except Exception:
            self.master.geometry("1080x780")
        try:
            self.master.minsize(960, 720)
        except Exception:
            pass
    
    def _setup_style(self):
        """设置ttk样式"""
        style = ttk.Style()
        names = style.theme_names()
        if "vista" in names:
            style.theme_use("vista")
        elif "clam" in names:
            style.theme_use("clam")
        
        try:
            default_font = tkfont.nametofont("TkDefaultFont")
            text_font = tkfont.nametofont("TkTextFont")
            fixed_font = tkfont.nametofont("TkFixedFont")
            for f in (default_font, text_font, fixed_font):
                f.configure(family="Segoe UI", size=10)
        except Exception:
            pass
    
    # ==================== UI创建 ====================
    
    def _create_header(self):
        """创建顶部标题栏"""
        # 顶部标题栏
        header = ttk.Frame(self.master)
        header.pack(fill='x', padx=10, pady=(10, 0))
        
        # 左侧标题和版本信息
        left_box = ttk.Frame(header)
        left_box.pack(side='left')
        title_lbl = ttk.Label(left_box, text="📊 Excel 工具箱", font=("Segoe UI", 16, "bold"))
        title_lbl.pack(side='left')
        version_lbl = ttk.Label(left_box, text=f"v{self.VERSION}", font=("Segoe UI", 9))
        version_lbl.pack(side='left', padx=(8, 0), pady=(4, 0))
        
        # 右侧控制栏
        right_box = ttk.Frame(header)
        right_box.pack(side='right')
        
        # 主题选择
        theme_frame = ttk.Frame(right_box)
        theme_frame.pack(side='left', padx=(0, 10))
        ttk.Label(theme_frame, text="🎨 主题:").pack(side='left', padx=(0, 4))
        theme_box = ttk.Combobox(theme_frame, state="readonly", width=6,
                                textvariable=self.theme_mode_var,
                                values=["浅色", "深色", "系统"])
        theme_box.pack(side='left')
        
        # 窗口置顶选项
        topmost_check = ttk.Checkbutton(right_box, text="📌 置顶", variable=self.topmost_var)
        topmost_check.pack(side='left')
        
        # 绑定事件
        self.theme_mode_var.trace_add("write", lambda *a: self._apply_theme(self.theme_mode_var.get()))
        self.topmost_var.trace_add("write", lambda *a: self._on_topmost_change())
        
        # 分隔线
        ttk.Separator(self.master, orient='horizontal').pack(fill='x', padx=10, pady=10)
        
        # 重要提示框
        tip_frame = ttk.Frame(self.master)
        tip_frame.pack(fill='x', padx=10, pady=(0, 5))
        
        tip_inner = ttk.Frame(tip_frame)
        tip_inner.pack(fill='x', padx=5, pady=5)
        
        warning_label = ttk.Label(tip_inner, text="⚠️", font=("Segoe UI", 14))
        warning_label.pack(side='left', padx=(5, 10))
        
        tip_text = ttk.Label(tip_inner,
                            text="重要提示：处理文件前请确保已关闭 Excel/WPS，避免保存失败！",
                            font=("Segoe UI", 10))
        tip_text.pack(side='left')
        
        # 帮助和关于按钮
        about_btn = ttk.Button(tip_inner, text="ℹ️ 关于", width=8,
                              command=self._show_about)
        about_btn.pack(side='right', padx=5)
        create_tooltip(about_btn, "查看软件版本和作者信息")
        
        help_btn = ttk.Button(tip_inner, text="❓ 帮助", width=8,
                             command=self._show_help)
        help_btn.pack(side='right', padx=5)
        create_tooltip(help_btn, "打开帮助文档（快捷键: F1）")
    
    def _create_tabs(self):
        """创建所有标签页（优化：延迟加载Tab内容）"""
        self.notebook = ttk.Notebook(self.master)
        self.notebook.pack(pady=10, padx=10, fill="both", expand=True)
        
        # 创建各个标签页框架（但不立即创建内容）
        self.tab1 = ttk.Frame(self.notebook, padding=10)
        self.tab2 = ttk.Frame(self.notebook, padding=10)
        self.tab3 = ttk.Frame(self.notebook, padding=10)
        self.tab4 = ttk.Frame(self.notebook, padding=10)
        self.tab5 = ttk.Frame(self.notebook, padding=10)
        self.tab6 = ttk.Frame(self.notebook, padding=10)
        self.tab7 = ttk.Frame(self.notebook, padding=10)
        self.tab8 = ttk.Frame(self.notebook, padding=10)
        self.tab9 = ttk.Frame(self.notebook, padding=10)
        self.tab10 = ttk.Frame(self.notebook, padding=10)
        self.tab11 = ttk.Frame(self.notebook, padding=10)
        self.tab12 = ttk.Frame(self.notebook, padding=10)
        self.tab13 = ttk.Frame(self.notebook, padding=10)
        self.tab14 = ttk.Frame(self.notebook, padding=10)
        
        # 添加标签页（优化：平衡文本长度，既清晰又节省空间）
        self.notebook.add(self.tab1, text="[1] 州名转换")
        self.notebook.add(self.tab2, text="[2] SKU填充")
        self.notebook.add(self.tab3, text="[3] 高亮重复")
        self.notebook.add(self.tab4, text="[4] 插入行")
        self.notebook.add(self.tab5, text="[5] 对比列")
        self.notebook.add(self.tab6, text="[6] PDF拆分")
        self.notebook.add(self.tab7, text="[7] 前缀填充")
        self.notebook.add(self.tab8, text="[8] 面单页脚")
        self.notebook.add(self.tab9, text="[9] 仓库推荐")
        self.notebook.add(self.tab10, text="[10] 录入库存")
        self.notebook.add(self.tab11, text="[11] 模板填充")
        self.notebook.add(self.tab12, text="[12] PPT转PDF")
        self.notebook.add(self.tab13, text="[13] 图片压缩")
        self.notebook.add(self.tab14, text="[14] 删除列")
        
        # 记录Tab是否已初始化
        self._tabs_initialized = set()
        
        # 【修复持久化】预初始化所有Tab的变量（不创建UI）
        self._initialize_all_variables()
        
        # 绑定Tab切换事件，实现延迟加载
        self.notebook.bind('<<NotebookTabChanged>>', self._on_tab_changed)
        
        # 立即初始化第一个Tab（用户最可能使用的）
        self._initialize_tab(0)
        
        # 延迟执行自动加载（等待UI完全创建）
        self.master.after(500, self._auto_load_persisted_data)
    
    def _create_status_bar(self):
        """创建底部状态栏"""
        status_frame = ttk.Frame(self.master)
        status_frame.pack(fill='x', side='bottom', padx=10, pady=(0, 6))
        
        status_inner = ttk.Frame(status_frame)
        status_inner.pack(fill='x')
        
        self.status_icon = ttk.Label(status_inner, text="✅", font=("Segoe UI", 10))
        self.status_icon.pack(side='left', padx=(0, 5))
        
        self.status_var = tk.StringVar(value="就绪")
        self.status_label = ttk.Label(status_inner, textvariable=self.status_var, font=("Segoe UI", 9))
        self.status_label.pack(side='left')
        
        # 进度条（默认隐藏）
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(status_frame,
                                           variable=self.progress_var,
                                           mode='indeterminate',
                                           length=200)
        
        # 快捷键提示
        shortcut_label = ttk.Label(status_inner,
                                  text="快捷键: Ctrl+O=打开文件 | Ctrl+R=运行 | Ctrl+L=清空日志 | F1=帮助",
                                  font=("Segoe UI", 8))
        shortcut_label.pack(side='right', padx=5)
    
    # ==================== 快捷键 ====================
    
    def _bind_shortcuts(self):
        """绑定全局快捷键"""
        self.master.bind('<F1>', lambda e: self._show_help())
        self.master.bind('<Control-h>', lambda e: self._show_help())
        self.master.bind('<Control-q>', lambda e: self.master.quit())
        self.master.bind('<Control-Tab>', lambda e: self._next_tab())
        self.master.bind('<Control-Shift-Tab>', lambda e: self._prev_tab())
    
    def _next_tab(self):
        """切换到下一个标签页"""
        current = self.notebook.index(self.notebook.select())
        total = self.notebook.index('end')
        next_tab = (current + 1) % total
        self.notebook.select(next_tab)
    
    def _prev_tab(self):
        """切换到上一个标签页"""
        current = self.notebook.index(self.notebook.select())
        total = self.notebook.index('end')
        prev_tab = (current - 1) % total
        self.notebook.select(prev_tab)
    
    # ==================== Tab延迟加载 ====================
    
    def _initialize_all_variables(self):
        """预初始化所有Tab的变量（修复持久化问题）"""
        try:
            # Tab1 - 州名转换
            if not hasattr(self, 'file1_var'):
                self.file1_var = tk.StringVar(value="未选择文件")
                self.sheet1_var = tk.StringVar()
                self.col1_var = tk.StringVar(value="G")
                self._trace_persist(self.file1_var)
                self._trace_persist(self.sheet1_var)
                self._trace_persist(self.col1_var)
            
            # Tab2 - SKU填充
            if not hasattr(self, 'file2_var'):
                self.file2_var = tk.StringVar(value="未选择文件")
                self.sku_db2_var = tk.StringVar(value="未选择SKU数据库")
                self.order2_sheet_var = tk.StringVar()
                self.sku_db2_sheet_var = tk.StringVar()
                self.db_sku_col = tk.StringVar()
                self.db_l_col = tk.StringVar()
                self.db_w_col = tk.StringVar()
                self.db_h_col = tk.StringVar()
                self.db_wt_col = tk.StringVar()
                self.template2_var = tk.StringVar(value="默认")
                self.target_sku_col = tk.StringVar(value="A")
                self.target_qty_col = tk.StringVar(value="B")
                self.target_l_col = tk.StringVar(value="C")
                self.target_w_col = tk.StringVar(value="D")
                self.target_h_col = tk.StringVar(value="E")
                self.target_wt_col = tk.StringVar(value="F")
                
                # file2_var、sku_db2_var和sku_db2_sheet_var使用独立的持久化机制，在tab2_skus.py中管理
                # 其他变量继续使用通用的持久化配置
                self._trace_persist(self.order2_sheet_var)
                self._trace_persist(self.db_sku_col)
                self._trace_persist(self.db_l_col)
                self._trace_persist(self.db_w_col)
                self._trace_persist(self.db_h_col)
                self._trace_persist(self.db_wt_col)
                self._trace_persist(self.template2_var)
                self._trace_persist(self.target_sku_col)
                self._trace_persist(self.target_qty_col)
                self._trace_persist(self.target_l_col)
                self._trace_persist(self.target_w_col)
                self._trace_persist(self.target_h_col)
                self._trace_persist(self.target_wt_col)
            
            # Tab4 - 插入行
            if not hasattr(self, 'file_x_var'):
                self.file_x_var = tk.StringVar(value="未选择文件")
                self.file_y_var = tk.StringVar(value="未选择文件")
                self.sheet_x_var = tk.StringVar()
                self.sheet_y_var = tk.StringVar()
                self._trace_persist(self.file_x_var)
                self._trace_persist(self.file_y_var)
                self._trace_persist(self.sheet_x_var)
                self._trace_persist(self.sheet_y_var)
            
            # Tab11 - 模板填充
            if not hasattr(self, 'config11_var'):
                self.config11_var = tk.StringVar(value="未选择配置文件")
                self.order11_file_var = tk.StringVar(value="未选择订单文件")
                self.order11_sheet_var = tk.StringVar()
                self.template11_file_var = tk.StringVar(value="未选择模板文件")
                self.template11_sheet_var = tk.StringVar()
                self.mapping11_choice_var = tk.StringVar()
                
                self._trace_persist(self.config11_var)
                self._trace_persist(self.order11_file_var)
                self._trace_persist(self.order11_sheet_var)
                self._trace_persist(self.template11_file_var)
                self._trace_persist(self.template11_sheet_var)
                self._trace_persist(self.mapping11_choice_var)
            
            # Tab9 - 仓库推荐
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
            
            # Tab10 - 录入库存
            if not hasattr(self, 'wh10'):
                self.wh10 = {}  # 仓库->州 映射
                self.sku10 = {}  # 仓库->SKU集合 映射
                self.inv10_var = tk.StringVar(value="未选择库存文件")  # 添加文件路径变量
                self._trace_persist(self.inv10_var)
            
            # Tab3 - 高亮重复
            if not hasattr(self, 'file3_var'):
                self.file3_var = tk.StringVar(value="未选择文件")
                self.col3_var = tk.StringVar(value="A")
                self.sheet3_var = tk.StringVar()
                self._trace_persist(self.file3_var)
                self._trace_persist(self.col3_var)
                self._trace_persist(self.sheet3_var)
            
            # Tab5 - 对比列
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
            
            # Tab6 - PDF拆分
            if not hasattr(self, 'pdf_input_var'):
                self.pdf_input_var = tk.StringVar(value="未选择PDF")
                self.pdf_outdir_var = tk.StringVar(value="未选择输出目录")
                self.pdf_bbox_x = tk.StringVar(value="100")
                self.pdf_bbox_y = tk.StringVar(value="200")
                self.pdf_bbox_w = tk.StringVar(value="800")
                self.pdf_bbox_h = tk.StringVar(value="200")
                self.uniuni_mode_var = tk.BooleanVar(value=False)
                self.pdf_bbox2_x = tk.StringVar(value="120")
                self.pdf_bbox2_y = tk.StringVar(value="220")
                self.pdf_bbox2_w = tk.StringVar(value="800")
                self.pdf_bbox2_h = tk.StringVar(value="200")
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
                self.template_choice_var = tk.StringVar(value="请选择")
                
                # 持久化所有PDF相关变量
                for var in [self.pdf_input_var, self.pdf_outdir_var, self.pdf_bbox_x, self.pdf_bbox_y,
                           self.pdf_bbox_w, self.pdf_bbox_h, self.uniuni_mode_var, self.pdf_bbox2_x,
                           self.pdf_bbox2_y, self.pdf_bbox2_w, self.pdf_bbox2_h, self.three_region_mode_var,
                           self.pdf_bbox3_x, self.pdf_bbox3_y, self.pdf_bbox3_w, self.pdf_bbox3_h,
                           self.pdf_dpi_var, self.poppler_var, self.tesseract_var, self.regex_var,
                           self.prefix_var, self.ocr_engine_var, self.template_choice_var]:
                    self._trace_persist(var)
            
            # Tab7 - 前缀填充
            if not hasattr(self, 'file7_var'):
                self.file7_var = tk.StringVar(value="未选择文件")
                self.src7_var = tk.StringVar(value="A")
                self.dst7_var = tk.StringVar(value="B")
                self._trace_persist(self.file7_var)
                self._trace_persist(self.src7_var)
                self._trace_persist(self.dst7_var)
            
            # Tab8 - 面单页脚
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
                
                for var in [self.pdf8_input_var, self.pdf8_output_var, self.pdf8_map_excel_var,
                           self.pdf8_map_sheet_var, self.pdf8_short_col_var, self.pdf8_full_col_var,
                           self.pdf8_overwrite_var, self.pdf8_font_var, self.pdf8_fontsize_var]:
                    self._trace_persist(var)
                
            # Tab12 - PPT转PDF
            if not hasattr(self, 'ppt_files_var'):
                self.ppt_files_var = tk.StringVar(value="未选择文件")
                self.ppt_outdir_var = tk.StringVar(value="与原文件相同")
                self._ppt_file_list = []
                self._trace_persist(self.ppt_files_var)
                self._trace_persist(self.ppt_outdir_var)
            
            # Tab14 - 删除列
            if not hasattr(self, 'file14_var'):
                self.file14_var = tk.StringVar(value="未选择文件")
                self.sheet14_var = tk.StringVar()
                self.cols14_var = tk.StringVar(value="")
                self._trace_persist(self.file14_var)
                self._trace_persist(self.sheet14_var)
                self._trace_persist(self.cols14_var)
            
            print("[OK] 所有Tab变量预初始化完成，持久化功能已修复")

        except Exception as e:
            print(f"[WARNING] 变量预初始化失败: {e}")
    
    def _on_tab_changed(self, event):
        """Tab切换时的回调，实现延迟加载"""
        try:
            current_tab = self.notebook.index(self.notebook.select())
            self._initialize_tab(current_tab)
        except Exception as e:
            print(f"Tab切换错误: {e}")
    
    def _initialize_tab(self, tab_index):
        """初始化指定的Tab（如果尚未初始化）"""
        if tab_index in self._tabs_initialized:
            return
        
        try:
            # 根据Tab索引调用对应的创建方法
            if tab_index == 0:  # Tab1 - 州名转换
                self.create_tab1_states(self.tab1)
            elif tab_index == 1:  # Tab2 - SKU填充
                self.create_tab2_skus(self.tab2)
            elif tab_index == 2:  # Tab3 - 高亮重复
                self.create_tab3_highlight(self.tab3)
            elif tab_index == 3:  # Tab4 - 插入行
                self.create_tab4_insert(self.tab4)
            elif tab_index == 4:  # Tab5 - 对比列
                self.create_tab5_compare(self.tab5)
            elif tab_index == 5:  # Tab6 - PDF拆分
                self.create_tab6_pdf(self.tab6)
            elif tab_index == 6:  # Tab7 - 前缀填充
                self.create_tab7_prefix(self.tab7)
            elif tab_index == 7:  # Tab8 - 面单页脚
                self.create_tab8_pdf_footer(self.tab8)
            elif tab_index == 8:  # Tab9 - 仓库推荐
                self.create_tab9_router(self.tab9)
            elif tab_index == 9:  # Tab10 - 录入库存
                self.create_tab10_entry(self.tab10)
            elif tab_index == 10:  # Tab11 - 模板填充
                self.create_tab11_shipping(self.tab11)
            elif tab_index == 11:  # Tab12 - PPT转PDF
                self.create_tab12_ppt(self.tab12)
            elif tab_index == 12:  # Tab13 - 图片压缩
                self.create_tab13_image_compress(self.tab13)
            elif tab_index == 13:  # Tab14 - 删除列
                self.create_tab14_delete_cols(self.tab14)
            
            # 标记为已初始化
            self._tabs_initialized.add(tab_index)
            
        except Exception as e:
            print(f"初始化Tab {tab_index} 失败: {e}")
    
    # ==================== 状态更新 ====================
    
    def _update_status(self, message, icon="✅", show_progress=False):
        """更新状态栏显示"""
        self.status_var.set(message)
        self.status_icon.config(text=icon)
        
        if show_progress:
            self.progress_bar.pack(side='left', padx=10)
            self.progress_bar.start(10)
        else:
            self.progress_bar.stop()
            self.progress_bar.pack_forget()
        
        self.master.update()
    
    # ==================== 日志组件 ====================
    
    def create_log_widget(self, parent_frame):
        """创建日志组件"""
        log_frame = ttk.LabelFrame(parent_frame, text="日志", style="Section.TLabelframe")
        log_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        text_widget = ScrolledText(log_frame, height=12, state="disabled")
        text_widget.pack(fill="both", expand=True, padx=5, pady=5)
        try:
            text_widget.configure(bg="#F9FAFB", fg="#111827", insertbackground="#111827")
        except Exception:
            pass
        self._text_widgets.append(text_widget)
        
        def logger(text):
            text_widget.config(state="normal")
            text_widget.insert("end", str(text) + "\n")
            text_widget.see("end")
            text_widget.config(state="disabled")
        
        def clear_log():
            text_widget.config(state="normal")
            text_widget.delete("1.0", "end")
            logger("日志已清空。")
        
        return logger, clear_log
    
    # ==================== 文件选择辅助 ====================
    
    def _update_combobox_options(self, combobox, var, options):
        """更新下拉框选项"""
        combobox['values'] = options or []
        if options:
            var.set(options[0])
            try:
                combobox.current(0)
            except tk.TclError:
                pass
        else:
            var.set("")
    
    def _update_listbox_options(self, listbox, options):
        """更新列表框选项"""
        listbox.delete(0, 'end')
        for item in options or []:
            listbox.insert('end', item)
    
    def select_file_and_sheets(self, file_var, sheet_var, combobox, title):
        """选择文件并加载工作表列表"""
        path = filedialog.askopenfilename(
            title=title,
            filetypes=[("Excel Files", "*.xlsx *.xlsm *.xls"), ("All Files", "*.*")]
        )
        if not path:
            return
        file_var.set(path)
        if sheet_var is not None and combobox is not None:
            names = get_sheet_names(path)
            if names:
                self._update_combobox_options(combobox, sheet_var, names)
            else:
                self._update_combobox_options(combobox, sheet_var, [])
                messagebox.showerror("读取错误", "无法读取此文件的子表，请确认文件未被占用。")
    
    def select_file_and_listbox(self, file_var, listbox, title):
        """选择文件并加载工作表列表到列表框"""
        path = filedialog.askopenfilename(
            title=title,
            filetypes=[("Excel Files", "*.xlsx *.xlsm *.xls"), ("All Files", "*.*")]
        )
        if not path:
            return
        file_var.set(path)
        names = get_sheet_names(path)
        if names:
            self._update_listbox_options(listbox, names)
        else:
            self._update_listbox_options(listbox, [])
            messagebox.showerror("读取错误", "无法读取此文件的子表，请确认文件未被占用。")
    
    # ==================== 主题 ====================
    
    def _on_topmost_change(self):
        """置顶状态变化时的回调"""
        try:
            self.master.attributes("-topmost", self.topmost_var.get())
        except tk.TclError as e:
            print(f"警告：设置窗口置顶失败: {e}")
        try:
            self._persist_config()
        except (IOError, OSError) as e:
            print(f"警告：保存配置失败: {e}")
    
    def _apply_theme(self, mode: str):
        """应用主题"""
        style = ttk.Style()
        bg_light = "#FBFBFD"; fg_light = "#111827"; tab_light = "#E5E7EB"; sel_light = "#FFFFFF"
        bg_dark = "#0B1220"; fg_dark = "#E5E7EB"; tab_dark = "#111827"; sel_dark = "#0B1220"
        
        if mode == "浅色":
            bg, fg, tab_bg, sel_bg, acc = bg_light, fg_light, tab_light, sel_light, "#3b82f6"
        elif mode == "深色":
            bg, fg, tab_bg, sel_bg, acc = bg_dark, fg_dark, tab_dark, sel_dark, "#60A5FA"
        else:
            bg, fg, tab_bg, sel_bg, acc = bg_light, fg_light, tab_light, sel_light, self._accent
        
        self._accent = acc
        
        try:
            self.master.configure(bg=bg)
        except Exception:
            pass
        
        try:
            style.configure('TFrame', background=bg)
            style.configure('TLabelframe', background=bg)
            style.configure('TLabelframe.Label', background=bg, foreground=fg)
            style.configure('Section.TLabelframe', background=bg)
            style.configure('Section.TLabelframe.Label', background=bg, foreground=fg, font=("Segoe UI", 11, "bold"))
            style.configure('TLabel', background=bg, foreground=fg)
            style.configure('TButton', padding=(10, 6), foreground=fg)
            
            acc_fg = ('#111827' if mode != '深色' else '#FFFFFF')
            style.configure('Accent.TButton', padding=(12, 8), background=acc, foreground=acc_fg)
            style.map('Accent.TButton',
                     background=[('active', '#2563eb' if mode != '深色' else '#93C5FD'),
                                ('pressed', '#1e40af' if mode != '深色' else '#3B82F6')],
                     foreground=[('disabled', '#9CA3AF')])
            
            sec_bg = ('#EEF2FF' if mode != '深色' else '#1F2937')
            sec_active = ('#E0E7FF' if mode != '深色' else '#374151')
            sec_pressed = ('#C7D2FE' if mode != '深色' else '#111827')
            style.configure('Secondary.TButton', padding=(10, 6), background=sec_bg,
                          foreground=(fg if mode != '深色' else '#E5E7EB'))
            style.map('Secondary.TButton',
                     background=[('active', sec_active), ('pressed', sec_pressed)],
                     foreground=[('disabled', '#9CA3AF')])
            
            style.configure('TNotebook', background=bg, borderwidth=0)
            style.configure('TNotebook.Tab', 
                          padding=(12, 6),  # 优化：减小水平间距，缩小选项卡宽度
                          background=tab_bg, 
                          foreground=fg)
            style.map('TNotebook.Tab', 
                     background=[('selected', sel_bg)], 
                     foreground=[('selected', fg)])
        except Exception:
            pass
        
        for tw in getattr(self, '_text_widgets', []):
            try:
                tw.configure(bg=("#0F172A" if mode == "深色" else "#F9FAFB"),
                           fg=(fg if mode != "深色" else "#D1D5DB"),
                           insertbackground=fg)
            except Exception:
                pass
    
    # ==================== 帮助和关于 ====================
    
    def _show_help(self):
        """显示帮助对话框"""
        help_window = tk.Toplevel(self.master)
        help_window.title("帮助 - Excel 工具箱")
        help_window.geometry("600x500")
        help_window.transient(self.master)
        help_window.grab_set()
        
        # 居中显示
        help_window.update_idletasks()
        x = self.master.winfo_x() + (self.master.winfo_width() - 600) // 2
        y = self.master.winfo_y() + (self.master.winfo_height() - 500) // 2
        help_window.geometry(f"+{x}+{y}")
        
        # 内容区域
        help_frame = ttk.Frame(help_window, padding=20)
        help_frame.pack(fill='both', expand=True)
        
        title_label = ttk.Label(help_frame, text="📚 帮助文档", font=("Segoe UI", 16, "bold"))
        title_label.pack(pady=(0, 15))
        
        text_frame = ttk.Frame(help_frame)
        text_frame.pack(fill='both', expand=True)
        
        scrollbar = ttk.Scrollbar(text_frame)
        scrollbar.pack(side='right', fill='y')
        
        help_text = tk.Text(text_frame, wrap='word', yscrollcommand=scrollbar.set,
                           font=("Segoe UI", 10), padx=10, pady=10)
        help_text.pack(side='left', fill='both', expand=True)
        scrollbar.config(command=help_text.yview)
        
        help_content = f"""
欢迎使用 Excel 工具箱 V{self.VERSION}

✨ 功能介绍

[1] 转换州名 - 将美国州全名转换为缩写
[2] 填充SKU信息 - 智能填充商品SKU相关信息
[3] 高亮重复项 - 自动标记Excel中的重复数据
[4] 插入缺失行 - 检测并插入缺失的数据行
[5] 对比列数据 - 对比两个Excel表格的列数据差异
[6] 拆分订单PDF - 将合并的PDF按页拆分，OCR识别订单号
[7] 前缀填充承运商 - 根据前缀规则填充物流信息
[9] 建议发货仓库 - 智能推荐最近的发货仓库
[10] 录入发货信息 - 维护仓库和库存数据
[11] 发货模板填充 - 订单信息自动填充到发货模板
[12] 批量PPT转PDF - 快速将PPT/PPTX转换为PDF文档

⏰ 快捷键

F1 / Ctrl+H - 打开帮助
Ctrl+Q - 退出程序
Ctrl+Tab - 下一个标签页
Ctrl+Shift+Tab - 上一个标签页

⚠️ 重要提示

1. 处理Excel文件前，必须关闭Excel/WPS，否则无法保存
2. 建议备份原始文件，避免数据丢失
3. PDF功能需要Tesseract和Poppler支持
4. 配置文件位于 excel_toolkit/config.json

作者：{self.AUTHOR}
版本：V{self.VERSION}
感谢使用！🎉
        """
        
        help_text.insert('1.0', help_content)
        help_text.config(state='disabled')
        
        btn_frame = ttk.Frame(help_frame)
        btn_frame.pack(pady=(10, 0))
        
        close_btn = ttk.Button(btn_frame, text="关闭",
                              command=help_window.destroy,
                              style='Accent.TButton',
                              width=15)
        close_btn.pack()
    
    def _show_about(self):
        """显示关于对话框"""
        about_window = tk.Toplevel(self.master)
        about_window.title("关于 - Excel 工具箱")
        about_window.geometry("450x380")
        about_window.resizable(False, False)
        about_window.transient(self.master)
        about_window.grab_set()
        
        about_window.update_idletasks()
        x = self.master.winfo_x() + (self.master.winfo_width() - 450) // 2
        y = self.master.winfo_y() + (self.master.winfo_height() - 380) // 2
        about_window.geometry(f"+{x}+{y}")
        
        main_frame = ttk.Frame(about_window, padding=30)
        main_frame.pack(fill='both', expand=True)
        
        icon_label = ttk.Label(main_frame, text="📊", font=("Segoe UI", 48))
        icon_label.pack(pady=(0, 10))
        
        title_label = ttk.Label(main_frame, text="Excel 工具箱",
                               font=("Segoe UI", 20, "bold"))
        title_label.pack(pady=5)
        
        version_label = ttk.Label(main_frame,
                                 text=f"Version {self.VERSION}",
                                 font=("Segoe UI", 12))
        version_label.pack(pady=5)
        
        ttk.Separator(main_frame, orient='horizontal').pack(fill='x', pady=15)
        
        info_frame = ttk.Frame(main_frame)
        info_frame.pack(pady=10)
        
        info_items = [
            ("👨‍💻 作者", self.AUTHOR),
            ("💻 技术栈", "Python + Tkinter"),
            ("📦 功能数量", "10 个工具"),
            ("🌟 架构", "Mixin 模块化"),
        ]
        
        for label, value in info_items:
            row = ttk.Frame(info_frame)
            row.pack(fill='x', pady=3)
            ttk.Label(row, text=label, font=("Segoe UI", 10, "bold")).pack(side='left')
            ttk.Label(row, text=value, font=("Segoe UI", 10)).pack(side='left', padx=10)
        
        ttk.Separator(main_frame, orient='horizontal').pack(fill='x', pady=15)
        
        copyright_label = ttk.Label(main_frame,
                                   text="© 2025 All Rights Reserved",
                                   font=("Segoe UI", 9))
        copyright_label.pack(pady=5)
        
        ttk.Button(main_frame, text="关闭",
                  command=about_window.destroy,
                  style='Accent.TButton',
                  width=15).pack(pady=10)
    
    # ====================配置管理 ====================
        
    def _auto_load_persisted_data(self):
        """自动加载持久化的数据（程序启动时）"""
        # Tab11: 自动刷新配置文件的映射和仓库列表
        if hasattr(self, 'config11_var'):
            config_path = self.config11_var.get()
            if config_path and config_path != "未选择配置文件" and os.path.exists(config_path):
                try:
                    self._refresh_mapping_choices11()
                    self._refresh_warehouses11()
                    print(f"[自动加载] Tab11 配置文件: {config_path}")
                except Exception as e:
                    print(f"[自动加载] Tab11 配置文件失败: {e}")
        
        # Tab10/Tab9: 优先从数据库加载库存
        try:
            from excel_toolkit.db_operations import load_warehouse_inventory
            from excel_toolkit.db_config import get_db_manager
            
            db = get_db_manager()
            if db.config.is_enabled():
                data = load_warehouse_inventory()
                if data:
                    warehouse_data, sku_data = data
                    
                    # 加载到Tab10
                    if hasattr(self, 'wh10'):
                        self.wh10 = warehouse_data
                        self.sku10 = sku_data
                        
                        # 只有在Tab10 UI已创建时才更新UI
                        if hasattr(self, 'tree10'):
                            for w, st in sorted(self.wh10.items()):
                                self.tree10.insert('', 'end', values=(w, st or ''))
                        
                        print(f"[自动加载] 从数据库加载 {len(self.wh10)} 个仓库")
                        
                        # 设置数据库标识
                        if hasattr(self, 'inv10_var'):
                            self.inv10_var.set("[数据库]")
                    
                    # 同步到Tab9
                    if hasattr(self, 'inv9_var'):
                        self.inv9_var.set("[数据库]")
                        # 延迟刷新，等待UI完全初始化
                        self.master.after(200, self._refresh_block9_from_inventory)
                    
                    return  # 从数据库加载成功，跳过Excel文件加载
        except Exception as e:
            print(f"[自动加载] 数据库加载失败: {e}")
        
        # 降级：从Excel文件加载（原有逻辑）
        # Tab9: 自动刷新仓库列表
        if hasattr(self, 'inv9_var'):
            inv_path = self.inv9_var.get()
            if inv_path and inv_path != "未选择发货信息表格" and os.path.exists(inv_path):
                try:
                    # 只有在Tab9 UI已创建时才刷新
                    if hasattr(self, '_refresh_block9_from_inventory'):
                        self._refresh_block9_from_inventory()
                        print(f"[自动加载] Tab9 库存文件: {inv_path}")
                except Exception as e:
                    print(f"[自加载] Tab9 库存文件失败: {e}")
        
        # Tab10: 自动加载库存数据（从 inv10_var 或 inv9_var 同步）
        if hasattr(self, 'inv10_var'):
            inv_path = self.inv10_var.get()
            if inv_path and inv_path != "未选择库存文件" and inv_path != "[数据库]" and os.path.exists(inv_path):
                try:
                    sku_by_wh, wh_state = read_inventory(inv_path, logger=lambda x: None)
                    
                    # 更新数据（无论UI是否存在）
                    self.wh10 = {str(k): str(v) if v else '' for k, v in wh_state.items()}
                    self.sku10 = {str(k): set(v) for k, v in sku_by_wh.items()}
                    
                    # 只有在Tab10 UI已创建时才更新UI
                    if hasattr(self, 'tree10') and hasattr(self, 'list10'):
                        # 清空现有数据
                        for item in self.tree10.get_children():
                            self.tree10.delete(item)
                        self.list10.delete(0, 'end')
                        
                        # 更新UI
                        for w, st in sorted(self.wh10.items()):
                            self.tree10.insert('', 'end', values=(w, st or ''))
                    
                    # 同步到Tab9
                    if hasattr(self, 'inv9_var'):
                        self.inv9_var.set(inv_path)
                    
                    print(f"[自动加载] Tab10 库存数据: {len(self.wh10)} 个仓库")
                except Exception as e:
                    print(f"[自动加载] Tab10 库存数据失败: {e}")
        elif hasattr(self, 'inv9_var'):
            # 如果Tab10没有自己的路径，尝试从Tab9同步
            inv_path = self.inv9_var.get()
            if inv_path and inv_path != "未选择发货信息表格" and inv_path != "[数据库]" and os.path.exists(inv_path):
                try:
                    sku_by_wh, wh_state = read_inventory(inv_path, logger=lambda x: None)
                    
                    # 更新数据（无论UI是否存在）
                    self.wh10 = {str(k): str(v) if v else '' for k, v in wh_state.items()}
                    self.sku10 = {str(k): set(v) for k, v in sku_by_wh.items()}
                    
                    # 只有在Tab10 UI已创建时才更新UI
                    if hasattr(self, 'tree10') and hasattr(self, 'list10'):
                        # 清空现有数据
                        for item in self.tree10.get_children():
                            self.tree10.delete(item)
                        self.list10.delete(0, 'end')
                        
                        # 更新UI
                        for w, st in sorted(self.wh10.items()):
                            self.tree10.insert('', 'end', values=(w, st or ''))
                    
                    # 同步路径到Tab10
                    if hasattr(self, 'inv10_var'):
                        self.inv10_var.set(inv_path)
                    
                    print(f"[自动加载] Tab10 库存数据（从Tab9同步）: {len(self.wh10)} 个仓库")
                except Exception as e:
                    print(f"[自动加载] Tab10 库存数据失败: {e}")
        
    # ====================配置管理 ====================
    
    def _config_dir(self):
        """获取配置目录（使用用户目录，打包后也能正常工作）"""
        # 优先使用用户目录
        user_dir = os.path.expanduser("~")
        config_dir = os.path.join(user_dir, ".excel_toolkit")
        try:
            os.makedirs(config_dir, exist_ok=True)
        except Exception:
            # 回退到程序目录
            config_dir = os.path.dirname(os.path.abspath(__file__))
        return config_dir
    
    def _config_path(self):
        """获取配置文件路径"""
        return os.path.join(self._config_dir(), "config.json")
    
    def _trace_persist(self, var):
        """为变量添加配置保存追踪"""
        # 注册变量到持久化列表
        if not hasattr(self, '_persist_vars'):
            self._persist_vars = {}
        
        # 获取变量名（通过查找实例属性）
        var_name = None
        for name, value in self.__dict__.items():
            if value is var:
                var_name = name
                break
        
        if var_name:
            self._persist_vars[var_name] = var
            try:
                var.trace_add("write", lambda *a: self._persist_config())
            except Exception:
                pass
    
    def _load_config(self):
        """加载配置文件"""
        p = self._config_path()
        if not os.path.exists(p):
            return
        
        try:
            with open(p, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            return
        
        # 加载置顶状态
        atop = data.get("always_on_top")
        if atop is not None:
            try:
                self.topmost_var.set(bool(atop))
            except Exception:
                pass
        
        # 加载所有已注册的变量
        vars_data = data.get("vars", {})
        if hasattr(self, '_persist_vars'):
            for var_name, var in self._persist_vars.items():
                if var_name in vars_data:
                    try:
                        var.set(vars_data[var_name])
                    except Exception:
                        pass
    
    def _persist_config(self):
        """保存配置文件"""
        p = self._config_path()
        data = {
            "always_on_top": self.topmost_var.get() if hasattr(self, 'topmost_var') else False,
        }
        
        # 保存所有已注册的变量
        if hasattr(self, '_persist_vars'):
            vars_data = {}
            for var_name, var in self._persist_vars.items():
                try:
                    vars_data[var_name] = var.get()
                except Exception:
                    pass
            data["vars"] = vars_data
        
        try:
            with open(p, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存配置失败: {e}")
