"""
Tab1 - 州名转换功能 (优化版布局)

布局优化说明:
1. 使用卡片式分组
2. 统一间距和边距
3. 优化按钮排版
4. 改进视觉层次
"""
import tkinter as tk
from tkinter import ttk, messagebox
import threading
import os

from excel_toolkit.states import process_states
from excel_toolkit.exceptions import ExcelToolkitError
from excel_toolkit.error_handler import get_user_friendly_error, log_error
from excel_toolkit.ui_config import SPACING, FONTS, SIZES, ICONS, LAYOUT_TEMPLATES
from excel_toolkit.tooltip import create_tooltip


class Tab1StatesImprovedMixin:
    """Tab1 州名转换 Mixin (优化版)"""
    
    def create_tab1_states_improved(self, tab):
        """创建Tab1界面(优化版)"""
        # 初始化变量
        self.file1_var = tk.StringVar(value="未选择文件")
        self.sheet1_var = tk.StringVar()
        self.col1_var = tk.StringVar(value="G")
        self._trace_persist(self.file1_var)
        self._trace_persist(self.sheet1_var)
        self._trace_persist(self.col1_var)

        # 主容器 - 添加滚动条支持(防止内容过多)
        main_container = ttk.Frame(tab)
        main_container.pack(fill='both', expand=True, padx=0, pady=0)
        
        # ==================== 卡片1: 文件选择 ====================
        file_card = ttk.LabelFrame(
            main_container, 
            text=f"{ICONS['excel']} 文件选择",
            style="Section.TLabelframe",
            padding=SPACING['section_padding']
        )
        file_card.pack(**LAYOUT_TEMPLATES['section_frame'])
        
        # 文件选择行
        file_row = ttk.Frame(file_card)
        file_row.pack(**LAYOUT_TEMPLATES['form_row'])
        
        select_btn = ttk.Button(
            file_row, 
            text=f"{ICONS['folder']} 选择文件",
            width=SIZES['button_width_normal'],
            command=lambda: self.select_file_and_sheets(
                self.file1_var, self.sheet1_var, self.combo1, 
                "选择要转换州名的Excel文件"
            )
        )
        select_btn.pack(side='left', padx=(0, SPACING['control_padding_x']))
        create_tooltip(select_btn, "选择包含州名数据的Excel文件")
        
        file_label = ttk.Label(
            file_row, 
            textvariable=self.file1_var,
            foreground='#6B7280'
        )
        file_label.pack(side='left', fill='x', expand=True)
        
        # ==================== 卡片2: 参数配置 ====================
        param_card = ttk.LabelFrame(
            main_container,
            text=f"{ICONS['settings']} 参数配置",
            style="Section.TLabelframe",
            padding=SPACING['section_padding']
        )
        param_card.pack(**LAYOUT_TEMPLATES['section_frame'])
        
        # 工作表选择行
        sheet_row = ttk.Frame(param_card)
        sheet_row.pack(**LAYOUT_TEMPLATES['form_row'])
        
        ttk.Label(
            sheet_row, 
            text="工作表:",
            width=SIZES['label_width_small']
        ).pack(side='left', padx=(0, SPACING['control_padding_x']))
        
        self.combo1 = ttk.Combobox(
            sheet_row, 
            textvariable=self.sheet1_var,
            state="readonly",
            width=SIZES['combobox_width_normal']
        )
        self.combo1.pack(side='left', padx=(0, SPACING['control_padding_x'] * 2))
        create_tooltip(self.combo1, "选择包含州名数据的工作表")
        
        # 列号输入
        ttk.Label(
            sheet_row,
            text="目标列:",
            width=SIZES['label_width_small']
        ).pack(side='left', padx=(0, SPACING['control_padding_x']))
        
        col_entry = ttk.Entry(
            sheet_row,
            textvariable=self.col1_var,
            width=SIZES['entry_width_small']
        )
        col_entry.pack(side='left')
        create_tooltip(col_entry, "输入要转换的列号(如A、B、G等)")
        
        # 提示信息
        hint_row = ttk.Frame(param_card)
        hint_row.pack(**LAYOUT_TEMPLATES['form_row'])
        
        ttk.Label(
            hint_row,
            text=f"{ICONS['info']} 提示: 程序会将选中列中的州全名转换为两字母缩写(如 California → CA)",
            foreground='#6B7280',
            font=FONTS['hint']
        ).pack(side='left')
        
        # ==================== 卡片3: 操作按钮 ====================
        action_card = ttk.Frame(main_container)
        action_card.pack(**LAYOUT_TEMPLATES['button_group'])
        
        # 执行按钮
        run_btn = ttk.Button(
            action_card,
            text=f"{ICONS['play']} 开始转换",
            command=self.run_tool1,
            style='Accent.TButton',
            width=SIZES['button_width_large']
        )
        run_btn.pack(side='left', padx=(0, SPACING['button_spacing']))
        create_tooltip(run_btn, "开始执行州名转换(快捷键: Ctrl+R)")
        
        # 清空日志按钮
        clear_btn = ttk.Button(
            action_card,
            text=f"{ICONS['clear']} 清空日志",
            command=lambda: self.clear_log1() if hasattr(self, 'clear_log1') else None,
            style='Secondary.TButton',
            width=SIZES['button_width_normal']
        )
        clear_btn.pack(side='left')
        create_tooltip(clear_btn, "清空下方的日志记录")
        
        # ==================== 卡片4: 日志区域 ====================
        log_card = ttk.LabelFrame(
            main_container,
            text=f"{ICONS['log']} 执行日志",
            style="Section.TLabelframe",
            padding=SPACING['section_padding']
        )
        log_card.pack(**LAYOUT_TEMPLATES['log_frame'])
        
        # 创建日志组件
        from tkinter.scrolledtext import ScrolledText
        log_widget = ScrolledText(
            log_card,
            height=SIZES['log_height'],
            state="disabled",
            font=FONTS['log'],
            wrap='word'
        )
        log_widget.pack(fill='both', expand=True)
        
        # 配置日志样式
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
        
        # 注册到主题管理
        if hasattr(self, '_text_widgets'):
            self._text_widgets.append(log_widget)
        
        # 日志函数
        def logger(text):
            log_widget.config(state="normal")
            log_widget.insert("end", str(text) + "\n")
            log_widget.see("end")
            log_widget.config(state="disabled")
        
        def clear_log():
            log_widget.config(state="normal")
            log_widget.delete("1.0", "end")
            logger(f"{ICONS['success']} 日志已清空")
        
        self.logger1 = logger
        self.clear_log1 = clear_log

    def run_tool1(self):
        """执行州名转换"""
        file = self.file1_var.get()
        sheet = self.sheet1_var.get()
        col = self.col1_var.get()
        
        # 验证输入
        if not file or file == "未选择文件":
            messagebox.showwarning("⚠️ 警告", "请先选择一个Excel文件")
            return
        if not sheet:
            messagebox.showwarning("⚠️ 警告", "请选择一个工作表")
            return
        if not col:
            messagebox.showwarning("⚠️ 警告", "请输入目标列号")
            return

        self.logger1("=" * 60)
        self.logger1(f"{ICONS['play']} 开始执行州名转换...")
        self.logger1(f"  文件: {os.path.basename(file)}")
        self.logger1(f"  工作表: {sheet}")
        self.logger1(f"  目标列: {col}")
        self.logger1("=" * 60)
        
        self._update_status("正在处理州名转换...", icon=ICONS['loading'], show_progress=True)
        self.master.config(cursor="watch")
        
        def thread_target():
            try:
                def safe_logger(msg):
                    self.master.after(0, lambda m=msg: self.logger1(m))
                
                stats = process_states(file, sheet, col, safe_logger)
                
                def on_success():
                    self.master.config(cursor="")
                    self._update_status("就绪", icon=ICONS['success'], show_progress=False)
                    
                    msg = (
                        f"✅ 州名转换完成！\n\n"
                        f"总共处理: {stats['total']} 行\n"
                        f"成功转换: {stats['success']} 行\n"
                        f"未找到/保持原值: {stats['failed']} 行\n\n"
                        f"文件已保存: {os.path.basename(file)}"
                    )
                    
                    messagebox.showinfo("✅ 完成", msg)
                    self.logger1("\n" + "=" * 60)
                    self.logger1(f"{ICONS['success']} 转换完成")
                    self.logger1(f"  总计: {stats['total']} 行")
                    self.logger1(f"  成功: {stats['success']} 行")
                    self.logger1(f"  跳过: {stats['failed']} 行")
                    self.logger1("=" * 60)

                self.master.after(0, on_success)

            except ExcelToolkitError as e:
                # 自定义异常(包含友好信息)
                def on_custom_error():
                    self.master.config(cursor="")
                    self._update_status("错误", icon=ICONS['error'], show_progress=False)
                    messagebox.showerror("❌ 错误", e.get_user_message())
                    self.logger1(f"\n{ICONS['error']} {e.message}")
                    if e.solution:
                        self.logger1(f"{ICONS['info']} 解决方案: {e.solution}")
                
                self.master.after(0, on_custom_error)
            
            except Exception as e:
                # 未预期的异常
                log_error(e, "州名转换")
                error_msg = get_user_friendly_error(e)
                
                def on_error():
                    self.master.config(cursor="")
                    self._update_status("错误", icon=ICONS['error'], show_progress=False)
                    messagebox.showerror("❌ 错误", error_msg)
                    self.logger1(f"\n{ICONS['error']} 发生错误: {str(e)}")
                    self.logger1(f"{ICONS['info']} 请查看日志文件获取详细信息")
                
                self.master.after(0, on_error)

        threading.Thread(target=thread_target, daemon=True).start()
