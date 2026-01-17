"""
共享的 Mixin 类 - 提供日志、文件选择等通用功能
"""
import tkinter as tk
from tkinter import ttk, filedialog
from tkinter.scrolledtext import ScrolledText
from excel_toolkit.excel_lite import get_sheet_names


class LoggerMixin:
    """日志功能混入类"""
    
    def create_log_widget(self, parent_frame):
        """创建日志组件
        
        Returns:
            (logger_func, clear_func): 日志函数和清空函数
        """
        log_frame = ttk.LabelFrame(parent_frame, text="日志", style="Section.TLabelframe")
        log_frame.pack(fill="both", expand=True, padx=5, pady=5)

        text_widget = ScrolledText(log_frame, height=12, state="disabled")
        text_widget.pack(fill="both", expand=True, padx=5, pady=5)
        try:
            text_widget.configure(bg="#F9FAFB", fg="#111827", insertbackground="#111827")
        except Exception:
            pass
        
        # 注册到主题管理
        if hasattr(self, '_text_widgets'):
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


class FileSelectMixin:
    """文件选择功能混入类"""
    
    def select_file_and_sheets(self, file_var, sheet_var, combobox, title):
        """选择文件并加载工作表列表到下拉框"""
        path = filedialog.askopenfilename(
            title=title,
            filetypes=[("Excel文件", "*.xlsx;*.xlsm;*.xls"), ("所有文件", "*.*")]
        )
        if path:
            file_var.set(path)
            if combobox is not None:
                # 使用logger属性如果存在，否则使用print
                logger_func = getattr(self, 'logger2', print) if hasattr(self, 'logger2') else print
                names = get_sheet_names(path, logger=logger_func)
                if names:
                    self._update_combobox_options(combobox, sheet_var, names)

    def select_file_and_listbox(self, file_var, listbox, title):
        """选择文件并加载工作表列表到列表框"""
        path = filedialog.askopenfilename(
            title=title,
            filetypes=[("Excel文件", "*.xlsx;*.xlsm;*.xls"), ("所有文件", "*.*")]
        )
        if path:
            file_var.set(path)
            # 使用logger属性如果存在
            logger_func = getattr(self, 'logger9', print) if hasattr(self, 'logger9') else print
            names = get_sheet_names(path, logger=logger_func)
            self._update_listbox_options(listbox, names or [])
    
    def _update_combobox_options(self, combobox, var, options):
        """更新下拉框选项"""
        combobox['values'] = options
        if options:
            var.set(options[0])
    
    def _update_listbox_options(self, listbox, options):
        """更新列表框选项"""
        listbox.delete(0, 'end')
        for opt in options:
            listbox.insert('end', opt)





























