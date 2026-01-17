"""
工具提示（Tooltip）组件
为 GUI 控件添加悬停提示功能
"""
import tkinter as tk


class ToolTip:
    """工具提示类，为控件添加悬停提示"""
    
    def __init__(self, widget, text, delay=500):
        """
        初始化工具提示
        
        Args:
            widget: 要添加提示的控件
            text: 提示文本
            delay: 延迟显示时间（毫秒）
        """
        self.widget = widget
        self.text = text
        self.delay = delay
        self.tip_window = None
        self.schedule_id = None
        
        # 绑定事件
        self.widget.bind('<Enter>', self._on_enter)
        self.widget.bind('<Leave>', self._on_leave)
        self.widget.bind('<ButtonPress>', self._on_leave)
    
    def _on_enter(self, event=None):
        """鼠标进入控件"""
        self._schedule_show()
    
    def _on_leave(self, event=None):
        """鼠标离开控件"""
        self._cancel_schedule()
        self._hide_tip()
    
    def _schedule_show(self):
        """延迟显示提示"""
        self._cancel_schedule()
        self.schedule_id = self.widget.after(self.delay, self._show_tip)
    
    def _cancel_schedule(self):
        """取消延迟显示"""
        if self.schedule_id:
            self.widget.after_cancel(self.schedule_id)
            self.schedule_id = None
    
    def _show_tip(self):
        """显示工具提示"""
        if self.tip_window or not self.text:
            return
        
        # 获取控件位置
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 5
        
        # 创建提示窗口
        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)  # 无边框窗口
        tw.wm_geometry(f"+{x}+{y}")
        
        # 设置提示样式
        label = tk.Label(tw, 
                        text=self.text,
                        justify=tk.LEFT,
                        background="#2C3E50",
                        foreground="#ECF0F1",
                        relief=tk.SOLID,
                        borderwidth=1,
                        font=("Segoe UI", 9),
                        padx=8,
                        pady=4)
        label.pack()
    
    def _hide_tip(self):
        """隐藏工具提示"""
        if self.tip_window:
            self.tip_window.destroy()
            self.tip_window = None


def create_tooltip(widget, text, delay=500):
    """
    便捷函数：为控件创建工具提示
    
    Args:
        widget: 要添加提示的控件
        text: 提示文本
        delay: 延迟显示时间（毫秒）
    
    Returns:
        ToolTip 实例
    """
    return ToolTip(widget, text, delay)
