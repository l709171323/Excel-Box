"""
模板制作工具
用于从PDF渲染图像中截取特征区域作为模板
"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from PIL import Image, ImageTk, ImageOps
from typing import Optional, Tuple, Callable


class TemplateMaker:
    """模板制作对话框"""
    
    def __init__(self, parent, pdf_path: str, dpi: int, poppler_path: Optional[str], 
                 page_index: int = 0, callback: Optional[Callable] = None):
        """
        初始化模板制作器
        
        Args:
            parent: 父窗口
            pdf_path: PDF文件路径
            dpi: 渲染DPI
            poppler_path: Poppler路径
            page_index: 要渲染的页面索引（0开始）
            callback: 保存完成后的回调函数
        """
        self.parent = parent
        self.pdf_path = pdf_path
        self.dpi = dpi
        self.poppler_path = poppler_path
        self.page_index = page_index
        self.callback = callback
        
        # 创建窗口
        self.window = tk.Toplevel(parent)
        self.window.title("模板制作工具 - 正在加载...")
        self.window.geometry("1000x800")
        self.window.transient(parent)
        
        # 状态变量
        self.original_image = None
        self.display_image = None
        self.photo_image = None
        self.canvas_scale = 1.0
        
        # 选择框状态
        self.rect_start = None
        self.rect_id = None
        self.rect_coords = None
        
        # 创建界面
        self._create_ui()
        
        # 延迟加载图像，让窗口先显示
        self.window.after(100, self._load_image)
    
    def _create_ui(self):
        """创建用户界面"""
        # 顶部工具栏
        toolbar = ttk.Frame(self.window, padding=10)
        toolbar.pack(fill='x')
        
        ttk.Label(toolbar, text="🎯 拖拽鼠标选择模板区域", 
                 font=("Segoe UI", 11, "bold")).pack(side='left')
        
        ttk.Button(toolbar, text="🔄 重置选择", 
                  command=self._reset_selection).pack(side='right', padx=5)
        
        ttk.Button(toolbar, text="💾 保存模板", 
                  command=self._save_template,
                  style='Accent.TButton').pack(side='right', padx=5)
        
        # 分隔线
        ttk.Separator(self.window, orient='horizontal').pack(fill='x', padx=10)
        
        # 画布区域
        canvas_frame = ttk.Frame(self.window)
        canvas_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # 创建滚动条
        v_scroll = ttk.Scrollbar(canvas_frame, orient='vertical')
        v_scroll.pack(side='right', fill='y')
        
        h_scroll = ttk.Scrollbar(canvas_frame, orient='horizontal')
        h_scroll.pack(side='bottom', fill='x')
        
        # 创建画布
        self.canvas = tk.Canvas(canvas_frame, 
                               bg='#2C3E50',
                               xscrollcommand=h_scroll.set,
                               yscrollcommand=v_scroll.set)
        self.canvas.pack(side='left', fill='both', expand=True)
        
        h_scroll.config(command=self.canvas.xview)
        v_scroll.config(command=self.canvas.yview)
        
        # 绑定鼠标事件
        self.canvas.bind('<ButtonPress-1>', self._on_press)
        self.canvas.bind('<B1-Motion>', self._on_drag)
        self.canvas.bind('<ButtonRelease-1>', self._on_release)
        
        # 底部信息栏
        info_frame = ttk.Frame(self.window, padding=10)
        info_frame.pack(fill='x')
        
        self.info_label = ttk.Label(info_frame, 
                                    text="提示：拖拽鼠标框选Logo或标题等固定特征区域",
                                    font=("Segoe UI", 9))
        self.info_label.pack(side='left')
        
        self.coord_label = ttk.Label(info_frame, 
                                     text="坐标: 未选择",
                                     font=("Segoe UI", 9, "bold"))
        self.coord_label.pack(side='right')
    
    def _load_image(self):
        """加载并渲染PDF页面"""
        try:
            # 显示加载提示
            self.info_label.config(text="⏳ 正在渲染PDF页面，请稍候...")
            self.window.update_idletasks()
            
            # 导入依赖
            try:
                from pdf2image import convert_from_path
            except ImportError as e:
                messagebox.showerror("缺少依赖", 
                    "pdf2image 未安装！\n\n"
                    "请运行: pip install pdf2image\n\n"
                    "并确保已安装 Poppler 工具。")
                self.window.destroy()
                return
            
            # 如果没有指定 poppler_path，尝试自动检测
            if not self.poppler_path:
                from excel_toolkit.pdf_ocr import find_poppler
                auto_poppler = find_poppler()
                if auto_poppler:
                    self.poppler_path = auto_poppler
                    self.info_label.config(text=f"✅ 自动检测到 Poppler: {auto_poppler[:50]}...")
                    self.window.update_idletasks()
            
            # 渲染PDF（可能需要几秒钟）
            poppler_info = self.poppler_path if self.poppler_path else "系统PATH"
            self.info_label.config(text=f"⏳ 正在以 {self.dpi} DPI 渲染第 {self.page_index+1} 页...")
            self.window.update_idletasks()
            
            images = convert_from_path(
                self.pdf_path,
                dpi=self.dpi,
                first_page=self.page_index + 1,
                last_page=self.page_index + 1,
                poppler_path=self.poppler_path if self.poppler_path else None
            )
            
            if not images:
                raise RuntimeError(f"无法渲染PDF第{self.page_index + 1}页")
            
            self.original_image = images[0]
            
            # 计算缩放比例以适应窗口
            canvas_width = 950
            canvas_height = 650
            img_width, img_height = self.original_image.size
            
            scale_w = canvas_width / img_width
            scale_h = canvas_height / img_height
            self.canvas_scale = min(scale_w, scale_h, 1.0)  # 不放大，只缩小
            
            # 缩放用于显示
            display_width = int(img_width * self.canvas_scale)
            display_height = int(img_height * self.canvas_scale)
            self.display_image = self.original_image.resize(
                (display_width, display_height),
                Image.LANCZOS
            )
            
            # 转换为Tkinter格式
            self.photo_image = ImageTk.PhotoImage(self.display_image)
            
            # 显示在画布上
            self.canvas.config(scrollregion=(0, 0, display_width, display_height))
            self.canvas.create_image(0, 0, anchor='nw', image=self.photo_image)
            
            self.info_label.config(
                text=f"✅ 已加载第{self.page_index + 1}页 (原始:{img_width}x{img_height}, "
                     f"显示:{display_width}x{display_height})"
            )
            
            # 更新窗口标题
            self.window.title("模板制作工具")
            
        except Exception as e:
            error_msg = str(e)
            
            # 检查是否是 Poppler 相关错误
            if "poppler" in error_msg.lower() or "Unable to get page count" in error_msg:
                messagebox.showerror("Poppler 未安装", 
                    "❌ 无法渲染PDF：Poppler 工具未找到！\n\n"
                    "解决方法：\n"
                    "1. 下载 Poppler:\n"
                    "   https://github.com/oschwartz10612/poppler-windows/releases\n\n"
                    "2. 解压到项目的 vendor/poppler/ 目录\n\n"
                    "3. 在功能6中设置 Poppler 路径\n\n"
                    "详细说明请查看: INSTALL_POPPLER.md")
            else:
                messagebox.showerror("加载错误", 
                    f"无法渲染PDF:\n\n{e}\n\n"
                    "请检查:\n"
                    "1. PDF文件是否完整\n"
                    "2. 是否有足够内存\n"
                    "3. Poppler 是否正确安装")
            
            self.window.destroy()
    
    def _on_press(self, event):
        """鼠标按下"""
        # 记录起始点（画布坐标）
        x = self.canvas.canvasx(event.x)
        y = self.canvas.canvasy(event.y)
        self.rect_start = (x, y)
        
        # 删除旧的矩形
        if self.rect_id:
            self.canvas.delete(self.rect_id)
            self.rect_id = None
    
    def _on_drag(self, event):
        """鼠标拖拽"""
        if not self.rect_start:
            return
        
        # 获取当前位置
        x = self.canvas.canvasx(event.x)
        y = self.canvas.canvasy(event.y)
        
        # 删除旧的矩形
        if self.rect_id:
            self.canvas.delete(self.rect_id)
        
        # 绘制新的矩形
        x0, y0 = self.rect_start
        self.rect_id = self.canvas.create_rectangle(
            x0, y0, x, y,
            outline='#3498DB',
            width=3,
            dash=(5, 5)
        )
        
        # 更新坐标显示
        width = abs(x - x0)
        height = abs(y - y0)
        self.coord_label.config(
            text=f"选择中: {int(width)}x{int(height)} px"
        )
    
    def _on_release(self, event):
        """鼠标释放"""
        if not self.rect_start:
            return
        
        # 获取结束位置
        x = self.canvas.canvasx(event.x)
        y = self.canvas.canvasy(event.y)
        
        x0, y0 = self.rect_start
        
        # 确保坐标从左上到右下
        x1 = min(x0, x)
        y1 = min(y0, y)
        x2 = max(x0, x)
        y2 = max(y0, y)
        
        # 检查有效性
        if x2 - x1 < 10 or y2 - y1 < 10:
            self.info_label.config(text="⚠️ 选择区域太小，请重新选择")
            if self.rect_id:
                self.canvas.delete(self.rect_id)
                self.rect_id = None
            return
        
        # 保存坐标（显示坐标）
        self.rect_coords = (int(x1), int(y1), int(x2), int(y2))
        
        # 更新显示
        width = int(x2 - x1)
        height = int(y2 - y1)
        self.coord_label.config(
            text=f"✅ 已选择: {width}x{height} px"
        )
        
        self.info_label.config(
            text=f"✅ 区域已选择，点击'保存模板'继续"
        )
    
    def _reset_selection(self):
        """重置选择"""
        if self.rect_id:
            self.canvas.delete(self.rect_id)
            self.rect_id = None
        self.rect_start = None
        self.rect_coords = None
        self.coord_label.config(text="坐标: 未选择")
        self.info_label.config(text="提示：拖拽鼠标框选Logo或标题等固定特征区域")
    
    def _save_template(self):
        """保存模板"""
        if not self.rect_coords:
            messagebox.showwarning("提示", "请先拖拽鼠标选择模板区域")
            return
        
        # 请求保存路径
        file_path = filedialog.asksaveasfilename(
            title="保存模板",
            defaultextension=".png",
            filetypes=[
                ("PNG图片", "*.png"),
                ("所有文件", "*.*")
            ],
            initialfile="Template.png"
        )
        
        if not file_path:
            return
        
        try:
            # 转换坐标到原始图像
            x1, y1, x2, y2 = self.rect_coords
            
            # 缩放回原始坐标
            orig_x1 = int(x1 / self.canvas_scale)
            orig_y1 = int(y1 / self.canvas_scale)
            orig_x2 = int(x2 / self.canvas_scale)
            orig_y2 = int(y2 / self.canvas_scale)
            
            # 裁剪原始图像
            template = self.original_image.crop((orig_x1, orig_y1, orig_x2, orig_y2))
            
            # 预处理（与匹配时一致）
            # 1. 转换为灰度
            template = ImageOps.grayscale(template)
            
            # 2. 缩放到标准宽度（180px）
            w, h = template.size
            if w > 180:
                ratio = 180 / w
                new_h = int(h * ratio)
                template = template.resize((180, new_h), Image.LANCZOS)
            
            # 3. 保存为PNG
            template.save(file_path, 'PNG', optimize=True)
            
            # 显示预览
            preview_window = tk.Toplevel(self.window)
            preview_window.title("模板预览")
            preview_window.geometry("400x300")
            
            preview_frame = ttk.Frame(preview_window, padding=20)
            preview_frame.pack(fill='both', expand=True)
            
            ttk.Label(preview_frame, 
                     text="✅ 模板已保存", 
                     font=("Segoe UI", 14, "bold")).pack(pady=10)
            
            # 显示模板预览
            preview_img = ImageTk.PhotoImage(template)
            preview_label = ttk.Label(preview_frame, image=preview_img)
            preview_label.image = preview_img  # 保持引用
            preview_label.pack(pady=10)
            
            # 显示信息
            info_text = (
                f"文件: {file_path}\n"
                f"原始区域: {orig_x2-orig_x1}x{orig_y2-orig_y1} px\n"
                f"模板尺寸: {template.size[0]}x{template.size[1]} px\n"
                f"格式: 灰度PNG\n"
                f"DPI: {self.dpi}"
            )
            ttk.Label(preview_frame, text=info_text, justify='left').pack(pady=10)
            
            ttk.Button(preview_frame, text="确定", 
                      command=preview_window.destroy,
                      style='Accent.TButton').pack(pady=10)
            
            # 调用回调
            if self.callback:
                self.callback(file_path)
            
            self.info_label.config(text=f"✅ 模板已保存到: {file_path}")
            
        except Exception as e:
            messagebox.showerror("保存失败", f"保存模板时出错:\n{e}")


def open_template_maker(parent, pdf_path: str, dpi: int = 300, 
                       poppler_path: Optional[str] = None, 
                       page_index: int = 0,
                       callback: Optional[Callable] = None):
    """
    打开模板制作工具
    
    Args:
        parent: 父窗口
        pdf_path: PDF文件路径
        dpi: 渲染DPI
        poppler_path: Poppler路径
        page_index: 页面索引（0开始）
        callback: 保存完成回调
    """
    TemplateMaker(parent, pdf_path, dpi, poppler_path, page_index, callback)
