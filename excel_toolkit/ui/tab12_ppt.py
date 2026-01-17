"""
Tab12 - PPT转PDF功能
"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import os

from excel_toolkit.ppt_to_pdf import batch_ppt_to_pdf


class Tab12PptMixin:
    """Tab12 PPT转PDF Mixin"""
    
    def create_tab12_ppt(self, tab):
        """创建Tab12界面"""
        # 初始化变量
        if not hasattr(self, 'ppt_files_var'):
            self.ppt_files_var = tk.StringVar(value="未选择文件")
            self.ppt_outdir_var = tk.StringVar(value="与原文件相同")
            self._ppt_file_list = [] # 存储实际路径列表
            self._trace_persist(self.ppt_files_var)
            self._trace_persist(self.ppt_outdir_var)

        # 文件选择
        f1 = ttk.Frame(tab)
        f1.pack(fill='x', pady=5)
        ttk.Button(f1, text="选择 PPT/PPTX 文件", 
                  command=self._select_ppt_files).pack(side='left', padx=5)
        ttk.Label(f1, textvariable=self.ppt_files_var, wraplength=600).pack(side='left', padx=5)

        # 输出目录
        f2 = ttk.Frame(tab)
        f2.pack(fill='x', pady=5)
        ttk.Button(f2, text="选择输出目录 (可选)", 
                  command=self._select_ppt_outdir).pack(side='left', padx=5)
        ttk.Label(f2, textvariable=self.ppt_outdir_var).pack(side='left', padx=5)
        ttk.Button(f2, text="重置", command=lambda: self.ppt_outdir_var.set("与原文件相同")).pack(side='left', padx=5)

        # 说明
        tip_f = ttk.Frame(tab)
        tip_f.pack(fill='x', pady=5)
        ttk.Label(tip_f, text="提示: 转换功能需要电脑已安装 Microsoft PowerPoint。", 
                  foreground="#6b7280", font=("Segoe UI", 9)).pack(side='left', padx=5)

        # 执行按钮
        f3 = ttk.Frame(tab)
        f3.pack(fill='x', pady=10)
        ttk.Button(f3, text="[12] 开始批量转换", command=self.run_tool12, 
                  style='Accent.TButton').pack(side='left', padx=5)
        
        self.logger12, clear_log12 = self.create_log_widget(tab)
        ttk.Button(f3, text="清空日志", command=clear_log12, 
                  style='Secondary.TButton').pack(side='left', padx=5)

    def _select_ppt_files(self):
        paths = filedialog.askopenfilenames(
            title="选择 PPT/PPTX 文件",
            filetypes=[("PowerPoint 文件", "*.pptx *.ppt"), ("所有文件", "*.*")]
        )
        if paths:
            self._ppt_file_list = list(paths)
            if len(paths) == 1:
                self.ppt_files_var.set(os.path.basename(paths[0]))
            else:
                self.ppt_files_var.set(f"已选择 {len(paths)} 个文件")

    def _select_ppt_outdir(self):
        path = filedialog.askdirectory(title="选择输出目录")
        if path:
            self.ppt_outdir_var.set(path)

    def run_tool12(self):
        """执行 PPT 转 PDF"""
        if not self._ppt_file_list:
            messagebox.showwarning("⚠️ 警告", "请先选择需要转换的 PPT 文件。")
            return

        out_dir = self.ppt_outdir_var.get()
        if out_dir == "与原文件相同":
            out_dir = None
        
        self.logger12("=" * 50)
        self.logger12(f"▶️ 开始运行 [12] PPT 批量转 PDF...")
        self.logger12(f"文件数: {len(self._ppt_file_list)}")
        
        self._update_status("正在转换 PPT...", icon="⏳", show_progress=True)
        self.master.config(cursor="watch")
        
        def thread_target():
            try:
                def safe_logger(msg):
                    self.master.after(0, lambda m=msg: self.logger12(m))
                
                stats = batch_ppt_to_pdf(self._ppt_file_list, out_dir, safe_logger)
                
                def on_success():
                    self.master.config(cursor="")
                    self._update_status("就绪", icon="✅", show_progress=False)
                    
                    msg = (
                        f"转换完成！\n\n"
                        f"成功: {stats['success']}\n"
                        f"失败: {stats['fail']}\n"
                    )
                    messagebox.showinfo("✅ 完成", msg)
                    self.logger12(msg.replace("\n\n", "\n"))
                    if stats['success'] > 0:
                        # 尝试打开第一个转换成功的文件所在的文件夹
                        try:
                            os.startfile(os.path.dirname(stats['files'][0]))
                        except:
                            pass
                
                self.master.after(0, on_success)
                
            except Exception as e:
                error_msg = str(e)
                def on_error(msg=error_msg):
                    self.master.config(cursor="")
                    self._update_status("错误", icon="❌", show_progress=False)
                    messagebox.showerror("❌ 错误", f"PowerPoint 接口调用失败，请确保已安装 PowerPoint 且未被限制。\n错误: {msg}")
                    self.logger12(f"❌ 发生严重错误: {msg}")
                self.master.after(0, on_error)

        threading.Thread(target=thread_target, daemon=True).start()
