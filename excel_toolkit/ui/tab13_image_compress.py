# """
# Tab13 - Image Compression Feature
# """

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
from PIL import Image

class Tab13ImageCompressMixin:
    """Mixin providing UI and logic for batch image compression.
    Supports JPEG, PNG, WebP formats with adjustable quality.
    """

    def create_tab13_image_compress(self, tab):
        """Create UI elements for the Image Compression tab."""
        # Initialize variables if not already present
        if not hasattr(self, "image_files_var"):
            self.image_files_var = tk.StringVar(value="未选择文件")
            self.image_outdir_var = tk.StringVar(value="与原文件相同")
            self._image_file_list = []  # store actual paths
            self.quality_var = tk.IntVar(value=85)  # default quality
            self.quality_display_var = tk.StringVar(value="85")  # for display
            self.format_var = tk.StringVar(value="JPEG")
            self._trace_persist(self.image_files_var)
            self._trace_persist(self.image_outdir_var)
            self._trace_persist(self.quality_var)
            self._trace_persist(self.format_var)
            # Update display when quality changes
            self.quality_var.trace_add("write", lambda *args: self.quality_display_var.set(str(self.quality_var.get())))

        # File selection frame
        f1 = ttk.Frame(tab)
        f1.pack(fill="x", pady=5)
        ttk.Button(
            f1,
            text="选择图片文件",
            command=self._select_image_files,
        ).pack(side="left", padx=5)
        ttk.Label(f1, textvariable=self.image_files_var, wraplength=600).pack(
            side="left", padx=5
        )

        # Output directory frame
        f2 = ttk.Frame(tab)
        f2.pack(fill="x", pady=5)
        ttk.Button(
            f2,
            text="选择输出目录 (可选)",
            command=self._select_image_outdir,
        ).pack(side="left", padx=5)
        ttk.Label(f2, textvariable=self.image_outdir_var).pack(side="left", padx=5)
        ttk.Button(
            f2,
            text="重置",
            command=lambda: self.image_outdir_var.set("与原文件相同"),
        ).pack(side="left", padx=5)

        # Compression options frame
        opt_f = ttk.Frame(tab)
        opt_f.pack(fill="x", pady=5)
        ttk.Label(opt_f, text="质量 (1-100):").pack(side="left", padx=5)
        ttk.Scale(
            opt_f,
            from_=1,
            to=100,
            orient="horizontal",
            variable=self.quality_var,
            length=150,
        ).pack(side="left", padx=5)
        # Real-time quality value display
        ttk.Label(opt_f, textvariable=self.quality_display_var, font=("Segoe UI", 10, "bold"), foreground="#3b82f6").pack(side="left", padx=5)
        ttk.Label(opt_f, text="格式:").pack(side="left", padx=(15, 5))
        fmt_cb = ttk.Combobox(
            opt_f,
            textvariable=self.format_var,
            values=["JPEG", "PNG", "WebP"],
            state="readonly",
            width=6,
        )
        fmt_cb.pack(side="left", padx=5)

        # Run button frame
        f3 = ttk.Frame(tab)
        f3.pack(fill="x", pady=10)
        ttk.Button(
            f3,
            text="[13] 开始压缩",
            command=self.run_image_compress,
            style="Accent.TButton",
        ).pack(side="left", padx=5)
        self.logger13, clear_log13 = self.create_log_widget(tab)
        ttk.Button(
            f3,
            text="清空日志",
            command=clear_log13,
            style="Secondary.TButton",
        ).pack(side="left", padx=5)

    def _select_image_files(self):
        paths = filedialog.askopenfilenames(
            title="选择图片文件",
            filetypes=[
                ("Image Files", "*.jpg *.jpeg *.png *.webp *.bmp *.tiff"),
                ("All Files", "*.*"),
            ],
        )
        if paths:
            self._image_file_list = list(paths)
            if len(paths) == 1:
                self.image_files_var.set(os.path.basename(paths[0]))
            else:
                self.image_files_var.set(f"已选择 {len(paths)} 张图片")

    def _select_image_outdir(self):
        path = filedialog.askdirectory(title="选择输出目录")
        if path:
            self.image_outdir_var.set(path)

    def run_image_compress(self):
        """Execute batch image compression in a background thread."""
        if not self._image_file_list:
            messagebox.showwarning("⚠️ 警告", "请先选择需要压缩的图片文件。")
            return
        out_dir = self.image_outdir_var.get()
        if out_dir == "与原文件相同":
            out_dir = None
        fmt = self.format_var.get().lower()
        quality = self.quality_var.get()

        self.logger13("=" * 50)
        self.logger13(f"▶️ 开始批量压缩图片 ({fmt.upper()}, 质量={quality}) ...")
        self.logger13(f"文件数: {len(self._image_file_list)}")
        self._update_status("正在压缩图片...", icon="⏳", show_progress=True)
        self.master.config(cursor="watch")

        def thread_target():
            try:
                success = 0
                fail = 0
                total = len(self._image_file_list)
                for idx, src_path in enumerate(self._image_file_list, 1):
                    try:
                        # Update progress in status bar
                        def update_progress(i=idx, t=total, f=os.path.basename(src_path)):
                            self._update_status(f"正在压缩 ({i}/{t}): {f}", icon="⏳", show_progress=True)
                        self.master.after(0, update_progress)
                        
                        # Log progress
                        def log_progress(i=idx, t=total, f=src_path):
                            self.logger13(f"[{i}/{t}] 正在处理: {f}")
                        self.master.after(0, log_progress)
                        
                        img = Image.open(src_path)
                        # Determine output path
                        base_name = os.path.splitext(os.path.basename(src_path))[0]
                        out_name = f"{base_name}.{fmt}"
                        out_path = os.path.join(out_dir or os.path.dirname(src_path), out_name)
                        save_kwargs = {}
                        if fmt == "jpeg":
                            save_kwargs["quality"] = quality
                            save_kwargs["optimize"] = True
                        elif fmt == "webp":
                            save_kwargs["quality"] = quality
                        elif fmt == "png":
                            # PNG uses compression level 0-9; map quality 1-100 to 9-0
                            level = max(0, min(9, 9 - int((quality - 1) / 11)))
                            save_kwargs["compress_level"] = level
                        img.convert("RGB").save(out_path, fmt.upper(), **save_kwargs)
                        success += 1
                        
                        # Get file sizes for logging
                        orig_size = os.path.getsize(src_path) / 1024  # KB
                        new_size = os.path.getsize(out_path) / 1024  # KB
                        ratio = ((orig_size - new_size) / orig_size * 100) if orig_size > 0 else 0
                        
                        def log_success(p=out_path, os=orig_size, ns=new_size, r=ratio):
                            self.logger13(f"✅ 完成: {p} ({os:.1f}KB -> {ns:.1f}KB, 减少{r:.1f}%)")
                        self.master.after(0, log_success)
                    except Exception as e:
                        fail += 1
                        def log_error(p=src_path, err=str(e)):
                            self.logger13(f"❌ {p}: {err}")
                        self.master.after(0, log_error)
                def on_success():
                    self.master.config(cursor="")
                    self._update_status("就绪", icon="✅", show_progress=False)
                    msg = f"压缩完成！\n成功: {success}\n失败: {fail}"
                    messagebox.showinfo("✅ 完成", msg)
                    self.logger13(msg)
                    if success > 0:
                        try:
                            # Open folder of first output file
                            first_out = os.path.join(out_dir or os.path.dirname(self._image_file_list[0]), f"{os.path.splitext(os.path.basename(self._image_file_list[0]))[0]}.{fmt}")
                            os.startfile(os.path.dirname(first_out))
                        except Exception:
                            pass
                self.master.after(0, on_success)
            except Exception as e:
                err_msg = str(e)
                def on_error():
                    self.master.config(cursor="")
                    self._update_status("错误", icon="❌", show_progress=False)
                    messagebox.showerror("❌ 错误", f"压缩过程中出现错误:\n{err_msg}")
                    self.logger13(f"❌ 发生错误: {err_msg}")
                self.master.after(0, on_error)

        threading.Thread(target=thread_target, daemon=True).start()
