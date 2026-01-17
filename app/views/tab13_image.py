"""
Tab 13 - Image Compression View

Batch compresses images with adjustable quality and format conversion.
"""

import os
import threading
from typing import Optional, List

import flet as ft

from app.views.base_view import TabView
from app.core.theme import ThemeMode
from app.core.constants import Text, FontSize, Spacing

# Import business logic
import sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "../.."))

try:
    from PIL import Image
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False
    Image = None

from excel_toolkit.error_handler import log_error, get_user_friendly_error


class ImageCompressView(TabView):
    """
    Tab 13: Image Compression

    Batch compresses images with adjustable quality settings.
    Supports JPEG, PNG, and WebP formats.
    """

    def get_tab_name(self) -> str:
        """Get the display name for this tab"""
        return "图片压缩"

    def get_tab_index(self) -> int:
        """Get the tab index"""
        return 12

    def build(self) -> ft.Control:
        """Build the tab content"""
        colors = self.get_colors()

        # Check if PIL is available
        if not PIL_AVAILABLE:
            return ft.Container(
                content=ft.Column([
                    ft.Icon(ft.Icons.ERROR, size=48, color=colors["error"]),
                    ft.Text(
                        "缺少依赖",
                        style=ft.TextStyle(size=FontSize.SUBTITLE, color=colors["text_primary"]),
                    ),
                    ft.Text(
                        "请安装 Pillow 库: pip install Pillow",
                        style=ft.TextStyle(size=FontSize.LABEL, color=colors["text_secondary"]),
                    ),
                ], horizontal_alignment=ft.CrossAxisAlignment.CENTER),
                alignment=ft.alignment.center,
                expand=True,
            )

        # Initialize variables
        self._file_list: List[str] = []
        self._output_dir = self.load_preference("output_dir", "")
        self._quality = self.load_preference("quality", 85)
        self._format = self.load_preference("format", "JPEG")

        # Multi-file picker button
        self.file_picker = ft.FilePicker(on_result=self._on_files_picked)

        # Directory picker button
        self.dir_picker = ft.FilePicker(on_result=self._on_dir_picked)

        # File list display
        self.file_list_display = ft.Text(
            "未选择文件",
            style=ft.TextStyle(size=FontSize.LABEL, color=colors["text_secondary"]),
            max_lines=2,
            overflow=ft.TextOverflow.ELLIPSIS,
        )

        # Output directory display
        self.output_dir_display = ft.Text(
            "与原文件相同",
            style=ft.TextStyle(size=FontSize.LABEL, color=colors["text_secondary"]),
        )

        # Quality slider
        self.quality_slider = ft.Slider(
            min=1,
            max=100,
            value=self._quality,
            label="质量",
            divisions=99,
            on_change=self._on_quality_change,
        )

        self.quality_display = ft.Text(
            str(self._quality),
            style=ft.TextStyle(size=FontSize.SUBTITLE, weight=ft.FontWeight.BOLD, color=colors["primary"]),
            width=40,
            text_align=ft.TextAlign.CENTER,
        )

        # Format dropdown
        self.format_dropdown = ft.Dropdown(
            label="输出格式",
            label_style=ft.TextStyle(size=FontSize.LABEL, color=colors["text_secondary"]),
            options=[
                ft.dropdown.Option("JPEG"),
                ft.dropdown.Option("PNG"),
                ft.dropdown.Option("WebP"),
            ],
            value=self._format,
            width=120,
            bgcolor=colors["surface"],
            border_color=colors["border"],
            focused_border_color=colors["primary"],
            text_style=ft.TextStyle(color=colors["on_surface"]),
            on_change=self._on_format_change,
        )

        # Select files button
        select_files_button = ft.ElevatedButton(
            "选择图片文件",
            icon=ft.Icons.FILE_UPLOAD,
            on_click=lambda _: self.file_picker.pick_files(
                allow_multiple=True,
                allowed_extensions=["jpg", "jpeg", "png", "webp", "bmp", "tiff"]
            ),
            style=ft.ButtonStyle(
                bgcolor=colors["primary"],
                color=colors["on_primary"],
            ),
        )

        # Select output directory button
        select_dir_button = ft.ElevatedButton(
            "选择输出目录（可选）",
            icon=ft.Icons.FOLDER_OPEN,
            on_click=lambda _: self.dir_picker.get_directory_path(),
            style=ft.ButtonStyle(
                bgcolor=colors["secondary_container"],
                color=colors["on_surface"],
            ),
        )

        # Reset directory button
        reset_button = ft.TextButton(
            "重置",
            on_click=self._on_reset_output_dir,
        )

        # File count display
        self.file_count_text = ft.Text(
            "",
            style=ft.TextStyle(size=9, color=colors["text_hint"]),
        )

        # File selection section
        file_section = ft.Column([
            ft.Row([
                select_files_button,
                self.file_count_text,
            ], spacing=Spacing.CONTROL_PADDING_X),
            ft.Container(
                content=self.file_list_display,
                bgcolor=colors["surface_variant"],
                border_radius=4,
                padding=ft.padding.symmetric(horizontal=10, vertical=8),
            ),
            ft.Row([
                select_dir_button,
                self.output_dir_display,
                ft.Container(expand=True),
                reset_button,
            ], spacing=Spacing.CONTROL_PADDING_X, alignment=ft.MainAxisAlignment.CENTER),
        ], spacing=Spacing.ROW_SPACING)

        # Parameters section
        param_section = ft.Column([
            ft.Row([
                ft.Text("压缩设置:", style=ft.TextStyle(size=FontSize.SECTION, weight=ft.FontWeight.BOLD)),
                ft.Container(expand=True),
            ], spacing=Spacing.CONTROL_PADDING_X),
            ft.Row([
                self.quality_slider,
                self.quality_display,
            ], spacing=Spacing.CONTROL_PADDING_X, alignment=ft.MainAxisAlignment.CENTER),
            ft.Row([
                self.format_dropdown,
            ], spacing=Spacing.CONTROL_PADDING_X),
            ft.Container(
                content=ft.Column([
                    ft.Text("说明：", style=ft.TextStyle(size=FontSize.LABEL, weight=ft.FontWeight.W_500)),
                    ft.Text("• 质量: 1-100，数值越小文件越小", style=ft.TextStyle(size=FontSize.LABEL)),
                    ft.Text("• JPEG: 有损压缩，适合照片", style=ft.TextStyle(size=FontSize.LABEL)),
                    ft.Text("• PNG: 无损压缩，适合图形", style=ft.TextStyle(size=FontSize.LABEL)),
                    ft.Text("• WebP: 现代格式，高压缩率", style=ft.TextStyle(size=FontSize.LABEL)),
                ], spacing=5),
                bgcolor=colors["info_container"],
                padding=ft.padding.all(12),
                border_radius=4,
            ),
        ], spacing=Spacing.ROW_SPACING)

        # Action buttons
        run_button = self.create_action_button(
            text="开始压缩",
            on_click=self._on_run_click,
            icon=ft.Icons.IMAGE_OUTLINED,
            variant="primary",
        )

        clear_button = self.create_action_button(
            text="清空日志",
            on_click=self._on_clear_log_click,
            icon=ft.Icons.CLEAR,
            variant="secondary",
        )

        action_section = ft.Row([
            run_button,
            clear_button,
        ], spacing=Spacing.BUTTON_SPACING)

        # Build standard layout
        content = self.build_standard_layout(
            file_section=file_section,
            param_section=param_section,
            action_section=action_section,
        )

        # Add file pickers to the page
        self.page.overlay.append(self.file_picker)
        self.page.overlay.append(self.dir_picker)

        return content

    # ==================== Event Handlers ====================

    def _on_files_picked(self, e: ft.FilePickerResultEvent) -> None:
        """Handle files pick"""
        if e.files:
            self._file_list = [f.path for f in e.files]
            count = len(self._file_list)

            if count == 1:
                self.file_list_display.text = os.path.basename(self._file_list[0])
            else:
                self.file_list_display.text = f"已选择 {count} 张图片"

            self.file_count_text.text = f"({count} 张)"
            self.file_list_display.color = self.get_colors()["text_primary"]

            self.file_list_display.update()
            self.file_count_text.update()

            self.log(f"已选择 {count} 张图片", "info")

    def _on_dir_picked(self, e: ft.FilePickerResultEvent) -> None:
        """Handle directory pick"""
        if e.path:
            self._output_dir = e.path
            self.output_dir_display.text = e.path
            self.output_dir_display.color = self.get_colors()["text_primary"]
            self.save_preference("output_dir", e.path)

            self.output_dir_display.update()
            self.log(f"输出目录: {e.path}", "info")

    def _on_reset_output_dir(self, e) -> None:
        """Reset output directory to default"""
        self._output_dir = ""
        self.output_dir_display.text = "与原文件相同"
        self.output_dir_display.color = self.get_colors()["text_secondary"]
        self.save_preference("output_dir", "")
        self.output_dir_display.update()
        self.log("输出目录已重置", "info")

    def _on_quality_change(self, e) -> None:
        """Handle quality slider change"""
        self._quality = int(e.control.value)
        self.quality_display.text = str(self._quality)
        self.quality_display.update()
        self.save_preference("quality", self._quality)

    def _on_format_change(self, e) -> None:
        """Handle format dropdown change"""
        self._format = e.control.value
        self.save_preference("format", self._format)

    def _on_run_click(self, e) -> None:
        """Handle run button click"""
        # Validate inputs
        if not self._file_list:
            self.show_error("验证错误", "请先选择图片文件！")
            return

        # Run processing in background thread
        self._run_processing()

    def _on_clear_log_click(self, e) -> None:
        """Handle clear log button click"""
        if self._log_widget:
            self._log_widget.clear()
            self.log("日志已清空", "info")

    # ==================== Processing ====================

    def _run_processing(self) -> None:
        """Run image compression in background thread"""
        self.set_processing("正在压缩...")
        self.page.cursor = ft.Cursor.WAIT

        output_dir = self._output_dir if self._output_dir else None
        fmt = self._format.lower()
        quality = self._quality

        # Log start
        self.log("=" * 60, "info")
        self.log("开始批量压缩图片...", "info")
        self.log(f"  格式: {fmt.upper()}, 质量: {quality}", "info")
        self.log(f"  文件数: {len(self._file_list)}", "info")
        self.log("=" * 60, "info")

        def process_thread():
            try:
                success = 0
                fail = 0
                total = len(self._file_list)

                for idx, src_path in enumerate(self._file_list, 1):
                    try:
                        # Update progress
                        progress_msg = f"正在压缩 ({idx}/{total}): {os.path.basename(src_path)}"
                        self.log(progress_msg, "info")
                        self.update_progress((idx - 1) / total * 100, progress_msg)

                        # Process image
                        img = Image.open(src_path)

                        # Determine output path
                        base_name = os.path.splitext(os.path.basename(src_path))[0]
                        out_name = f"{base_name}.{fmt}"
                        out_path = os.path.join(
                            output_dir or os.path.dirname(src_path),
                            out_name
                        )

                        # Save with compression
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

                        # Get file sizes
                        orig_size = os.path.getsize(src_path) / 1024  # KB
                        new_size = os.path.getsize(out_path) / 1024  # KB
                        ratio = ((orig_size - new_size) / orig_size * 100) if orig_size > 0 else 0

                        self.log(
                            f"✅ 完成: {os.path.basename(out_path)} ({orig_size:.1f}KB -> {new_size:.1f}KB, 减少{ratio:.1f}%)",
                            "success"
                        )

                    except Exception as e:
                        fail += 1
                        self.log(f"❌ {os.path.basename(src_path)}: {e}", "error")

                # Success callback
                def on_success():
                    self.page.cursor = ft.Cursor.DEFAULT
                    self.set_success("压缩完成")

                    self.log("", "info")
                    self.log("=" * 60, "success")
                    self.log("压缩完成！", "success")
                    self.log(f"  成功: {success}", "info")
                    self.log(f"  失败: {fail}", "error" if fail > 0 else "info")
                    self.log("=" * 60, "success")

                    # Show success dialog
                    self.show_success(f"压缩完成！\n\n成功: {success}\n失败: {fail}")

                    # Try to open output folder if there were successful conversions
                    if success > 0:
                        try:
                            import subprocess
                            first_out = os.path.join(
                                output_dir or os.path.dirname(self._file_list[0]),
                                f"{os.path.splitext(os.path.basename(self._file_list[0]))[0]}.{fmt}"
                            )
                            output_folder = os.path.dirname(first_out)
                            if os.name == 'nt':  # Windows
                                subprocess.Popen(['explorer', output_folder])
                        except Exception:
                            pass

                self.page.run_thread(on_success)

            except Exception as e:
                # Error handling
                log_error(e, "图片压缩")
                error_msg = get_user_friendly_error(e)

                def on_error():
                    self.page.cursor = ft.Cursor.DEFAULT
                    self.set_error("压缩失败")

                    self.log(f"", "info")
                    self.log(f"❌ 发生错误: {e}", "error")
                    self.log(f"💡 请查看日志文件获取详细信息", "warning")

                    self.show_error("压缩失败", error_msg)

                self.page.run_thread(on_error)

        # Start processing thread
        threading.Thread(target=process_thread, daemon=True).start()


def create_view(page: ft.Page) -> ImageCompressView:
    """
    Factory function to create the image compression view.

    Args:
        page: Flet page control

    Returns:
        ImageCompressView instance
    """
    view = ImageCompressView(page, tab_index=12)
    return view
