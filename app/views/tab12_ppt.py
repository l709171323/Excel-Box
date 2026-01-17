"""
Tab 12 - PPT to PDF View

Batch converts PowerPoint files to PDF format.
"""

import os
import threading
from typing import Optional, List

import flet as ft

from app.views.base_view import TabView
from app.core.theme import ThemeMode
from app.core.constants import Text, FontSize, Spacing

# Import business logic from existing module
import sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "../.."))

from excel_toolkit.ppt_to_pdf import batch_ppt_to_pdf
from excel_toolkit.error_handler import log_error, get_user_friendly_error


class PptToPdfView(TabView):
    """
    Tab 12: PPT to PDF

    Batch converts PowerPoint (.ppt, .pptx) files to PDF format.
    Requires Microsoft PowerPoint to be installed.
    """

    def get_tab_name(self) -> str:
        """Get the display name for this tab"""
        return "PPT转PDF"

    def get_tab_index(self) -> int:
        """Get the tab index"""
        return 11

    def build(self) -> ft.Control:
        """Build the tab content"""
        colors = self.get_colors()

        # Initialize variables
        self._file_list: List[str] = []
        self._output_dir = self.load_preference("output_dir", "")

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

        # Select files button
        select_files_button = ft.ElevatedButton(
            "选择 PPT/PPTX 文件",
            icon=ft.Icons.FILE_UPLOAD,
            on_click=lambda _: self.file_picker.pick_files(
                allow_multiple=True,
                allowed_extensions=["ppt", "pptx"]
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
            ft.Container(
                content=ft.Column([
                    ft.Text("说明：", style=ft.TextStyle(size=FontSize.LABEL, weight=ft.FontWeight.W_500)),
                    ft.Text("• 支持批量转换 .ppt 和 .pptx 文件", style=ft.TextStyle(size=FontSize.LABEL)),
                    ft.Text("• 需要安装 Microsoft PowerPoint", style=ft.TextStyle(size=FontSize.LABEL)),
                    ft.Text("• 输出目录默认与原文件相同", style=ft.TextStyle(size=FontSize.LABEL)),
                ], spacing=5),
                bgcolor=colors["warning_container"],
                padding=ft.padding.all(12),
                border_radius=4,
            ),
        ], spacing=Spacing.ROW_SPACING)

        # Action buttons
        run_button = self.create_action_button(
            text="开始转换",
            on_click=self._on_run_click,
            icon=ft.Icons.SLIDESHOW_OUTLINED,
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
                self.file_list_display.text = f"已选择 {count} 个文件"

            self.file_count_text.text = f"({count} 个文件)"
            self.file_list_display.color = self.get_colors()["text_primary"]

            self.file_list_display.update()
            self.file_count_text.update()

            self.log(f"已选择 {count} 个文件", "info")

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

    def _on_run_click(self, e) -> None:
        """Handle run button click"""
        # Validate inputs
        if not self._file_list:
            self.show_error("验证错误", "请先选择 PPT/PPTX 文件！")
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
        """Run PPT to PDF conversion in background thread"""
        self.set_processing("正在转换...")
        self.page.cursor = ft.Cursor.WAIT

        output_dir = self._output_dir if self._output_dir else None

        # Log start
        self.log("=" * 60, "info")
        self.log("开始执行 PPT 批量转 PDF...", "info")
        self.log(f"  文件数: {len(self._file_list)}", "info")
        self.log(f"  输出目录: {output_dir if output_dir else '与原文件相同'}", "info")
        self.log("=" * 60, "info")

        def process_thread():
            try:
                # Create thread-safe logger
                def thread_safe_logger(msg: str):
                    self.log(msg, "info")

                # Process conversion
                stats = batch_ppt_to_pdf(
                    self._file_list,
                    output_dir,
                    thread_safe_logger
                )

                # Success callback
                def on_success():
                    self.page.cursor = ft.Cursor.DEFAULT
                    self.set_success("转换完成")

                    self.log("", "info")
                    self.log("=" * 60, "success")
                    self.log("转换完成！", "success")
                    self.log(f"  成功: {stats['success']}", "info")
                    self.log(f"  失败: {stats['fail']}", "error" if stats['fail'] > 0 else "info")
                    self.log("=" * 60, "success")

                    # Show success dialog
                    self.show_success(
                        f"转换完成！\n\n"
                        f"成功: {stats['success']}\n"
                        f"失败: {stats['fail']}"
                    )

                    # Try to open output folder if there were successful conversions
                    if stats['success'] > 0 and stats['files']:
                        try:
                            import subprocess
                            output_folder = os.path.dirname(stats['files'][0])
                            if os.name == 'nt':  # Windows
                                subprocess.Popen(['explorer', output_folder])
                        except Exception:
                            pass

                self.page.run_thread(on_success)

            except Exception as e:
                # Error handling
                log_error(e, "PPT转PDF")
                error_msg = get_user_friendly_error(e)

                def on_error():
                    self.page.cursor = ft.Cursor.DEFAULT
                    self.set_error("转换失败")

                    self.log(f"", "info")
                    self.log(f"❌ 发生错误: {e}", "error")
                    self.log(f"💡 请确保已安装 Microsoft PowerPoint", "warning")

                    self.show_error(
                        "转换失败",
                        f"PowerPoint 接口调用失败。\n\n请确保已安装 PowerPoint 且未被限制。\n\n错误: {e}"
                    )

                self.page.run_thread(on_error)

        # Start processing thread
        threading.Thread(target=process_thread, daemon=True).start()


def create_view(page: ft.Page) -> PptToPdfView:
    """
    Factory function to create the PPT to PDF view.

    Args:
        page: Flet page control

    Returns:
        PptToPdfView instance
    """
    view = PptToPdfView(page, tab_index=11)
    return view
