"""
Tab 7 - Prefix Fill View

Fills carrier names based on tracking number prefixes.
"""

import os
import threading
from typing import Optional

import flet as ft

from app.views.base_view import TabView
from app.core.theme import ThemeMode
from app.core.constants import Text, FontSize, Spacing

# Import business logic from existing module
import sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "../.."))

from excel_toolkit.prefix_fill import process_prefix_fill
from excel_toolkit.error_handler import log_error, get_user_friendly_error


class PrefixFillView(TabView):
    """
    Tab 7: Prefix Fill

    Fills carrier names based on tracking number prefixes:
    - '9' prefix → 'usps'
    - 'G' prefix → 'GOFO'
    - 'U' prefix → 'UniUni'
    """

    def get_tab_name(self) -> str:
        """Get the display name for this tab"""
        return "前缀填充"

    def get_tab_index(self) -> int:
        """Get the tab index"""
        return 6

    def build(self) -> ft.Control:
        """Build the tab content"""
        colors = self.get_colors()

        # Initialize variables
        self._file_path = self.load_file_path("main_file")
        self._src_col = self.load_preference("src_col", "A")
        self._dst_col = self.load_preference("dst_col", "B")

        # File picker
        file_picker = self.create_file_picker(
            label="Excel 文件",
            config_key="main_file",
            on_pick=self._on_file_picked,
        )

        # Source column input
        self.src_col_input = ft.TextField(
            label="源列号（含单号）",
            label_style=ft.TextStyle(size=FontSize.LABEL, color=colors["text_secondary"]),
            value=self._src_col,
            width=100,
            bgcolor=colors["surface"],
            border_color=colors["border"],
            focused_border_color=colors["primary"],
            text_style=ft.TextStyle(color=colors["on_surface"]),
            hint_text="如 A",
        )

        # Destination column input
        self.dst_col_input = ft.TextField(
            label="目标列号（填充承运商）",
            label_style=ft.TextStyle(size=FontSize.LABEL, color=colors["text_secondary"]),
            value=self._dst_col,
            width=100,
            bgcolor=colors["surface"],
            border_color=colors["border"],
            focused_border_color=colors["primary"],
            text_style=ft.TextStyle(color=colors["on_surface"]),
            hint_text="如 B",
        )

        # File selection section
        file_section = ft.Column([
            file_picker,
        ], spacing=Spacing.ROW_SPACING)

        # Parameters section with rules
        rules_content = ft.Column([
            ft.Text("填充规则：", style=ft.TextStyle(size=FontSize.LABEL, weight=ft.FontWeight.W_500)),
            ft.Text("• 首字符 '9' → 填充 'usps'", style=ft.TextStyle(size=FontSize.LABEL)),
            ft.Text("• 首字符 'G' → 填充 'GOFO'", style=ft.TextStyle(size=FontSize.LABEL)),
            ft.Text("• 首字符 'U' → 填充 'UniUni'", style=ft.TextStyle(size=FontSize.LABEL)),
        ], spacing=5)

        param_section = ft.Column([
            ft.Row([
                self.src_col_input,
                self.dst_col_input,
            ], spacing=Spacing.CONTROL_PADDING_X),
            ft.Container(
                content=rules_content,
                bgcolor=colors["info_container"],
                padding=ft.padding.all(12),
                border_radius=4,
            ),
        ], spacing=Spacing.ROW_SPACING)

        # Action buttons
        run_button = self.create_action_button(
            text="开始前缀填充",
            on_click=self._on_run_click,
            icon=ft.Icons.TEXT_FIELDS_OUTLINED,
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
        return self.build_standard_layout(
            file_section=file_section,
            param_section=param_section,
            action_section=action_section,
        )

    # ==================== Event Handlers ====================

    def _on_file_picked(self, path: str) -> None:
        """Handle file pick"""
        self._file_path = path
        self.log(f"已选择文件: {os.path.basename(path)}", "info")

    def _on_run_click(self, e) -> None:
        """Handle run button click"""
        # Get current values
        file_path = self._file_path or ""
        src_col = self.src_col_input.value or ""
        dst_col = self.dst_col_input.value or ""

        # Validate inputs
        if not self._validate_inputs(file_path, src_col, dst_col):
            return

        # Save preferences
        self._src_col = src_col
        self._dst_col = dst_col
        self.save_preference("src_col", src_col)
        self.save_preference("dst_col", dst_col)

        # Run processing in background thread
        self._run_processing(file_path, src_col, dst_col)

    def _on_clear_log_click(self, e) -> None:
        """Handle clear log button click"""
        if self._log_widget:
            self._log_widget.clear()
            self.log("日志已清空", "info")

    # ==================== Processing ====================

    def _validate_inputs(self, file_path: str, src_col: str, dst_col: str) -> bool:
        """Validate input parameters"""
        if not file_path or file_path == Text.FILE_NOT_SELECTED:
            self.show_error("验证错误", "请选择 Excel 文件！")
            return False

        if not src_col:
            self.show_error("验证错误", "请输入源列号！")
            return False

        if not dst_col:
            self.show_error("验证错误", "请输入目标列号！")
            return False

        return True

    def _run_processing(self, file_path: str, src_col: str, dst_col: str) -> None:
        """Run prefix fill in background thread"""
        self.set_processing("正在处理...")
        self.page.cursor = ft.Cursor.WAIT

        # Log start
        self.log("=" * 60, "info")
        self.log("开始执行前缀填充...", "info")
        self.log(f"  文件: {os.path.basename(file_path)}", "info")
        self.log(f"  源列: {src_col}", "info")
        self.log(f"  目标列: {dst_col}", "info")
        self.log("=" * 60, "info")

        def process_thread():
            try:
                # Create thread-safe logger
                def thread_safe_logger(msg: str):
                    self.log(msg, "info")

                # Process prefix fill
                result = process_prefix_fill(
                    file_path,
                    src_col,
                    dst_col,
                    thread_safe_logger
                )

                # Success callback
                def on_success():
                    self.page.cursor = ft.Cursor.DEFAULT
                    self.set_success("填充完成")

                    self.log("", "info")
                    self.log("=" * 60, "success")
                    self.log("前缀填充完成！", "success")
                    self.log("=" * 60, "success")

                    # Show success dialog
                    self.show_success(result)

                self.page.run_thread(on_success)

            except Exception as e:
                # Error handling
                log_error(e, "前缀填充")
                error_msg = get_user_friendly_error(e)

                def on_error():
                    self.page.cursor = ft.Cursor.DEFAULT
                    self.set_error("填充失败")

                    self.log(f"", "info")
                    self.log(f"❌ 发生错误: {e}", "error")
                    self.log(f"💡 请查看日志文件获取详细信息", "warning")

                    self.show_error("填充失败", error_msg)

                self.page.run_thread(on_error)

        # Start processing thread
        threading.Thread(target=process_thread, daemon=True).start()


def create_view(page: ft.Page) -> PrefixFillView:
    """
    Factory function to create the prefix fill view.

    Args:
        page: Flet page control

    Returns:
        PrefixFillView instance
    """
    view = PrefixFillView(page, tab_index=6)
    return view
