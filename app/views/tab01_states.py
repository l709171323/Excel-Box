"""
Tab 1 - State Name Conversion View

Converts US state full names to two-letter abbreviations in Excel files.
"""

import os
import threading
from typing import Optional

import flet as ft

from app.views.base_view import TabView
from app.core.theme import ThemeMode
from app.core.constants import Text, FontSize, Spacing
from app.core.constants import Icon
from app.components.common import InputField


# Import business logic from existing module
import sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "../.."))

from excel_toolkit.states import process_states
from excel_toolkit.exceptions import ExcelToolkitError
from excel_toolkit.error_handler import get_user_friendly_error, log_error


class StateConversionView(TabView):
    """
    Tab 1: State Name Conversion

    Converts US state full names to two-letter abbreviations in Excel files.
    """

    def get_tab_name(self) -> str:
        """Get the display name for this tab"""
        return "州名转换"

    def get_tab_index(self) -> int:
        """Get the tab index"""
        return 0

    def build(self) -> ft.Control:
        """Build the tab content"""
        colors = self.get_colors()

        # Initialize variables
        self._file_path = self.load_file_path("main_file")
        self._sheet_name = self.load_preference("sheet_name", "")
        self._column_letter = self.load_preference("column_letter", "G")
        self._sheet_names: list[str] = []

        # File picker
        file_picker = self.create_file_picker(
            label="Excel 文件",
            config_key="main_file",
            on_pick=self._on_file_picked,
            file_type="xlsx",
        )

        # Sheet dropdown
        self.sheet_dropdown = ft.Dropdown(
            label="工作表",
            label_style=ft.TextStyle(size=FontSize.LABEL, color=colors["text_secondary"]),
            options=[],
            value=self._sheet_name if self._sheet_name else None,
            width=200,
            bgcolor=colors["surface"],
            border_color=colors["border"],
            focused_border_color=colors["primary"],
            text_style=ft.TextStyle(color=colors["on_surface"]),
            on_change=self._on_sheet_changed,
        )

        # Column input
        self.column_input = ft.TextField(
            label="目标列",
            label_style=ft.TextStyle(size=FontSize.LABEL, color=colors["text_secondary"]),
            value=self._column_letter,
            width=100,
            bgcolor=colors["surface"],
            border_color=colors["border"],
            focused_border_color=colors["primary"],
            text_style=ft.TextStyle(color=colors["on_surface"]),
            hint_text="如 G",
        )

        # File selection section
        file_section = ft.Column([
            file_picker,
        ], spacing=Spacing.ROW_SPACING)

        # Parameters section
        param_section = ft.Column([
            ft.Row([
                self.sheet_dropdown,
                ft.Text("提示：选择包含州名数据的工作表", style=ft.TextStyle(size=9, color=colors["text_hint"])),
            ], spacing=Spacing.CONTROL_PADDING_X),
            ft.Row([
                self.column_input,
                ft.Text("提示：输入要转换的列号（如 G）", style=ft.TextStyle(size=9, color=colors["text_hint"])),
            ], spacing=Spacing.CONTROL_PADDING_X),
            ft.Container(
                content=ft.Text(
                    "ℹ️ 程序会将选中列的州全名转换为两字母缩写（如 California → CA）",
                    style=ft.TextStyle(size=FontSize.LABEL, color=colors["text_secondary"]),
                ),
                bgcolor=colors["info_container"],
                padding=ft.padding.all(8),
                border_radius=4,
            ),
        ], spacing=Spacing.ROW_SPACING)

        # Action buttons
        run_button = self.create_action_button(
            text="开始转换",
            on_click=self._on_run_click,
            icon=ft.Icons.PLAY_ARROW,
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
        # TEMPORARY: Disable ALL processing to test if file picker dialog closes properly
        print(f"[DEBUG] File picked: {path}")
        self._file_path = path
        # Don't do anything else - just store the path

    def _on_sheet_changed(self, e) -> None:
        """Handle sheet dropdown change"""
        self._sheet_name = e.control.value
        self.save_preference("sheet_name", self._sheet_name)

    def _on_run_click(self, e) -> None:
        """Handle run button click"""
        # Get current values
        file_path = self._file_path or ""
        sheet_name = self.sheet_dropdown.value or ""
        column_letter = self.column_input.value or ""

        # Validate inputs
        if not self._validate_inputs(file_path, sheet_name, column_letter):
            return

        # Save preferences
        self._column_letter = column_letter
        self.save_preference("column_letter", column_letter)

        # Run processing in background thread
        self._run_processing(file_path, sheet_name, column_letter)

    def _on_clear_log_click(self, e) -> None:
        """Handle clear log button click"""
        if self._log_widget:
            self._log_widget.clear()
            self.log("日志已清空", "info")

    # ==================== Processing ====================

    def _validate_inputs(self, file_path: str, sheet_name: str, column_letter: str) -> bool:
        """Validate input parameters"""
        if not file_path or file_path == Text.FILE_NOT_SELECTED:
            self.show_error("验证错误", "请选择 Excel 文件！")
            return False

        if not sheet_name:
            self.show_error("验证错误", "请选择工作表！")
            return False

        if not column_letter:
            self.show_error("验证错误", "请输入目标列号！")
            return False

        return True

    def _run_processing(self, file_path: str, sheet_name: str, column_letter: str) -> None:
        """Run state name conversion in background thread"""
        self.set_processing("正在转换州名...")
        self.page.cursor = ft.Cursor.WAIT

        # Log start
        self.log("=" * 60, "info")
        self.log("开始执行州名转换...", "info")
        self.log(f"  文件: {os.path.basename(file_path)}", "info")
        self.log(f"  工作表: {sheet_name}", "info")
        self.log(f"  目标列: {column_letter}", "info")
        self.log("=" * 60, "info")

        def process_thread():
            try:
                # Create thread-safe logger
                def thread_safe_logger(msg: str):
                    self.log(msg, "info")

                # Process states
                stats = process_states(
                    file_path,
                    sheet_name,
                    column_letter,
                    thread_safe_logger
                )

                # Success callback
                def on_success():
                    self.page.cursor = ft.Cursor.DEFAULT
                    self.set_success("转换完成")

                    # Log results
                    self.log("", "info")
                    self.log("=" * 60, "success")
                    self.log("转换完成！", "success")
                    self.log(f"  总计: {stats['total']} 行", "info")
                    self.log(f"  成功: {stats['success']} 行", "success")
                    self.log(f"  跳过: {stats['failed']} 行", "warning")
                    self.log("=" * 60, "success")

                    # Show success dialog
                    self.show_success(
                        f"州名转换完成！\n\n"
                        f"总共处理: {stats['total']} 行\n"
                        f"成功转换: {stats['success']} 行\n"
                        f"未找到/保持原值: {stats['failed']} 行\n\n"
                        f"文件已保存: {os.path.basename(file_path)}"
                    )

                self.page.run_thread(on_success)

            except ExcelToolkitError as e:
                # Custom exception with user-friendly message
                def on_custom_error():
                    self.page.cursor = ft.Cursor.DEFAULT
                    self.set_error("转换失败")

                    self.log(f"", "info")
                    self.log(f"❌ {e.message}", "error")
                    if e.solution:
                        self.log(f"💡 解决方案: {e.solution}", "warning")

                    self.show_error("转换失败", e.get_user_message())

                self.page.run_thread(on_custom_error)

            except Exception as e:
                # Unexpected error
                log_error(e, "州名转换")
                error_msg = get_user_friendly_error(e)

                def on_error():
                    self.page.cursor = ft.Cursor.DEFAULT
                    self.set_error("转换失败")

                    self.log(f"", "info")
                    self.log(f"❌ 发生错误: {e}", "error")
                    self.log(f"💡 请查看日志文件获取详细信息", "warning")

                    self.show_error("转换失败", error_msg)

                self.page.run_thread(on_error)

        # Start processing thread
        threading.Thread(target=process_thread, daemon=True).start()


def create_view(page: ft.Page) -> StateConversionView:
    """
    Factory function to create the state conversion view.

    Args:
        page: Flet page control

    Returns:
        StateConversionView instance
    """
    view = StateConversionView(page, tab_index=0)
    return view
