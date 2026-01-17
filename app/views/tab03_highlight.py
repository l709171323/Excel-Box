"""
Tab 3 - Highlight Duplicates View

Highlights duplicate values in Excel columns with different colors.
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

from excel_toolkit.highlight import highlight_duplicates
from excel_toolkit.error_handler import log_error, get_user_friendly_error


class HighlightDuplicatesView(TabView):
    """
    Tab 3: Highlight Duplicates

    Detects and highlights duplicate values in Excel columns
    with different colors (yellow and orange).
    """

    def get_tab_name(self) -> str:
        """Get the display name for this tab"""
        return "高亮重复"

    def get_tab_index(self) -> int:
        """Get the tab index"""
        return 2

    def build(self) -> ft.Control:
        """Build the tab content"""
        colors = self.get_colors()

        # Initialize variables
        self._file_path = self.load_file_path("main_file")
        self._sheet_name = self.load_preference("sheet_name", "")
        self._column_letter = self.load_preference("column_letter", "A")
        self._sheet_names: list[str] = []

        # File picker
        file_picker = self.create_file_picker(
            label="Excel 文件",
            config_key="main_file",
            on_pick=self._on_file_picked,
        )

        # Sheet dropdown (optional - empty means all sheets)
        self.sheet_dropdown = ft.Dropdown(
            label="工作表（可选）",
            label_style=ft.TextStyle(size=FontSize.LABEL, color=colors["text_secondary"]),
            options=[],
            value=self._sheet_name if self._sheet_name else None,
            width=200,
            bgcolor=colors["surface"],
            border_color=colors["border"],
            focused_border_color=colors["primary"],
            text_style=ft.TextStyle(color=colors["on_surface"]),
            on_change=self._on_sheet_changed,
            hint_text="不选择=处理所有工作表",
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
            hint_text="如 A",
        )

        # File selection section
        file_section = ft.Column([
            file_picker,
            ft.Row([
                self.sheet_dropdown,
                ft.Text("提示：不选择则处理所有工作表", style=ft.TextStyle(size=9, color=colors["text_hint"])),
            ], spacing=Spacing.CONTROL_PADDING_X),
        ], spacing=Spacing.ROW_SPACING)

        # Parameters section
        param_section = ft.Column([
            ft.Row([
                self.column_input,
                ft.Text("提示：输入要检查重复的列号", style=ft.TextStyle(size=9, color=colors["text_hint"])),
            ], spacing=Spacing.CONTROL_PADDING_X),
            ft.Container(
                content=ft.Text(
                    "ℹ️ 程序会自动检测指定列的重复值，并用黄色和橙色高亮标记",
                    style=ft.TextStyle(size=FontSize.LABEL, color=colors["text_secondary"]),
                ),
                bgcolor=colors["info_container"],
                padding=ft.padding.all(8),
                border_radius=4,
            ),
        ], spacing=Spacing.ROW_SPACING)

        # Action buttons
        run_button = self.create_action_button(
            text="开始高亮",
            on_click=self._on_run_click,
            icon=ft.Icons.HIGHLIGHT_OUTLINED,
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

        # Load sheet names
        self._load_sheet_names(path)

        self.log(f"已选择文件: {os.path.basename(path)}", "info")

    def _load_sheet_names(self, file_path: str) -> None:
        """Load sheet names from Excel file"""
        try:
            from excel_toolkit.ui import get_sheet_names
            self._sheet_names = get_sheet_names(file_path)

            # Update dropdown options
            self.sheet_dropdown.options = [
                ft.dropdown.Option(name) for name in self._sheet_names
            ]

            # Keep current value or clear
            if self._sheet_name and self._sheet_name in self._sheet_names:
                self.sheet_dropdown.value = self._sheet_name
            else:
                self.sheet_dropdown.value = None
                self._sheet_name = ""

            self.sheet_dropdown.update()

            if self._sheet_names:
                self.log(f"  工作表: {', '.join(self._sheet_names)}", "info")

        except Exception as e:
            self.log(f"加载工作表列表失败: {e}", "error")
            self._sheet_names = []

    def _on_sheet_changed(self, e) -> None:
        """Handle sheet dropdown change"""
        self._sheet_name = e.control.value or ""
        self.save_preference("sheet_name", self._sheet_name)

    def _on_run_click(self, e) -> None:
        """Handle run button click"""
        # Get current values
        file_path = self._file_path or ""
        column_letter = self.column_input.value or ""

        # Validate inputs
        if not self._validate_inputs(file_path, column_letter):
            return

        # Save preferences
        self._column_letter = column_letter
        self.save_preference("column_letter", column_letter)

        # Run processing in background thread
        self._run_processing(file_path, column_letter)

    def _on_clear_log_click(self, e) -> None:
        """Handle clear log button click"""
        if self._log_widget:
            self._log_widget.clear()
            self.log("日志已清空", "info")

    # ==================== Processing ====================

    def _validate_inputs(self, file_path: str, column_letter: str) -> bool:
        """Validate input parameters"""
        if not file_path or file_path == Text.FILE_NOT_SELECTED:
            self.show_error("验证错误", "请选择 Excel 文件！")
            return False

        if not column_letter:
            self.show_error("验证错误", "请输入目标列号！")
            return False

        return True

    def _run_processing(self, file_path: str, column_letter: str) -> None:
        """Run highlight duplicates in background thread"""
        self.set_processing("正在高亮重复项...")
        self.page.cursor = ft.Cursor.WAIT

        # Log start
        self.log("=" * 60, "info")
        self.log("开始执行高亮重复项...", "info")
        self.log(f"  文件: {os.path.basename(file_path)}", "info")
        if self._sheet_name:
            self.log(f"  工作表: {self._sheet_name}", "info")
        else:
            self.log(f"  工作表: 全部", "info")
        self.log(f"  目标列: {column_letter}", "info")
        self.log("=" * 60, "info")

        def process_thread():
            try:
                # Create thread-safe logger
                def thread_safe_logger(msg: str):
                    self.log(msg, "info")

                # Process highlighting
                sheet_name_param = self._sheet_name if self._sheet_name else None
                stats = highlight_duplicates(
                    file_path,
                    column_letter,
                    thread_safe_logger,
                    sheet_name_param
                )

                # Success callback
                def on_success():
                    self.page.cursor = ft.Cursor.DEFAULT
                    self.set_success("高亮完成")

                    # Log results
                    self.log("", "info")
                    self.log("=" * 60, "success")
                    self.log("高亮完成！", "success")
                    self.log(f"  处理工作表: {stats['sheets_processed']} 个", "info")
                    self.log(f"  高亮单元格: {stats['cells_highlighted']} 个", "info")
                    self.log("=" * 60, "success")

                    # Show success dialog
                    self.show_success(
                        f"高亮完成！\n\n"
                        f"处理工作表数: {stats['sheets_processed']}\n"
                        f"高亮单元格数: {stats['cells_highlighted']}\n\n"
                        f"文件已保存: {os.path.basename(file_path)}"
                    )

                self.page.run_thread(on_success)

            except Exception as e:
                # Error handling
                log_error(e, "高亮重复项")
                error_msg = get_user_friendly_error(e)

                def on_error():
                    self.page.cursor = ft.Cursor.DEFAULT
                    self.set_error("高亮失败")

                    self.log(f"", "info")
                    self.log(f"❌ 发生错误: {e}", "error")
                    self.log(f"💡 请查看日志文件获取详细信息", "warning")

                    self.show_error("高亮失败", error_msg)

                self.page.run_thread(on_error)

        # Start processing thread
        threading.Thread(target=process_thread, daemon=True).start()


def create_view(page: ft.Page) -> HighlightDuplicatesView:
    """
    Factory function to create the highlight duplicates view.

    Args:
        page: Flet page control

    Returns:
        HighlightDuplicatesView instance
    """
    view = HighlightDuplicatesView(page, tab_index=2)
    return view
