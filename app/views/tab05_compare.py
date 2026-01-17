"""
Tab 5 - Compare Columns View

Compares column data between two Excel files to find differences.
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

from excel_toolkit.compare import process_compare_columns
from excel_toolkit.error_handler import log_error, get_user_friendly_error


class CompareColumnsView(TabView):
    """
    Tab 5: Compare Columns

    Compares column data between two Excel files and finds
    missing or extra values.
    """

    def get_tab_name(self) -> str:
        """Get the display name for this tab"""
        return "对比列"

    def get_tab_index(self) -> int:
        """Get the tab index"""
        return 4

    def build(self) -> ft.Control:
        """Build the tab content"""
        colors = self.get_colors()

        # Initialize variables
        self._file_x_path = self.load_file_path("file_x")
        self._file_y_path = self.load_file_path("file_y")
        self._col_x = self.load_preference("col_x", "A")
        self._col_y = self.load_preference("col_y", "A")
        self._sheet_y = self.load_preference("sheet_y", "")
        self._selected_sheets_x: List[str] = []
        self._ignore_dups = self.load_preference("ignore_dups", True)

        # File X picker
        file_x_picker = self.create_file_picker(
            label="表格 X 文件",
            config_key="file_x",
            on_pick=self._on_file_x_picked,
        )

        # Sheet X list (for multi-selection)
        self.sheet_x_listbox = ft.ListView(
            expand=True,
            spacing=5,
            height=100,
            item_extent=30,
        )

        # Column X input
        self.col_x_input = ft.TextField(
            label="X 列号",
            label_style=ft.TextStyle(size=FontSize.LABEL, color=colors["text_secondary"]),
            value=self._col_x,
            width=80,
            bgcolor=colors["surface"],
            border_color=colors["border"],
            focused_border_color=colors["primary"],
            text_style=ft.TextStyle(color=colors["on_surface"]),
            hint_text="如 A",
        )

        # File Y picker
        file_y_picker = self.create_file_picker(
            label="表格 Y 文件",
            config_key="file_y",
            on_pick=self._on_file_y_picked,
        )

        # Sheet Y dropdown
        self.sheet_y_dropdown = ft.Dropdown(
            label="Y 工作表",
            label_style=ft.TextStyle(size=FontSize.LABEL, color=colors["text_secondary"]),
            options=[],
            value=self._sheet_y if self._sheet_y else None,
            width=150,
            bgcolor=colors["surface"],
            border_color=colors["border"],
            focused_border_color=colors["primary"],
            text_style=ft.TextStyle(color=colors["on_surface"]),
            on_change=self._on_sheet_y_changed,
        )

        # Column Y input
        self.col_y_input = ft.TextField(
            label="Y 列号",
            label_style=ft.TextStyle(size=FontSize.LABEL, color=colors["text_secondary"]),
            value=self._col_y,
            width=80,
            bgcolor=colors["surface"],
            border_color=colors["border"],
            focused_border_color=colors["primary"],
            text_style=ft.TextStyle(color=colors["on_surface"]),
            hint_text="如 A",
        )

        # Ignore duplicates checkbox
        self.ignore_dups_checkbox = ft.Checkbox(
            label="忽略重复值（集合比较）",
            value=self._ignore_dups,
            on_change=self._on_ignore_dups_changed,
        )

        # File selection section
        file_section = ft.Column([
            ft.Text("表格 X", style=ft.TextStyle(size=FontSize.SECTION, weight=ft.FontWeight.BOLD)),
            file_x_picker,
            ft.Row([
                ft.Text("X 工作表（可多选）:", style=ft.TextStyle(size=FontSize.LABEL)),
                self.col_x_input,
            ], spacing=Spacing.CONTROL_PADDING_X),
            ft.Container(
                content=self.sheet_x_listbox,
                bgcolor=colors["surface_variant"],
                border_radius=4,
                padding=5,
                height=120,
            ),
            ft.Divider(height=20),
            ft.Text("表格 Y", style=ft.TextStyle(size=FontSize.SECTION, weight=ft.FontWeight.BOLD)),
            file_y_picker,
            ft.Row([
                self.sheet_y_dropdown,
                self.col_y_input,
            ], spacing=Spacing.CONTROL_PADDING_X),
        ], spacing=Spacing.ROW_SPACING)

        # Parameters section
        param_section = ft.Column([
            self.ignore_dups_checkbox,
            ft.Container(
                content=ft.Text(
                    "ℹ️ 对比两个Excel文件指定列的数据差异，找出Y中缺失或多余的值",
                    style=ft.TextStyle(size=FontSize.LABEL, color=colors["text_secondary"]),
                ),
                bgcolor=colors["info_container"],
                padding=ft.padding.all(8),
                border_radius=4,
            ),
        ], spacing=Spacing.ROW_SPACING)

        # Action buttons
        run_button = self.create_action_button(
            text="开始对比",
            on_click=self._on_run_click,
            icon=ft.Icons.COMPARE_OUTLINED,
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

    def _on_file_x_picked(self, path: str) -> None:
        """Handle file X pick"""
        self._file_x_path = path
        self._load_sheets_x(path)
        self.log(f"已选择表格 X: {os.path.basename(path)}", "info")

    def _load_sheets_x(self, file_path: str) -> None:
        """Load sheet names for file X"""
        try:
            from excel_toolkit.ui import get_sheet_names
            sheet_names = get_sheet_names(file_path)

            # Clear and update listbox
            self.sheet_x_listbox.controls.clear()

            for name in sheet_names:
                checkbox = ft.Checkbox(
                    label=name,
                    value=False,
                    on_change=self._on_sheet_x_selection_change,
                )
                self.sheet_x_listbox.controls.append(checkbox)

            self.sheet_x_listbox.update()
            self._selected_sheets_x = []

        except Exception as e:
            self.log(f"加载X工作表列表失败: {e}", "error")

    def _on_sheet_x_selection_change(self, e) -> None:
        """Handle sheet X selection change"""
        selected = []
        for control in self.sheet_x_listbox.controls:
            if isinstance(control, ft.Checkbox) and control.value:
                selected.append(control.label)
        self._selected_sheets_x = selected

    def _on_file_y_picked(self, path: str) -> None:
        """Handle file Y pick"""
        self._file_y_path = path
        self._load_sheets_y(path)
        self.log(f"已选择表格 Y: {os.path.basename(path)}", "info")

    def _load_sheets_y(self, file_path: str) -> None:
        """Load sheet names for file Y"""
        try:
            from excel_toolkit.ui import get_sheet_names
            sheet_names = get_sheet_names(file_path)

            # Update dropdown options
            self.sheet_y_dropdown.options = [
                ft.dropdown.Option(name) for name in sheet_names
            ]

            # Select first sheet if available
            if sheet_names:
                self.sheet_y_dropdown.value = sheet_names[0]
                self._sheet_y = sheet_names[0]
                self.save_preference("sheet_y", self._sheet_y)

            self.sheet_y_dropdown.update()

        except Exception as e:
            self.log(f"加载Y工作表列表失败: {e}", "error")

    def _on_sheet_y_changed(self, e) -> None:
        """Handle sheet Y dropdown change"""
        self._sheet_y = e.control.value or ""
        self.save_preference("sheet_y", self._sheet_y)

    def _on_ignore_dups_changed(self, e) -> None:
        """Handle ignore duplicates checkbox change"""
        self._ignore_dups = e.control.value
        self.save_preference("ignore_dups", self._ignore_dups)

    def _on_run_click(self, e) -> None:
        """Handle run button click"""
        # Get current values
        file_x = self._file_x_path or ""
        file_y = self._file_y_path or ""
        col_x = self.col_x_input.value or ""
        col_y = self.col_y_input.value or ""

        # Validate inputs
        if not self._validate_inputs(file_x, file_y, col_x, col_y):
            return

        # Save preferences
        self._col_x = col_x
        self._col_y = col_y
        self.save_preference("col_x", col_x)
        self.save_preference("col_y", col_y)

        # Run processing in background thread
        self._run_processing(file_x, file_y, col_x, col_y)

    def _on_clear_log_click(self, e) -> None:
        """Handle clear log button click"""
        if self._log_widget:
            self._log_widget.clear()
            self.log("日志已清空", "info")

    # ==================== Processing ====================

    def _validate_inputs(self, file_x: str, file_y: str, col_x: str, col_y: str) -> bool:
        """Validate input parameters"""
        if not file_x or file_x == Text.FILE_NOT_SELECTED:
            self.show_error("验证错误", "请选择表格 X 文件！")
            return False

        if not file_y or file_y == Text.FILE_NOT_SELECTED:
            self.show_error("验证错误", "请选择表格 Y 文件！")
            return False

        if not self._selected_sheets_x:
            self.show_error("验证错误", "请至少选择一个 X 工作表！")
            return False

        if not self._sheet_y:
            self.show_error("验证错误", "请选择 Y 工作表！")
            return False

        if not col_x:
            self.show_error("验证错误", "请输入 X 列号！")
            return False

        if not col_y:
            self.show_error("验证错误", "请输入 Y 列号！")
            return False

        return True

    def _run_processing(self, file_x: str, file_y: str, col_x: str, col_y: str) -> None:
        """Run column comparison in background thread"""
        self.set_processing("正在对比...")
        self.page.cursor = ft.Cursor.WAIT

        # Log start
        self.log("=" * 60, "info")
        self.log("开始执行对比列数据...", "info")
        self.log(f"  X 工作表: {self._selected_sheets_x}", "info")
        self.log(f"  X 列号: {col_x}", "info")
        self.log(f"  Y 工作表: {self._sheet_y}", "info")
        self.log(f"  Y 列号: {col_y}", "info")
        self.log(f"  忽略重复: {self._ignore_dups}", "info")
        self.log("=" * 60, "info")

        def process_thread():
            try:
                # Create thread-safe logger
                def thread_safe_logger(msg: str):
                    self.log(msg, "info")

                # Process comparison
                result = process_compare_columns(
                    file_x,
                    self._selected_sheets_x,
                    col_x,
                    file_y,
                    self._sheet_y,
                    col_y,
                    thread_safe_logger,
                    self._ignore_dups
                )

                # Success callback
                def on_success():
                    self.page.cursor = ft.Cursor.DEFAULT
                    self.set_success("对比完成")

                    self.log("", "info")
                    self.log("=" * 60, "success")
                    self.log("对比完成！", "success")
                    self.log("=" * 60, "success")

                    # Show success dialog with result
                    self.show_info("对比完成", result)

                self.page.run_thread(on_success)

            except Exception as e:
                # Error handling
                log_error(e, "对比列")
                error_msg = get_user_friendly_error(e)

                def on_error():
                    self.page.cursor = ft.Cursor.DEFAULT
                    self.set_error("对比失败")

                    self.log(f"", "info")
                    self.log(f"❌ 发生错误: {e}", "error")
                    self.log(f"💡 请查看日志文件获取详细信息", "warning")

                    self.show_error("对比失败", error_msg)

                self.page.run_thread(on_error)

        # Start processing thread
        threading.Thread(target=process_thread, daemon=True).start()


def create_view(page: ft.Page) -> CompareColumnsView:
    """
    Factory function to create the compare columns view.

    Args:
        page: Flet page control

    Returns:
        CompareColumnsView instance
    """
    view = CompareColumnsView(page, tab_index=4)
    return view
