"""
Tab 4 - Insert Rows View

Compares two Excel files and inserts missing rows from file X into file Y.
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

from excel_toolkit.insert_rows import process_insert_rows
from excel_toolkit.error_handler import log_error, get_user_friendly_error


class InsertRowsView(TabView):
    """
    Tab 4: Insert Rows

    Compares two Excel files and inserts rows that exist in X but not in Y.
    Uses columns A (product ID) and B (item ID) for comparison.
    """

    def get_tab_name(self) -> str:
        """Get the display name for this tab"""
        return "插入行"

    def get_tab_index(self) -> int:
        """Get the tab index"""
        return 3

    def build(self) -> ft.Control:
        """Build the tab content"""
        colors = self.get_colors()

        # Initialize variables
        self._file_x_path = self.load_file_path("file_x")
        self._file_y_path = self.load_file_path("file_y")
        self._sheet_x = self.load_preference("sheet_x", "")
        self._sheet_y = self.load_preference("sheet_y", "")

        # File X picker
        file_x_picker = self.create_file_picker(
            label="表格 X 文件（源）",
            config_key="file_x",
            on_pick=self._on_file_x_picked,
        )

        # Sheet X dropdown
        self.sheet_x_dropdown = ft.Dropdown(
            label="X 工作表",
            label_style=ft.TextStyle(size=FontSize.LABEL, color=colors["text_secondary"]),
            options=[],
            value=self._sheet_x if self._sheet_x else None,
            width=200,
            bgcolor=colors["surface"],
            border_color=colors["border"],
            focused_border_color=colors["primary"],
            text_style=ft.TextStyle(color=colors["on_surface"]),
            on_change=self._on_sheet_x_changed,
        )

        # File Y picker
        file_y_picker = self.create_file_picker(
            label="表格 Y 文件（目标）",
            config_key="file_y",
            on_pick=self._on_file_y_picked,
        )

        # Sheet Y dropdown
        self.sheet_y_dropdown = ft.Dropdown(
            label="Y 工作表",
            label_style=ft.TextStyle(size=FontSize.LABEL, color=colors["text_secondary"]),
            options=[],
            value=self._sheet_y if self._sheet_y else None,
            width=200,
            bgcolor=colors["surface"],
            border_color=colors["border"],
            focused_border_color=colors["primary"],
            text_style=ft.TextStyle(color=colors["on_surface"]),
            on_change=self._on_sheet_y_changed,
        )

        # File selection section
        file_section = ft.Column([
            ft.Text("表格 X（源）", style=ft.TextStyle(size=FontSize.SECTION, weight=ft.FontWeight.BOLD)),
            file_x_picker,
            self.sheet_x_dropdown,
            ft.Divider(height=20),
            ft.Text("表格 Y（目标）", style=ft.TextStyle(size=FontSize.SECTION, weight=ft.FontWeight.BOLD)),
            file_y_picker,
            self.sheet_y_dropdown,
        ], spacing=Spacing.ROW_SPACING)

        # Parameters section
        param_section = ft.Column([
            ft.Container(
                content=ft.Column([
                    ft.Text("说明：", style=ft.TextStyle(size=FontSize.LABEL, weight=ft.FontWeight.W_500)),
                    ft.Text("• 程序会对比表格 X 和 Y 的 A、B 列", style=ft.TextStyle(size=FontSize.LABEL)),
                    ft.Text("• 将 X 中存在但 Y 中缺失的行插入到 Y", style=ft.TextStyle(size=FontSize.LABEL)),
                    ft.Text("• 插入的行会用蓝色标记", style=ft.TextStyle(size=FontSize.LABEL)),
                ], spacing=5),
                bgcolor=colors["info_container"],
                padding=ft.padding.all(12),
                border_radius=4,
            ),
        ], spacing=Spacing.ROW_SPACING)

        # Action buttons
        run_button = self.create_action_button(
            text="开始插入",
            on_click=self._on_run_click,
            icon=ft.Icons.PLAYLIST_ADD_OUTLINED,
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

            # Update dropdown options
            self.sheet_x_dropdown.options = [
                ft.dropdown.Option(name) for name in sheet_names
            ]

            # Select first sheet if available
            if sheet_names:
                self.sheet_x_dropdown.value = sheet_names[0]
                self._sheet_x = sheet_names[0]
                self.save_preference("sheet_x", self._sheet_x)

            self.sheet_x_dropdown.update()

        except Exception as e:
            self.log(f"加载X工作表列表失败: {e}", "error")

    def _on_sheet_x_changed(self, e) -> None:
        """Handle sheet X dropdown change"""
        self._sheet_x = e.control.value or ""
        self.save_preference("sheet_x", self._sheet_x)

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

    def _on_run_click(self, e) -> None:
        """Handle run button click"""
        # Get current values
        file_x = self._file_x_path or ""
        file_y = self._file_y_path or ""

        # Validate inputs
        if not self._validate_inputs(file_x, file_y):
            return

        # Run processing in background thread
        self._run_processing(file_x, file_y)

    def _on_clear_log_click(self, e) -> None:
        """Handle clear log button click"""
        if self._log_widget:
            self._log_widget.clear()
            self.log("日志已清空", "info")

    # ==================== Processing ====================

    def _validate_inputs(self, file_x: str, file_y: str) -> bool:
        """Validate input parameters"""
        if not file_x or file_x == Text.FILE_NOT_SELECTED:
            self.show_error("验证错误", "请选择表格 X 文件！")
            return False

        if not file_y or file_y == Text.FILE_NOT_SELECTED:
            self.show_error("验证错误", "请选择表格 Y 文件！")
            return False

        if not self._sheet_x:
            self.show_error("验证错误", "请选择表格 X 的工作表！")
            return False

        if not self._sheet_y:
            self.show_error("验证错误", "请选择表格 Y 的工作表！")
            return False

        return True

    def _run_processing(self, file_x: str, file_y: str) -> None:
        """Run insert rows in background thread"""
        self.set_processing("正在处理...")
        self.page.cursor = ft.Cursor.WAIT

        # Log start
        self.log("=" * 60, "info")
        self.log("开始执行插入缺失行...", "info")
        self.log(f"  表格 X: {os.path.basename(file_x)} / {self._sheet_x}", "info")
        self.log(f"  表格 Y: {os.path.basename(file_y)} / {self._sheet_y}", "info")
        self.log("=" * 60, "info")

        def process_thread():
            try:
                # Create thread-safe logger
                def thread_safe_logger(msg: str):
                    self.log(msg, "info")

                # Process insert rows
                stats = process_insert_rows(
                    file_x,
                    self._sheet_x,
                    file_y,
                    self._sheet_y,
                    thread_safe_logger
                )

                # Success callback
                def on_success():
                    self.page.cursor = ft.Cursor.DEFAULT
                    self.set_success("插入完成")

                    self.log("", "info")
                    self.log("=" * 60, "success")
                    self.log("插入完成！", "success")
                    self.log(f"  缺失行数: {stats['missing_count']}", "info")
                    self.log(f"  已插入行数: {stats['inserted_rows']}", "info")
                    self.log("=" * 60, "success")

                    # Show success dialog
                    self.show_success(
                        f"插入完成！\n\n"
                        f"缺失行数: {stats['missing_count']}\n"
                        f"已插入行数: {stats['inserted_rows']}\n\n"
                        f"文件已保存: {os.path.basename(file_y)}"
                    )

                self.page.run_thread(on_success)

            except Exception as e:
                # Error handling
                log_error(e, "插入缺失行")
                error_msg = get_user_friendly_error(e)

                def on_error():
                    self.page.cursor = ft.Cursor.DEFAULT
                    self.set_error("插入失败")

                    self.log(f"", "info")
                    self.log(f"❌ 发生错误: {e}", "error")
                    self.log(f"💡 请查看日志文件获取详细信息", "warning")

                    self.show_error("插入失败", error_msg)

                self.page.run_thread(on_error)

        # Start processing thread
        threading.Thread(target=process_thread, daemon=True).start()


def create_view(page: ft.Page) -> InsertRowsView:
    """
    Factory function to create the insert rows view.

    Args:
        page: Flet page control

    Returns:
        InsertRowsView instance
    """
    view = InsertRowsView(page, tab_index=3)
    return view
