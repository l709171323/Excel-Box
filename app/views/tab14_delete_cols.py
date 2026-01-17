"""
Tab 14 - Delete Columns View

Deletes specified columns from Excel files with template management.
"""

import os
import json
import threading
from typing import Optional, Dict

import flet as ft

from app.views.base_view import TabView
from app.core.theme import ThemeMode
from app.core.constants import Text, FontSize, Spacing

# Import business logic from existing module
import sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "../.."))

from excel_toolkit.delete_cols import delete_columns, parse_column_input
from excel_toolkit.error_handler import log_error, get_user_friendly_error


class DeleteColumnsView(TabView):
    """
    Tab 14: Delete Columns

    Deletes specified columns from Excel files.
    Supports column templates for quick reuse.
    """

    def get_tab_name(self) -> str:
        """Get the display name for this tab"""
        return "删除列"

    def get_tab_index(self) -> int:
        """Get the tab index"""
        return 13

    def _get_templates_path(self) -> str:
        """Get the templates configuration file path"""
        config_dir = os.path.expanduser("~")
        return os.path.join(config_dir, ".excel_toolkit_flet", "delete_cols_templates.json")

    def _load_templates(self) -> Dict[str, str]:
        """Load column templates from file"""
        path = self._get_templates_path()
        if os.path.exists(path):
            try:
                with open(path, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception:
                return {}
        return {}

    def _save_templates(self, templates: Dict[str, str]) -> None:
        """Save column templates to file"""
        path = self._get_templates_path()
        os.makedirs(os.path.dirname(path), exist_ok=True)
        try:
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(templates, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self.show_error("保存模板失败", str(e))

    def build(self) -> ft.Control:
        """Build the tab content"""
        colors = self.get_colors()

        # Initialize variables
        self._file_path = self.load_file_path("main_file")
        self._sheet_name = self.load_preference("sheet_name", "")
        self._columns_input = self.load_preference("columns", "")
        self._selected_template = "(选择模板)"

        # Load templates
        self._templates = self._load_templates()

        # File picker
        file_picker = self.create_file_picker(
            label="Excel 文件",
            config_key="main_file",
            on_pick=self._on_file_picked,
        )

        # Sheet dropdown
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

        # Template dropdown
        self.template_dropdown = ft.Dropdown(
            label="模板",
            label_style=ft.TextStyle(size=FontSize.LABEL, color=colors["text_secondary"]),
            options=[ft.dropdown.Option("(选择模板)")] + [
                ft.dropdown.Option(name) for name in self._templates.keys()
            ],
            value="(选择模板)",
            width=150,
            bgcolor=colors["surface"],
            border_color=colors["border"],
            focused_border_color=colors["primary"],
            text_style=ft.TextStyle(color=colors["on_surface"]),
            on_change=self._on_template_changed,
        )

        # Save template button
        save_template_btn = ft.IconButton(
            icon=ft.Icons.SAVE,
            tooltip="保存当前配置为模板",
            icon_color=colors["text_secondary"],
            on_click=self._on_save_template,
        )

        # Delete template button
        delete_template_btn = ft.IconButton(
            icon=ft.Icons.DELETE,
            tooltip="删除当前模板",
            icon_color=colors["text_secondary"],
            on_click=self._on_delete_template,
        )

        # Refresh templates button
        refresh_template_btn = ft.IconButton(
            icon=ft.Icons.REFRESH,
            tooltip="刷新模板列表",
            icon_color=colors["text_secondary"],
            on_click=self._on_refresh_templates,
        )

        # Columns input
        self.columns_text_field = ft.TextField(
            label="要删除的列",
            label_style=ft.TextStyle(size=FontSize.LABEL, color=colors["text_secondary"]),
            value=self._columns_input,
            width=200,
            bgcolor=colors["surface"],
            border_color=colors["border"],
            focused_border_color=colors["primary"],
            text_style=ft.TextStyle(color=colors["on_surface"]),
            hint_text="如: D,E 或 D-F 或 A C E",
        )

        # File selection section
        file_section = ft.Column([
            file_picker,
            self.sheet_dropdown,
        ], spacing=Spacing.ROW_SPACING)

        # Parameters section
        param_section = ft.Column([
            # Template management row
            ft.Row([
                ft.Text("模板管理:", style=ft.TextStyle(size=FontSize.LABEL, weight=ft.FontWeight.W_500)),
                self.template_dropdown,
                save_template_btn,
                delete_template_btn,
                refresh_template_btn,
            ], spacing=Spacing.CONTROL_PADDING_X),
            # Columns input row
            ft.Row([
                self.columns_text_field,
            ], spacing=Spacing.CONTROL_PADDING_X),
            # Help text
            ft.Container(
                content=ft.Column([
                    ft.Text("支持的格式:", style=ft.TextStyle(size=FontSize.LABEL, weight=ft.FontWeight.W_500)),
                    ft.Text("• \"D,E\" - 逗号分隔", style=ft.TextStyle(size=FontSize.LABEL)),
                    ft.Text("• \"D-F\" - 范围格式", style=ft.TextStyle(size=FontSize.LABEL)),
                    ft.Text("• \"A C E\" - 空格分隔", style=ft.TextStyle(size=FontSize.LABEL)),
                ], spacing=3),
                bgcolor=colors["info_container"],
                padding=ft.padding.all(10),
                border_radius=4,
            ),
            # Warning
            ft.Container(
                content=ft.Row([
                    ft.Icon(ft.Icons.WARNING, color=colors["error"], size=16),
                    ft.Text(
                        "注意: 删除操作会直接修改原文件，建议提前备份！",
                        style=ft.TextStyle(size=FontSize.LABEL, color=colors["error"]),
                    ),
                ], spacing=5),
                bgcolor=colors["error_container"],
                padding=ft.padding.all(8),
                border_radius=4,
            ),
        ], spacing=Spacing.ROW_SPACING)

        # Action buttons
        run_button = self.create_action_button(
            text="开始删除",
            on_click=self._on_run_click,
            icon=ft.Icons.DELETE_OUTLINED,
            variant="danger",
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
        self._load_sheets(path)
        self.log(f"已选择文件: {os.path.basename(path)}", "info")

    def _load_sheets(self, file_path: str) -> None:
        """Load sheet names from Excel file"""
        try:
            from excel_toolkit.ui import get_sheet_names
            sheet_names = get_sheet_names(file_path)

            # Update dropdown options
            self.sheet_dropdown.options = [
                ft.dropdown.Option(name) for name in sheet_names
            ]

            # Keep current value or clear
            if self._sheet_name and self._sheet_name in sheet_names:
                self.sheet_dropdown.value = self._sheet_name
            else:
                self.sheet_dropdown.value = None

            self.sheet_dropdown.update()

        except Exception as e:
            self.log(f"加载工作表列表失败: {e}", "error")

    def _on_sheet_changed(self, e) -> None:
        """Handle sheet dropdown change"""
        self._sheet_name = e.control.value or ""
        self.save_preference("sheet_name", self._sheet_name)

    def _on_template_changed(self, e) -> None:
        """Handle template dropdown change"""
        template_name = e.control.value
        self._selected_template = template_name

        if template_name != "(选择模板)" and template_name in self._templates:
            self.columns_text_field.value = self._templates[template_name]
            self.columns_text_field.update()
            self.log(f"已加载模板 [{template_name}]: {self._templates[template_name]}", "info")

    def _on_save_template(self, e) -> None:
        """Save current configuration as a template"""
        columns = self.columns_text_field.value.strip()
        if not columns:
            self.show_error("保存模板", "请先输入要删除的列标识")
            return

        # Show dialog to get template name
        template_name_dialog = ft.AlertDialog(
            modal=True,
            title=ft.Text("保存模板"),
            content=ft.TextField(
                label="模板名称",
                hint="输入模板名称",
                autofocus=True,
            ),
            actions=[
                ft.TextButton("取消", on_click=lambda _: self._close_dialog()),
                ft.TextButton("保存", on_click=lambda _: self._confirm_save_template(template_name_dialog)),
            ],
        )

        self.template_name_input = template_name_dialog.content
        self.page.dialog = template_name_dialog
        template_name_dialog.open = True
        self.page.update()

    def _confirm_save_template(self, dialog: ft.AlertDialog) -> None:
        """Confirm and save template"""
        name = self.template_name_input.value.strip()
        self._close_dialog()

        if not name:
            self.show_error("保存模板", "模板名称不能为空")
            return

        # Save template
        self._templates[name] = self.columns_text_field.value.strip()
        self._save_templates(self._templates)
        self._refresh_template_list()

        # Select the new template
        self.template_dropdown.value = name
        self._selected_template = name

        self.log(f"✅ 模板 [{name}] 已保存", "success")
        self.show_info("保存成功", f"模板 [{name}] 已保存")

    def _on_delete_template(self, e) -> None:
        """Delete current template"""
        if self._selected_template == "(选择模板)":
            self.show_error("删除模板", "请先选择要删除的模板")
            return

        if self._selected_template not in self._templates:
            return

        # Confirm deletion
        dialog = ft.AlertDialog(
            modal=True,
            title=ft.Text("确认删除"),
            content=ft.Text(f"确定要删除模板 [{self._selected_template}] 吗？"),
            actions=[
                ft.TextButton("取消", on_click=lambda _: self._close_dialog()),
                ft.TextButton("删除", on_click=lambda _: self._confirm_delete_template()),
            ],
        )

        self.page.dialog = dialog
        dialog.open = True
        self.page.update()

    def _confirm_delete_template(self) -> None:
        """Confirm and delete template"""
        template_name = self._selected_template
        self._close_dialog()

        if template_name in self._templates:
            del self._templates[template_name]
            self._save_templates(self._templates)
            self._refresh_template_list()
            self._selected_template = "(选择模板)"

            self.log(f"✅ 模板 [{template_name}] 已删除", "info")

    def _on_refresh_templates(self, e) -> None:
        """Refresh template list"""
        self._templates = self._load_templates()
        self._refresh_template_list()
        self.log("模板列表已刷新", "info")

    def _refresh_template_list(self) -> None:
        """Refresh the template dropdown options"""
        options = [ft.dropdown.Option("(选择模板)")]
        for name in sorted(self._templates.keys()):
            options.append(ft.dropdown.Option(name))

        self.template_dropdown.options = options
        self.template_dropdown.update()

    def _on_run_click(self, e) -> None:
        """Handle run button click"""
        file_path = self._file_path or ""
        columns_str = self.columns_text_field.value.strip()

        # Validate inputs
        if not self._validate_inputs(file_path, columns_str):
            return

        # Parse columns
        try:
            columns = parse_column_input(columns_str)
            if not columns:
                self.show_error("验证错误", "无法解析列标识，请检查输入格式")
                return
        except Exception as e:
            self.show_error("解析错误", f"无法解析列标识: {e}")
            return

        # Confirm deletion
        cols_display = ", ".join(sorted(columns))
        dialog = ft.AlertDialog(
            modal=True,
            title=ft.Row([
                ft.Icon(ft.Icons.WARNING, color=ft.Colors.RED),
                ft.Text("确认删除", color=ft.Colors.RED),
            ]),
            content=ft.Text(
                f"确定要删除以下列吗？\n\n列: {cols_display}\n\n⚠️ 此操作会直接修改原文件！"
            ),
            actions=[
                ft.TextButton("取消", on_click=lambda _: self._close_dialog()),
                ft.TextButton("确认删除", on_click=lambda _: self._confirm_run(columns)),
            ],
            bgcolor=ft.Colors.RED_50,
        )

        self.page.dialog = dialog
        dialog.open = True
        self.page.update()

    def _confirm_run(self, columns: list) -> None:
        """Confirm and run column deletion"""
        self._close_dialog()
        self._run_processing(self._file_path, columns)

    def _on_clear_log_click(self, e) -> None:
        """Handle clear log button click"""
        if self._log_widget:
            self._log_widget.clear()
            self.log("日志已清空", "info")

    # ==================== Processing ====================

    def _validate_inputs(self, file_path: str, columns_str: str) -> bool:
        """Validate input parameters"""
        if not file_path or file_path == Text.FILE_NOT_SELECTED:
            self.show_error("验证错误", "请选择 Excel 文件！")
            return False

        if not columns_str:
            self.show_error("验证错误", "请输入要删除的列标识！")
            return False

        return True

    def _run_processing(self, file_path: str, columns: list) -> None:
        """Run column deletion in background thread"""
        self.set_processing("正在删除列...")
        self.page.cursor = ft.Cursor.WAIT

        sheet_name = self._sheet_name if self._sheet_name else None

        # Log start
        self.log("=" * 60, "info")
        self.log("开始执行批量删除列...", "info")
        self.log(f"  文件: {os.path.basename(file_path)}", "info")
        self.log(f"  工作表: {sheet_name if sheet_name else '全部'}", "info")
        self.log(f"  删除列: {', '.join(sorted(columns))}", "info")
        self.log("=" * 60, "info")

        def process_thread():
            try:
                # Create thread-safe logger
                def thread_safe_logger(msg: str):
                    self.log(msg, "info")

                # Process deletion
                stats = delete_columns(
                    file_path,
                    columns,
                    thread_safe_logger,
                    sheet_name
                )

                # Success callback
                def on_success():
                    self.page.cursor = ft.Cursor.DEFAULT
                    self.set_success("删除完成")

                    self.log("", "info")
                    self.log("=" * 60, "success")
                    self.log("删除完成！", "success")
                    self.log(f"  处理工作表: {stats['sheets_processed']} 个", "info")
                    self.log(f"  删除列: {stats['columns_deleted']} 个", "info")
                    self.log("=" * 60, "success")

                    # Show success dialog
                    self.show_success(
                        f"删除完成！\n\n"
                        f"处理工作表数: {stats['sheets_processed']}\n"
                        f"删除列数: {stats['columns_deleted']}\n\n"
                        f"文件已保存: {os.path.basename(file_path)}"
                    )

                self.page.run_thread(on_success)

            except Exception as e:
                # Error handling
                log_error(e, "删除列")
                error_msg = get_user_friendly_error(e)

                def on_error():
                    self.page.cursor = ft.Cursor.DEFAULT
                    self.set_error("删除失败")

                    self.log(f"", "info")
                    self.log(f"❌ 发生错误: {e}", "error")
                    self.log(f"💡 请查看日志文件获取详细信息", "warning")

                    self.show_error("删除失败", error_msg)

                self.page.run_thread(on_error)

        # Start processing thread
        threading.Thread(target=process_thread, daemon=True).start()


def create_view(page: ft.Page) -> DeleteColumnsView:
    """
    Factory function to create the delete columns view.

    Args:
        page: Flet page control

    Returns:
        DeleteColumnsView instance
    """
    view = DeleteColumnsView(page, tab_index=13)
    return view
