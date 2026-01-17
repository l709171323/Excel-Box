"""
Tab 2 - SKU Filling View

Intelligently fills SKU information (length, width, height, weight) from a database
into order files based on SKU matching.
"""

import os
import json
import threading
from typing import Optional, Dict, List

import flet as ft

from app.views.base_view import TabView
from app.core.theme import ThemeMode
from app.core.constants import Text, FontSize, Spacing

# Import business logic
import sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "../.."))

from excel_toolkit.sku_fill import process_skus, identify_header_mapping
from excel_toolkit.error_handler import log_error, get_user_friendly_error


class SkuFillView(TabView):
    """
    Tab 2: SKU Filling

    Intelligently fills SKU information from a database into order files.
    Supports flexible column mapping and template management.
    """

    def get_tab_name(self) -> str:
        """Get the display name for this tab"""
        return "SKU填充"

    def get_tab_index(self) -> int:
        """Get the tab index"""
        return 1

    def _get_config_dir(self) -> str:
        """Get the configuration directory"""
        config_dir = os.path.expanduser("~")
        dir_path = os.path.join(config_dir, ".excel_toolkit_flet")
        os.makedirs(dir_path, exist_ok=True)
        return dir_path

    def build(self) -> ft.Control:
        """Build the tab content"""
        colors = self.get_colors()

        # Initialize variables
        self._order_file_path = self._load_order_file()
        self._sku_db_file_path = self._load_sku_db_file()
        self._order_sheet = self.load_preference("order_sheet", "")
        self._sku_db_sheet = self.load_preference("sku_db_sheet", "")

        # Database column mapping
        self._db_sku_col = self.load_preference("db_sku_col", "")
        self._db_l_col = self.load_preference("db_l_col", "")
        self._db_w_col = self.load_preference("db_w_col", "")
        self._db_h_col = self.load_preference("db_h_col", "")
        self._db_wt_col = self.load_preference("db_wt_col", "")

        # Target column mapping
        self._target_sku_col = self.load_preference("target_sku_col", "A")
        self._target_qty_col = self.load_preference("target_qty_col", "B")
        self._target_l_col = self.load_preference("target_l_col", "C")
        self._target_w_col = self.load_preference("target_w_col", "D")
        self._target_h_col = self.load_preference("target_h_col", "E")
        self._target_wt_col = self.load_preference("target_wt_col", "F")

        self._ignore_qty = self.load_preference("ignore_qty", False)
        self._template_name = self.load_preference("template_name", "默认")

        # File pickers
        self.order_file_picker = ft.FilePicker(on_result=self._on_order_file_picked)
        self.sku_db_file_picker = ft.FilePicker(on_result=self._on_sku_db_file_picked)

        # Order file section
        order_file_section = ft.Column([
            self._create_file_picker_button(
                "选择订单表格",
                self.order_file_picker,
                self._order_file_path
            ),
            self._create_sheet_dropdown(
                "订单工作表",
                self._order_sheet,
                on_change=self._on_order_sheet_changed,
            ),
        ], spacing=Spacing.ROW_SPACING)

        # SKU database section
        sku_db_section = ft.Column([
            self._create_file_picker_button(
                "选择SKU数据库",
                self.sku_db_file_picker,
                self._sku_db_file_path
            ),
            self._create_sheet_dropdown(
                "SKU数据库工作表",
                self._sku_db_sheet,
                on_change=self._on_sku_db_sheet_changed,
            ),
        ], spacing=Spacing.ROW_SPACING)

        # Database column mapping section
        db_map_section = ft.Column([
            ft.Text("SKU数据库列映射（自动识别）", style=ft.TextStyle(
                size=FontSize.SECTION, weight=ft.FontWeight.BOLD
            )),
            ft.Row([
                self._create_column_dropdown("SKU列", "db_sku"),
                ft.Container(expand=True),
                self._create_column_dropdown("长列", "db_l"),
            ], spacing=Spacing.CONTROL_PADDING_X),
            ft.Row([
                self._create_column_dropdown("宽列", "db_w"),
                ft.Container(expand=True),
                self._create_column_dropdown("高列", "db_h"),
            ], spacing=Spacing.CONTROL_PADDING_X),
            ft.Row([
                self._create_column_dropdown("重量列", "db_wt"),
            ], spacing=Spacing.CONTROL_PADDING_X),
        ], spacing=Spacing.ROW_SPACING)

        # Target column mapping section
        target_map_section = ft.Column([
            ft.Row([
                ft.Text("目标表格列配置", style=ft.TextStyle(
                    size=FontSize.SECTION, weight=ft.FontWeight.BOLD
                )),
                ft.Container(expand=True),
                self._create_template_dropdown(),
                ft.IconButton(
                    icon=ft.Icons.SAVE,
                    icon_color=colors["text_secondary"],
                    tooltip="保存当前配置为模板",
                    on_click=self._on_save_template,
                ),
                ft.IconButton(
                    icon=ft.Icons.DELETE,
                    icon_color=colors["text_secondary"],
                    tooltip="删除当前模板",
                    on_click=self._on_delete_template,
                ),
            ], spacing=Spacing.CONTROL_PADDING_X, alignment=ft.MainAxisAlignment.CENTER),
            ft.Divider(height=1, color=colors["border"]),
            ft.Text("输入列号 (如: A, B, C) 或列名", style=ft.TextStyle(
                size=9, color=colors["text_hint"]
            )),
            ft.Row([
                self._create_column_input("SKU列", "target_sku_col", 50),
                ft.Container(expand=True),
                self._create_column_input("数量列", "target_qty_col", 50),
            ], spacing=Spacing.CONTROL_PADDING_X),
            ft.Row([
                self._create_column_input("长填充到", "target_l_col", 50),
                ft.Container(expand=True),
                self._create_column_input("宽填充到", "target_w_col", 50),
            ], spacing=Spacing.CONTROL_PADDING_X),
            ft.Row([
                self._create_column_input("高填充到", "target_h_col", 50),
                ft.Container(expand=True),
                self._create_column_input("重量填充到", "target_wt_col", 50),
            ], spacing=Spacing.CONTROL_PADDING_X),
            ft.Row([
                ft.Checkbox(
                    label="计算单个SKU数据（忽略数量列）",
                    value=self._ignore_qty,
                    on_change=self._on_ignore_qty_changed,
                ),
            ], spacing=Spacing.CONTROL_PADDING_X),
        ], spacing=Spacing.ROW_SPACING)

        # File selection section (combined)
        file_section = ft.Column([
            ft.Text("订单文件", style=ft.TextStyle(size=FontSize.SECTION, weight=ft.FontWeight.BOLD)),
            order_file_section,
            ft.Divider(height=1, color=colors["border"]),
            ft.Text("SKU数据库", style=ft.TextStyle(size=FontSize.SECTION, weight=ft.FontWeight.BOLD)),
            sku_db_section,
        ], spacing=Spacing.ROW_SPACING)

        # Parameters section
        param_section = ft.Column([
            db_map_section,
            ft.Divider(height=1, color=colors["border"]),
            target_map_section,
        ], spacing=Spacing.ROW_SPACING)

        # Action buttons
        run_button = self.create_action_button(
            text="开始填充",
            on_click=self._on_run_click,
            icon=ft.Icons.TAG_OUTLINED,
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

        # Add file pickers to page overlay
        self.page.overlay.append(self.order_file_picker)
        self.page.overlay.append(self.sku_db_file_picker)

        # Load column dropdown headers
        self._db_column_headers: List[str] = []

        return content

    # ==================== Helper Methods ====================

    def _create_file_picker_button(
        self,
        label: str,
        picker: ft.FilePicker,
        current_path: str
    ) -> ft.Row:
        """Create a file picker button row"""
        colors = self.get_colors()

        path_display = ft.Text(
            current_path if current_path and current_path != "未选择文件" else "未选择文件",
            style=ft.TextStyle(size=FontSize.LABEL, color=colors["text_secondary"]),
            max_lines=1,
            overflow=ft.TextOverflow.ELLIPSIS,
            expand=True,
        )

        # Store reference for updates
        setattr(self, f"{label}_path_display", path_display)

        button = ft.ElevatedButton(
            label,
            icon=ft.Icons.FILE_UPLOAD,
            on_click=lambda _: picker.pick_files(),
            style=ft.ButtonStyle(
                bgcolor=colors["primary"],
                color=colors["on_primary"],
            ),
        )

        return ft.Row([button, path_display], spacing=Spacing.CONTROL_PADDING_X)

    def _create_sheet_dropdown(
        self,
        label: str,
        current_value: str,
        on_change: Optional[callable] = None
    ) -> ft.Dropdown:
        """Create a sheet selection dropdown"""
        colors = self.get_colors()

        dropdown = ft.Dropdown(
            label=label,
            label_style=ft.TextStyle(size=FontSize.LABEL, color=colors["text_secondary"]),
            options=[],
            value=current_value if current_value else None,
            width=200,
            bgcolor=colors["surface"],
            border_color=colors["border"],
            focused_border_color=colors["primary"],
            text_style=ft.TextStyle(color=colors["on_surface"]),
            on_change=on_change,
            hint_text="不选择=处理所有工作表",
        )

        # Store reference
        setattr(self, f"{label}_dropdown", dropdown)
        return dropdown

    def _create_column_dropdown(self, label: str, attr_name: str) -> ft.Dropdown:
        """Create a column mapping dropdown for database"""
        colors = self.get_colors()

        current_value = getattr(self, f"_{attr_name}_col")
        dropdown = ft.Dropdown(
            label=label,
            label_style=ft.TextStyle(size=FontSize.LABEL, color=colors["text_secondary"]),
            options=[],
            value=current_value if current_value else None,
            width=150,
            bgcolor=colors["surface"],
            border_color=colors["border"],
            focused_border_color=colors["primary"],
            text_style=ft.TextStyle(color=colors["on_surface"]),
            on_change=lambda e, attr=attr_name: self._on_db_col_changed(attr_name, e),
        )

        # Store reference
        setattr(self, f"{attr_name}_dropdown", dropdown)
        return dropdown

    def _create_column_input(self, label: str, attr_name: str, width: int) -> ft.TextField:
        """Create a target column input field"""
        colors = self.get_colors()

        # attr_name already includes _col suffix (e.g., "target_sku_col")
        current_value = getattr(self, f"_{attr_name}", "")
        text_field = ft.TextField(
            label=label,
            label_style=ft.TextStyle(size=FontSize.LABEL, color=colors["text_secondary"]),
            value=current_value,
            width=width,
            bgcolor=colors["surface"],
            border_color=colors["border"],
            focused_border_color=colors["primary"],
            text_style=ft.TextStyle(color=colors["on_surface"]),
            hint_text="如 A",
            on_change=lambda e, attr=attr_name: self._on_target_col_changed(attr, e),
        )

        # Store reference
        setattr(self, f"{attr_name}_input", text_field)
        return text_field

    def _create_template_dropdown(self) -> ft.Dropdown:
        """Create template dropdown"""
        colors = self.get_colors()

        templates = self._load_templates()
        options = [ft.dropdown.Option("默认")]
        for name in templates.keys():
            options.append(ft.dropdown.Option(name))

        dropdown = ft.Dropdown(
            options=options,
            value=self._template_name if self._template_name in templates else "默认",
            width=150,
            bgcolor=colors["surface"],
            border_color=colors["border"],
            focused_border_color=colors["primary"],
            text_style=ft.TextStyle(color=colors["on_surface"]),
            on_change=self._on_template_changed,
        )

        self.template_dropdown = dropdown
        return dropdown

    # ==================== Configuration Persistence ====================

    def _get_order_file_path(self) -> str:
        """Get order file path storage"""
        return os.path.join(self._get_config_dir(), "tab2_order_file.json")

    def _get_sku_db_file_path(self) -> str:
        """Get SKU database file path storage"""
        return os.path.join(self._get_config_dir(), "tab2_sku_db_file.json")

    def _get_templates_file_path(self) -> str:
        """Get templates file path"""
        return os.path.join(self._get_config_dir(), "tab2_target_templates.json")

    def _load_order_file(self) -> str:
        """Load order file path from storage"""
        path = self._get_order_file_path()
        if os.path.exists(path):
            try:
                with open(path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    return data.get("order_file", "")
            except Exception:
                pass
        return ""

    def _save_order_file(self, path: str) -> None:
        """Save order file path to storage"""
        try:
            with open(self._get_order_file_path(), 'w', encoding='utf-8') as f:
                json.dump({"order_file": path}, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self.log(f"保存订单文件路径失败: {e}", "warning")

    def _load_sku_db_file(self) -> str:
        """Load SKU database file path from storage"""
        path = self._get_sku_db_file_path()
        if os.path.exists(path):
            try:
                with open(path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    return data.get("sku_db_file", "")
            except Exception:
                pass
        return ""

    def _save_sku_db_file(self, path: str, sheet: str) -> None:
        """Save SKU database file path and sheet to storage"""
        try:
            with open(self._get_sku_db_file_path(), 'w', encoding='utf-8') as f:
                json.dump({"sku_db_file": path, "sku_db_sheet": sheet}, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self.log(f"保存SKU数据库文件路径失败: {e}", "warning")

    def _load_templates(self) -> Dict[str, Dict]:
        """Load column templates"""
        path = self._get_templates_file_path()
        if os.path.exists(path):
            try:
                with open(path, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception:
                pass
        return {}

    def _save_templates(self, templates: Dict[str, Dict]) -> None:
        """Save column templates"""
        try:
            with open(self._get_templates_file_path(), 'w', encoding='utf-8') as f:
                json.dump(templates, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self.show_error("保存模板失败", str(e))

    # ==================== Event Handlers ====================

    def _on_order_file_picked(self, e: ft.FilePickerResultEvent) -> None:
        """Handle order file pick"""
        if e.files:
            path = e.files[0].path
            self._order_file_path = path
            self._save_order_file(path)

            # Update display
            if hasattr(self, "选择订单表格_path_display"):
                colors = self.get_colors()
                self.选择订单表格_path_display.text = path
                self.选择订单表格_path_display.color = colors["text_primary"]
                self.选择订单表格_path_display.update()

            # Load sheets
            self._load_order_sheets(path)
            self.log(f"已选择订单文件: {os.path.basename(path)}", "info")

    def _load_order_sheets(self, file_path: str) -> None:
        """Load sheet names from order file"""
        try:
            from excel_toolkit.ui import get_sheet_names
            sheet_names = get_sheet_names(file_path)

            dropdown = getattr(self, "订单工作表_dropdown", None)
            if dropdown:
                dropdown.options = [ft.dropdown.Option(name) for name in sheet_names]
                if sheet_names:
                    dropdown.value = sheet_names[0]
                    self._order_sheet = sheet_names[0]
                    self.save_preference("order_sheet", self._order_sheet)
                dropdown.update()

        except Exception as e:
            self.log(f"加载工作表列表失败: {e}", "error")

    def _on_order_sheet_changed(self, e) -> None:
        """Handle order sheet dropdown change"""
        self._order_sheet = e.control.value or ""
        self.save_preference("order_sheet", self._order_sheet)

    def _on_sku_db_file_picked(self, e: ft.FilePickerResultEvent) -> None:
        """Handle SKU database file pick"""
        if e.files:
            path = e.files[0].path
            self._sku_db_file_path = path

            # Update display
            if hasattr(self, "选择SKU数据库_path_display"):
                colors = self.get_colors()
                self.选择SKU数据库_path_display.text = path
                self.选择SKU数据库_path_display.color = colors["text_primary"]
                self.选择SKU数据库_path_display.update()

            # Load sheets and headers
            self._load_sku_db_sheets(path)
            self.log(f"已选择SKU数据库: {os.path.basename(path)}", "info")

    def _load_sku_db_sheets(self, file_path: str) -> None:
        """Load SKU database sheets and headers"""
        try:
            from excel_toolkit.ui import get_sheet_names
            sheet_names = get_sheet_names(file_path)

            dropdown = getattr(self, "SKU数据库工作表_dropdown", None)
            if dropdown:
                dropdown.options = [ft.dropdown.Option(name) for name in sheet_names]

                # Smart default selection
                default_sheet = None
                if "商品资料" in sheet_names:
                    default_sheet = "商品资料"
                elif "SKU" in sheet_names:
                    default_sheet = "SKU"
                else:
                    default_sheet = sheet_names[0] if sheet_names else None

                if default_sheet:
                    dropdown.value = default_sheet
                    self._sku_db_sheet = default_sheet
                    self.save_preference("sku_db_sheet", self._sku_db_sheet)

                dropdown.update()

            # Load headers
            self._load_sku_headers(file_path, self._sku_db_sheet or default_sheet)

        except Exception as e:
            self.log(f"加载SKU数据库工作表失败: {e}", "error")

    def _on_sku_db_sheet_changed(self, e) -> None:
        """Handle SKU database sheet dropdown change"""
        self._sku_db_sheet = e.control.value or ""
        self.save_preference("sku_db_sheet", self._sku_db_sheet)
        self._save_sku_db_file(self._sku_db_file_path, self._sku_db_sheet)

        # Reload headers
        if self._sku_db_file_path:
            self._load_sku_headers(self._sku_db_file_path, self._sku_db_sheet)

    def _load_sku_headers(self, file_path: str, sheet_name: str) -> None:
        """Load and auto-identify SKU database headers"""
        try:
            from excel_toolkit.excel_lite import ExcelReader

            wb = ExcelReader(file_path, read_only=True, data_only=True)
            if sheet_name not in wb.sheetnames:
                wb.close()
                self.log(f"工作表 '{sheet_name}' 不存在", "warning")
                return

            ws = wb[sheet_name]
            headers = [cell.value for cell in list(ws.rows)[0] if cell.value]
            wb.close()

            if not headers:
                self.log("未找到表头，请手动配置列映射", "warning")
                return

            self._db_column_headers = headers

            # Update all database dropdowns
            for attr in ["sku", "l", "w", "h", "wt"]:
                dropdown = getattr(self, f"{attr}_dropdown", None)
                if dropdown:
                    dropdown.options = [ft.dropdown.Option(h) for h in headers]

                    # Auto-select based on intelligent mapping
                    current_value = getattr(self, f"_{attr}_col")
                    if not current_value:
                        mapping = identify_header_mapping(headers)
                        if mapping[attr]:
                            dropdown.value = mapping[attr]
                            setattr(self, f"_{attr}_col", mapping[attr])
                            dropdown.update()

            self.log("已自动识别列映射", "success")
            self.log(f"  SKU={self._db_sku_col}, 长={self._db_l_col}, 宽={self._db_w_col}, 高={self._db_h_col}, 重量={self._db_wt_col}", "info")
            self.log("  如有误，请手动调整下拉框", "info")

        except Exception as e:
            self.log(f"读取表头失败: {e}", "error")

    def _on_db_col_changed(self, attr: str, e) -> None:
        """Handle database column dropdown change"""
        value = e.control.value
        setattr(self, f"_{attr}_col", value)
        self.save_preference(f"db_{attr}_col", value)

    def _on_target_col_changed(self, attr: str, e) -> None:
        """Handle target column input change"""
        value = e.control.value
        setattr(self, f"_{attr}_col", value)
        self.save_preference(f"target_{attr}_col", value)

    def _on_ignore_qty_changed(self, e) -> None:
        """Handle ignore quantity checkbox change"""
        self._ignore_qty = e.control.value
        self.save_preference("ignore_qty", self._ignore_qty)

    def _on_template_changed(self, e) -> None:
        """Handle template dropdown change"""
        template_name = e.control.value
        self._template_name = template_name

        if template_name == "默认":
            return

        templates = self._load_templates()
        if template_name not in templates:
            self.log(f"模板 '{template_name}' 不存在", "warning")
            return

        config = templates[template_name]
        self._target_sku_col = config.get('sku', 'A')
        self._target_qty_col = config.get('qty', 'B')
        self._target_l_col = config.get('l', 'C')
        self._target_w_col = config.get('w', 'D')
        self._target_h_col = config.get('h', 'E')
        self._target_wt_col = config.get('wt', 'F')

        # Update input fields
        for attr in ["sku", "qty", "l", "w", "h", "wt"]:
            input_field = getattr(self, f"target_{attr}_input", None)
            if input_field:
                input_field.value = getattr(self, f"_target_{attr}_col")
                input_field.update()

        self.log(f"已加载模板 '{template_name}'", "info")

    def _on_save_template(self, e) -> None:
        """Save current configuration as template"""
        # Show dialog to get template name
        self._template_name_input = ft.TextField(label="模板名称", value=self._template_name)
        save_btn = ft.TextButton("保存", on_click=lambda _: self._confirm_save_template())
        cancel_btn = ft.TextButton("取消", on_click=lambda _: self._close_dialog())

        dialog = ft.AlertDialog(
            modal=True,
            title=ft.Text("保存配置模板"),
            content=self._template_name_input,
            actions=[save_btn, cancel_btn],
        )

        self.page.dialog = dialog
        dialog.open = True
        self.page.update()

    def _confirm_save_template(self) -> None:
        """Confirm and save template"""
        template_name = self._template_name_input.value.strip()
        self._close_dialog()

        if not template_name:
            self.show_error("保存模板", "模板名称不能为空")
            return

        # Collect current configuration
        config = {
            'sku': self._target_sku_col,
            'qty': self._target_qty_col,
            'l': self._target_l_col,
            'w': self._target_w_col,
            'h': self._target_h_col,
            'wt': self._target_wt_col,
        }

        # Save template
        templates = self._load_templates()
        templates[template_name] = config
        self._save_templates(templates)

        # Refresh and select template
        self._refresh_template_list()
        self.template_dropdown.value = template_name
        self._template_name = template_name
        self.template_dropdown.update()

        self.log(f"配置模板 '{template_name}' 已保存", "success")
        self.show_info("保存成功", f"配置模板 '{template_name}' 已保存！")

    def _on_delete_template(self, e) -> None:
        """Delete current template"""
        if self._template_name == "默认":
            self.show_error("删除模板", "不能删除默认模板")
            return

        # Confirm deletion
        dialog = ft.AlertDialog(
            modal=True,
            title=ft.Text("确认删除"),
            content=ft.Text(f"确定要删除模板 '{self._template_name}' 吗？"),
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
        template_name = self._template_name
        self._close_dialog()

        templates = self._load_templates()
        if template_name in templates:
            del templates[template_name]
            self._save_templates(templates)

            # Refresh template list
            self._refresh_template_list()
            self.template_dropdown.value = "默认"
            self._template_name = "默认"
            self.template_dropdown.update()

            self.log(f"模板 '{template_name}' 已删除", "info")

    def _refresh_template_list(self) -> None:
        """Refresh template dropdown options"""
        templates = self._load_templates()
        options = [ft.dropdown.Option("默认")]
        for name in templates.keys():
            options.append(ft.dropdown.Option(name))

        self.template_dropdown.options = options
        self.template_dropdown.update()

    # ==================== Processing ====================

    def _on_run_click(self, e) -> None:
        """Handle run button click"""
        # Validate inputs
        if not self._validate_inputs():
            return

        # Build column mappings
        db_col_map = {
            'sku': self._db_sku_col,
            'l': self._db_l_col,
            'w': self._db_w_col,
            'h': self._db_h_col,
            'wt': self._db_wt_col,
        }

        target_col_map = {
            'sku': self._target_sku_col,
            'qty': self._target_qty_col,
            'l': self._target_l_col,
            'w': self._target_w_col,
            'h': self._target_h_col,
            'wt': self._target_wt_col,
        }

        # Check for missing mappings
        missing_db = [k for k, v in db_col_map.items() if not v]
        if missing_db:
            self.show_error("验证错误", f"请配置SKU数据库的列映射: {', '.join(missing_db)}")
            return

        missing_target = [k for k, v in target_col_map.items() if not v]
        if missing_target:
            self.show_error("验证错误", f"请配置目标表格的列: {', '.join(missing_target)}")
            return

        # Run processing
        self._run_processing(self._order_file_path, self._sku_db_file_path, db_col_map, target_col_map)

    def _validate_inputs(self) -> bool:
        """Validate all inputs"""
        if not self._order_file_path or self._order_file_path == "未选择文件":
            self.show_error("验证错误", "请选择订单表格！")
            return False

        if not self._sku_db_file_path or self._sku_db_file_path == "未选择SKU数据库":
            self.show_error("验证错误", "请选择SKU数据库表格！")
            return False

        if not self._sku_db_sheet:
            self.show_error("验证错误", "请选择SKU数据库工作表！")
            return False

        return True

    def _run_processing(self, order_file: str, sku_db_file: str, db_col_map: Dict, target_col_map: Dict) -> None:
        """Run SKU filling in background thread"""
        self.set_processing("正在填充...")
        self.page.cursor = ft.Cursor.WAIT

        # Log start
        self.log("=" * 60, "info")
        self.log("开始执行SKU智能填充...", "info")
        self.log(f"  订单文件: {os.path.basename(order_file)}", "info")
        self.log(f"  SKU数据库: {os.path.basename(sku_db_file)}", "info")
        self.log(f"  数据库工作表: {self._sku_db_sheet}", "info")
        self.log("=" * 60, "info")

        def process_thread():
            try:
                # Create thread-safe logger
                def thread_safe_logger(msg: str):
                    self.log(msg, "info")

                # Process SKU filling
                stats = process_skus(
                    order_file,
                    sku_db_file,
                    db_col_map,
                    target_col_map,
                    thread_safe_logger,
                    self._sku_db_sheet if self._sku_db_sheet else None,
                    self._ignore_qty,
                    self._order_sheet if self._order_sheet else None
                )

                # Success callback
                def on_success():
                    self.page.cursor = ft.Cursor.DEFAULT
                    self.set_success("填充完成")

                    self.log("", "info")
                    self.log("=" * 60, "success")
                    self.log("SKU填充完成！", "success")
                    self.log(f"  处理工作表: {stats['sheets_processed']}", "info")
                    self.log(f"  填充行数: {stats['rows_filled']}", "info")
                    self.log("=" * 60, "success")

                    # Show success dialog
                    if stats['sheets_processed'] > 0:
                        self.show_success(
                            f"SKU填充完成！\n\n"
                            f"处理工作表数: {stats['sheets_processed']}\n"
                            f"填充行数: {stats['rows_filled']}\n\n"
                            f"文件已保存: {os.path.basename(order_file)}"
                        )
                    else:
                        self.show_info(
                            "SKU填充完成",
                            "未处理任何数据（可能未找到匹配列或数据为空）"
                        )

                self.page.run_thread(on_success)

            except Exception as e:
                # Error handling
                log_error(e, "SKU填充")
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

    def _on_clear_log_click(self, e) -> None:
        """Handle clear log button click"""
        if self._log_widget:
            self._log_widget.clear()
            self.log("日志已清空", "info")


def create_view(page: ft.Page) -> SkuFillView:
    """
    Factory function to create the SKU filling view.

    Args:
        page: Flet page control

    Returns:
        SkuFillView instance
    """
    view = SkuFillView(page, tab_index=1)
    return view
