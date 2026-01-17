"""
Excel Toolkit - Flet-based Application Entry Point

Modern GUI application using Flet framework with Windows 11/macOS flat design style.
"""

import os
import sys

# Add project root to path for imports
project_root = os.path.dirname(os.path.abspath(__file__))
if project_root not in sys.path:
    sys.path.insert(0, project_root)

import flet as ft

from app.core.theme import AppTheme, ThemeMode
from app.core.constants import Text, FontSize, Spacing, Size
from app.core.config import ConfigManager, get_config_manager
from app.core.app_state import get_state, AppState
from app.views.tab01_states import StateConversionView
from app.views.tab02_skus import SkuFillView
from app.views.tab03_highlight import HighlightDuplicatesView
from app.views.tab04_insert import InsertRowsView
from app.views.tab05_compare import CompareColumnsView
from app.views.tab07_prefix import PrefixFillView
from app.views.tab12_ppt import PptToPdfView
from app.views.tab13_image import ImageCompressView
from app.views.tab14_delete_cols import DeleteColumnsView


class ExcelToolkitApp:
    """
    Main Excel Toolkit application using Flet framework.

    Features:
    - Modern flat design (Windows 11/macOS style)
    - Theme switching (light/dark/system)
    - 14 tool tabs for Excel operations
    - Configuration persistence
    """

    def __init__(self):
        """Initialize the application"""
        self._state = get_state()
        self._config = get_config_manager()
        self._theme_mode = ThemeMode.SYSTEM

        # Register state change listeners
        self._state.add_listener("theme_changed", self._on_theme_changed)
        self._state.add_listener("topmost_changed", self._on_topmost_changed)
        self._state.add_listener("status_changed", self._on_status_changed)

    def run(self) -> None:
        """Start the Flet application"""
        ft.app(target=self.main)

    def main(self, page: ft.Page) -> None:
        """
        Main entry point for the Flet application.

        Args:
            page: Flet page control
        """
        self.page = page

        # Load configuration
        config = self._config.load()
        self._theme_mode = ThemeMode(config.theme_mode)

        # Apply theme
        page.theme = AppTheme.get_flet_theme(self._theme_mode)
        page.theme_mode = ft.ThemeMode.SYSTEM if self._theme_mode == ThemeMode.SYSTEM else (
            ft.ThemeMode.DARK if self._theme_mode == ThemeMode.DARK else ft.ThemeMode.LIGHT
        )

        # Page settings
        page.title = f"{Text.APP_TITLE} V{Text.APP_VERSION}"
        page.window_width = Size.WINDOW_WIDTH
        page.window_height = Size.WINDOW_HEIGHT
        page.window_min_width = Size.WINDOW_MIN_WIDTH
        page.window_min_height = Size.WINDOW_MIN_HEIGHT
        page.window_always_on_top = config.always_on_top
        page.padding = 0
        page.spacing = 0

        # Apply window size from config
        if config.window_left and config.window_top:
            page.window_left = config.window_left
            page.window_top = config.window_top

        # Create UI
        self._create_ui()

        # Sync state with loaded config
        self._state.always_on_top = config.always_on_top
        self._state.theme_mode = self._theme_mode

        # Update page
        page.update()

    def _create_ui(self) -> None:
        """Create the main UI layout"""
        colors = AppTheme.get_colors(self._theme_mode)

        # App bar (header)
        self.app_bar = ft.AppBar(
            leading=ft.Icon(ft.Icons.TABLE_CHART, color=colors["on_primary"]),
            title=ft.Text(
                f"{Text.APP_TITLE} V{Text.APP_VERSION}",
                style=ft.TextStyle(
                    size=FontSize.TITLE,
                    weight=ft.FontWeight.BOLD,
                ),
            ),
            bgcolor=colors["primary"],
            color=colors["on_primary"],
            actions=[
                # Theme dropdown
                ft.Dropdown(
                    options=[
                        ft.dropdown.Option(Text.THEME_LIGHT),
                        ft.dropdown.Option(Text.THEME_DARK),
                        ft.dropdown.Option(Text.THEME_SYSTEM),
                    ],
                    width=100,
                    value=Text.THEME_SYSTEM if self._theme_mode == ThemeMode.SYSTEM else (
                        Text.THEME_DARK if self._theme_mode == ThemeMode.DARK else Text.THEME_LIGHT
                    ),
                    on_change=self._on_theme_dropdown_change,
                    text_style=ft.TextStyle(size=FontSize.LABEL),
                    bgcolor=colors["primary_container"],
                    color=colors["on_background"] if self._theme_mode == ThemeMode.DARK else colors["text_primary"],
                ),
                # Topmost toggle
                ft.IconButton(
                    icon=ft.Icons.PUSH_PIN,
                    tooltip="窗口置顶",
                    icon_color=colors["on_primary"],
                    selected=self.page.window_always_on_top,
                    selected_icon=ft.Icons.PUSH_PIN,
                    on_click=self._on_topmost_click,
                ),
                # About button
                ft.IconButton(
                    icon=ft.Icons.INFO_OUTLINE,
                    tooltip="关于",
                    icon_color=colors["on_primary"],
                    on_click=self._on_about_click,
                ),
            ],
        )

        # Navigation rail
        self.navigation = ft.NavigationRail(
            selected_index=0,
            label_type=ft.NavigationRailLabelType.ALL,
            min_width=100,
            min_extended_width=200,
            destinations=self._create_nav_destinations(),
            on_change=self._on_nav_change,
            bgcolor=colors["background"],
        )

        # Main content area (tabs)
        self.content_area = ft.Tabs(
            selected_index=0,
            animation_duration=300,
            tabs=self._create_tab_contents(),
            expand=True,
        )

        # Status bar
        self.status_bar = ft.Container(
            content=ft.Row([
                ft.Icon(
                    ft.Icons.CHECK_CIRCLE,
                    size=16,
                    color=colors["success"],
                ),
                ft.Text(
                    Text.STATUS_READY,
                    style=ft.TextStyle(size=FontSize.STATUS),
                    color=colors["text_secondary"],
                    expand=True,
                ),
                ft.Text(
                    f"快捷键: F1=帮助",
                    style=ft.TextStyle(size=8),
                    color=colors["text_hint"],
                ),
            ], spacing=Spacing.CONTROL_PADDING_X),
            bgcolor=colors["surface_variant"],
            padding=ft.padding.symmetric(
                horizontal=Spacing.SECTION_PADDING,
                vertical=5,
            ),
        )

        # Warning banner
        self.warning_banner = ft.Container(
            content=ft.Row([
                ft.Icon(ft.Icons.WARNING, color=colors["warning"], size=16),
                ft.Text(
                    "重要提示：处理文件前请确保已关闭 Excel/WPS，避免保存失败！",
                    style=ft.TextStyle(size=FontSize.LABEL),
                    color=colors["text_primary"],
                    expand=True,
                ),
            ], spacing=Spacing.CONTROL_PADDING_X),
            bgcolor=colors["warning_container"],
            padding=ft.padding.symmetric(
                horizontal=Spacing.SECTION_PADDING,
                vertical=8,
            ),
        )

        # Main layout
        self.page.appbar = self.app_bar
        self.page.add(
            ft.Row([
                self.navigation,
                ft.VerticalDivider(width=1, color=colors["border"]),
                ft.Column([
                    self.warning_banner,
                    self.content_area,
                ], expand=True, spacing=0),
            ], expand=True, spacing=0),
            self.status_bar,
        )

    def _create_nav_destinations(self) -> list[ft.NavigationRailDestination]:
        """Create navigation rail destinations"""
        colors = AppTheme.get_colors(self._theme_mode)
        destinations = []

        # Create destinations for each tab
        tab_labels = [
            "州名转换", "SKU填充", "高亮重复", "插入行",
            "对比列", "PDF拆分", "前缀填充", "面单页脚",
            "仓库推荐", "录入库存", "模板填充", "PPT转PDF",
            "图片压缩", "删除列",
        ]

        icons = [
            ft.Icons.MAP_OUTLINED, ft.Icons.TAG_OUTLINED, ft.Icons.HIGHLIGHT_OUTLINED, ft.Icons.PLAYLIST_ADD_OUTLINED,
            ft.Icons.COMPARE_OUTLINED, ft.Icons.PICTURE_AS_PDF_OUTLINED, ft.Icons.TEXT_FIELDS_OUTLINED, ft.Icons.INSERT_PAGE_BREAK_OUTLINED,
            ft.Icons.LOCATION_ON_OUTLINED, ft.Icons.EDIT_NOTE_OUTLINED, ft.Icons.LOCAL_SHIPPING_OUTLINED, ft.Icons.SLIDESHOW_OUTLINED,
            ft.Icons.IMAGE_OUTLINED, ft.Icons.DELETE_OUTLINED,
        ]

        for i, (label, icon) in enumerate(zip(tab_labels, icons)):
            destinations.append(
                ft.NavigationRailDestination(
                    icon=icon,
                    selected_icon=icon,
                    label=f"[{i+1}] {label}",
                )
            )

        return destinations

    def _create_tab_contents(self) -> list[ft.Tab]:
        """Create tab contents for the main content area"""
        colors = AppTheme.get_colors(self._theme_mode)
        tabs = []

        tab_names = [
            "州名转换", "SKU填充", "高亮重复", "插入行",
            "对比列", "PDF拆分", "前缀填充", "面单页脚",
            "仓库推荐", "录入库存", "模板填充", "PPT转PDF",
            "图片压缩", "删除列",
        ]

        # Map tab index to view class (implemented tabs)
        view_classes = {
            0: StateConversionView,    # Tab 1: 州名转换 ✅
            1: SkuFillView,            # Tab 2: SKU填充 ✅
            2: HighlightDuplicatesView, # Tab 3: 高亮重复 ✅
            3: InsertRowsView,         # Tab 4: 插入行 ✅
            4: CompareColumnsView,     # Tab 5: 对比列 ✅
            6: PrefixFillView,         # Tab 7: 前缀填充 ✅
            11: PptToPdfView,          # Tab 12: PPT转PDF ✅
            12: ImageCompressView,     # Tab 13: 图片压缩 ✅
            13: DeleteColumnsView,     # Tab 14: 删除列 ✅
        }

        for i, name in enumerate(tab_names):
            # Check if this tab has an implemented view
            if i in view_classes:
                view_class = view_classes[i]
                view = view_class(self.page, tab_index=i)
                view.theme_mode = self._theme_mode
                tab_content = view.build()
            else:
                # Placeholder content for unimplemented tabs
                tab_content = ft.Container(
                    content=ft.Column([
                        ft.Icon(ft.Icons.CONSTRUCTION, size=48, color=colors["text_hint"]),
                        ft.Text(
                            f"{name} - 即将推出",
                            style=ft.TextStyle(
                                size=FontSize.SUBTITLE,
                                color=colors["text_hint"]
                            ),
                        ),
                        ft.Text(
                            "此功能正在开发中，请使用原 Tkinter 版本。",
                            style=ft.TextStyle(
                                size=FontSize.LABEL,
                                color=colors["text_secondary"]
                            ),
                        ),
                    ], horizontal_alignment=ft.CrossAxisAlignment.CENTER),
                    alignment=ft.alignment.center,
                    expand=True,
                )

            tab = ft.Tab(
                text=f"[{i+1}] {name}",
                content=tab_content,
            )
            tabs.append(tab)

        return tabs

    # ==================== Event Handlers ====================

    def _on_theme_dropdown_change(self, e):
        """Handle theme dropdown change"""
        value = e.control.value

        if value == Text.THEME_LIGHT:
            new_mode = ThemeMode.LIGHT
        elif value == Text.THEME_DARK:
            new_mode = ThemeMode.DARK
        else:
            new_mode = ThemeMode.SYSTEM

        self._state.theme_mode = new_mode

    def _on_theme_changed(self, mode: ThemeMode):
        """Handle theme change from state"""
        self._theme_mode = mode

        # Update page theme
        self.page.theme = AppTheme.get_flet_theme(mode)
        self.page.theme_mode = ft.ThemeMode.SYSTEM if mode == ThemeMode.SYSTEM else (
            ft.ThemeMode.DARK if mode == ThemeMode.DARK else ft.ThemeMode.LIGHT
        )

        # Save to config
        self._config.set_theme_mode(mode)
        self._config.save()

        # Update theme dropdown
        theme_value = Text.THEME_SYSTEM if mode == ThemeMode.SYSTEM else (
            Text.THEME_DARK if mode == ThemeMode.DARK else Text.THEME_LIGHT
        )
        for action in self.app_bar.actions:
            if isinstance(action, ft.Dropdown):
                action.value = theme_value
                action.update()
                break

    def _on_topmost_click(self, e):
        """Handle topmost toggle button click"""
        current = self.page.window_always_on_top
        self.page.window_always_on_top = not current
        self._state.always_on_top = not current
        e.control.selected = not current
        e.control.update()

    def _on_topmost_changed(self, value: bool):
        """Handle topmost state change from state"""
        self.page.window_always_on_top = value
        self._config.set_always_on_top(value)
        self._config.save()

    def _on_status_changed(self, status):
        """Handle status change from state"""
        colors = AppTheme.get_colors(self._theme_mode)

        # Update status bar content
        status_row = self.status_bar.content
        if isinstance(status_row, ft.Row) and len(status_row.controls) >= 2:
            # Update icon
            status_row.controls[0].icon = self._status_to_icon(status.icon)
            status_row.controls[0].color = self._status_to_color(status.icon, colors)

            # Update text
            status_row.controls[1].text = status.message
            status_row.controls[1].update()

        self.status_bar.update()

    def _status_to_icon(self, icon: str) -> str:
        """Convert status icon string to Flet icon name"""
        icon_map = {
            "✅": ft.Icons.CHECK_CIRCLE,
            "⏳": ft.Icons.HOURGLASS_EMPTY,
            "❌": ft.Icons.ERROR,
            "⚠️": ft.Icons.WARNING,
            "ℹ️": ft.Icons.INFO,
        }
        return icon_map.get(icon, ft.Icons.INFO)

    def _status_to_color(self, icon: str, colors: dict) -> str:
        """Convert status icon to color"""
        color_map = {
            "✅": colors["success"],
            "⏳": colors["info"],
            "❌": colors["error"],
            "⚠️": colors["warning"],
            "ℹ️": colors["info"],
        }
        return color_map.get(icon, colors["info"])

    def _on_nav_change(self, e):
        """Handle navigation rail change"""
        selected_index = e.control.selected_index
        self.content_area.selected_index = selected_index
        self.content_area.update()
        self._state.current_tab = selected_index

    def _on_about_click(self, e):
        """Handle about button click"""
        colors = AppTheme.get_colors(self._theme_mode)

        # Create about dialog
        about_dialog = ft.AlertDialog(
            modal=True,
            title=ft.Row([
                ft.Icon(ft.Icons.TABLE_CHART, size=24, color=colors["primary"]),
                ft.Text(f"{Text.APP_TITLE}", size=20, weight=ft.FontWeight.BOLD),
            ]),
            content=ft.Column([
                ft.Text(f"Version {Text.APP_VERSION}", size=FontSize.SUBTITLE),
                ft.Divider(),
                ft.Text(f"作者: {Text.APP_AUTHOR}"),
                ft.Text("技术栈: Python + Flet"),
                ft.Text(f"功能数量: 14 个工具"),
                ft.Divider(),
                ft.Text("© 2025 All Rights Reserved", size=8),
            ], spacing=10, tight=True),
            actions=[
                ft.TextButton("关闭", on_click=lambda e: self.close_dialog()),
            ],
            actions_alignment=ft.MainAxisAlignment.END,
        )

        self.page.dialog = about_dialog
        about_dialog.open = True
        self.page.update()

    def close_dialog(self):
        """Close the current dialog"""
        if self.page.dialog:
            self.page.dialog.open = False
            self.page.update()


def main():
    """Entry point for the application"""
    app = ExcelToolkitApp()
    app.run()


if __name__ == "__main__":
    main()
