"""
Common Components Module - Reusable UI components for the Flet-based Excel Toolkit

Provides file picker, log display, card container, and other common UI elements.
"""

import os
from typing import Callable, Optional, List
from datetime import datetime

import flet as ft

from app.core.theme import AppTheme, ThemeMode
from app.core.constants import Text, FontSize, Spacing, Icon


class Card(ft.Container):
    """
    A card container with consistent styling.

    Provides a Windows 11/macOS style card with rounded corners,
    border, and optional shadow.
    """

    def __init__(
        self,
        content: ft.Control,
        title: Optional[str] = None,
        title_icon: Optional[str] = None,
        padding: int = Spacing.SECTION_PADDING,
        theme_mode: ThemeMode = ThemeMode.LIGHT,
        **kwargs
    ):
        """
        Initialize a card component.

        Args:
            content: Card content control
            title: Optional card title
            title_icon: Optional title icon (Flet icon name)
            padding: Internal padding
            theme_mode: Theme mode for styling
            **kwargs: Additional Container properties
        """
        colors = AppTheme.get_colors(theme_mode)

        # Build card content
        if title:
            title_row_items = []
            if title_icon:
                title_row_items.append(ft.Icon(title_icon, size=16, color=colors["primary"]))
            title_row_items.append(ft.Text(
                title,
                style=ft.TextStyle(
                    size=FontSize.SECTION,
                    weight=ft.FontWeight.BOLD,
                    color=colors["text_primary"]
                )
            ))
            card_content = ft.Column([
                # Title row
                ft.Row(title_row_items, spacing=5),
                ft.Divider(height=1, color=colors["border"]),
                content,
            ], spacing=Spacing.SECTION_SPACING, expand=True)
        else:
            card_content = content

        super().__init__(
            content=card_content,
            bgcolor=colors["surface"],
            border_radius=8,
            padding=padding,
            border=ft.border.all(1, colors["border"]),
            **kwargs
        )


class FilePickerCard(ft.Container):
    """
    A file picker component with button and path display.

    Provides a consistent file selection UI with icon, button,
    and path display with tooltip.
    """

    def __init__(
        self,
        label: str,
        on_pick: Callable[[str], None],
        file_type: Optional[str] = None,
        initial_path: str = Text.FILE_NOT_SELECTED,
        theme_mode: ThemeMode = ThemeMode.LIGHT,
        dialog_title: str = "选择文件",
        page: Optional[ft.Page] = None,
    ):
        """
        Initialize a file picker card.

        Args:
            label: Label text for the file picker
            on_pick: Callback when file is selected (receives file path)
            file_type: File type filter (e.g., "xlsx", "pdf"), None for all files
            initial_path: Initial path display text
            theme_mode: Theme mode for styling
            dialog_title: Title for the file picker dialog
            page: Flet page object (required for FilePicker to work)
        """
        self._on_pick = on_pick
        self._file_type = file_type
        self._theme_mode = theme_mode
        self._dialog_title = dialog_title
        self._page_ref = page  # Store page reference for async callbacks

        colors = AppTheme.get_colors(theme_mode)

        # Path display
        self.path_display = ft.Text(
            initial_path,
            style=ft.TextStyle(
                size=FontSize.LABEL,
                color=colors["text_secondary"] if initial_path == Text.FILE_NOT_SELECTED else colors["text_primary"],
            ),
            max_lines=1,
            overflow=ft.TextOverflow.ELLIPSIS,
            expand=True,
            tooltip=initial_path if initial_path != Text.FILE_NOT_SELECTED else None,
        )

        # File picker (hidden, triggered by button)
        self.file_picker = ft.FilePicker(on_result=self._on_pick_result)

        # Add FilePicker to page overlay if page is provided
        # Don't update here - let the page update when it's ready
        if page is not None:
            page.overlay.append(self.file_picker)

        # Browse button
        self.browse_button = ft.ElevatedButton(
            "浏览",
            icon=ft.Icons.FOLDER_OPEN,
            on_click=self._on_browse_click,
            style=ft.ButtonStyle(
                bgcolor=colors["primary"],
                color=colors["on_primary"],
            ),
        )

        # Main layout
        super().__init__(
            content=ft.Column([
                # Label row
                ft.Row([
                    ft.Text(
                        label,
                        style=ft.TextStyle(
                            size=FontSize.LABEL,
                            weight=ft.FontWeight.W_500,
                            color=colors["text_secondary"]
                        )
                    ),
                ]),
                # Path and button row
                ft.Row([
                    ft.Container(
                        content=self.path_display,
                        bgcolor=colors["surface_variant"],
                        border_radius=4,
                        padding=ft.padding.symmetric(horizontal=10, vertical=8),
                        expand=True,
                    ),
                    self.browse_button,
                ], spacing=Spacing.CONTROL_PADDING_X),
            ], spacing=Spacing.ROW_SPACING),
        )

    def _on_browse_click(self, e):
        """Handle browse button click"""
        # Set up file type filter
        if self._file_type == "xlsx":
            # Excel files filter
            self.file_picker.pick_files(
                dialog_title=self._dialog_title,
                allowed_extensions=["xlsx", "xls"],
            )
        elif self._file_type == "pdf":
            # PDF files filter
            self.file_picker.pick_files(
                dialog_title=self._dialog_title,
                allowed_extensions=["pdf"],
            )
        elif self._file_type == "ppt":
            # PowerPoint files filter
            self.file_picker.pick_files(
                dialog_title=self._dialog_title,
                allowed_extensions=["pptx", "ppt"],
            )
        elif self._file_type == "image":
            # Image files filter
            self.file_picker.pick_files(
                dialog_title=self._dialog_title,
                allowed_extensions=["jpg", "jpeg", "png", "gif", "bmp", "webp"],
            )
        else:
            # Any file type
            self.file_picker.pick_files(
                dialog_title=self._dialog_title,
            )

    def _get_file_types(self) -> str:
        """Get file type description"""
        type_map = {
            "xlsx": "Excel 文件",
            "pdf": "PDF 文件",
            "ppt": "PowerPoint 文件",
            "image": "图片文件",
        }
        return type_map.get(self._file_type, "所有文件")

    def _on_pick_result(self, e: ft.FilePickerResultEvent):
        """Handle file picker result"""
        # TEMPORARY: Minimal callback to test if dialog closes
        if e.files:
            path = e.files[0].path
            print(f"[DEBUG FilePickerCard] File selected: {path}")
            # Only call the callback - don't update UI yet
            self._on_pick(path)

    def set_path(self, path: str, update_callback: bool = True) -> None:
        """
        Set the displayed file path.

        Args:
            path: File path to display
            update_callback: Whether to call the on_pick callback
        """
        colors = AppTheme.get_colors(self._theme_mode)

        self.path_display.text = path
        self.path_display.color = colors["text_primary"]
        self.path_display.tooltip = path
        self.path_display.update()

        if update_callback:
            self._on_pick(path)

    def get_path(self) -> str:
        """Get the current file path"""
        return self.path_display.text

    def clear(self) -> None:
        """Clear the file picker"""
        colors = AppTheme.get_colors(self._theme_mode)
        self.path_display.text = Text.FILE_NOT_SELECTED
        self.path_display.color = colors["text_secondary"]
        self.path_display.tooltip = None
        self.path_display.update()


class DirectoryPickerCard(ft.Container):
    """
    A directory picker component with button and path display.
    """

    def __init__(
        self,
        label: str,
        on_pick: Callable[[str], None],
        initial_path: str = Text.DIR_NOT_SELECTED,
        theme_mode: ThemeMode = ThemeMode.LIGHT,
        dialog_title: str = "选择目录",
        page: Optional[ft.Page] = None,
    ):
        """
        Initialize a directory picker card.

        Args:
            label: Label text for the directory picker
            on_pick: Callback when directory is selected
            initial_path: Initial path display text
            theme_mode: Theme mode for styling
            dialog_title: Title for the directory picker dialog
            page: Flet page object (required for FilePicker to work)
        """
        self._on_pick = on_pick
        self._theme_mode = theme_mode
        self._dialog_title = dialog_title

        colors = AppTheme.get_colors(theme_mode)

        # Path display
        self.path_display = ft.Text(
            initial_path,
            style=ft.TextStyle(
                size=FontSize.LABEL,
                color=colors["text_secondary"] if initial_path == Text.DIR_NOT_SELECTED else colors["text_primary"],
            ),
            max_lines=1,
            overflow=ft.TextOverflow.ELLIPSIS,
            expand=True,
            tooltip=initial_path if initial_path != Text.DIR_NOT_SELECTED else None,
        )

        # Directory picker (hidden)
        self.dir_picker = ft.FilePicker(on_result=self._on_pick_result)

        # Add FilePicker to page overlay if page is provided
        # Don't update here - let the page update when it's ready
        if page is not None:
            page.overlay.append(self.dir_picker)

        # Browse button
        self.browse_button = ft.ElevatedButton(
            "浏览",
            icon=ft.Icons.FOLDER_OPEN,
            on_click=self._on_browse_click,
            style=ft.ButtonStyle(
                bgcolor=colors["primary"],
                color=colors["on_primary"],
            ),
        )

        # Main layout
        super().__init__(
            content=ft.Column([
                ft.Row([
                    ft.Text(
                        label,
                        style=ft.TextStyle(
                            size=FontSize.LABEL,
                            weight=ft.FontWeight.W_500,
                            color=colors["text_secondary"]
                        )
                    ),
                ]),
                ft.Row([
                    ft.Container(
                        content=self.path_display,
                        bgcolor=colors["surface_variant"],
                        border_radius=4,
                        padding=ft.padding.symmetric(horizontal=10, vertical=8),
                        expand=True,
                    ),
                    self.browse_button,
                ], spacing=Spacing.CONTROL_PADDING_X),
            ], spacing=Spacing.ROW_SPACING),
        )

    def _on_browse_click(self, e):
        """Handle browse button click"""
        self.dir_picker.get_directory_path(dialog_title=self._dialog_title)

    def _on_pick_result(self, e: ft.FilePickerResultEvent):
        """Handle directory picker result"""
        if e.path:
            path = e.path
            self.set_path(path)
            self._on_pick(path)

    def set_path(self, path: str, update_callback: bool = True) -> None:
        """Set the displayed directory path"""
        colors = AppTheme.get_colors(self._theme_mode)

        self.path_display.text = path
        self.path_display.color = colors["text_primary"]
        self.path_display.tooltip = path
        self.path_display.update()

        if update_callback:
            self._on_pick(path)

    def get_path(self) -> str:
        """Get the current directory path"""
        return self.path_display.text

    def clear(self) -> None:
        """Clear the directory picker"""
        colors = AppTheme.get_colors(self._theme_mode)
        self.path_display.text = Text.DIR_NOT_SELECTED
        self.path_display.color = colors["text_secondary"]
        self.path_display.tooltip = None
        self.path_display.update()


class LogWidget(ft.Container):
    """
    A log display widget with timestamp, colored output, and clear button.
    """

    def __init__(
        self,
        height: int = 200,
        theme_mode: ThemeMode = ThemeMode.LIGHT,
        title: str = "日志",
        show_timestamp: bool = True,
        **kwargs
    ):
        """
        Initialize a log widget.

        Args:
            height: Height of the log area
            theme_mode: Theme mode for styling
            title: Title for the log widget
            show_timestamp: Whether to show timestamps
            **kwargs: Additional Container properties
        """
        self._theme_mode = theme_mode
        self._show_timestamp = show_timestamp

        colors = AppTheme.get_colors(theme_mode)

        # Log text column
        self.log_column = ft.Column(
            [],
            spacing=2,
            scroll=ft.ScrollMode.AUTO,
            expand=True,
        )

        # Clear button
        self.clear_button = ft.IconButton(
            icon=ft.Icons.CLEAR,
            tooltip="清空日志",
            icon_color=colors["text_secondary"],
            on_click=self._on_clear,
        )

        # Header row
        header = ft.Row([
            ft.Text(
                title,
                style=ft.TextStyle(
                    size=FontSize.SECTION,
                    weight=ft.FontWeight.BOLD,
                    color=colors["text_primary"]
                )
            ),
            ft.Container(expand=True),
            self.clear_button,
        ])

        # Main layout
        super().__init__(
            content=ft.Column([
                header,
                ft.Divider(height=1, color=colors["border"]),
                self.log_column,
            ], spacing=Spacing.SECTION_SPACING),
            bgcolor=colors["surface_variant"],
            border_radius=8,
            padding=Spacing.SECTION_PADDING,
            border=ft.border.all(1, colors["border"]),
            height=height,
            **kwargs
        )

    def _get_timestamp(self) -> str:
        """Get current timestamp string"""
        if self._show_timestamp:
            return datetime.now().strftime("%H:%M:%S")
        return ""

    def log(self, message: str, level: str = "info") -> None:
        """
        Add a log message.

        Args:
            message: Log message text
            level: Log level - "info", "success", "warning", "error"
        """
        colors = AppTheme.get_colors(self._theme_mode)

        # Get color based on level
        level_colors = {
            "info": colors["text_primary"],
            "success": colors["success"],
            "warning": colors["warning"],
            "error": colors["error"],
        }
        text_color = level_colors.get(level, colors["text_primary"])

        # Create timestamp
        timestamp = self._get_timestamp()

        # Build log entry
        if timestamp:
            log_text = ft.Text(
                f"[{timestamp}] {message}",
                style=ft.TextStyle(
                    font_family="Consolas",
                    size=FontSize.LOG,
                    color=text_color,
                ),
                selectable=True,
            )
        else:
            log_text = ft.Text(
                message,
                style=ft.TextStyle(
                    font_family="Consolas",
                    size=FontSize.LOG,
                    color=text_color,
                ),
                selectable=True,
            )

        self.log_column.controls.append(log_text)
        self.log_column.update()

        # Auto-scroll to bottom
        self.log_column.scroll_to(offset=-1, duration=100)

    def info(self, message: str) -> None:
        """Add an info log message"""
        self.log(message, "info")

    def success(self, message: str) -> None:
        """Add a success log message"""
        self.log(message, "success")

    def warning(self, message: str) -> None:
        """Add a warning log message"""
        self.log(message, "warning")

    def error(self, message: str) -> None:
        """Add an error log message"""
        self.log(message, "error")

    def _on_clear(self, e) -> None:
        """Handle clear button click"""
        self.clear()

    def clear(self) -> None:
        """Clear all log messages"""
        self.log_column.controls.clear()
        self.log_column.update()

    def get_log_text(self) -> str:
        """Get all log messages as a single string"""
        return "\n".join(
            control.text for control in self.log_column.controls
            if isinstance(control, ft.Text)
        )


class InputField(ft.Row):
    """
    A labeled input field with consistent styling.
    """

    def __init__(
        self,
        label: str,
        value: str = "",
        placeholder: Optional[str] = None,
        width: int = 200,
        theme_mode: ThemeMode = ThemeMode.LIGHT,
        password: bool = False,
        **kwargs
    ):
        """
        Initialize an input field.

        Args:
            label: Label text
            value: Initial value
            placeholder: Placeholder text
            width: Input field width
            theme_mode: Theme mode for styling
            password: Whether to hide input (password field)
            **kwargs: Additional TextField properties
        """
        colors = AppTheme.get_colors(theme_mode)

        self.label = ft.Text(
            label,
            style=ft.TextStyle(
                size=FontSize.LABEL,
                color=colors["text_secondary"],
            ),
            width=80,
        )

        self.text_field = ft.TextField(
            value=value,
            placeholder=placeholder,
            width=width,
            password=password,
            bgcolor=colors["surface"],
            border_color=colors["border"],
            focused_border_color=colors["primary"],
            text_style=ft.TextStyle(color=colors["on_surface"]),
            **kwargs
        )

        super().__init__(
            controls=[self.label, self.text_field],
            spacing=Spacing.CONTROL_PADDING_X,
            alignment=ft.MainAxisAlignment.START,
        )

    @property
    def value(self) -> str:
        """Get the current value"""
        return self.text_field.value

    @value.setter
    def value(self, v: str) -> None:
        """Set the value"""
        self.text_field.value = v
        self.text_field.update()


class ActionButton(ft.ElevatedButton):
    """
    A styled action button with consistent appearance.
    """

    def __init__(
        self,
        text: str,
        on_click: Optional[Callable] = None,
        icon: Optional[str] = None,
        variant: str = "primary",
        theme_mode: ThemeMode = ThemeMode.LIGHT,
        width: Optional[int] = None,
        **kwargs
    ):
        """
        Initialize an action button.

        Args:
            text: Button text
            on_click: Click handler
            icon: Flet icon name
            variant: Button variant - "primary", "secondary", "success", "danger"
            theme_mode: Theme mode for styling
            width: Button width
            **kwargs: Additional Button properties
        """
        colors = AppTheme.get_colors(theme_mode)

        # Variant styles
        variant_styles = {
            "primary": {
                "bgcolor": colors["primary"],
                "color": colors["on_primary"],
            },
            "secondary": {
                "bgcolor": colors["secondary_container"],
                "color": colors["on_surface"],
            },
            "success": {
                "bgcolor": colors["success"],
                "color": colors["on_primary"],
            },
            "danger": {
                "bgcolor": colors["error"],
                "color": colors["on_primary"],
            },
        }

        style = variant_styles.get(variant, variant_styles["primary"])

        super().__init__(
            text=text,
            icon=icon,
            on_click=on_click,
            width=width,
            style=ft.ButtonStyle(
                bgcolor=style["bgcolor"],
                color=style["color"],
                padding=ft.padding.symmetric(
                    horizontal=Spacing.BUTTON_PADDING_X,
                    vertical=Spacing.BUTTON_PADDING_Y,
                ),
            ),
            **kwargs
        )


# Helper function to create a file picker that can be reused
def create_file_picker(
    page: ft.Page,
    on_selected: Callable[[str], None],
    file_type: Optional[str] = None,
    initial_path: str = Text.FILE_NOT_SELECTED,
) -> tuple[FilePickerCard, ft.FilePicker]:
    """
    Helper function to create a file picker and add it to the page overlay.

    Args:
        page: Flet page to add the file picker to
        on_selected: Callback when file is selected
        file_type: File type filter
        initial_path: Initial path display text

    Returns:
        Tuple of (FilePickerCard, ft.FilePicker)
    """
    picker = FilePickerCard(
        label="选择文件",
        on_pick=on_selected,
        file_type=file_type,
        initial_path=initial_path,
    )

    # Add the hidden file picker to the page overlay
    page.overlay.append(picker.file_picker)

    return picker, picker.file_picker
