"""
Base View Module - Base class for all tab views in the Flet-based Excel Toolkit

Provides common functionality for all tabs including logging, status updates,
and shared UI components.
"""

from abc import ABC, abstractmethod
from typing import Optional, Callable

import flet as ft

from app.core.theme import AppTheme, ThemeMode
from app.core.constants import Text, FontSize, Spacing, FontFamily
from app.core.app_state import get_state, AppState
from app.core.config import get_config_manager
from app.components.common import FilePickerCard, DirectoryPickerCard, LogWidget, Card, ActionButton


class BaseView(ABC):
    """
    Base class for all tab views.

    Provides common functionality including:
    - Theme-aware styling
    - Logging with timestamps
    - Status updates
    - File picking integration
    - Configuration persistence
    """

    def __init__(self, page: ft.Page, tab_index: int = 0):
        """
        Initialize the base view.

        Args:
            page: Flet page control
            tab_index: Tab index (0-13)
        """
        self.page = page
        self.tab_index = tab_index
        self._state = get_state()
        self._config = get_config_manager()
        self._theme_mode = ThemeMode.LIGHT

    # ==================== Abstract Methods ====================

    @abstractmethod
    def get_tab_name(self) -> str:
        """
        Get the display name for this tab.

        Returns:
            Tab display name
        """
        pass

    @abstractmethod
    def build(self) -> ft.Control:
        """
        Build the tab content.

        Returns:
            Flet control for the tab content
        """
        pass

    # ==================== Theme Management ====================

    @property
    def theme_mode(self) -> ThemeMode:
        """Get current theme mode"""
        return self._theme_mode

    @theme_mode.setter
    def theme_mode(self, value: ThemeMode) -> None:
        """Set theme mode"""
        self._theme_mode = value

    def get_colors(self) -> dict:
        """Get current color scheme"""
        return AppTheme.get_colors(self._theme_mode)

    # ==================== Status Updates ====================

    def set_status(self, message: str, icon: str = "ℹ️", show_progress: bool = False) -> None:
        """
        Update the status bar.

        Args:
            message: Status message
            icon: Status icon
            show_progress: Whether to show progress bar
        """
        self._state.set_status(message, icon, show_progress)

    def set_ready(self, message: str = "就绪") -> None:
        """Set status to ready"""
        self._state.set_ready(message)

    def set_processing(self, message: str = "处理中...") -> None:
        """Set status to processing"""
        self._state.set_processing(message)

    def set_success(self, message: str = "完成") -> None:
        """Set status to success"""
        self._state.set_success(message)

    def set_error(self, message: str = "错误") -> None:
        """Set status to error"""
        self._state.set_error(message)

    def set_warning(self, message: str) -> None:
        """Set status to warning"""
        self._state.set_warning(message)

    def update_progress(self, value: float, message: Optional[str] = None) -> None:
        """
        Update progress bar value.

        Args:
            value: Progress value (0-100)
            message: Optional new status message
        """
        self._state.update_progress(value, message)

    # ==================== Logging ====================

    def create_logger(self, log_widget: Optional[LogWidget] = None) -> Callable[[str, str], None]:
        """
        Create a logger function that writes to a log widget.

        Args:
            log_widget: Optional existing log widget

        Returns:
            Logger function that accepts (message, level) parameters
        """
        def log(message: str, level: str = "info") -> None:
            """Log a message"""
            if log_widget:
                log_widget.log(message, level)
            else:
                print(f"[{level.upper()}] {message}")

        return log

    # ==================== Configuration Persistence ====================

    def save_file_path(self, key: str, path: str) -> None:
        """
        Save a file path to configuration.

        Args:
            key: Configuration key
            path: File path to save
        """
        self._config.set_file_path(f"{self.get_tab_name()}_{key}", path)
        self._state.set_file(f"{self.get_tab_name()}_{key}", path)

    def load_file_path(self, key: str, default: str = "") -> str:
        """
        Load a file path from configuration.

        Args:
            key: Configuration key
            default: Default value if not found

        Returns:
            Saved file path or default value
        """
        return self._config.get_file_path(f"{self.get_tab_name()}_{key}", default)

    def save_preference(self, key: str, value: any) -> None:
        """
        Save a preference to configuration.

        Args:
            key: Preference key
            value: Preference value (must be JSON-serializable)
        """
        self._config.set_preference(f"{self.get_tab_name()}_{key}", value)

    def load_preference(self, key: str, default: any = None) -> any:
        """
        Load a preference from configuration.

        Args:
            key: Preference key
            default: Default value if not found

        Returns:
            Saved preference value or default value
        """
        return self._config.get_preference(f"{self.get_tab_name()}_{key}", default)

    # ==================== UI Building Helpers ====================

    def create_section_card(
        self,
        title: str,
        content: ft.Control,
        icon: Optional[str] = None,
    ) -> Card:
        """
        Create a section card with consistent styling.

        Args:
            title: Card title
            content: Card content
            icon: Optional title icon

        Returns:
            Card component
        """
        return Card(
            content=content,
            title=title,
            title_icon=icon,
            theme_mode=self._theme_mode,
        )

    def create_file_picker(
        self,
        label: str,
        config_key: str,
        on_pick: Optional[Callable[[str], None]] = None,
        file_type: Optional[str] = None,
    ) -> FilePickerCard:
        """
        Create a file picker with configuration persistence.

        Args:
            label: Label for the file picker
            config_key: Key for saving/loading the path
            on_pick: Optional callback when file is selected
            file_type: Optional file type filter

        Returns:
            FilePickerCard component
        """
        # Load saved path
        saved_path = self.load_file_path(config_key)
        initial_path = saved_path if saved_path else Text.FILE_NOT_SELECTED

        # Create callback that saves to config
        def handle_pick(path: str):
            self.save_file_path(config_key, path)
            if on_pick:
                on_pick(path)

        picker = FilePickerCard(
            label=label,
            on_pick=handle_pick,
            file_type=file_type,
            initial_path=initial_path,
            theme_mode=self._theme_mode,
            page=self.page,  # Pass page for FilePicker to work
        )

        return picker

    def create_directory_picker(
        self,
        label: str,
        config_key: str,
        on_pick: Optional[Callable[[str], None]] = None,
    ) -> DirectoryPickerCard:
        """
        Create a directory picker with configuration persistence.

        Args:
            label: Label for the directory picker
            config_key: Key for saving/loading the path
            on_pick: Optional callback when directory is selected

        Returns:
            DirectoryPickerCard component
        """
        # Load saved path
        saved_path = self.load_file_path(config_key)
        initial_path = saved_path if saved_path else Text.DIR_NOT_SELECTED

        # Create callback that saves to config
        def handle_pick(path: str):
            self.save_file_path(config_key, path)
            if on_pick:
                on_pick(path)

        picker = DirectoryPickerCard(
            label=label,
            on_pick=handle_pick,
            initial_path=initial_path,
            theme_mode=self._theme_mode,
            page=self.page,  # Pass page for FilePicker to work
        )

        return picker

    def create_action_button(
        self,
        text: str = Text.BTN_RUN,
        on_click: Optional[Callable] = None,
        icon: Optional[str] = None,
        variant: str = "primary",
    ) -> ActionButton:
        """
        Create an action button with consistent styling.

        Args:
            text: Button text
            on_click: Click handler
            icon: Optional icon
            variant: Button variant (primary, secondary, success, danger)

        Returns:
            ActionButton component
        """
        return ActionButton(
            text=text,
            on_click=on_click,
            icon=icon,
            variant=variant,
            theme_mode=self._theme_mode,
        )

    # ==================== Error Handling ====================

    def show_error(self, title: str, message: str) -> None:
        """
        Show an error dialog.

        Args:
            title: Dialog title
            message: Error message
        """
        colors = self.get_colors()

        dialog = ft.AlertDialog(
            modal=True,
            title=ft.Text(title, color=colors["error"]),
            content=ft.Text(message),
            actions=[
                ft.TextButton("确定", on_click=lambda e: self._close_dialog(dialog)),
            ],
        )

        self.page.dialog = dialog
        dialog.open = True
        self.page.update()

    def show_info(self, title: str, message: str) -> None:
        """
        Show an info dialog.

        Args:
            title: Dialog title
            message: Info message
        """
        colors = self.get_colors()

        dialog = ft.AlertDialog(
            modal=True,
            title=ft.Text(title, color=colors["info"]),
            content=ft.Text(message),
            actions=[
                ft.TextButton("确定", on_click=lambda e: self._close_dialog(dialog)),
            ],
        )

        self.page.dialog = dialog
        dialog.open = True
        self.page.update()

    def show_success(self, message: str) -> None:
        """
        Show a success dialog.

        Args:
            message: Success message
        """
        colors = self.get_colors()

        dialog = ft.AlertDialog(
            modal=True,
            title=ft.Row([
                ft.Icon(ft.Icons.CHECK_CIRCLE, color=colors["success"]),
                ft.Text("操作成功", color=colors["success"]),
            ]),
            content=ft.Text(message),
            actions=[
                ft.TextButton("确定", on_click=lambda e: self._close_dialog(dialog)),
            ],
        )

        self.page.dialog = dialog
        dialog.open = True
        self.page.update()

    def _close_dialog(self, dialog: ft.AlertDialog) -> None:
        """Close a dialog"""
        dialog.open = False
        self.page.update()

    # ==================== Validation ====================

    def validate_file_selected(self, path: str, file_description: str = "文件") -> bool:
        """
        Validate that a file has been selected.

        Args:
            path: File path to validate
            file_description: Description of the file for error messages

        Returns:
            True if file is selected, False otherwise
        """
        if not path or path == Text.FILE_NOT_SELECTED:
            self.show_error("验证错误", f"请选择{file_description}！")
            return False
        return True

    def validate_column_input(self, col: str) -> bool:
        """
        Validate that a column input is a valid Excel column letter.

        Args:
            col: Column letter to validate

        Returns:
            True if valid, False otherwise
        """
        if not col or not col.strip():
            self.show_error("验证错误", "请输入列号！")
            return False
        return True


class TabView(BaseView):
    """
    Base class for tab views with standard layout.

    Provides a standard 3-section layout:
    1. File selection section
    2. Parameters/Options section
    3. Log output section
    """

    def __init__(self, page: ft.Page, tab_index: int = 0):
        """Initialize the tab view"""
        super().__init__(page, tab_index)
        self._log_widget: Optional[LogWidget] = None
        self._logger: Optional[Callable] = None

    def build_standard_layout(
        self,
        file_section: ft.Control,
        param_section: ft.Control,
        action_section: ft.Control,
    ) -> ft.Column:
        """
        Build standard tab layout with file, params, action, and log sections.

        Args:
            file_section: File selection section
            param_section: Parameters/options section
            action_section: Action buttons section

        Returns:
            Column with standard layout
        """
        colors = self.get_colors()

        # Create log widget
        self._log_widget = LogWidget(
            height=200,
            theme_mode=self._theme_mode,
        )
        self._logger = self.create_logger(self._log_widget)

        # Main layout
        return ft.Column([
            # File selection section
            self.create_section_card(
                title="文件选择",
                content=file_section,
                icon=ft.Icons.FOLDER_OPEN,
            ),
            # Parameters section
            self.create_section_card(
                title="参数配置",
                content=param_section,
                icon=ft.Icons.SETTINGS,
            ),
            # Action buttons
            action_section,
            # Log section
            self._log_widget,
        ], spacing=Spacing.SECTION_SPACING, expand=True)

    @property
    def logger(self) -> Callable[[str, str], None]:
        """Get the logger function"""
        if self._logger is None:
            self._logger = self.create_logger()
        return self._logger

    def log(self, message: str, level: str = "info") -> None:
        """
        Log a message to the tab's log widget.

        Args:
            message: Message to log
            level: Log level (info, success, warning, error)
        """
        self.logger(message, level)
