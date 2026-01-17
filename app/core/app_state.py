"""
Application State Management Module - Centralized state for the Flet-based Excel Toolkit

Manages application-wide state including theme mode, window state, and shared data.
"""

import threading
from typing import Any, Callable, Dict, Optional, Set
from dataclasses import dataclass, field
from enum import Enum

from app.core.theme import ThemeMode


class StatusIcon(str, Enum):
    """Status icon enumeration"""
    READY = "✅"
    PROCESSING = "⏳"
    SUCCESS = "✅"
    ERROR = "❌"
    WARNING = "⚠️"
    INFO = "ℹ️"


@dataclass
class StatusState:
    """Status bar state"""
    message: str = "就绪"
    icon: str = StatusIcon.READY
    show_progress: bool = False
    progress_value: float = 0.0


@dataclass
class WindowState:
    """Window state"""
    always_on_top: bool = True
    width: int = 1080
    height: int = 780
    left: Optional[int] = None
    top: Optional[int] = None


class AppState:
    """
    Central application state manager.

    Thread-safe state management with change notifications.
    """

    _instance: Optional['AppState'] = None
    _lock = threading.Lock()

    def __new__(cls) -> 'AppState':
        """Singleton pattern implementation"""
        if cls._instance is None:
            with cls._lock:
                if cls._instance is None:
                    cls._instance = super().__new__(cls)
        return cls._instance

    def __init__(self):
        """Initialize application state"""
        # Avoid re-initialization
        if hasattr(self, '_initialized'):
            return

        self._initialized = True
        self._lock = threading.Lock()
        self._listeners: Dict[str, Set[Callable]] = {}

        # Theme state
        self._theme_mode: ThemeMode = ThemeMode.SYSTEM

        # Window state
        self._window = WindowState()

        # Status state
        self._status = StatusState()

        # Current tab index
        self._current_tab: int = 0

        # File selection tracking
        self._selected_files: Dict[str, str] = {}

        # Sheet names cache for Excel files
        self._sheet_cache: Dict[str, list] = {}

    # ==================== Event System ====================

    def add_listener(self, event: str, callback: Callable) -> None:
        """
        Add a listener for an event.

        Args:
            event: Event name
            callback: Callback function to call when event is triggered
        """
        with self._lock:
            if event not in self._listeners:
                self._listeners[event] = set()
            self._listeners[event].add(callback)

    def remove_listener(self, event: str, callback: Callable) -> None:
        """
        Remove a listener for an event.

        Args:
            event: Event name
            callback: Callback function to remove
        """
        with self._lock:
            if event in self._listeners:
                self._listeners[event].discard(callback)

    def _notify(self, event: str, *args, **kwargs) -> None:
        """
        Notify all listeners of an event.

        Args:
            event: Event name
            *args: Positional arguments to pass to callbacks
            **kwargs: Keyword arguments to pass to callbacks
        """
        with self._lock:
            callbacks = self._listeners.get(event, set()).copy()

        for callback in callbacks:
            try:
                callback(*args, **kwargs)
            except Exception as e:
                print(f"Error in event listener for {event}: {e}")

    # ==================== Theme State ====================

    @property
    def theme_mode(self) -> ThemeMode:
        """Get current theme mode"""
        return self._theme_mode

    @theme_mode.setter
    def theme_mode(self, value: ThemeMode) -> None:
        """
        Set theme mode and notify listeners.

        Args:
            value: New theme mode
        """
        if self._theme_mode != value:
            self._theme_mode = value
            self._notify("theme_changed", value)

    def toggle_theme(self) -> None:
        """Toggle between light and dark theme"""
        if self._theme_mode == ThemeMode.LIGHT:
            self.theme_mode = ThemeMode.DARK
        elif self._theme_mode == ThemeMode.DARK:
            self.theme_mode = ThemeMode.LIGHT
        else:
            # Default to light if currently system
            self.theme_mode = ThemeMode.LIGHT

    # ==================== Window State ====================

    @property
    def window(self) -> WindowState:
        """Get window state"""
        return self._window

    @window.setter
    def window(self, value: WindowState) -> None:
        """
        Set window state and notify listeners.

        Args:
            value: New window state
        """
        self._window = value
        self._notify("window_changed", value)

    @property
    def always_on_top(self) -> bool:
        """Get always-on-top state"""
        return self._window.always_on_top

    @always_on_top.setter
    def always_on_top(self, value: bool) -> None:
        """
        Set always-on-top state and notify listeners.

        Args:
            value: New always-on-top state
        """
        if self._window.always_on_top != value:
            self._window.always_on_top = value
            self._notify("topmost_changed", value)

    # ==================== Status State ====================

    @property
    def status(self) -> StatusState:
        """Get current status state"""
        return self._status

    def set_status(self, message: str, icon: str = StatusIcon.READY,
                   show_progress: bool = False, progress_value: float = 0.0) -> None:
        """
        Update status bar state.

        Args:
            message: Status message text
            icon: Status icon
            show_progress: Whether to show progress bar
            progress_value: Progress bar value (0-100)
        """
        self._status = StatusState(
            message=message,
            icon=icon,
            show_progress=show_progress,
            progress_value=progress_value
        )
        self._notify("status_changed", self._status)

    def set_ready(self, message: str = "就绪") -> None:
        """Set status to ready"""
        self.set_status(message, StatusIcon.READY, False, 0.0)

    def set_processing(self, message: str = "处理中...") -> None:
        """Set status to processing"""
        self.set_status(message, StatusIcon.PROCESSING, True, 0.0)

    def set_success(self, message: str = "完成") -> None:
        """Set status to success"""
        self.set_status(message, StatusIcon.SUCCESS, False, 100.0)

    def set_error(self, message: str = "错误") -> None:
        """Set status to error"""
        self.set_status(message, StatusIcon.ERROR, False, 0.0)

    def set_warning(self, message: str) -> None:
        """Set status to warning"""
        self.set_status(message, StatusIcon.WARNING, False, 0.0)

    def update_progress(self, value: float, message: Optional[str] = None) -> None:
        """
        Update progress bar value.

        Args:
            value: Progress value (0-100)
            message: Optional new status message
        """
        if message is not None:
            self._status.message = message
        self._status.progress_value = value
        self._notify("status_changed", self._status)

    # ==================== Tab State ====================

    @property
    def current_tab(self) -> int:
        """Get current tab index"""
        return self._current_tab

    @current_tab.setter
    def current_tab(self, value: int) -> None:
        """
        Set current tab index and notify listeners.

        Args:
            value: New tab index (0-13)
        """
        if 0 <= value <= 13 and self._current_tab != value:
            self._current_tab = value
            self._notify("tab_changed", value)

    # ==================== File Selection State ====================

    def set_file(self, key: str, path: str) -> None:
        """
        Store a selected file path.

        Args:
            key: Identifier for the file (e.g., "file1", "pdf_input")
            path: File path
        """
        self._selected_files[key] = path
        self._notify("file_selected", key, path)

    def get_file(self, key: str, default: str = "") -> str:
        """
        Get a stored file path.

        Args:
            key: Identifier for the file
            default: Default value if not found

        Returns:
            File path or default value
        """
        return self._selected_files.get(key, default)

    def clear_file(self, key: str) -> None:
        """
        Clear a stored file path.

        Args:
            key: Identifier for the file
        """
        if key in self._selected_files:
            del self._selected_files[key]
            self._notify("file_cleared", key)

    # ==================== Sheet Cache ====================

    def cache_sheets(self, file_path: str, sheets: list) -> None:
        """
        Cache sheet names for an Excel file.

        Args:
            file_path: Path to the Excel file
            sheets: List of sheet names
        """
        self._sheet_cache[file_path] = sheets

    def get_cached_sheets(self, file_path: str) -> Optional[list]:
        """
        Get cached sheet names for an Excel file.

        Args:
            file_path: Path to the Excel file

        Returns:
            List of sheet names or None if not cached
        """
        return self._sheet_cache.get(file_path)

    def clear_sheet_cache(self, file_path: Optional[str] = None) -> None:
        """
        Clear sheet cache.

        Args:
            file_path: Specific file to clear, or None to clear all
        """
        if file_path:
            self._sheet_cache.pop(file_path, None)
        else:
            self._sheet_cache.clear()

    # ==================== Utility Methods ====================

    def reset(self) -> None:
        """Reset all state to defaults (except theme)"""
        with self._lock:
            theme = self._theme_mode
            self.__init__()
            self._theme_mode = theme
            self._notify("state_reset")


# Global state instance
_state: Optional[AppState] = None


def get_state() -> AppState:
    """
    Get the global application state instance.

    Returns:
        AppState instance
    """
    global _state
    if _state is None:
        _state = AppState()
    return _state


def reset_state() -> None:
    """Reset the global application state"""
    global _state
    _state = None
