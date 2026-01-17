"""
Configuration Persistence Module - Save and load application configuration

Handles persistent storage of application settings, file paths, and user preferences.
"""

import json
import os
from pathlib import Path
from typing import Any, Dict, Optional
from dataclasses import dataclass, asdict

from app.core.theme import ThemeMode


@dataclass
class AppConfig:
    """Application configuration data class"""

    # Version info
    version: str = "2.3"
    config_version: int = 1

    # Window settings
    window_width: int = 1080
    window_height: int = 780
    window_left: Optional[int] = None
    window_top: Optional[int] = None
    always_on_top: bool = True

    # Theme setting
    theme_mode: str = ThemeMode.SYSTEM

    # File paths (persisted file selections)
    file_paths: Dict[str, str] = None

    # User preferences
    preferences: Dict[str, Any] = None

    def __post_init__(self):
        """Initialize mutable default values"""
        if self.file_paths is None:
            self.file_paths = {}
        if self.preferences is None:
            self.preferences = {}


class ConfigManager:
    """
    Configuration manager for persisting application settings.

    Handles loading and saving configuration to a JSON file in the user's
    home directory, ensuring settings persist across application restarts.
    """

    def __init__(self, app_name: str = "excel_toolkit_flet"):
        """
        Initialize configuration manager.

        Args:
            app_name: Application name for config directory naming
        """
        self._app_name = app_name
        self._config_dir = self._get_config_dir()
        self._config_file = os.path.join(self._config_dir, "config.json")
        self._config: AppConfig = AppConfig()

        # Ensure config directory exists
        os.makedirs(self._config_dir, exist_ok=True)

    def _get_config_dir(self) -> str:
        """
        Get the configuration directory path.

        Uses user's home directory to create a hidden config folder.
        Falls back to current directory if home directory is not writable.

        Returns:
            Path to configuration directory
        """
        # Try user's home directory first
        user_home = os.path.expanduser("~")
        config_dir = os.path.join(user_home, f".{self._app_name}")

        try:
            # Test if directory is writable
            os.makedirs(config_dir, exist_ok=True)
            test_file = os.path.join(config_dir, ".write_test")
            with open(test_file, "w") as f:
                f.write("test")
            os.remove(test_file)
            return config_dir
        except (OSError, IOError):
            # Fall back to current directory
            return os.path.join(os.getcwd(), f".{self._app_name}")

    def load(self) -> AppConfig:
        """
        Load configuration from file.

        Returns:
            Loaded AppConfig object, or default config if file doesn't exist
        """
        if not os.path.exists(self._config_file):
            self._config = AppConfig()
            return self._config

        try:
            with open(self._config_file, "r", encoding="utf-8") as f:
                data = json.load(f)

            # Convert loaded data to AppConfig
            self._config = AppConfig(
                version=data.get("version", "2.3"),
                config_version=data.get("config_version", 1),
                window_width=data.get("window_width", 1080),
                window_height=data.get("window_height", 780),
                window_left=data.get("window_left"),
                window_top=data.get("window_top"),
                always_on_top=data.get("always_on_top", True),
                theme_mode=data.get("theme_mode", ThemeMode.SYSTEM),
                file_paths=data.get("file_paths", {}),
                preferences=data.get("preferences", {}),
            )

            # Validate theme mode
            if self._config.theme_mode not in (ThemeMode.LIGHT, ThemeMode.DARK, ThemeMode.SYSTEM):
                self._config.theme_mode = ThemeMode.SYSTEM

        except (json.JSONDecodeError, IOError, OSError) as e:
            print(f"Warning: Failed to load config file, using defaults: {e}")
            self._config = AppConfig()

        return self._config

    def save(self) -> bool:
        """
        Save current configuration to file.

        Returns:
            True if save was successful, False otherwise
        """
        try:
            data = {
                "version": self._config.version,
                "config_version": self._config.config_version,
                "window_width": self._config.window_width,
                "window_height": self._config.window_height,
                "window_left": self._config.window_left,
                "window_top": self._config.window_top,
                "always_on_top": self._config.always_on_top,
                "theme_mode": self._config.theme_mode,
                "file_paths": self._config.file_paths,
                "preferences": self._config.preferences,
            }

            with open(self._config_file, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)

            return True

        except (IOError, OSError) as e:
            print(f"Error: Failed to save config file: {e}")
            return False

    @property
    def config(self) -> AppConfig:
        """Get current configuration"""
        return self._config

    # ==================== Window Settings ====================

    def get_window_size(self) -> tuple[int, int]:
        """Get window size as (width, height) tuple"""
        return (self._config.window_width, self._config.window_height)

    def set_window_size(self, width: int, height: int) -> None:
        """Set window size"""
        self._config.window_width = width
        self._config.window_height = height

    def get_window_position(self) -> Optional[tuple[int, int]]:
        """Get window position as (left, top) tuple, or None if not set"""
        if self._config.window_left is not None and self._config.window_top is not None:
            return (self._config.window_left, self._config.window_top)
        return None

    def set_window_position(self, left: int, top: int) -> None:
        """Set window position"""
        self._config.window_left = left
        self._config.window_top = top

    def get_always_on_top(self) -> bool:
        """Get always-on-top setting"""
        return self._config.always_on_top

    def set_always_on_top(self, value: bool) -> None:
        """Set always-on-top setting"""
        self._config.always_on_top = value

    # ==================== Theme Settings ====================

    def get_theme_mode(self) -> ThemeMode:
        """Get theme mode"""
        return ThemeMode(self._config.theme_mode)

    def set_theme_mode(self, mode: ThemeMode) -> None:
        """Set theme mode"""
        self._config.theme_mode = mode

    # ==================== File Path Persistence ====================

    def set_file_path(self, key: str, path: str) -> None:
        """
        Store a file path in configuration.

        Args:
            key: Identifier for the file path
            path: File path to store
        """
        self._config.file_paths[key] = path

    def get_file_path(self, key: str, default: str = "") -> str:
        """
        Get a stored file path.

        Args:
            key: Identifier for the file path
            default: Default value if not found

        Returns:
            Stored file path or default value
        """
        return self._config.file_paths.get(key, default)

    def clear_file_path(self, key: str) -> None:
        """Remove a stored file path"""
        self._config.file_paths.pop(key, None)

    def get_all_file_paths(self) -> Dict[str, str]:
        """Get all stored file paths"""
        return self._config.file_paths.copy()

    # ==================== Preferences ====================

    def set_preference(self, key: str, value: Any) -> None:
        """
        Store a preference value.

        Args:
            key: Preference key
            value: Preference value (must be JSON-serializable)
        """
        self._config.preferences[key] = value

    def get_preference(self, key: str, default: Any = None) -> Any:
        """
        Get a preference value.

        Args:
            key: Preference key
            default: Default value if not found

        Returns:
            Stored preference value or default
        """
        return self._config.preferences.get(key, default)

    def clear_preference(self, key: str) -> None:
        """Remove a preference"""
        self._config.preferences.pop(key, None)

    # ==================== Utility Methods ====================

    def reset_to_defaults(self) -> None:
        """Reset configuration to default values"""
        self._config = AppConfig()

    def export_config(self, path: str) -> bool:
        """
        Export configuration to a specific file.

        Args:
            path: Destination file path

        Returns:
            True if export was successful
        """
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(asdict(self._config), f, ensure_ascii=False, indent=2)
            return True
        except (IOError, OSError) as e:
            print(f"Error: Failed to export config: {e}")
            return False

    def import_config(self, path: str) -> bool:
        """
        Import configuration from a file.

        Args:
            path: Source file path

        Returns:
            True if import was successful
        """
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)

            # Merge with current config
            for key, value in data.items():
                if hasattr(self._config, key):
                    setattr(self._config, key, value)

            return True
        except (json.JSONDecodeError, IOError, OSError) as e:
            print(f"Error: Failed to import config: {e}")
            return False


# Global config manager instance
_config_manager: Optional[ConfigManager] = None


def get_config_manager() -> ConfigManager:
    """
    Get the global configuration manager instance.

    Returns:
        ConfigManager instance
    """
    global _config_manager
    if _config_manager is None:
        _config_manager = ConfigManager()
    return _config_manager


def load_config() -> AppConfig:
    """
    Load configuration from file.

    Returns:
        Loaded AppConfig object
    """
    return get_config_manager().load()


def save_config() -> bool:
    """
    Save current configuration to file.

    Returns:
        True if save was successful
    """
    return get_config_manager().save()
