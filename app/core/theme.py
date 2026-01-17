"""
Theme Configuration Module - Theme system for the Flet-based Excel Toolkit

Defines light and dark theme colors with Windows 11/macOS flat design style.
"""

from enum import Enum
from typing import Dict
import flet as ft


class ThemeMode(str, Enum):
    """Theme mode enumeration"""
    LIGHT = "light"
    DARK = "dark"
    SYSTEM = "system"


class AppTheme:
    """
    Application theme configuration

    Provides color schemes for light and dark themes with Windows 11/macOS
    flat design style.
    """

    # Light theme colors - Windows 11 inspired
    LIGHT: Dict[str, str] = {
        # Primary colors
        "primary": "#0078D4",           # Windows 11 blue
        "primary_hover": "#106EBE",
        "primary_container": "#DEECF9",
        "on_primary": "#FFFFFF",

        # Background colors
        "background": "#F3F3F3",        # Light gray background
        "surface": "#FFFFFF",           # White card surface
        "surface_variant": "#F9F9F9",
        "on_background": "#1A1A1A",     # Dark text
        "on_surface": "#1A1A1A",        # Dark text on surface
        "on_surface_variant": "#5C5C5C",

        # Border colors
        "border": "#E0E0E0",
        "outline": "#747775",
        "outline_variant": "#CAC4D0",

        # Secondary colors
        "secondary": "#6B7280",
        "secondary_container": "#E4E7E9",
        "on_secondary": "#FFFFFF",

        # Status colors
        "success": "#10B981",           # Green
        "success_container": "#D1FAE5",
        "warning": "#F59E0B",           # Orange
        "warning_container": "#FEF3C7",
        "error": "#EF4444",             # Red
        "error_container": "#FEE2E2",
        "info": "#3B82F6",              # Blue
        "info_container": "#DBEAFE",

        # Text colors
        "text_primary": "#111827",
        "text_secondary": "#6B7280",
        "text_hint": "#9CA3AF",
        "text_disabled": "#D1D5DB",

        # Elevations (shadows)
        "elevation_1": "0 1px 3px rgba(0,0,0,0.12), 0 1px 2px rgba(0,0,0,0.24)",
        "elevation_2": "0 3px 6px rgba(0,0,0,0.16), 0 3px 6px rgba(0,0,0,0.23)",
        "elevation_3": "0 10px 20px rgba(0,0,0,0.19), 0 6px 6px rgba(0,0,0,0.23)",
    }

    # Dark theme colors
    DARK: Dict[str, str] = {
        # Primary colors
        "primary": "#60CDFF",           # Bright cyan blue for dark mode
        "primary_hover": "#4FB8EB",
        "primary_container": "#004C6C",
        "on_primary": "#001F29",

        # Background colors
        "background": "#202020",        # Dark gray background
        "surface": "#2B2B2B",           # Dark card surface
        "surface_variant": "#252525",
        "on_background": "#FFFFFF",     # Light text
        "on_surface": "#FFFFFF",        # Light text on surface
        "on_surface_variant": "#C4C7C5",

        # Border colors
        "border": "#404040",
        "outline": "#9CA3AF",
        "outline_variant": "#49454F",

        # Secondary colors
        "secondary": "#A8A8A8",
        "secondary_container": "#434748",
        "on_secondary": "#FFFFFF",

        # Status colors
        "success": "#34D399",           # Brighter green for dark mode
        "success_container": "#064E3B",
        "warning": "#FBBF24",           # Brighter orange
        "warning_container": "#78350F",
        "error": "#F87171",             # Brighter red
        "error_container": "#7F1D1D",
        "info": "#60A5FA",              # Brighter blue
        "info_container": "#1E3A8A",

        # Text colors
        "text_primary": "#F9FAFB",
        "text_secondary": "#D1D5DB",
        "text_hint": "#9CA3AF",
        "text_disabled": "#6B7280",

        # Elevations (shadows) - less visible in dark mode
        "elevation_1": "0 1px 3px rgba(0,0,0,0.3), 0 1px 2px rgba(0,0,0,0.4)",
        "elevation_2": "0 3px 6px rgba(0,0,0,0.4), 0 3px 6px rgba(0,0,0,0.5)",
        "elevation_3": "0 10px 20px rgba(0,0,0,0.5), 0 6px 6px rgba(0,0,0,0.6)",
    }

    @classmethod
    def get_colors(cls, mode: ThemeMode = ThemeMode.LIGHT) -> Dict[str, str]:
        """
        Get color scheme for the specified theme mode.

        Args:
            mode: Theme mode (light, dark, or system)

        Returns:
            Dictionary of color values
        """
        if mode == ThemeMode.DARK:
            return cls.DARK
        return cls.LIGHT

    @classmethod
    def get_flet_theme(cls, mode: ThemeMode = ThemeMode.LIGHT) -> ft.Theme:
        """
        Get Flet Theme object for the specified mode.

        Args:
            mode: Theme mode (light, dark, or system)

        Returns:
            Flet Theme object
        """
        colors = cls.get_colors(mode)

        return ft.Theme(
            color_scheme=ft.ColorScheme(
                primary=colors["primary"],
                on_primary=colors["on_primary"],
                primary_container=colors["primary_container"],
                on_primary_container=colors["on_background"] if mode == ThemeMode.DARK else colors["text_primary"],
                secondary=colors["secondary"],
                on_secondary=colors["on_secondary"],
                secondary_container=colors["secondary_container"],
                on_secondary_container=colors["on_background"] if mode == ThemeMode.DARK else colors["text_primary"],
                background=colors["background"],
                on_background=colors["on_background"],
                surface=colors["surface"],
                on_surface=colors["on_surface"],
                surface_variant=colors["surface_variant"],
                on_surface_variant=colors["on_surface_variant"],
                outline=colors["outline"],
                outline_variant=colors["outline_variant"],
                error=colors["error"],
                on_error=colors["on_primary"],
                error_container=colors["error_container"],
                on_error_container=colors["on_background"] if mode == ThemeMode.DARK else colors["text_primary"],
            ),
            visual_density=ft.VisualDensity.COMFORTABLE,
            use_material3=True,
        )

    @classmethod
    def get_card_style(cls, mode: ThemeMode = ThemeMode.LIGHT) -> Dict[str, any]:
        """
        Get common card container style.

        Args:
            mode: Theme mode

        Returns:
            Dictionary of style properties for ft.Container
        """
        colors = cls.get_colors(mode)

        return {
            "bgcolor": colors["surface"],
            "border_radius": 8,
            "padding": 12,
            "border": ft.border.all(1, colors["border"]),
        }

    @classmethod
    def get_button_style(cls, mode: ThemeMode = ThemeMode.LIGHT, variant: str = "filled") -> Dict[str, any]:
        """
        Get button style.

        Args:
            mode: Theme mode
            variant: Button variant - "filled", "outlined", or "text"

        Returns:
            Dictionary of style properties
        """
        colors = cls.get_colors(mode)

        if variant == "filled":
            return {
                "bgcolor": colors["primary"],
                "color": colors["on_primary"],
            }
        elif variant == "outlined":
            return {
                "style": ft.ButtonStyle(
                    side=ft.border.BorderSide(1, colors["outline"]),
                ),
            }
        else:  # text
            return {
                "style": ft.ButtonStyle(),
            }

    @classmethod
    def get_input_style(cls, mode: ThemeMode = ThemeMode.LIGHT) -> Dict[str, any]:
        """
        Get input field (TextField, Dropdown) style.

        Args:
            mode: Theme mode

        Returns:
            Dictionary of style properties
        """
        colors = cls.get_colors(mode)

        return {
            "bgcolor": colors["surface"],
            "border_color": colors["border"],
            "focused_border_color": colors["primary"],
            "text_style": ft.TextStyle(color=colors["on_surface"]),
            "label_style": ft.TextStyle(color=colors["text_secondary"]),
            "hint_style": ft.TextStyle(color=colors["text_hint"]),
        }

    @classmethod
    def get_text_style(cls, mode: ThemeMode = ThemeMode.LIGHT, variant: str = "body") -> ft.TextStyle:
        """
        Get text style.

        Args:
            mode: Theme mode
            variant: Text variant - "title", "subtitle", "body", "caption", "hint"

        Returns:
            Flet TextStyle object
        """
        colors = cls.get_colors(mode)

        styles = {
            "title": ft.TextStyle(
                size=16,
                weight=ft.FontWeight.BOLD,
                color=colors["text_primary"],
            ),
            "subtitle": ft.TextStyle(
                size=12,
                weight=ft.FontWeight.BOLD,
                color=colors["text_primary"],
            ),
            "section": ft.TextStyle(
                size=11,
                weight=ft.FontWeight.BOLD,
                color=colors["text_primary"],
            ),
            "body": ft.TextStyle(
                size=10,
                color=colors["text_primary"],
            ),
            "caption": ft.TextStyle(
                size=9,
                color=colors["text_secondary"],
            ),
            "hint": ft.TextStyle(
                size=9,
                color=colors["text_hint"],
            ),
        }

        return styles.get(variant, styles["body"])

    @classmethod
    def get_log_style(cls, mode: ThemeMode = ThemeMode.LIGHT) -> Dict[str, any]:
        """
        Get log display area style.

        Args:
            mode: Theme mode

        Returns:
            Dictionary of style properties
        """
        colors = cls.get_colors(mode)

        return {
            "bgcolor": colors["surface_variant"],
            "color": colors["on_surface"],
            "text_style": ft.TextStyle(
                font_family="Consolas",
                size=9,
                color=colors["on_surface_variant"],
            ),
        }
