"""
Simple test to verify Flet application startup
"""
import flet as ft

def main(page: ft.Page):
    page.title = "Excel 工具箱 V2.3 - 测试"
    page.window_width = 1080
    page.window_height = 780
    page.padding = 0
    page.spacing = 0
    page.theme_mode = ft.ThemeMode.LIGHT
    page.bgcolor = ft.colors.BLUE_50
    page.vertical_alignment = ft.MainAxisAlignment.CENTER
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER

    # Create a simple test UI
    page.add(
        ft.Column([
            ft.Icon(ft.icons.TABLE_CHART, size=64, color=ft.colors.BLUE),
            ft.Text(
                "📊 Excel 工具箱 V2.3",
                size=32,
                weight=ft.FontWeight.BOLD,
                color=ft.colors.WHITE
            ),
            ft.Text(
                "Flet 框架测试 - 如果看到此窗口，说明 Flet 工作正常",
                size=14,
                color=ft.colors.WHITE
            ),
            ft.ElevatedButton(
                "关闭",
                on_click=lambda _: page.window.close()
            ),
        ],
            spacing=20,
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
        )
    )

if __name__ == "__main__":
    ft.app(target=main)
