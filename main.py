import os
import tkinter as tk


os.environ.setdefault("DISABLE_MODEL_SOURCE_CHECK", "True")


def show_splash(root):
    """显示启动画面（使用Toplevel避免双Tk冲突）"""
    splash = tk.Toplevel(root)
    splash.title("Excel 工具箱")
    splash.overrideredirect(True)  # 无边框窗口
    
    # 窗口大小和居中
    width, height = 380, 160
    screen_w = splash.winfo_screenwidth()
    screen_h = splash.winfo_screenheight()
    x = (screen_w - width) // 2
    y = (screen_h - height) // 2
    splash.geometry(f"{width}x{height}+{x}+{y}")
    
    # 设置背景色
    splash.configure(bg="#2563eb")
    
    # 主容器
    frame = tk.Frame(splash, bg="#2563eb")
    frame.pack(expand=True, fill="both", padx=20, pady=20)
    
    # 图标和标题
    title_label = tk.Label(
        frame, 
        text="📊 Excel 工具箱", 
        font=("Segoe UI", 20, "bold"),
        fg="white",
        bg="#2563eb"
    )
    title_label.pack(pady=(10, 5))
    
    # 版本信息
    version_label = tk.Label(
        frame,
        text="V2.3",
        font=("Segoe UI", 10),
        fg="#93c5fd",
        bg="#2563eb"
    )
    version_label.pack()
    
    # 加载提示
    status_label = tk.Label(
        frame,
        text="正在加载模块，请稍候...",
        font=("Segoe UI", 10),
        fg="#bfdbfe",
        bg="#2563eb"
    )
    status_label.pack(pady=(15, 5))
    
    # 进度条样式
    progress_frame = tk.Frame(frame, bg="#1e40af", height=6)
    progress_frame.pack(fill="x", pady=(5, 0))
    progress_frame.pack_propagate(False)
    
    progress_bar = tk.Frame(progress_frame, bg="#60a5fa", width=0)
    progress_bar.pack(side="left", fill="y")
    
    # 动画函数
    def animate_progress(current_width=0):
        if splash.winfo_exists() and current_width < width - 40:
            progress_bar.configure(width=current_width + 8)
            splash.after(50, lambda: animate_progress(current_width + 8))
    
    splash.after(100, animate_progress)
    splash.update()
    
    return splash, status_label


def main():
    # 先创建主窗口（隐藏）
    root = tk.Tk()
    root.withdraw()
    
    # 显示启动画面（作为Toplevel）
    splash, status_label = show_splash(root)
    
    # 异步加载模块，避免阻塞启动画面
    def load_modules_async():
        try:
            # 分阶段加载，每个阶段更新状态
            status_label.configure(text="正在加载基础模块...")
            splash.update()
            
            # 导入基础模块
            import excel_toolkit
            
            status_label.configure(text="正在加载UI组件...")
            splash.update()
            
            # 导入主应用（这里会加载所有依赖）
            from excel_toolkit.app import ToolkitAppRefactored as ToolkitApp
            
            status_label.configure(text="初始化界面...")
            splash.update()
            
            # 初始化应用
            app = ToolkitApp(root)
            
            status_label.configure(text="启动完成...")
            splash.update()
            
            # 短暂延迟后关闭启动画面
            def finish_startup():
                splash.destroy()
                root.deiconify()  # 显示主窗口
            
            root.after(500, finish_startup)
            
        except Exception as e:
            try:
                splash.destroy()
            except:
                pass
            import traceback
            with open("error.log", "w") as f:
                traceback.print_exc(file=f)
            print("!!! 程序发生严重错误 !!!")
            traceback.print_exc()
            input("按回车键退出...")
    
    # 使用after方法异步执行加载，避免阻塞启动画面
    root.after(100, load_modules_async)
    
    # 启动主循环
    root.mainloop()


if __name__ == "__main__":
    main()