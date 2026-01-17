#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel工具箱 优化打包脚本
专门用于生成启动速度优化的exe文件
"""
import os
import shutil
import subprocess
import sys

print("=" * 60)
print("Excel工具箱 优化打包程序")
print("=" * 60)

# 1. 检查 PyInstaller
print("\n[1/4] 检查 PyInstaller...")
try:
    import PyInstaller
    print(f"✅ PyInstaller {PyInstaller.__version__} 已安装")
except ImportError:
    print("❌ PyInstaller 未安装，正在安装...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
    print("✅ PyInstaller 安装完成")

# 2. 清理旧文件
print("\n[2/4] 清理旧文件...")
for folder in ["build", "dist"]:
    if os.path.exists(folder):
        shutil.rmtree(folder)
        print(f"✅ 删除 {folder} 目录")

# 3. 创建优化的spec文件
print("\n[3/4] 创建优化配置...")

spec_content = '''# -*- mode: python ; coding: utf-8 -*-
# Excel工具箱 优化打包配置

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('excel_toolkit/state_coords.json', 'excel_toolkit'),
    ],
    hiddenimports=[
        # 核心模块
        'tkinter',
        'tkinter.ttk',
        'tkinter.scrolledtext',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'xlrd',
        'xlsxwriter',
        'defusedxml',
        'lxml',
        'PIL.Image',
        'PIL._tkinter_finder',
        # Excel工具箱模块
        'excel_toolkit',
        'excel_toolkit.app',
        'excel_toolkit.ui',
        'excel_toolkit.ui.mixins',
        'excel_toolkit.ui.tab13_image_compress', # Explicitly include new tab
        'excel_toolkit.states',
        'excel_toolkit.sku_fill',
        'excel_toolkit.highlight',
        'excel_toolkit.insert_rows',
        'excel_toolkit.compare',
        'excel_toolkit.pdf_ocr',
        'excel_toolkit.prefix_fill',
        'excel_toolkit.warehouse_router',
        'excel_toolkit.shipping_fill',
        'excel_toolkit.db_config',
        'excel_toolkit.db_models',
        'excel_toolkit.db_operations',
        'excel_toolkit.tooltip',
        'excel_toolkit.template_maker',
        'excel_toolkit.delete_cols',
        'excel_toolkit.ui.tab14_delete_cols',
        # 数据库相关
        'sqlalchemy',
        'sqlalchemy.engine',
        'sqlalchemy.orm',
        'pydantic',
        # OCR相关（轻量级）
        'pytesseract',
        'pdf2image',
        'pypdf',
        'rapidocr_onnxruntime',
        'onnxruntime',
        # 支持旧版Excel
        'xlrd',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # 完全排除numpy相关模块以避免兼容性问题
        'numpy',
        'numpy.core',
        'numpy.core.multiarray',
        'numpy.random',
        'numpy.linalg',
        'numpy.fft',
        'numpy.polynomial',
        'numpy.random._pickle',
        'numpy.random._bounded_integers',
        'numpy.distutils',
        'numpy.f2py',
        'numpy.testing',
        # 排除大型模块以减少启动时间
        'matplotlib',
        'scipy',
        'pandas',
        'tensorflow',
        'torch',
        'torchvision',
        'cv2',
        'easyocr',
        'paddleocr',
        'paddle',
        'unittest',
        'pytest',
        'IPython',
        'jupyter',
    ],
    noarchive=False,
    optimize=2,  # 最高优化级别
)

pyz = PYZ(a.pure, a.zipped_data, cipher=None)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='Excel工具箱-优化版',
    debug=False,
    bootloader_ignore_signals=False,
    strip=True,  # 移除符号表
    upx=True,    # 启用压缩
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico' if os.path.exists('icon.ico') else None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=True,
    upx=True,
    upx_exclude=[],
    name='Excel工具箱-优化版',
)
'''

with open("Excel工具箱-优化版.spec", "w", encoding="utf-8") as f:
    f.write(spec_content)
print("✅ 优化配置文件已创建")

# 4. 执行打包
print("\n[4/4] 开始优化打包...")
print("-" * 60)
print("⏳ 正在打包，请稍候...")
print("-" * 60)

try:
    cmd = [
        sys.executable, 
        "-m", 
        "PyInstaller",
        "--clean",
        "--noconfirm",
        "Excel工具箱-优化版.spec"
    ]
    
    result = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8')
    
    if result.returncode == 0:
        print("\n" + "=" * 60)
        print("✅ 优化打包成功！")
        print("=" * 60)
        
        exe_path = os.path.join("dist", "Excel工具箱-优化版", "Excel工具箱-优化版.exe")
        
        if os.path.exists(exe_path):
            size_mb = os.path.getsize(exe_path) / (1024 * 1024)
            print(f"\n📦 优化版可执行文件:")
            print(f"   {os.path.abspath(exe_path)}")
            print(f"   大小: {size_mb:.1f} MB")
            
            dist_dir = os.path.join("dist", "Excel工具箱-优化版")
            total_size = sum(
                os.path.getsize(os.path.join(root, f))
                for root, _, files in os.walk(dist_dir)
                for f in files
            ) / (1024 * 1024)
            print(f"\n📁 完整程序目录:")
            print(f"   {os.path.abspath(dist_dir)}")
            print(f"   总大小: {total_size:.1f} MB")
            
            print("\n🚀 优化特性:")
            print("   ✅ 延迟加载Tab页面")
            print("   ✅ 异步模块导入")
            print("   ✅ 字节码优化")
            print("   ✅ 排除不必要模块")
            print("   ✅ 启动画面优化")
            
            print("\n📋 使用方法:")
            print("   1. 将整个 dist/Excel工具箱-优化版 文件夹复制到任何地方")
            print("   2. 双击 Excel工具箱-优化版.exe 运行")
            print("   3. 首次启动应该明显更快！")
        else:
            print("\n⚠️  exe文件未找到，请检查打包日志")
    else:
        print("\n❌ 打包失败")
        print("错误输出:")
        print(result.stderr)
    
except Exception as e:
    print(f"\n❌ 发生错误: {e}")

print("\n按回车键退出...")
input()