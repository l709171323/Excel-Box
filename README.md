# Excel 工具箱

> 一款功能强大的 Excel 批量处理工具，专为电商和仓储管理设计。

[![Python Version](https://img.shields.io/badge/python-3.9+-blue.svg)](https://www.python.org/downloads/)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
[![GUI](https://img.shields.io/badge/GUI-Tkinter-orange.svg)](https://docs.python.org/3/library/tkinter.html)

---

## 功能概览

本工具提供 **14 个实用功能模块**，覆盖 Excel 数据处理的常见需求：

| 功能 | 描述 |
|------|------|
| 1️⃣ 州名转换 | 美国州全名与两字母缩写互转 |
| 2️⃣ SKU 填充 | 根据数据库自动匹配并填充 SKU 信息 |
| 3️⃣ 高亮重复 | 标记指定列中的重复数据行（黄色高亮） |
| 4️⃣ 插入行 | 按 SKU 变化自动插入分隔行 |
| 5️⃣ 对比列 | 比较两列数据差异，支持忽略重复项 |
| 6️⃣ PDF 拆分 | 拆分 PDF 文件并 OCR 识别订单号 |
| 7️⃣ 前缀填充 | 批量为指定列添加前缀 |
| 8️⃣ 面单页脚 | 为 PDF 面单添加页脚信息 |
| 9️⃣ 仓库路由 | 基于距离和库存推荐最优发货仓库 |
| 🔟 库存录入 | 管理和编辑仓库库存信息 |
| 1️⃣1️⃣ 模板填充 | 根据配置自动填充发货模板 |
| 1️⃣2️⃣ PPT 转 PDF | 批量转换 PowerPoint 文件为 PDF |
| 1️⃣3️⃣ 图片压缩 | 批量压缩图片文件 |
| 1️⃣4️⃣ 删除列 | 批量删除 Excel 中的指定列 |

---

## 快速开始

### 环境要求

- Python 3.9 或更高版本
- Windows 操作系统

### 安装步骤

```bash
# 1. 克隆仓库
git clone https://github.com/l709171323/Excel-Box.git
cd Excel-Box

# 2. 安装依赖
pip install -r requirements.txt

# 3. 运行程序
python main.py
```

### 打包为 EXE

```bash
# 一键打包（默认配置）
python build_exe.py

# 精简版（不含 PaddleOCR，减小约 500MB）
python build_exe.py --no-paddle
```

---

## 项目结构

```
Python-O/
├── main.py                 # 程序入口
├── build_exe.py            # 一键打包脚本
├── icon.ico                # 程序图标
├── requirements.txt        # Python 依赖
├── excel_toolkit/          # 核心模块
│   ├── app.py             # 主界面
│   ├── states.py          # 功能1：州名转换
│   ├── sku_fill.py        # 功能2：SKU填充
│   ├── highlight.py       # 功能3：高亮重复
│   ├── insert_rows.py     # 功能4：插入行
│   ├── compare.py         # 功能5：对比列
│   ├── pdf_ocr.py         # 功能6：PDF OCR
│   ├── prefix_fill.py     # 功能7：前缀填充
│   ├── warehouse_router.py # 功能9：仓库路由
│   ├── shipping_fill.py   # 功能11：发货填充
│   └── ui/                # UI 组件
└── vendor/                # 第三方依赖
    ├── poppler/           # PDF 处理
    └── tesseract/         # OCR 引擎
```

---

## 使用指南

### 基本操作流程

1. 选择功能标签页
2. 点击"浏览"选择输入文件
3. 配置处理参数（列号、工作表等）
4. 点击"开始处理"
5. 查看日志输出

### 注意事项

> ⚠️ **处理前请关闭所有 Excel 文件，避免保存失败**

---

## 核心功能详解

### 功能 2：SKU 填充

根据 SKU 数据库自动匹配并填充产品信息。

**支持的数据源**：
- Excel 数据库文件（SKU编号、产品名称、规格等）

### 功能 6：PDF 拆分 + OCR

将多页 PDF 拆分为单页文件，并自动 OCR 识别订单号重命名。

**依赖组件**：
- Tesseract OCR（已内置）
- Poppler（已内置）

### 功能 9：仓库路由

基于地理距离和 SKU 库存，智能推荐最优发货仓库。

**计算逻辑**：
1. 筛选有库存的仓库
2. 计算收件地址到各仓库的距离
3. 推荐最近且有货的仓库

---

## 技术栈

| 组件 | 用途 |
|------|------|
| Python 3.9+ | 核心语言 |
| Tkinter | 图形界面 |
| OpenPyXL | Excel 文件处理 |
| PyPDF2/pdf2image | PDF 处理 |
| Tesseract OCR | 文字识别 |
| PyInstaller | EXE 打包 |

---

## 常见问题

<details>
<summary><b>Q: 程序启动慢？</b></summary>

首次启动需要加载 OCR 引擎，约 10-30 秒。这是正常现象。
</details>

<details>
<summary><b>Q: Excel 保存失败？</b></summary>

请确保关闭所有正在使用的 Excel 文件，避免文件被占用。
</details>

<details>
<summary><b>Q: 如何更换程序图标？</b></summary>

准备 `.ico` 格式图标文件，重命名为 `icon.ico`，放到项目根目录后重新打包。
</details>

<details>
<summary><b>Q: 打包文件太大？</b></summary>

使用 `python build_exe.py --no-paddle` 可排除 PaddleOCR，减小约 500MB 体积。
</details>

---

## 贡献

欢迎提交 Issue 和 Pull Request！

1. Fork 本仓库
2. 创建功能分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 开启 Pull Request

---

## 许可证

本项目采用 MIT 许可证 - 详见 [LICENSE](LICENSE) 文件

---

## 联系方式

- GitHub: [@l709171323](https://github.com/l709171323)
- 仓库地址: [https://github.com/l709171323/Excel-Box](https://github.com/l709171323/Excel-Box)

---

⭐ 如果这个项目对你有帮助，请给个 Star！