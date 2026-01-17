import os
import re
import sys
import io
import threading
from fractions import Fraction
from typing import Tuple, Optional, Callable, Dict

# 防止 PaddleOCR 多线程冲突导致崩溃
os.environ["KMP_DUPLICATE_LIB_OK"] = "TRUE"
os.environ["FLAGS_allocator_strategy"] = 'auto_growth'

# ===== PaddleOCR 内存优化配置 =====
# 限制CPU线程数，防止过度占用
os.environ["OMP_NUM_THREADS"] = "2"
os.environ["MKL_NUM_THREADS"] = "2"
# 限制 Paddle 内部线程
os.environ["CPU_NUM"] = "2"
# 禁用 MKL-DNN 加速(减少内存，但会稍慢)
os.environ["FLAGS_use_mkldnn"] = "0"
# 内存增长策略：按需分配而非预分配
os.environ["FLAGS_initial_cpu_memory_in_mb"] = "100"
os.environ["FLAGS_fraction_of_cpu_memory_to_use"] = "0.3"

# PaddleOCR 线程锁
_PADDLE_LOCK = threading.Lock()

try:
    from pypdf import PdfReader, PdfWriter
    from pdf2image import convert_from_path
    import pytesseract
    from PIL import Image, ImageOps, ImageChops, ImageStat, ImageEnhance, ImageFilter
except Exception as e:
    # Defer import errors until runtime with a clearer message
    PdfReader = PdfWriter = convert_from_path = pytesseract = Image = ImageOps = ImageChops = ImageStat = None  # type: ignore
    _IMPORT_ERROR = e
else:
    _IMPORT_ERROR = None

# PaddleOCR 支持 (兼容 2.x 和 3.x)
_PADDLE_AVAILABLE = False
_PADDLE_VERSION = None  # 'v2' or 'v3'
_PADDLE_IMPORT_ERROR = None
PaddleOCR = None
np = None

try:
    import numpy as np
    # 尝试导入 PaddleOCR 3.x
    try:
        from paddleocr import PaddleOCR
        # 检测版本：3.x 有 predict 方法
        import paddleocr
        _ver = getattr(paddleocr, '__version__', '2.0.0')
        if int(_ver.split('.')[0]) >= 3:
            _PADDLE_VERSION = 'v3'
        else:
            _PADDLE_VERSION = 'v2'
        _PADDLE_AVAILABLE = True
    except Exception as e:
        _PADDLE_IMPORT_ERROR = e
except Exception as e:
    _PADDLE_IMPORT_ERROR = e

_PADDLE_INSTANCE = None
_PADDLE_LOCK = None

# ===== RapidOCR 支持（轻量级 ONNX 推理）=====
_RAPID_AVAILABLE = False
_RAPID_IMPORT_ERROR = None
RapidOCR = None

try:
    from rapidocr_onnxruntime import RapidOCR
    _RAPID_AVAILABLE = True
except Exception as e:
    _RAPID_IMPORT_ERROR = e

_RAPID_INSTANCE = None
_RAPID_LOCK = None

def get_rapid_instance():
    """获取 RapidOCR 实例（懒加载 + 线程安全）"""
    global _RAPID_INSTANCE, _RAPID_LOCK
    
    if not _RAPID_AVAILABLE:
        return None
    
    # 快速检查：如果已初始化，直接返回
    if _RAPID_INSTANCE is not None:
        return _RAPID_INSTANCE
    
    # 初始化线程锁
    if _RAPID_LOCK is None:
        import threading
        _RAPID_LOCK = threading.Lock()
    
    # 线程安全的初始化
    with _RAPID_LOCK:
        # 双重检查：防止多线程同时初始化
        if _RAPID_INSTANCE is None:
            print("⚡ 正在初始化 RapidOCR [轻量ONNX模式]...")
            try:
                _RAPID_INSTANCE = RapidOCR()
                print("✓ RapidOCR 初始化成功！")
            except Exception as e:
                print(f"❌ RapidOCR 初始化失败: {e}")
                return None
    
    return _RAPID_INSTANCE


def release_rapid_instance():
    """释放 RapidOCR 实例以回收内存"""
    global _RAPID_INSTANCE
    if _RAPID_INSTANCE is not None:
        print("正在释放 RapidOCR 实例...")
        del _RAPID_INSTANCE
        _RAPID_INSTANCE = None
        import gc
        gc.collect()
        print("✓ RapidOCR 实例已释放")


def get_paddle_instance():
    """获取 PaddleOCR 实例（兼容 2.x 和 3.x）- 内存优化版 + 线程安全"""
    global _PADDLE_INSTANCE, _PADDLE_LOCK
    
    if not _PADDLE_AVAILABLE:
        return None
    
    # 快速检查：如果已初始化，直接返回
    if _PADDLE_INSTANCE is not None:
        return _PADDLE_INSTANCE
    
    # 初始化线程锁
    if _PADDLE_LOCK is None:
        import threading
        _PADDLE_LOCK = threading.Lock()
    
    # 线程安全的初始化
    with _PADDLE_LOCK:
        # 双重检查
        if _PADDLE_INSTANCE is None:
            print(f"⚡ 正在初始化 PaddleOCR ({_PADDLE_VERSION}) [轻量模式]...")
            try:
                import logging
                logging.getLogger('ppocr').setLevel(logging.ERROR)
                logging.getLogger('paddle').setLevel(logging.ERROR)
                
                if _PADDLE_VERSION == 'v3':
                    # PaddleOCR 3.x API - 轻量优化配置
                    _PADDLE_INSTANCE = PaddleOCR(
                        use_doc_orientation_classify=False,  # 禁用文档方向分类
                        use_doc_unwarping=False,             # 禁用文档矫正
                        use_textline_orientation=False,      # 禁用文本行方向
                        lang='en',
                        # 3.x 内存优化
                        text_det_limit_side_len=960,         # 限制检测尺寸(默认960)
                        text_det_limit_type='max',
                    )
                else:
                    # PaddleOCR 2.x API - 轻量优化配置
                    _PADDLE_INSTANCE = PaddleOCR(
                        use_angle_cls=False,      # 禁用角度分类器（减少~200MB）
                        lang='en',
                        show_log=False,
                        use_gpu=False,            # 强制CPU模式
                        enable_mkldnn=False,      # 禁用MKL-DNN（减少内存但稍慢）
                        cpu_threads=2,            # 限制CPU线程数
                        det_limit_side_len=960,   # 限制检测尺寸
                        det_limit_type='max',
                        rec_batch_num=1,          # 批处理大小=1（减少内存）
                        # 使用轻量模型 (如果可用)
                        # det_model_dir='ch_PP-OCRv3_det_slim',
                        # rec_model_dir='en_PP-OCRv3_rec_slim',
                    )
                print(f"✓ PaddleOCR {_PADDLE_VERSION} 初始化成功！[轻量模式]")
            except Exception as e:
                print(f"❌ PaddleOCR 初始化失败: {e}")
                import traceback
                traceback.print_exc()
                return None
    
    return _PADDLE_INSTANCE


def release_paddle_instance():
    """释放 PaddleOCR 实例以回收内存"""
    global _PADDLE_INSTANCE
    if _PADDLE_INSTANCE is not None:
        print("正在释放 PaddleOCR 实例...")
        del _PADDLE_INSTANCE
        _PADDLE_INSTANCE = None
        # 强制垃圾回收
        import gc
        gc.collect()
        print("✓ PaddleOCR 实例已释放")




def _resource_path(rel: str) -> str:
    """
    Resolve bundled data path under PyInstaller (sys._MEIPASS) or package directory.
    """
    base = getattr(sys, "_MEIPASS", os.path.dirname(__file__))
    return os.path.join(base, rel)


def find_poppler(user_path: Optional[str] = None) -> Optional[str]:
    if user_path:
        return user_path
    candidates = [
        # Windows 预编译版本的标准路径
        _resource_path(os.path.join("vendor", "poppler", "Library", "bin")),
        os.path.join(os.getcwd(), "vendor", "poppler", "Library", "bin"),
        # 备用路径
        _resource_path(os.path.join("vendor", "poppler", "bin")),
        os.path.join(os.getcwd(), "vendor", "poppler", "bin"),
        r"C:\\poppler-23.11.0\\bin",
        r"C:\\Program Files\\poppler\\bin",
        r"C:\\Program Files (x86)\\poppler\\bin",
    ]
    for p in candidates:
        if p and os.path.isdir(p):
            return p
    return None


def find_tesseract(user_path: Optional[str] = None) -> Optional[str]:
    if user_path:
        return user_path
    candidates = [
        _resource_path(os.path.join("vendor", "tesseract", "tesseract.exe")),
        os.path.join(os.getcwd(), "vendor", "tesseract", "tesseract.exe"),
        r"C:\\Program Files\\Tesseract-OCR\\tesseract.exe",
        r"C:\\Program Files (x86)\\Tesseract-OCR\\tesseract.exe",
    ]
    for p in candidates:
        if p and os.path.isfile(p):
            return p
    return None


def _set_tessdata_prefix(tess_cmd: Optional[str] = None) -> None:
    """Set TESSDATA_PREFIX to directory that contains the 'tessdata' folder (parent).
    Newer docs suggest prefix should be parent and actual data in prefix/tessdata.
    We'll set parent and also pass --tessdata-dir explicitly to avoid ambiguity.
    """
    exe_dir = os.path.dirname(tess_cmd) if tess_cmd else None
    candidates = [c for c in [
        exe_dir,
        os.path.join(exe_dir, "share") if exe_dir else None,
        _resource_path(os.path.join("vendor", "tesseract")),
        os.path.join(os.getcwd(), "vendor", "tesseract"),
    ] if c]

    for base in candidates:
        td = os.path.join(base, "tessdata")
        if os.path.isdir(td):
            os.environ["TESSDATA_PREFIX"] = base
            return


def _get_tessdata_dir() -> Optional[str]:
    prefix = os.environ.get("TESSDATA_PREFIX")
    if not prefix:
        return None
    # If env points directly to tessdata
    if os.path.basename(prefix).lower() == "tessdata" and os.path.isdir(prefix):
        return prefix
    candidate = os.path.join(prefix, "tessdata")
    if os.path.isdir(candidate):
        return candidate
    # Fallback: if prefix exists, return as-is
    return prefix if os.path.isdir(prefix) else None


def _ensure_tesseract_cmd(user_path: Optional[str]) -> Optional[str]:
    """Resolve tesseract.exe and set pytesseract command accordingly.
    Returns resolved path or None if not found.
    """
    resolved = find_tesseract(user_path)
    if resolved:
        pytesseract.pytesseract.tesseract_cmd = resolved
    # Always try to set TESSDATA_PREFIX even if command is None
    _set_tessdata_prefix(resolved)
    return resolved


def parse_bbox(bbox_str: str) -> Tuple[int, int, int, int]:
    x, y, w, h = [int(v.strip()) for v in bbox_str.split(",")]
    if w <= 0 or h <= 0:
        raise ValueError("width/height must be > 0")
    return x, y, w, h


def crop_region(img: "Image.Image", bbox: Tuple[int, int, int, int]) -> "Image.Image":
    x, y, w, h = bbox
    # 简单的边界保护
    if x < 0: x = 0
    if y < 0: y = 0
    if x + w > img.width: w = img.width - x
    if y + h > img.height: h = img.height - y
    if w <= 0 or h <= 0:
        # 返回一个极小的空白图防止报错
        return Image.new('RGB', (1, 1), (255, 255, 255))
    return img.crop((x, y, x + w, y + h))


def preprocess_image_for_ocr(img: "Image.Image", enhance_level: int = 2) -> "Image.Image":
    """
    预处理图像以提高 OCR 识别准确率
    
    Args:
        img: 原始图像
        enhance_level: 增强级别 (1=轻微, 2=中等, 3=强烈)
    
    Returns:
        处理后的图像
    """
    # 转换为灰度图
    img = img.convert('L')
    
    # 根据级别调整对比度
    contrast_factor = 1.5 + (enhance_level * 0.5)  # 2.0, 2.5, 3.0
    enhancer = ImageEnhance.Contrast(img)
    img = enhancer.enhance(contrast_factor)
    
    # 锐化图像
    if enhance_level >= 2:
        img = img.filter(ImageFilter.SHARPEN)
    
    # 二值化处理（阈值法）
    if enhance_level >= 2:
        threshold = 128
        img = img.point(lambda x: 255 if x > threshold else 0, mode='1')
    
    return img


def correct_ocr_confusion(text: str, context: str = 'mixed') -> str:
    """
    智能纠正 OCR 常见混淆字符
    
    Args:
        text: OCR 识别的文本
        context: 上下文类型 ('numeric'=纯数字, 'alpha'=纯字母, 'mixed'=混合)
    
    Returns:
        纠正后的文本
    """
    if not text:
        return text
    
    # 常见混淆字符映射
    corrections = {
        'O': '0',  # 字母O -> 数字0
        'o': '0',  # 小写o -> 数字0
        'I': '1',  # 字母I -> 数字1
        'l': '1',  # 小写l -> 数字1
        'Z': '2',  # 某些字体下 Z/2 混淆
        'B': '8',  # 某些情况 B/8 混淆
    }
    
    # 根据上下文智能纠正
    result = []
    for i, char in enumerate(text):
        prev_char = text[i-1] if i > 0 else ''
        next_char = text[i+1] if i < len(text)-1 else ''
        
        # 如果前后都是数字，则当前字符应该也是数字
        if prev_char.isdigit() and next_char.isdigit():
            if char in corrections:
                result.append(corrections[char])
            else:
                result.append(char)
        # 启发式规则：如果周围主要是数字
        elif char in corrections and context == 'mixed':
            surrounding = (prev_char + next_char)
            digit_count = sum(1 for c in surrounding if c.isdigit())
            if digit_count >= 1:
                result.append(corrections[char])
            else:
                result.append(char)
        else:
            result.append(char)
    
    return ''.join(result)


def render_page_to_image(pdf_path: str, page_index: int, dpi: int, poppler_path: Optional[str]) -> "Image.Image":
    images = convert_from_path(
        pdf_path,
        dpi=dpi,
        first_page=page_index + 1,
        last_page=page_index + 1,
        poppler_path=find_poppler(poppler_path),
    )
    if not images:
        raise RuntimeError(f"Failed to render page {page_index}")
    return images[0]


def ocr_with_umi(img: "Image.Image", host="127.0.0.1", port=1224) -> str:
    """
    调用 Umi-OCR 的 HTTP API 进行识别
    """
    import base64
    import io
    import json
    import urllib.request
    
    # 图片转Base64
    buffered = io.BytesIO()
    img.save(buffered, format="PNG")
    img_b64 = base64.b64encode(buffered.getvalue()).decode('utf-8')
    
    # 构造请求
    url = f"http://{host}:{port}/api/ocr"
    data = {
        "base64": img_b64,
        "options": {
            "data.format": "text", # 直接返回文本
        }
    }
    
    try:
        req = urllib.request.Request(url, json.dumps(data).encode('utf-8'), headers={'Content-Type': 'application/json'})
        with urllib.request.urlopen(req) as response:
            res_json = json.loads(response.read().decode('utf-8'))
            if res_json["code"] == 100:
                # 提取文本
                return "".join([item["text"] for item in res_json["data"]])
            else:
                return ""
    except Exception as e:
        # print(f"Umi-OCR Error: {e}")
        return ""


def ocr_with_paddle(img: "Image.Image") -> str:
    """
    使用 PaddleOCR 库进行识别（兼容 2.x 和 3.x）- 内存优化版
    """
    import gc
    
    ocr = get_paddle_instance()
    if ocr is None:
        print("❌ 无法获取 PaddleOCR 实例")
        return ""
    
    img_np = None
    result = None
    
    try:
        # 转为 RGB 格式
        if img.mode != 'RGB':
            img = img.convert('RGB')
        
        # 图像预处理：缩小尺寸以减少内存（面单识别不需要太高分辨率）
        max_side = 1200  # 限制最大边长
        w, h = img.size
        if max(w, h) > max_side:
            scale = max_side / max(w, h)
            new_w, new_h = int(w * scale), int(h * scale)
            img = img.resize((new_w, new_h), Image.LANCZOS)
        
        img_np = np.array(img)
        
        if _PADDLE_VERSION == 'v3':
            # PaddleOCR 3.x: 使用 predict() 方法
            with _PADDLE_LOCK:
                result = ocr.predict(img_np)
            
            if not result:
                return ""
            txts = []
            for res in result:
                if 'rec_texts' in res:
                    txts.extend(res['rec_texts'])
                elif hasattr(res, 'rec_texts'):
                    txts.extend(res.rec_texts)
            combined = "".join(txts)
        else:
            # PaddleOCR 2.x: 使用 ocr() 方法
            with _PADDLE_LOCK:
                result = ocr.ocr(img_np, cls=False)
                
            if not result or not result[0]:
                return ""
            txts = [line[1][0] for line in result[0]]
            combined = "".join(txts)
        
        return combined
        
    except Exception as e:
        print(f"❌ PaddleOCR 识别错误: {e}")
        import traceback
        traceback.print_exc()
        return ""
    finally:
        # 及时释放numpy数组内存
        if img_np is not None:
            del img_np
        if result is not None:
            del result
        # 每次识别后执行轻量GC
        gc.collect(generation=0)


def ocr_with_rapid(img: "Image.Image") -> str:
    """
    使用 RapidOCR 进行识别（轻量级 ONNX 推理，打包体积小）
    这是 PaddleOCR 模型的 ONNX 版本，效果相同但体积从 600MB 减少到 ~70MB
    """
    import gc
    
    ocr = get_rapid_instance()
    if ocr is None:
        print("❌ 无法获取 RapidOCR 实例")
        return ""
    
    img_np = None
    result = None
    
    try:
        # 转为 RGB 格式
        if img.mode != 'RGB':
            img = img.convert('RGB')
        
        # 图像预处理：缩小尺寸以减少内存
        max_side = 1200
        w, h = img.size
        if max(w, h) > max_side:
            scale = max_side / max(w, h)
            new_w, new_h = int(w * scale), int(h * scale)
            img = img.resize((new_w, new_h), Image.LANCZOS)
        
        img_np = np.array(img)
        
        # RapidOCR 调用
        result, _ = ocr(img_np)
        
        if not result:
            return ""
        
        # RapidOCR 返回格式: [[bbox, (text, confidence)], ...]
        txts = [item[1] for item in result]
        combined = "".join(txts)
        
        return combined
        
    except Exception as e:
        print(f"❌ RapidOCR 识别错误: {e}")
        import traceback
        traceback.print_exc()
        return ""
    finally:
        if img_np is not None:
            del img_np
        if result is not None:
            del result
        gc.collect(generation=0)


def ocr_order_number(
    img: "Image.Image",
    regex: Optional[str] = None,
    tesseract_cmd: Optional[str] = None,
    enable_preprocessing: bool = True,
    enable_correction: bool = True,
    engine: str = "tesseract",  # 新增参数: tesseract, umi, paddle, rapid
) -> Optional[str]:
    """
    OCR 订单号识别（增强版）
    
    Args:
        img: 图像对象
        regex: 自定义正则表达式
        tesseract_cmd: Tesseract 路径
        enable_preprocessing: 是否启用图像预处理（提高准确率）
        enable_correction: 是否启用字符纠错（修正 O/0 等混淆）
        engine: OCR引擎 ('tesseract', 'umi', 'paddle', 'rapid')
    
    Returns:
        识别到的订单号，失败返回 None
    """
    
    # 图像预处理（可选）
    # 注意：对于Umi-OCR和PaddleOCR，通常不需要太强的预处理
    if enable_preprocessing and engine == "tesseract":
        img = preprocess_image_for_ocr(img, enhance_level=2)
    
    text = ""
    
    if engine == "umi":
        text = ocr_with_umi(img)
    elif engine == "paddle":
        text = ocr_with_paddle(img)
    elif engine == "rapid":
        text = ocr_with_rapid(img)
    else:
        # Tesseract 逻辑
        resolved = _ensure_tesseract_cmd(tesseract_cmd)
        
        # 优化的 OCR 配置
        # --psm 7: 单行文本模式
        # --oem 3: 使用 LSTM OCR 引擎（更准确）
        base_config = "--psm 7 --oem 3 -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-#"
        td = _get_tessdata_dir()
        td_arg = td.replace("\\", "/") if td else None
        config = f"{base_config} --tessdata-dir {td_arg}" if td_arg else base_config
        
        try:
            text = pytesseract.image_to_string(img, config=config, lang="eng")
        except Exception:
            text = pytesseract.image_to_string(img, config=config)
    
    cleaned = re.sub(r"\s+", "", text)
    
    # 智能字符纠错（可选）
    if enable_correction and cleaned:
        # 检测是否是纯数字上下文（USPS运单号：9开头的纯数字）
        # 如果大部分是数字，使用numeric上下文以更严格地纠错
        digit_ratio = sum(1 for c in cleaned if c.isdigit()) / len(cleaned) if cleaned else 0
        context = 'numeric' if digit_ratio > 0.7 else 'mixed'
        cleaned = correct_ocr_confusion(cleaned, context=context)
    
    pattern = regex or r"[A-Za-z0-9#-]{6,32}"
    m = re.search(pattern, cleaned)
    return m.group(0) if m else None





def ocr_text_simple(
    img: "Image.Image",
    tesseract_cmd: Optional[str] = None,
) -> str:
    _ensure_tesseract_cmd(tesseract_cmd)
    base_config = "--psm 7 --oem 1 -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
    td = _get_tessdata_dir()
    td_arg = td.replace("\\", "/") if td else None
    config = f"{base_config} --tessdata-dir {td_arg}" if td_arg else base_config
    try:
        text = pytesseract.image_to_string(img, config=config, lang="eng")
    except Exception:
        text = pytesseract.image_to_string(img, config=config)
    return re.sub(r"\s+", " ", text).strip()


def sanitize_filename(name: str) -> str:
    return re.sub(r"[^A-Za-z0-9#\-_.]", "_", name)


def parse_shipping_label_spec(spec: str) -> Tuple[str, int, int]:
    s = str(spec).strip()
    m = re.match(r"^(.+?)-(\d+)\s*单\s*(\d+)\s*个\s*$", s) or re.match(r"^(.+?)-(\d+)单(\d+)个$", s)
    if not m:
        raise ValueError(f"面单格式不正确: {spec}")
    sku_short = m.group(1).strip()
    x = int(m.group(2))
    y = int(m.group(3))
    if x <= 0 or y <= 0:
        raise ValueError(f"面单数量必须为正整数: {spec}")
    return sku_short, x, y


def format_sku_footer_text(
    sku_full_name: str,
    x: int,
    y: int,
    hide_multiplier_if_one: bool = True,
) -> str:
    ratio = Fraction(y, x)
    if hide_multiplier_if_one and ratio == 1:
        return sku_full_name
    if ratio.denominator == 1:
        return f"{sku_full_name}*{ratio.numerator}"
    return f"{sku_full_name}*{ratio.numerator}/{ratio.denominator}"


def load_sku_full_name_map_from_excel(
    excel_path: str,
    sheet_name: Optional[str] = None,
    sku_short_col: Optional[str] = None,
    sku_full_col: Optional[str] = None,
) -> Dict[str, str]:
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"SKU映射Excel不存在: {excel_path}")

    try:
        from excel_toolkit.excel_lite import ExcelReader, ExcelWriter
        from excel_toolkit.excel_lite import column_index_from_string
    except Exception as e:
        raise RuntimeError(f"读取SKU映射Excel需要 openpyxl: {e}")

    def _resolve_col(headers, col_spec: Optional[str]) -> Optional[int]:
        if not col_spec:
            return None
        spec = str(col_spec).strip()
        if not spec:
            return None
        if spec.isalpha() and len(spec) <= 3:
            try:
                return column_index_from_string(spec.upper())
            except Exception:
                return None
        if spec in headers:
            return headers.index(spec) + 1
        return None

    wb = ExcelReader(excel_path, read_only=True, data_only=True)
    try:
        ws = wb[sheet_name] if sheet_name and sheet_name in wb.sheetnames else wb.worksheets[0]

        header_row = None
        for row in ws.iter_rows(min_row=1, max_row=1, values_only=True):
            header_row = list(row)
            break
        headers = [str(h).strip() if h is not None else "" for h in (header_row or [])]

        col_short = _resolve_col(headers, sku_short_col)
        col_full = _resolve_col(headers, sku_full_col)

        if col_short is None:
            for key in ["SKU简称", "简称", "Short", "short", "sku_short", "SKU_SHORT", "SKU"]:
                if key in headers:
                    col_short = headers.index(key) + 1
                    break

        if col_full is None:
            for key in ["SKU全称", "全称", "Full", "full", "sku_full", "SKU_FULL", "商品名称", "品名", "名称"]:
                if key in headers:
                    col_full = headers.index(key) + 1
                    break

        if col_short is None or col_full is None:
            raise ValueError(
                "SKU映射Excel缺少必要列，请指定 sku_short_col/sku_full_col 或确保表头包含 'SKU简称'/'SKU全称'。"
            )

        mapping: Dict[str, str] = {}
        for r in ws.iter_rows(min_row=2, values_only=True):
            try:
                short_val = r[col_short - 1] if col_short - 1 < len(r) else None
                full_val = r[col_full - 1] if col_full - 1 < len(r) else None
            except Exception:
                continue
            if short_val is None or full_val is None:
                continue
            s_raw = str(short_val).strip()
            f = str(full_val).strip()
            if not s_raw or not f:
                continue

            parts = [p.strip() for p in s_raw.split("||")] if "||" in s_raw else [s_raw]
            for s in parts:
                if not s:
                    continue
                mapping[s] = f

        return mapping
    finally:
        wb.close()


def add_pdf_footer_for_shipping_label(
    input_pdf: str,
    output_pdf: str,
    label_spec: str,
    sku_full_name_map: Optional[Dict[str, str]] = None,
    sku_mapping_excel: Optional[str] = None,
    sku_mapping_sheet_name: Optional[str] = None,
    sku_short_col: Optional[str] = None,
    sku_full_col: Optional[str] = None,
    sku_full_name: Optional[str] = None,
    hide_multiplier_if_one: bool = True,
    font_name: Optional[str] = None,
    font_size: int = 10,
    margin_bottom: float = 18,
    logger_func: Optional[Callable[[str], None]] = None,
) -> str:
    if _IMPORT_ERROR is not None:
        raise RuntimeError(
            "PDF 依赖未安装，请先安装: pip install pypdf\n"
            f"ImportError: {_IMPORT_ERROR}"
        )

    try:
        from reportlab.pdfgen import canvas
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    except Exception as e:
        raise RuntimeError(
            "写入PDF页脚需要 reportlab 依赖，请先安装: pip install reportlab\n"
            f"ImportError: {e}"
        )

    sku_short, x, y = parse_shipping_label_spec(label_spec)
    if logger_func:
        logger_func(f"解析面单规格: {label_spec} -> SKU简称={sku_short}, X={x}, Y={y}")

    if sku_full_name is None:
        if sku_full_name_map is None and sku_mapping_excel:
            sku_full_name_map = load_sku_full_name_map_from_excel(
                excel_path=sku_mapping_excel,
                sheet_name=sku_mapping_sheet_name,
                sku_short_col=sku_short_col,
                sku_full_col=sku_full_col,
            )
        sku_full_name = (sku_full_name_map or {}).get(sku_short) or sku_short

    footer_text = format_sku_footer_text(
        sku_full_name=sku_full_name,
        x=x,
        y=y,
        hide_multiplier_if_one=hide_multiplier_if_one,
    )

    reader = PdfReader(input_pdf)
    writer = PdfWriter()

    # 使用传入的字体名称
    actual_font_name = font_name or "STSong-Light"
    
    # 注册字体
    if actual_font_name in ["STSong-Light", "STHeiti-Regular"]:
        try:
            pdfmetrics.registerFont(UnicodeCIDFont(actual_font_name))
        except Exception:
            actual_font_name = "Helvetica"
    elif actual_font_name == "msyh":
        # 微软雅黑 - 从系统字体注册
        try:
            from reportlab.pdfbase.ttfonts import TTFont
            # Windows 系统字体路径
            font_paths = [
                "C:/Windows/Fonts/msyh.ttc",
                "C:/Windows/Fonts/msyh.ttf",
                os.path.expanduser("~/AppData/Local/Microsoft/Windows/Fonts/msyh.ttc"),
            ]
            font_registered = False
            for font_path in font_paths:
                if os.path.exists(font_path):
                    pdfmetrics.registerFont(TTFont("msyh", font_path))
                    font_registered = True
                    break
            if not font_registered:
                if logger_func:
                    logger_func("⚠️ 未找到微软雅黑字体，使用宋体替代")
                pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
                actual_font_name = "STSong-Light"
        except Exception as e:
            if logger_func:
                logger_func(f"⚠️ 注册微软雅黑失败: {e}，使用宋体替代")
            try:
                pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
                actual_font_name = "STSong-Light"
            except:
                actual_font_name = "Helvetica"
    # 其他字体直接使用（reportlab内置字体）

    for idx, page in enumerate(reader.pages):
        page_width = float(page.mediabox.width)
        page_height = float(page.mediabox.height)

        buf = io.BytesIO()
        c = canvas.Canvas(buf, pagesize=(page_width, page_height))
        c.setFont(actual_font_name, font_size)
        c.drawCentredString(page_width / 2.0, margin_bottom, footer_text)
        c.save()
        buf.seek(0)

        overlay_reader = PdfReader(buf)
        overlay_page = overlay_reader.pages[0]
        page.merge_page(overlay_page)
        writer.add_page(page)

        if logger_func:
            logger_func(f"已写入页脚: {idx+1}/{len(reader.pages)}")

    out_dir = os.path.dirname(output_pdf)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    with open(output_pdf, "wb") as f:
        writer.write(f)

    return f"完成！已输出PDF: {output_pdf}"


def split_pdf_with_ocr(
    input_pdf: str,
    out_dir: str,
    bbox: Tuple[int, int, int, int],
    bbox2: Optional[Tuple[int, int, int, int]] = None,
    bbox3: Optional[Tuple[int, int, int, int]] = None,
    uniuni_mode: bool = False,
    three_region_mode: bool = False,
    dpi: int = 300,
    poppler_path: Optional[str] = None,
    tesseract_cmd: Optional[str] = None,
    regex: Optional[str] = None,
    prefix: str = "",
    logger_func: Optional[Callable[[str], None]] = None,
    ocr_engine: str = "tesseract",
    progress_callback: Optional[Callable[[int, int, str], None]] = None,
) -> None:
    """
    Split a merged orders PDF into pages, OCR a fixed region for order number,
    and save each page as a new PDF named by the detected order number.
    
    Modes:
    - Standard: OCR bbox only
    - UniUni mode: Check bbox for UUS, fallback to bbox2 if not found
    - Three-region mode: OCR all three regions (USPS/GOFO/Uni), auto-detect carrier
    
    Returns a summary message.
    """
    if _IMPORT_ERROR is not None:
        raise RuntimeError(
            "PDF/OCR 依赖未安装，请先安装: pip install pypdf pdf2image pytesseract pillow; "
            "并在 Windows 上安装 Poppler 与 Tesseract。详细见 README。\n"
            f"ImportError: {_IMPORT_ERROR}"
        )

    # Resolve external tools early and log diagnostic info
    poppler_resolved = find_poppler(poppler_path)
    tess_resolved = _ensure_tesseract_cmd(tesseract_cmd)
    tess_prefix = os.environ.get("TESSDATA_PREFIX")
    tessdata_dir = _get_tessdata_dir()
    if logger_func:
        logger_func(f"Poppler 路径: {poppler_resolved or '未找到，将依赖系统PATH'}")
        logger_func(f"Tesseract 路径: {tess_resolved or '未找到，将依赖系统PATH'}")
        logger_func(f"TESSDATA_PREFIX: {tess_prefix or '未设置'}")
        if tessdata_dir:
            exists_eng = os.path.isfile(os.path.join(tessdata_dir, "eng.traineddata"))
            logger_func(f"tessdata 目录: {tessdata_dir} (eng={exists_eng})")

    # 检查 OCR 引擎状态
    if ocr_engine == "paddle":
        if not _PADDLE_AVAILABLE:
            msg = f"⚠️ PaddleOCR 不可用，将回退到 Tesseract。\n  错误原因: {_PADDLE_IMPORT_ERROR}"
            if logger_func: logger_func(msg)
            ocr_engine = "tesseract"
        else:
            if logger_func: logger_func(f"✓ 使用 PaddleOCR ({_PADDLE_VERSION}) 引擎")
    elif ocr_engine == "rapid":
        if not _RAPID_AVAILABLE:
            msg = f"⚠️ RapidOCR 不可用，将回退到 Tesseract。\n  错误原因: {_RAPID_IMPORT_ERROR}"
            if logger_func: logger_func(msg)
            ocr_engine = "tesseract"
        else:
            if logger_func: logger_func("✓ 使用 RapidOCR 引擎 [轻量ONNX模式]")

    reader = PdfReader(input_pdf)
    total_pages = len(reader.pages)
    os.makedirs(out_dir, exist_ok=True)

    import concurrent.futures
    
    # Helper function for processing a single page
    def process_page(idx):
        try:
            # Re-open reader inside thread if needed, but pypdf objects are generally not thread-safe for writing
            # However, we are reading here. To be safe and avoid pickling issues with PdfReader objects,
            # we might need to open the file again or be careful.
            # Actually, PdfReader is lazy. Let's try passing the path and opening it locally or just rendering.
            # render_page_to_image uses pdf2image which takes the path, so that's thread-safe.
            
            img = render_page_to_image(input_pdf, idx, dpi, poppler_resolved)
            region1 = crop_region(img, bbox)
            use_region = 1
            order_no = None
            
            log_msgs = [] # Collect logs to print at once to avoid interleaving
            
            if three_region_mode and bbox2 and bbox3:
                # 三区域智能识别模式：OCR三个区域，根据首字符判断承运商
                log_msgs.append(f"第{idx+1}页 三区域模式: 开始识别...")
                
                # OCR三个区域
                region_usps = crop_region(img, bbox)   # bbox = USPS区域
                region_gofo = crop_region(img, bbox3)  # bbox3 = GOFO区域
                region_uni = crop_region(img, bbox2)   # bbox2 = Uni区域
                
                # 尝试识别每个区域
                candidates = []
                
                # 检查USPS (9开头，22位数字)
                usps_result = ocr_order_number(
                    region_usps, 
                    regex=r"9\d{21}", 
                    tesseract_cmd=tesseract_cmd,
                    enable_preprocessing=True,
                    engine=ocr_engine
                )
                if usps_result and usps_result[0] == '9':
                    candidates.append(("USPS", usps_result, 1))
                    log_msgs.append(f"  区域1(USPS): {usps_result}")
                
                # 检查GOFO (G开头，GFUS+纯数字)
                # 正则说明：GFUS后跟14位字符（数字或O/o，容错）
                # 尝试关闭预处理，直接识别原生图像，避免预处理引入噪点
                gofo_result = ocr_order_number(
                    region_gofo, 
                    regex=r"GFUS[0-9O]{14}", 
                    tesseract_cmd=tesseract_cmd,
                    enable_preprocessing=False,
                    engine=ocr_engine
                )
                if gofo_result and gofo_result[0].upper() == 'G':
                    # 修正常见的OCR错误：将O替换为0
                    gofo_result = gofo_result.replace('O', '0').replace('o', '0')
                    # 再次验证格式：必须是GFUS+14位纯数字
                    if re.match(r"^GFUS\d{14}$", gofo_result):
                        candidates.append(("GOFO", gofo_result, 3))
                        log_msgs.append(f"  区域3(GOFO): {gofo_result}")
                
                # 检查UniUni (U开头，UUS+16位字符)
                uni_result = ocr_order_number(
                    region_uni, 
                    regex=r"UUS[A-Za-z0-9]{16}", 
                    tesseract_cmd=tesseract_cmd,
                    enable_preprocessing=True,
                    engine=ocr_engine
                )
                if uni_result and uni_result[0].upper() == 'U':
                    candidates.append(("Uni", uni_result, 2))
                    log_msgs.append(f"  区域2(Uni): {uni_result}")
                
                # 选择最佳结果
                if candidates:
                    # 优先选择符合格式的结果
                    carrier, order_no, use_region = candidates[0]
                    log_msgs.append(f"第{idx+1}页 识别为 {carrier} 单号: {order_no}")
                else:
                    # 都没识别到，使用默认区域
                    order_no = ocr_order_number(region_usps, regex=regex, tesseract_cmd=tesseract_cmd)
                    use_region = 1
                    log_msgs.append(f"第{idx+1}页 三区域均未识别，使用默认区域1")
                    
            elif uniuni_mode:
                txt1 = ocr_text_simple(region1, tesseract_cmd=tesseract_cmd)
                has_uus = "UUS" in txt1
                
                short = txt1.replace("\n", " ")
                if len(short) > 120: short = short[:120] + "…"
                log_msgs.append(f"第{idx+1}页 UniUni: 区域1文本='{short}', 匹配UUS={has_uus}")
                
                if has_uus or not bbox2:
                    if has_uus:
                        log_msgs.append(f"第{idx+1}页 UniUni: 区域1含UUS，使用区域1进行命名")
                    elif not bbox2:
                        log_msgs.append(f"第{idx+1}页 UniUni: 区域1不含UUS，但未设置区域2，继续使用区域1")
                    order_no = ocr_order_number(region1, regex=regex, tesseract_cmd=tesseract_cmd)
                else:
                    region2 = crop_region(img, bbox2)
                    use_region = 2
                    log_msgs.append(f"第{idx+1}页 UniUni: 区域1不含UUS，改用区域2进行命名")
                    order_no = ocr_order_number(region2, regex=regex, tesseract_cmd=tesseract_cmd)
                log_msgs.append(f"第{idx+1}页 UniUni: 最终使用区域{use_region}")
            else:
                order_no = ocr_order_number(region1, regex=regex, tesseract_cmd=tesseract_cmd)
            
            # Smart Optimization - 已注释(存在细节错误)
            # if order_no:
            #     optimized = None
            #     detected_type = None
            #     
            #     # 检测 USPS (9开头的纯数字, 22位)
            #     # 示例: 9400136208423282801755
            #     # 宽松匹配: 只要包含符合格式的串即可，不强制要求字符串以9开头
            #     if '9' in order_no:
            #         usps_match = re.search(r'9\d{21}', order_no)
            #         if usps_match: optimized = usps_match.group(0); detected_type = "USPS"
            #     
            #     # 检测 GOFO (GFUS开头 + 14位数字)
            #     # 示例: GFUS01020467935616
            #     if not optimized and 'GFUS' in order_no.upper():
            #         gofo_match = re.search(r'GFUS\d{14}', order_no, re.IGNORECASE)
            #         if gofo_match: optimized = gofo_match.group(0); detected_type = "GOFO"
            #     
            #     # 检测 UniUni (UUS开头 + 16位字符)
            #     # 示例: UUS5BN3825409025413 (总长19位)
            #     # 用户提示: 第六位和第七位是大写英文 (示例中是第5、6位为BN，可能存在计数差异，这里放宽匹配)
            #     if not optimized and 'UUS' in order_no.upper():
            #         uni_match = re.search(r'UUS[A-Za-z0-9]{16}', order_no, re.IGNORECASE)
            #         if uni_match: optimized = uni_match.group(0); detected_type = "UniUni"
            #     
            #     if optimized:
            #         log_msgs.append(f"第{idx+1}页: 智能优化 {detected_type} '{order_no}' → '{optimized}'")
            #         order_no = optimized
            
            if not order_no:
                order_no = f"page_{idx+1:03d}"
            safe_name = sanitize_filename(order_no)
            
            # PDF Writing needs to be careful. 
            # We can create a new PdfWriter and add the page from a fresh reader to avoid concurrency issues with the shared reader.
            # Or we can lock the shared reader.
            # Safer approach: Open a local reader for this thread.
            local_reader = PdfReader(input_pdf)
            writer = PdfWriter()
            writer.add_page(local_reader.pages[idx])
            
            filename = f"{prefix}{safe_name}"
            if not filename.lower().endswith(".pdf"):
                filename += ".pdf"
            out_path = os.path.join(out_dir, filename)
            
            with open(out_path, "wb") as f:
                writer.write(f)
                
            log_msgs.append(f"已保存: {out_path}")
            return True, log_msgs
            
        except Exception as e:
            return False, [f"第{idx+1}页处理失败: {e}"]

    saved = 0
    completed = 0
    
    # 动态调整线程数
    # PaddleOCR 必须加锁串行，开多线程只会抢占 PDF 渲染资源，且容易内存积压
    if ocr_engine == "paddle":
        max_workers = 2
    else:
        max_workers = min(8, os.cpu_count() or 4)
        
    if logger_func: logger_func(f"启动多线程处理，线程数: {max_workers} (引擎: {ocr_engine})")
    
    if progress_callback:
        progress_callback(0, total_pages, "正在初始化...")
    
    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
        # Submit all tasks
        future_to_idx = {executor.submit(process_page, i): i for i in range(total_pages)}
        
        for future in concurrent.futures.as_completed(future_to_idx):
            idx = future_to_idx[future]
            try:
                success, msgs = future.result()
                if logger_func:
                    for msg in msgs: logger_func(msg)
                if success:
                    saved += 1
                completed += 1
                # 更新进度
                if progress_callback:
                    progress_callback(completed, total_pages, f"正在处理第 {idx+1}/{total_pages} 页...")
            except Exception as exc:
                if logger_func: logger_func(f"第{idx+1}页 发生未捕获异常: {exc}")
                completed += 1
                if progress_callback:
                    progress_callback(completed, total_pages, f"第 {idx+1} 页处理失败...")

    # 处理完成后释放OCR实例以回收内存
    if ocr_engine == "paddle":
        release_paddle_instance()
        import gc
        gc.collect()
        if logger_func: logger_func("✓ 已释放PaddleOCR引擎内存")
    elif ocr_engine == "rapid":
        release_rapid_instance()
        import gc
        gc.collect()
        if logger_func: logger_func("✓ 已释放RapidOCR引擎内存")

    return f"完成！共输出 {saved} 个PDF到: {out_dir}"


def _prep_gray_resize(img: "Image.Image", width: int) -> "Image.Image":
    g = ImageOps.grayscale(img)
    w = width
    h = int(g.height * (w / g.width))
    return g.resize((w, h))


def _template_present(page_img: "Image.Image", tpl_path: str, step: int = 12, page_w: int = 900, tpl_w: int = 180, thresh: float = 18.0) -> bool:
    try:
        tpl = Image.open(tpl_path)
    except Exception:
        return False
    P = _prep_gray_resize(page_img, page_w)
    T = _prep_gray_resize(tpl, tpl_w)
    pw, ph = P.size
    tw, th = T.size
    if tw > pw or th > ph:
        return False
    best = 255.0
    for y in range(0, ph - th + 1, step):
        for x in range(0, pw - tw + 1, step):
            patch = P.crop((x, y, x + tw, y + th))
            diff = ImageChops.difference(patch, T)
            m = ImageStat.Stat(diff).mean[0]
            if m < best:
                best = m
                if best <= thresh:
                    return True
    return best <= thresh


def _template_best_diff(page_img: "Image.Image", tpl_path: str, step: int = 12, page_w: int = 900, tpl_w: int = 180) -> float:
    try:
        tpl = Image.open(tpl_path)
    except Exception:
        return 255.0
    P = _prep_gray_resize(page_img, page_w)
    T = _prep_gray_resize(tpl, tpl_w)
    pw, ph = P.size
    tw, th = T.size
    if tw > pw or th > ph:
        return 255.0
    best = 255.0
    for y in range(0, ph - th + 1, step):
        for x in range(0, pw - tw + 1, step):
            patch = P.crop((x, y, x + tw, y + th))
            diff = ImageChops.difference(patch, T)
            m = ImageStat.Stat(diff).mean[0]
            if m < best:
                best = m
    return best


def _template_present_ncc(page_img: "Image.Image", tpl_path: str, step: int = 12, page_w: int = 900, tpl_w: int = 180, thresh: float = 0.75) -> bool:
    try:
        tpl = Image.open(tpl_path)
    except Exception:
        return False
    P = _prep_gray_resize(page_img, page_w)
    T = _prep_gray_resize(tpl, tpl_w)
    pw, ph = P.size
    tw, th = T.size
    if tw > pw or th > ph:
        return False
    t_data = list(T.getdata())
    n = len(t_data)
    sum_t = 0.0
    sum_t2 = 0.0
    for v in t_data:
        sum_t += v
        sum_t2 += v * v
    best = -1.0
    for y in range(0, ph - th + 1, step):
        for x in range(0, pw - tw + 1, step):
            patch = P.crop((x, y, x + tw, y + th))
            p_data = list(patch.getdata())
            sum_p = 0.0
            sum_p2 = 0.0
            sum_pt = 0.0
            for i in range(n):
                pv = p_data[i]
                tv = t_data[i]
                sum_p += pv
                sum_p2 += pv * pv
                sum_pt += pv * tv
            denom = ((n * sum_p2 - sum_p * sum_p) * (n * sum_t2 - sum_t * sum_t)) ** 0.5
            if denom <= 1e-9:
                corr = -1.0
            else:
                corr = (n * sum_pt - sum_p * sum_t) / denom
            if corr > best:
                best = corr
                if best >= thresh:
                    return True
    return best >= thresh


def _template_best_ncc(page_img: "Image.Image", tpl_path: str, step: int = 12, page_w: int = 900, tpl_w: int = 180) -> float:
    try:
        tpl = Image.open(tpl_path)
    except Exception:
        return -1.0
    P = _prep_gray_resize(page_img, page_w)
    T = _prep_gray_resize(tpl, tpl_w)
    pw, ph = P.size
    tw, th = T.size
    if tw > pw or th > ph:
        return -1.0
    t_data = list(T.getdata())
    n = len(t_data)
    sum_t = 0.0
    sum_t2 = 0.0
    for v in t_data:
        sum_t += v
        sum_t2 += v * v
    best = -1.0
    for y in range(0, ph - th + 1, step):
        for x in range(0, pw - tw + 1, step):
            patch = P.crop((x, y, x + tw, y + th))
            p_data = list(patch.getdata())
            sum_p = 0.0
            sum_p2 = 0.0
            sum_pt = 0.0
            for i in range(n):
                pv = p_data[i]
                tv = t_data[i]
                sum_p += pv
                sum_p2 += pv * pv
                sum_pt += pv * tv
            denom = ((n * sum_p2 - sum_p * sum_p) * (n * sum_t2 - sum_t * sum_t)) ** 0.5
            if denom <= 1e-9:
                corr = -1.0
            else:
                corr = (n * sum_pt - sum_p * sum_t) / denom
            if corr > best:
                best = corr
    return best


def split_pdf_with_ocr_v8(
    input_pdf: str,
    out_dir: str,
    bbox_usps: Tuple[int, int, int, int],
    bbox_gofo: Tuple[int, int, int, int],
    bbox_uni: Tuple[int, int, int, int],
    dpi: int = 300,
    poppler_path: Optional[str] = None,
    tesseract_cmd: Optional[str] = None,
    regex: Optional[str] = None,
    prefix: str = "",
    logger: Callable[[str], None] = print,
    template_dir: Optional[str] = None,
    template_usps: Optional[str] = None,
    template_gofo: Optional[str] = None,
    template_uni: Optional[str] = None,
    template_threshold: Optional[float] = None,
    template_step: Optional[int] = None,
    page_resize_w: Optional[int] = None,
    template_resize_w: Optional[int] = None,
    template_mode: str = "diff",
) -> str:
    if _IMPORT_ERROR is not None:
        raise RuntimeError(
            "PDF/OCR 依赖未安装，请先安装: pip install pypdf pdf2image pytesseract pillow; 并在 Windows 上安装 Poppler 与 Tesseract。详细见 README。\n"
            f"ImportError: {_IMPORT_ERROR}"
        )
    poppler_resolved = find_poppler(poppler_path)
    _ensure_tesseract_cmd(tesseract_cmd)
    td = _get_tessdata_dir()
    reader = PdfReader(input_pdf)
    total_pages = len(reader.pages)
    os.makedirs(out_dir, exist_ok=True)
    if not template_dir:
        template_dir = os.path.join(os.getcwd(), "Png")
    tpl_usps = template_usps or os.path.join(template_dir, "USPS-Pn.png")
    tpl_gofo = template_gofo or os.path.join(template_dir, "GOFO-Pn.png")
    tpl_uni = template_uni or os.path.join(template_dir, "Uni-Pn.png")
    saved = 0
    thr = template_threshold if template_threshold is not None else (0.75 if template_mode == "ncc" else 18.0)
    stp = template_step if template_step is not None else 12
    pgw = page_resize_w if page_resize_w is not None else 900
    tpw = template_resize_w if template_resize_w is not None else 180
    for idx in range(total_pages):
        img = render_page_to_image(input_pdf, idx, dpi, poppler_resolved)
        label = None
        if template_mode == "ncc":
            match_usps = _template_present_ncc(img, tpl_usps, step=stp, page_w=pgw, tpl_w=tpw, thresh=thr)
            match_gofo = _template_present_ncc(img, tpl_gofo, step=stp, page_w=pgw, tpl_w=tpw, thresh=thr)
            match_uni = _template_present_ncc(img, tpl_uni, step=stp, page_w=pgw, tpl_w=tpw, thresh=thr)
        else:
            match_usps = _template_present(img, tpl_usps, step=stp, page_w=pgw, tpl_w=tpw, thresh=thr)
            match_gofo = _template_present(img, tpl_gofo, step=stp, page_w=pgw, tpl_w=tpw, thresh=thr)
            match_uni = _template_present(img, tpl_uni, step=stp, page_w=pgw, tpl_w=tpw, thresh=thr)
        if match_usps:
            label = "USPS"
            bbox = bbox_usps
            # USPS: 9开头，22位纯数字
            target_regex = r"9\d{21}"
        elif match_gofo:
            label = "GOFO"
            bbox = bbox_gofo
            # GOFO: GFUS+14位数字 (兼容偶发的15位)
            target_regex = r"GFUS\d{14,15}"
        elif match_uni:
            label = "Uni"
            bbox = bbox_uni
            # UniUni: UUS+16位混合 (总长19位)
            target_regex = r"UUS[A-Za-z0-9]{16}"
        else:
            label = "USPS"
            bbox = bbox_usps
            target_regex = regex  # Fallback to generic regex

        region = crop_region(img, bbox)
        # Use the specific regex if matched, otherwise use the generic one passed in
        current_regex = target_regex if label != "USPS" or match_usps else regex
        # Note: If no template matched (else block), we default to USPS bbox but maybe we should keep generic regex?
        # The logic above: if else (no match), label=USPS, bbox=USPS. 
        # Let's assume generic regex if no match found, but specific if match found.
        
        if label == "USPS" and not match_usps:
             # No match found, use generic regex
             current_regex = regex

        order_no = ocr_order_number(region, regex=current_regex, tesseract_cmd=tesseract_cmd)
        if not order_no:
            order_no = f"page_{idx+1:03d}"
        safe_name = sanitize_filename(order_no)
        writer = PdfWriter()
        writer.add_page(reader.pages[idx])
        filename = f"{prefix}{safe_name}"
        if not filename.lower().endswith(".pdf"):
            filename += ".pdf"
        out_path = os.path.join(out_dir, filename)
        with open(out_path, "wb") as f:
            writer.write(f)
        saved += 1
        if logger:
            logger(f"第{idx+1}页: 模板={label} 输出= {out_path}")
    return f"完成！共输出 {saved} 个PDF到: {out_dir}"


def detect_pdf_template_matches(
    input_pdf: str,
    dpi: int = 300,
    poppler_path: Optional[str] = None,
    template_dir: Optional[str] = None,
    template_usps: Optional[str] = None,
    template_gofo: Optional[str] = None,
    template_uni: Optional[str] = None,
    template_threshold: Optional[float] = None,
    template_mode: str = "diff",
    logger: Callable[[str], None] = print,
) -> str:
    if _IMPORT_ERROR is not None:
        raise RuntimeError(
            "PDF/OCR 依赖未安装，请先安装: pip install pypdf pdf2image pytesseract pillow; 并在 Windows 上安装 Poppler 与 Tesseract。详细见 README。\n"
            f"ImportError: {_IMPORT_ERROR}"
        )
    poppler_resolved = find_poppler(poppler_path)
    reader = PdfReader(input_pdf)
    total_pages = len(reader.pages)
    if not template_dir:
        template_dir = os.path.join(os.getcwd(), "Png")
    tpl_usps = template_usps or os.path.join(template_dir, "USPS-Pn.png")
    tpl_gofo = template_gofo or os.path.join(template_dir, "GOFO-Pn.png")
    tpl_uni = template_uni or os.path.join(template_dir, "Uni-Pn.png")
    thr = template_threshold if template_threshold is not None else (0.75 if template_mode == "ncc" else 18.0)
    stp = 12
    pgw = 900
    tpw = 180
    c_usps = c_gofo = c_uni = c_none = 0
    for idx in range(total_pages):
        img = render_page_to_image(input_pdf, idx, dpi, poppler_resolved)
        if template_mode == "ncc":
            m_usps = _template_present_ncc(img, tpl_usps, step=stp, page_w=pgw, tpl_w=tpw, thresh=thr)
            m_gofo = _template_present_ncc(img, tpl_gofo, step=stp, page_w=pgw, tpl_w=tpw, thresh=thr)
            m_uni = _template_present_ncc(img, tpl_uni, step=stp, page_w=pgw, tpl_w=tpw, thresh=thr)
            b_usps = _template_best_ncc(img, tpl_usps, step=stp, page_w=pgw, tpl_w=tpw)
            b_gofo = _template_best_ncc(img, tpl_gofo, step=stp, page_w=pgw, tpl_w=tpw)
            b_uni = _template_best_ncc(img, tpl_uni, step=stp, page_w=pgw, tpl_w=tpw)
        else:
            m_usps = _template_present(img, tpl_usps, step=stp, page_w=pgw, tpl_w=tpw, thresh=thr)
            m_gofo = _template_present(img, tpl_gofo, step=stp, page_w=pgw, tpl_w=tpw, thresh=thr)
            m_uni = _template_present(img, tpl_uni, step=stp, page_w=pgw, tpl_w=tpw, thresh=thr)
            b_usps = _template_best_diff(img, tpl_usps, step=stp, page_w=pgw, tpl_w=tpw)
            b_gofo = _template_best_diff(img, tpl_gofo, step=stp, page_w=pgw, tpl_w=tpw)
            b_uni = _template_best_diff(img, tpl_uni, step=stp, page_w=pgw, tpl_w=tpw)
        label = None
        if m_usps:
            c_usps += 1; label = "USPS"
        elif m_gofo:
            c_gofo += 1; label = "GOFO"
        elif m_uni:
            c_uni += 1; label = "Uni"
        else:
            c_none += 1; label = "无匹配"
        if logger:
            logger(f"第{idx+1}页: 匹配={label} (USPS={m_usps}/best={b_usps:.3f} GOFO={m_gofo}/best={b_gofo:.3f} Uni={m_uni}/best={b_uni:.3f})")
    summary = f"USPS: {c_usps}  GOFO: {c_gofo}  Uni: {c_uni}  未匹配: {c_none}  总页数: {total_pages}"
    return summary
