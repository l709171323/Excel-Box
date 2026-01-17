import math
import random
import json
import os
from typing import Dict, Set, Tuple, Optional, Callable, Any
from excel_toolkit.excel_lite import ExcelReader, ExcelWriter
from excel_toolkit.excel_lite import column_index_from_string
# PatternFill functionality moved to ExcelWriter
from excel_toolkit.states import get_state_abbreviation

# 常量定义
EARTH_RADIUS_KM = 6371.0  # 地球半径（公里）
DISTANCE_EPSILON = 1e-9  # 浮点数比较容差（公里）
DEFAULT_HIGHLIGHT_COLOR = "FFF59E"  # 已填充行的高亮颜色

# 全局变量：延迟加载州坐标数据
_STATE_COORDS: Optional[Dict[str, Tuple[float, float]]] = None


def _load_state_coords() -> Dict[str, Tuple[float, float]]:
    """从 JSON 文件加载州坐标数据（单例模式）"""
    global _STATE_COORDS
    if _STATE_COORDS is not None:
        return _STATE_COORDS
    
    # 获取配置文件路径
    current_dir = os.path.dirname(os.path.abspath(__file__))
    config_path = os.path.join(current_dir, 'state_coords.json')
    
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
            # 转换列表为元组
            _STATE_COORDS = {k: tuple(v) for k, v in data.items()}
            return _STATE_COORDS
    except FileNotFoundError:
        # 如果文件不存在，使用内置的默认数据（向后兼容）
        _STATE_COORDS = {
            "AL": (32.806671, -86.791130), "AK": (61.385000, -152.268300), "AZ": (34.048928, -111.093731),
            "AR": (34.969704, -92.373123), "CA": (36.778259, -119.417931), "CO": (39.550051, -105.782067),
            "CT": (41.603221, -73.087749), "DE": (38.910832, -75.527670), "FL": (27.994402, -81.760254),
            "GA": (32.157435, -82.907123), "HI": (19.741755, -155.844437), "ID": (44.068203, -114.742041),
            "IL": (40.633125, -89.398528), "IN": (40.551217, -85.602364), "IA": (41.878003, -93.097702),
            "KS": (39.011902, -98.484246), "KY": (37.839333, -84.270019), "LA": (30.984298, -91.962333),
            "ME": (45.253783, -69.445469), "MD": (39.045755, -76.641271), "MA": (42.407210, -71.382437),
            "MI": (44.314844, -85.602364), "MN": (46.729553, -94.685900), "MS": (32.354668, -89.398528),
            "MO": (37.964253, -91.831833), "MT": (46.879682, -110.362566), "NE": (41.492537, -99.901813),
            "NV": (38.802610, -116.419389), "NH": (43.193852, -71.572395), "NJ": (40.058324, -74.405661),
            "NM": (34.519940, -105.870090), "NY": (43.299428, -74.217933), "NC": (35.759573, -79.019300),
            "ND": (47.551493, -101.002012), "OH": (40.417287, -82.907123), "OK": (35.007751, -97.092877),
            "OR": (43.804133, -120.554201), "PA": (41.203322, -77.194525), "RI": (41.580095, -71.477429),
            "SC": (33.836081, -81.163725), "SD": (43.969515, -99.901813), "TN": (35.517491, -86.580447),
            "TX": (31.968599, -99.901813), "UT": (39.320980, -111.093731), "VT": (44.558803, -72.577841),
            "VA": (37.431573, -78.656894), "WA": (47.751074, -120.740139), "WV": (38.597626, -80.454903),
            "WI": (43.784440, -88.787868), "WY": (43.075968, -107.290284), "DC": (38.907192, -77.036873),
            "PR": (18.220833, -66.590149), "GU": (13.444304, 144.793731), "VI": (18.335765, -64.896335)
        }
        return _STATE_COORDS
    except Exception as e:
        raise RuntimeError(f"加载州坐标配置失败: {e}")

def _haversine(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    """使用哈弗辛公式计算两点间的球面距离
    
    Args:
        lat1: 第一个点的纬度（度）
        lon1: 第一个点的经度（度）
        lat2: 第二个点的纬度（度）
        lon2: 第二个点的经度（度）
    
    Returns:
        两点间的距离（公里）
    """
    dlat = math.radians(lat2 - lat1)
    dlon = math.radians(lon2 - lon1)
    a = math.sin(dlat / 2.0) ** 2 + math.cos(math.radians(lat1)) * math.cos(math.radians(lat2)) * math.sin(dlon / 2.0) ** 2
    c = 2.0 * math.atan2(math.sqrt(a), math.sqrt(1.0 - a))
    return EARTH_RADIUS_KM * c

def _state_to_abbr(value: Any) -> Optional[str]:
    if value is None:
        return None
    s = str(value).strip()
    if len(s) == 2:
        return s.upper()
    ab = get_state_abbreviation(s)
    return ab if ab != "未找到" else None

def _load_inventory(
    inventory_file: str, 
    logger: Callable[[str], None] = print, 
    blocked_states: Optional[Set[str]] = None, 
    blocked_names: Optional[Set[str]] = None
) -> Tuple[Dict[str, Set[str]], Dict[str, Optional[str]]]:
    import os
    file_ext = os.path.splitext(inventory_file)[1].lower()
    
    sku_by_wh = {}
    wh_state = {}
    
    if file_ext in ['.xlsx', '.xlsm']:
        # 使用openpyxl读取新版Excel
        wb = ExcelReader(inventory_file)
        for name in wb.sheetnames:
            ws = wb[name]
            nm = str(name).strip()
            if nm == "仓库名和地址":
                for r in range(1, ws.max_row + 1):
                    wname = ws.cell(row=r, column=1).value
                    wstate = ws.cell(row=r, column=2).value
                    if wname:
                        abbr = _state_to_abbr(wstate)
                        wh_state[str(wname).strip()] = abbr
                        if logger:
                            mark_state = " (屏蔽)" if blocked_states and abbr in blocked_states else ""
                            mark_name = " (屏蔽名称)" if blocked_names and str(wname).strip() in blocked_names else ""
                            logger(f"映射：仓库='{str(wname).strip()}' 州='{abbr or wstate}{mark_state}{mark_name}'")
                continue
            s = set()
            for r in range(1, ws.max_row + 1):
                v = ws.cell(row=r, column=1).value
                if v is None:
                    continue
                s.add(str(v).strip())
            sku_by_wh[nm] = s
            if logger:
                logger(f"仓库表：{nm} SKU数量={len(s)}")
                if s:
                    samples = list(s)
                    samples.sort()
                    samples = samples[:10]
                    logger(f"示例SKU：{', '.join(samples)}")
        wb.close()
    
    elif file_ext == '.xls':
        # 使用xlrd读取旧版Excel
        import xlrd
        wb = xlrd.open_workbook(inventory_file)
        for name in wb.sheetnames:
            ws = wb.sheet_by_name(name)
            nm = str(name).strip()
            if nm == "仓库名和地址":
                for r in range(ws.nrows):
                    wname = ws.cell_value(r, 0) if ws.ncols > 0 else None
                    wstate = ws.cell_value(r, 1) if ws.ncols > 1 else None
                    if wname:
                        abbr = _state_to_abbr(wstate)
                        wh_state[str(wname).strip()] = abbr
                        if logger:
                            mark_state = " (屏蔽)" if blocked_states and abbr in blocked_states else ""
                            mark_name = " (屏蔽名称)" if blocked_names and str(wname).strip() in blocked_names else ""
                            logger(f"映射：仓库='{str(wname).strip()}' 州='{abbr or wstate}{mark_state}{mark_name}'")
                continue
            s = set()
            for r in range(ws.nrows):
                v = ws.cell_value(r, 0) if ws.ncols > 0 else None
                if v is None or v == '':
                    continue
                s.add(str(v).strip())
            sku_by_wh[nm] = s
            if logger:
                logger(f"仓库表：{nm} SKU数量={len(s)}")
                if s:
                    samples = list(s)
                    samples.sort()
                    samples = samples[:10]
                    logger(f"示例SKU：{', '.join(samples)}")
    else:
        raise ValueError(f"不支持的文件格式: {file_ext}，请使用 .xlsx, .xlsm 或 .xls 文件")
    
    if logger:
        logger(f"库存加载完成：仓库数={len(sku_by_wh)}；州映射数={sum(1 for k in wh_state if wh_state[k])}")
        missing = [w for w in sku_by_wh.keys() if not wh_state.get(w)]
        if missing:
            logger(f"缺少州映射的仓库：{', '.join(missing)}")
        all_names = sorted(set(list(sku_by_wh.keys()) + list(wh_state.keys())))
        logger("仓库\t州\tSKU")
        for nm in all_names:
            st = wh_state.get(nm) or ""
            skus = sku_by_wh.get(nm) or set()
            sku_str = ", ".join(sorted(skus)) if skus else "-"
            mark_state = " (屏蔽)" if blocked_states and st in blocked_states else ""
            mark_name = " (屏蔽名称)" if blocked_names and nm in blocked_names else ""
            logger(f"{nm}\t{(st or '未知')}{mark_state}{mark_name}\t{sku_str}")
    return sku_by_wh, wh_state

def process_warehouse_routing(
    file_name: str, 
    sheet_name: str, 
    sku_col_letter: str, 
    state_col_letter: str, 
    dst_col_letter: str, 
    inventory_file: str, 
    logger: Callable[[str], None] = print, 
    block_tech_states: bool = False, 
    blocked_warehouses: Optional[list] = None
) -> str:
    try:
        wb = ExcelReader(file_name)
    except Exception as e:
        return f"加载收件信息表格失败: {e}"
    try:
        ws = wb[sheet_name]
    except Exception as e:
        wb.close(); return f"获取子表失败: {e}"
    try:
        sku_col = column_index_from_string(sku_col_letter)
        state_col = column_index_from_string(state_col_letter)
        dst_col = column_index_from_string(dst_col_letter)
    except Exception:
        wb.close(); return "列号无效，请检查SKU列、州列与输出列。"
    blocked_states = {"GA", "TX"} if block_tech_states else set()
    names_set = set(blocked_warehouses or [])
    
    # 加载州坐标数据
    STATE_COORDS = _load_state_coords()
    
    sku_by_wh, wh_state = _load_inventory(inventory_file, logger=logger, blocked_states=blocked_states, blocked_names=names_set)
    if logger:
        if block_tech_states:
            bwh = [w for w, st in wh_state.items() if st in blocked_states]
            logger(f"屏蔽科技单(州)已启用：屏蔽仓库={', '.join(bwh) if bwh else '无'}")
        if names_set:
            logger(f"按名称屏蔽已启用：屏蔽仓库={', '.join(sorted(names_set))}")
    changes = 0
    for r in range(1, ws.max_row + 1):
        dst_current = ws.cell(row=r, column=dst_col).value
        if dst_current is not None and str(dst_current).strip() != "":
            try:
                ws.cell(row=r, column=dst_col).fill = PatternFill(
                    start_color=DEFAULT_HIGHLIGHT_COLOR, 
                    end_color=DEFAULT_HIGHLIGHT_COLOR, 
                    fill_type="solid"
                )
            except Exception:
                pass
            if logger:
                logger(f"第{r}行：目标列已有内容，已跳过并高亮")
            continue
        sku = ws.cell(row=r, column=sku_col).value
        st_raw = ws.cell(row=r, column=state_col).value
        if sku is None or str(sku).strip() == "":
            continue
        st = _state_to_abbr(st_raw)
        if not st or st not in STATE_COORDS:
            if logger:
                logger(f"第{r}行：州无效或未知，跳过")
            continue
        candidates = []
        for wname, sset in sku_by_wh.items():
            if str(sku).strip() in sset:
                stw = wh_state.get(wname)
                if stw in blocked_states:
                    continue
                if wname in names_set:
                    continue
                candidates.append(wname)
        if not candidates:
            ws.cell(row=r, column=dst_col).value = "无可发货仓"
            if logger:
                logger(f"第{r}行：SKU={sku} 无仓库可发货")
            continue
        lat1, lon1 = STATE_COORDS[st]
        best_w = None
        best_d = None
        best_list = []
        for wname in candidates:
            wst = wh_state.get(wname)
            if not wst or wst not in STATE_COORDS:
                continue
            lat2, lon2 = STATE_COORDS[wst]
            d = _haversine(lat1, lon1, lat2, lon2)
            if best_d is None or d < best_d - DISTANCE_EPSILON:
                best_d = d
                best_w = wname
                best_list = [wname]
            elif best_d is not None and abs(d - best_d) <= DISTANCE_EPSILON:
                best_list.append(wname)
        if len(best_list) > 1:
            best_w = random.choice(best_list)
        if not best_w:
            ws.cell(row=r, column=dst_col).value = "缺少仓库州映射"
            if logger:
                logger(f"第{r}行：SKU={sku} 候选={len(candidates)} 但缺少仓库州映射，未能计算距离")
            continue
        ws.cell(row=r, column=dst_col).value = best_w
        changes += 1
        if logger:
            logger(f"第{r}行：SKU={sku} 候选={len(candidates)} 选择={best_w}")
    try:
        wb.save(file_name)
        wb.close()
        return f"路由完成！共写入 {changes} 行。"
    except PermissionError:
        return f"错误：保存失败！请关闭Excel文件 '{file_name}' 后再运行。"
    except Exception as e:
        return f"保存文件时发生错误: {e}"

def read_inventory(
    file_path: str, 
    logger: Callable[[str], None] = print
) -> Tuple[Dict[str, Set[str]], Dict[str, Optional[str]]]:
    sku_by_wh, wh_state = _load_inventory(file_path, logger=logger)
    return sku_by_wh, wh_state

def write_inventory(
    file_path: str, 
    wh_state: Dict[str, Optional[str]], 
    sku_by_wh: Dict[str, Set[str]], 
    logger: Callable[[str], None] = print
) -> str:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "仓库名和地址"
    rows = sorted(list(wh_state.items()), key=lambda x: x[0])
    for i, (w, st) in enumerate(rows, start=1):
        ws.cell(row=i, column=1).value = str(w).strip()
        ws.cell(row=i, column=2).value = _state_to_abbr(st)
    for w, skus in sku_by_wh.items():
        nm = str(w).strip()
        if nm == "":
            continue
        ws2 = wb.create_sheet(nm)
        if isinstance(skus, (set, list, tuple)):
            items = sorted([str(x).strip() for x in skus if str(x).strip() != ""]) 
        elif isinstance(skus, dict):
            items = sorted([str(x).strip() for x in skus.keys() if str(x).strip() != ""]) 
        else:
            items = []
        for i, sku in enumerate(items, start=1):
            ws2.cell(row=i, column=1).value = sku
    wb.save(file_path)
    wb.close()
    if logger:
        logger(f"已生成库存文件：{file_path}")
    return f"已生成库存文件：{file_path}"