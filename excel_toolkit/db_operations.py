"""
数据库操作辅助函数
"""
from typing import List, Optional, Dict, Any
from excel_toolkit.db_config import get_db_manager
from excel_toolkit.db_models import OCRTemplate, ShippingConfig

# ==================== OCR模板操作（功能6） ====================

def save_ocr_template(name: str, region: int, bbox: Dict[str, int], 
                     description: Optional[str] = None):
    """
    保存OCR模板到数据库
    
    Args:
        name: 模板名称
        region: 区域编号（1=USPS, 2=Uni, 3=GOFO）
        bbox: 边界框 {"x": int, "y": int, "width": int, "height": int}
        description: 描述
    
    Returns:
        (成功标志, 消息)
    """
    db = get_db_manager()
    if not db.is_connected():
        success, msg = db.connect()
        if not success:
            return False, f"无法连接数据库: {msg}"
    
    session = db.get_session()
    try:
        # 检查是否已存在同名同区域的模板
        existing = session.query(OCRTemplate).filter(
            OCRTemplate.name == name,
            OCRTemplate.region == region
        ).first()
        
        if existing:
            # 更新现有模板
            existing.bbox_x = bbox['x']
            existing.bbox_y = bbox['y']
            existing.bbox_width = bbox['width']
            existing.bbox_height = bbox['height']
            if description:
                existing.description = description
            session.commit()
            return True, f"模板 '{name}' 已更新"
        else:
            # 创建新模板
            template = OCRTemplate(
                name=name,
                region=region,
                bbox_x=bbox['x'],
                bbox_y=bbox['y'],
                bbox_width=bbox['width'],
                bbox_height=bbox['height'],
                description=description
            )
            session.add(template)
            session.commit()
            return True, f"模板 '{name}' 已保存"
    
    except Exception as e:
        session.rollback()
        return False, f"保存失败: {str(e)}"
    finally:
        session.close()


def load_ocr_template(name: str, region: int) -> Optional[Dict[str, Any]]:
    """
    从数据库加载OCR模板
    
    Args:
        name: 模板名称
        region: 区域编号
    
    Returns:
        模板数据（字典格式）或None
    """
    db = get_db_manager()
    if not db.is_connected():
        success, msg = db.connect()
        if not success:
            return None
    
    session = db.get_session()
    try:
        template = session.query(OCRTemplate).filter(
            OCRTemplate.name == name,
            OCRTemplate.region == region
        ).first()
        
        if template:
            return template.to_json_format()
        return None
    
    except Exception as e:
        print(f"加载模板失败: {e}")
        return None
    finally:
        session.close()


def list_ocr_templates(region: Optional[int] = None) -> List[Dict[str, Any]]:
    """
    列出所有OCR模板
    
    Args:
        region: 可选，筛选特定区域
    
    Returns:
        模板列表
    """
    db = get_db_manager()
    if not db.is_connected():
        success, msg = db.connect()
        if not success:
            return []
    
    session = db.get_session()
    try:
        query = session.query(OCRTemplate)
        if region is not None:
            query = query.filter(OCRTemplate.region == region)
        
        templates = query.order_by(OCRTemplate.name, OCRTemplate.region).all()
        return [t.to_dict() for t in templates]
    
    except Exception as e:
        print(f"列出模板失败: {e}")
        return []
    finally:
        session.close()


def delete_ocr_template(name: str, region: int):
    """删除OCR模板"""
    db = get_db_manager()
    if not db.is_connected():
        success, msg = db.connect()
        if not success:
            return False, f"无法连接数据库: {msg}"
    
    session = db.get_session()
    try:
        template = session.query(OCRTemplate).filter(
            OCRTemplate.name == name,
            OCRTemplate.region == region
        ).first()
        
        if template:
            session.delete(template)
            session.commit()
            return True, f"模板 '{name}' 已删除"
        else:
            return False, "模板不存在"
    
    except Exception as e:
        session.rollback()
        return False, f"删除失败: {str(e)}"
    finally:
        session.close()


# ==================== 发货配置操作（功能11） ====================

def save_shipping_mapping(name: str, mapping_type: str, mapping_data: Dict[str, str],
                         description: Optional[str] = None):
    """
    保存发货列映射（映射1或映射2）
    
    Args:
        name: 配置名称（如："默认映射1"）
        mapping_type: "mapping1" 或 "mapping2"
        mapping_data: 映射数据 {"订单列名": "模板列名", ...}
        description: 描述
    
    Returns:
        (成功标志, 消息)
    """
    db = get_db_manager()
    if not db.is_connected():
        success, msg = db.connect()
        if not success:
            return False, f"无法连接数据库: {msg}"
    
    session = db.get_session()
    try:
        # 检查是否已存在
        existing = session.query(ShippingConfig).filter(
            ShippingConfig.name == name,
            ShippingConfig.config_type == mapping_type
        ).first()
        
        if existing:
            existing.config_data = mapping_data
            if description:
                existing.description = description
            session.commit()
            return True, f"映射 '{name}' 已更新"
        else:
            config = ShippingConfig(
                name=name,
                config_type=mapping_type,
                config_data=mapping_data,
                description=description
            )
            session.add(config)
            session.commit()
            return True, f"映射 '{name}' 已保存"
    
    except Exception as e:
        session.rollback()
        return False, f"保存失败: {str(e)}"
    finally:
        session.close()


def save_shipping_warehouse(warehouse_name: str, carrier_mapping: Dict[str, str],
                           description: Optional[str] = None):
    """
    保存仓库的物流渠道映射
    
    Args:
        warehouse_name: 仓库名称
        carrier_mapping: 承运商到物流渠道的映射 {"承运商": "物流渠道", ...}
        description: 描述
    
    Returns:
        (成功标志, 消息)
    """
    db = get_db_manager()
    if not db.is_connected():
        success, msg = db.connect()
        if not success:
            return False, f"无法连接数据库: {msg}"
    
    session = db.get_session()
    try:
        # 检查是否已存在
        existing = session.query(ShippingConfig).filter(
            ShippingConfig.warehouse_name == warehouse_name,
            ShippingConfig.config_type == "warehouse"
        ).first()
        
        if existing:
            existing.config_data = carrier_mapping
            if description:
                existing.description = description
            session.commit()
            return True, f"仓库 '{warehouse_name}' 的配置已更新"
        else:
            config = ShippingConfig(
                name=f"仓库-{warehouse_name}",
                config_type="warehouse",
                warehouse_name=warehouse_name,
                config_data=carrier_mapping,
                description=description
            )
            session.add(config)
            session.commit()
            return True, f"仓库 '{warehouse_name}' 的配置已保存"
    
    except Exception as e:
        session.rollback()
        return False, f"保存失败: {str(e)}"
    finally:
        session.close()


def load_shipping_config(name: str, config_type: str) -> Optional[Dict[str, str]]:
    """
    加载发货配置
    
    Args:
        name: 配置名称
        config_type: "mapping1", "mapping2", 或 "warehouse"
    
    Returns:
        配置数据（字典）或None
    """
    db = get_db_manager()
    if not db.is_connected():
        success, msg = db.connect()
        if not success:
            return None
    
    session = db.get_session()
    try:
        config = session.query(ShippingConfig).filter(
            ShippingConfig.name == name,
            ShippingConfig.config_type == config_type
        ).first()
        
        if config:
            return config.config_data
        return None
    
    except Exception as e:
        print(f"加载配置失败: {e}")
        return None
    finally:
        session.close()


def load_shipping_warehouse(warehouse_name: str) -> Optional[Dict[str, str]]:
    """加载仓库的物流渠道映射"""
    db = get_db_manager()
    if not db.is_connected():
        success, msg = db.connect()
        if not success:
            return None
    
    session = db.get_session()
    try:
        config = session.query(ShippingConfig).filter(
            ShippingConfig.warehouse_name == warehouse_name,
            ShippingConfig.config_type == "warehouse"
        ).first()
        
        if config:
            return config.config_data
        return None
    
    except Exception as e:
        print(f"加载仓库配置失败: {e}")
        return None
    finally:
        session.close()


def list_shipping_configs(config_type: Optional[str] = None) -> List[Dict[str, Any]]:
    """列出所有发货配置"""
    db = get_db_manager()
    if not db.is_connected():
        success, msg = db.connect()
        if not success:
            return []
    
    session = db.get_session()
    try:
        query = session.query(ShippingConfig)
        if config_type:
            query = query.filter(ShippingConfig.config_type == config_type)
        
        configs = query.order_by(ShippingConfig.name).all()
        return [c.to_dict() for c in configs]
    
    except Exception as e:
        print(f"列出配置失败: {e}")
        return []
    finally:
        session.close()


def get_all_warehouses() -> List[str]:
    """获取所有仓库名称列表"""
    db = get_db_manager()
    if not db.is_connected():
        success, msg = db.connect()
        if not success:
            return []
    
    session = db.get_session()
    try:
        warehouses = session.query(ShippingConfig.warehouse_name).filter(
            ShippingConfig.config_type == "warehouse",
            ShippingConfig.warehouse_name.isnot(None)
        ).distinct().all()
        
        return [w[0] for w in warehouses if w[0]]
    
    except Exception as e:
        print(f"获取仓库列表失败: {e}")
        return []
    finally:
        session.close()


def load_shipping_config_from_db(config_name: str = "默认配置") -> Optional[Dict[str, Any]]:
    """
    从数据库加载完整的发货配置（包括映射1、映射2和所有仓库）
    
    Returns:
        {
            "column_mapping_1": {...},
            "column_mapping_2": {...},
            "warehouses": [...],
            "shipping_map": {仓库名: {承运商: 物流渠道}}
        }
    """
    db = get_db_manager()
    if not db.is_connected():
        success, msg = db.connect()
        if not success:
            return None
    
    result = {
        "column_mapping_1": {},
        "column_mapping_2": {},
        "warehouses": [],
        "shipping_map": {}
    }
    
    session = db.get_session()
    try:
        # 加载映射1
        mapping1 = load_shipping_config(f"{config_name}-映射1", "mapping1")
        if mapping1:
            result["column_mapping_1"] = mapping1
        
        # 加载映射2
        mapping2 = load_shipping_config(f"{config_name}-映射2", "mapping2")
        if mapping2:
            result["column_mapping_2"] = mapping2
        
        # 加载所有仓库配置
        warehouses = get_all_warehouses()
        result["warehouses"] = warehouses
        
        for wh in warehouses:
            carrier_map = load_shipping_warehouse(wh)
            if carrier_map:
                result["shipping_map"][wh] = carrier_map
        
        return result
    
    except Exception as e:
        print(f"加载完整配置失败: {e}")
        return None
    finally:
        session.close()


# ==================== 仓库库存操作（功能9、10） ====================

from excel_toolkit.db_models import WarehouseInventory, WarehouseSKU


def save_warehouse_inventory(warehouse_data: Dict[str, str], 
                             sku_data: Dict[str, set]):
    """
    保存仓库库存到数据库
    
    Args:
        warehouse_data: {仓库名: 州代码, ...}
        sku_data: {仓库名: set([SKU列表]), ...}
    
    Returns:
        (成功标志, 消息)
    """
    db = get_db_manager()
    if not db.is_connected():
        success, msg = db.connect()
        if not success:
            return False, f"无法连接数据库: {msg}"
    
    session = db.get_session()
    try:
        # 1. 保存仓库信息
        for wh_name, state_code in warehouse_data.items():
            existing = session.query(WarehouseInventory).filter(
                WarehouseInventory.warehouse_name == wh_name
            ).first()
            
            if existing:
                existing.state_code = state_code if state_code else None
            else:
                warehouse = WarehouseInventory(
                    warehouse_name=wh_name,
                    state_code=state_code if state_code else None
                )
                session.add(warehouse)
        
        # 2. 清除旧的SKU数据（全量更新）
        session.query(WarehouseSKU).delete()
        
        # 3. 保存新的SKU数据
        for wh_name, skus in sku_data.items():
            for sku in skus:
                sku_obj = WarehouseSKU(
                    warehouse_name=wh_name,
                    sku=str(sku)
                )
                session.add(sku_obj)
        
        session.commit()
        total_wh = len(warehouse_data)
        total_sku = sum(len(skus) for skus in sku_data.values())
        return True, f"已保存 {total_wh} 个仓库，{total_sku} 个SKU"
    
    except Exception as e:
        session.rollback()
        return False, f"保存失败: {str(e)}"
    finally:
        session.close()


def load_warehouse_inventory() -> Optional[tuple]:
    """
    从数据库加载仓库库存
    
    Returns:
        (warehouse_data, sku_data) 或 None
        - warehouse_data: {仓库名: 州代码}
        - sku_data: {仓库名: set([SKU列表])}
    """
    db = get_db_manager()
    if not db.is_connected():
        success, msg = db.connect()
        if not success:
            return None
    
    session = db.get_session()
    try:
        # 加载仓库信息
        warehouses = session.query(WarehouseInventory).all()
        warehouse_data = {wh.warehouse_name: wh.state_code or '' for wh in warehouses}
        
        # 加载SKU信息
        skus = session.query(WarehouseSKU).all()
        sku_data = {}
        for sku_obj in skus:
            wh_name = sku_obj.warehouse_name
            if wh_name not in sku_data:
                sku_data[wh_name] = set()
            sku_data[wh_name].add(sku_obj.sku)
        
        return warehouse_data, sku_data
    
    except Exception as e:
        print(f"加载库存失败: {e}")
        return None
    finally:
        session.close()


def export_inventory_to_excel(output_path: str):
    """
    将数据库中的库存导出为Excel文件
    
    Args:
        output_path: 输出Excel文件路径
    
    Returns:
        (成功标志, 消息)
    """
    data = load_warehouse_inventory()
    if not data:
        return False, "无法从数据库加载库存数据"
    
    warehouse_data, sku_data = data
    
    # 调用现有的write_inventory函数
    from excel_toolkit.warehouse_router import write_inventory
    try:
        msg = write_inventory(output_path, warehouse_data, sku_data, lambda x: None)
        return True, msg
    except Exception as e:
        return False, f"导出失败: {str(e)}"


def import_inventory_from_excel(excel_path: str):
    """
    从Excel文件导入库存到数据库
    
    Args:
        excel_path: Excel文件路径
    
    Returns:
        (成功标志, 消息)
    """
    from excel_toolkit.warehouse_router import read_inventory
    
    try:
        sku_by_wh, wh_state = read_inventory(excel_path, logger=lambda x: None)
        
        # 转换格式
        warehouse_data = wh_state  # 已经是 {仓库名: 州代码}
        sku_data = sku_by_wh  # 已经是 {仓库名: set([SKU])}
        
        # 保存到数据库
        return save_warehouse_inventory(warehouse_data, sku_data)
    
    except Exception as e:
        return False, f"导入失败: {str(e)}"


def export_shipping_config_to_excel(output_file: str, config_name: str = "默认配置"):
    """
    将数据库中的发货配置导出为Excel文件
    
    Args:
        output_file: 输出Excel文件路径
        config_name: 配置名称
    
    Returns:
        (成功标志, 消息)
    """
    try:
        from excel_toolkit.excel_lite import ExcelWriter
        
        # 从数据库加载配置
        config = load_shipping_config_from_db(config_name)
        if not config:
            return False, "数据库中没有配置数据"
        
        wb = Workbook()
        
        # 子表1: 映射1
        if config.get("column_mapping_1"):
            ws1 = wb.active
            ws1.title = "子表1"
            ws1.append(["订单列名", "模板列名"])
            for order_col, template_col in config["column_mapping_1"].items():
                ws1.append([order_col, template_col])
        
        # 子表2: 映射2
        if config.get("column_mapping_2"):
            ws2 = wb.create_sheet("子表2")
            ws2.append(["订单列名", "模板列名"])
            for order_col, template_col in config["column_mapping_2"].items():
                ws2.append([order_col, template_col])
        
        # 仓库配置子表
        shipping_map = config.get("shipping_map", {})
        for wh_name, carrier_mapping in shipping_map.items():
            ws_wh = wb.create_sheet(wh_name)
            ws_wh.append(["承运商", "物流渠道"])
            for carrier, service in carrier_mapping.items():
                ws_wh.append([carrier, service])
        
        wb.save(output_file)
        return True, f"已导出Excel配置文件: {output_file}"
    
    except Exception as e:
        return False, f"导出失败: {str(e)}"

