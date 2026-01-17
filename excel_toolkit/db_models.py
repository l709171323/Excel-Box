"""
数据库模型定义
"""
from datetime import datetime
from typing import Optional
import json

try:
    from sqlalchemy import Column, Integer, String, Text, DateTime, JSON, Index
    from sqlalchemy.orm import declarative_base
    from excel_toolkit.db_config import Base
    _MODELS_AVAILABLE = True
except ImportError:
    _MODELS_AVAILABLE = False
    Base = None
    Column = Integer = String = Text = DateTime = JSON = Index = None


if _MODELS_AVAILABLE and Base:
    
    class OCRTemplate(Base):
        """OCR区域模板表（功能6）"""
        __tablename__ = 'ocr_templates'
        
        id = Column(Integer, primary_key=True, autoincrement=True)
        name = Column(String(100), nullable=False, comment='模板名称（如：USPS区域、GOFO区域）')
        region = Column(Integer, nullable=False, comment='区域编号：1=USPS, 2=Uni, 3=GOFO')
        bbox_x = Column(Integer, nullable=False, comment='X坐标')
        bbox_y = Column(Integer, nullable=False, comment='Y坐标')
        bbox_width = Column(Integer, nullable=False, comment='宽度')
        bbox_height = Column(Integer, nullable=False, comment='高度')
        description = Column(Text, nullable=True, comment='模板描述')
        created_at = Column(DateTime, default=datetime.now, comment='创建时间')
        updated_at = Column(DateTime, default=datetime.now, onupdate=datetime.now, comment='更新时间')
        
        # 索引
        __table_args__ = (
            Index('idx_name_region', 'name', 'region'),
        )
        
        def to_dict(self) -> dict:
            """转换为字典"""
            return {
                'id': self.id,
                'name': self.name,
                'region': self.region,
                'bbox': {
                    'x': self.bbox_x,
                    'y': self.bbox_y,
                    'width': self.bbox_width,
                    'height': self.bbox_height
                },
                'description': self.description,
                'created_at': self.created_at.isoformat() if self.created_at else None,
                'updated_at': self.updated_at.isoformat() if self.updated_at else None
            }
        
        def to_json_format(self) -> dict:
            """转换为JSON格式（兼容本地文件格式）"""
            return {
                'region': self.region,
                'name': self.name,
                'bbox': {
                    'x': self.bbox_x,
                    'y': self.bbox_y,
                    'width': self.bbox_width,
                    'height': self.bbox_height
                }
            }
    
    
    class ShippingConfig(Base):
        """发货配置表（功能11）"""
        __tablename__ = 'shipping_configs'
        
        id = Column(Integer, primary_key=True, autoincrement=True)
        name = Column(String(100), nullable=False, comment='配置名称')
        config_type = Column(String(20), nullable=False, comment='配置类型：mapping1, mapping2, warehouse')
        warehouse_name = Column(String(100), nullable=True, comment='仓库名称（仅warehouse类型需要）')
        
        # 列映射数据（JSON格式）
        # mapping1/mapping2: {"订单列名": "模板列名", ...}
        # warehouse: {"承运商": "物流渠道", ...}
        config_data = Column(JSON, nullable=False, comment='配置数据（JSON格式）')
        
        description = Column(Text, nullable=True, comment='配置描述')
        created_at = Column(DateTime, default=datetime.now, comment='创建时间')
        updated_at = Column(DateTime, default=datetime.now, onupdate=datetime.now, comment='更新时间')
        
        # 索引
        __table_args__ = (
            Index('idx_name_type', 'name', 'config_type'),
            Index('idx_warehouse', 'warehouse_name'),
        )
        
        def to_dict(self) -> dict:
            """转换为字典"""
            return {
                'id': self.id,
                'name': self.name,
                'config_type': self.config_type,
                'warehouse_name': self.warehouse_name,
                'config_data': self.config_data,
                'description': self.description,
                'created_at': self.created_at.isoformat() if self.created_at else None,
                'updated_at': self.updated_at.isoformat() if self.updated_at else None
            }
    
    
    class WarehouseInventory(Base):
        """仓库库存表（功能9、10）"""
        __tablename__ = 'warehouse_inventory'
        
        id = Column(Integer, primary_key=True, autoincrement=True)
        warehouse_name = Column(String(100), nullable=False, comment='仓库名称')
        state_code = Column(String(2), nullable=True, comment='州代码（2位缩写）')
        description = Column(Text, nullable=True, comment='仓库描述')
        created_at = Column(DateTime, default=datetime.now, comment='创建时间')
        updated_at = Column(DateTime, default=datetime.now, onupdate=datetime.now, comment='更新时间')
        
        # 索引
        __table_args__ = (
            Index('idx_warehouse_name', 'warehouse_name', unique=True),
            Index('idx_state_code', 'state_code'),
        )
        
        def to_dict(self) -> dict:
            return {
                'id': self.id,
                'warehouse_name': self.warehouse_name,
                'state_code': self.state_code,
                'description': self.description,
                'created_at': self.created_at.isoformat() if self.created_at else None,
                'updated_at': self.updated_at.isoformat() if self.updated_at else None
            }
    
    
    class WarehouseSKU(Base):
        """仓库SKU表（功能9、10）"""
        __tablename__ = 'warehouse_skus'
        
        id = Column(Integer, primary_key=True, autoincrement=True)
        warehouse_name = Column(String(100), nullable=False, comment='仓库名称')
        sku = Column(String(100), nullable=False, comment='SKU编号')
        created_at = Column(DateTime, default=datetime.now, comment='创建时间')
        
        # 索引
        __table_args__ = (
            Index('idx_warehouse_sku', 'warehouse_name', 'sku', unique=True),
            Index('idx_sku', 'sku'),
        )
        
        def to_dict(self) -> dict:
            return {
                'id': self.id,
                'warehouse_name': self.warehouse_name,
                'sku': self.sku,
                'created_at': self.created_at.isoformat() if self.created_at else None
            }

else:
    # 如果SQLAlchemy不可用，创建占位类
    class OCRTemplate:
        pass
    
    class ShippingConfig:
        pass
    
    class WarehouseInventory:
        pass
    
    class WarehouseSKU:
        pass

