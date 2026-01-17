"""
数据库配置和连接模块
支持MySQL、PostgreSQL、SQLite
"""
import os
import json
from typing import Optional, Dict, Any
from pathlib import Path

# 尝试导入数据库驱动
_DB_DRIVERS = {}
try:
    import sqlite3
    _DB_DRIVERS['sqlite'] = True
except ImportError:
    _DB_DRIVERS['sqlite'] = False

try:
    import pymysql
    _DB_DRIVERS['mysql'] = True
except ImportError:
    _DB_DRIVERS['mysql'] = False

try:
    import psycopg2
    _DB_DRIVERS['postgresql'] = True
except ImportError:
    _DB_DRIVERS['postgresql'] = False

try:
    from sqlalchemy import create_engine, text, inspect
    from sqlalchemy.orm import sessionmaker, declarative_base
    from sqlalchemy.pool import QueuePool
    _SQLALCHEMY_AVAILABLE = True
except ImportError:
    _SQLALCHEMY_AVAILABLE = False
    create_engine = text = inspect = sessionmaker = declarative_base = QueuePool = None

Base = declarative_base() if _SQLALCHEMY_AVAILABLE else None


class DatabaseConfig:
    """数据库配置类"""
    
    def __init__(self):
        self.config_file = os.path.join(os.path.dirname(__file__), "..", "db_config.json")
        self._config = self._load_config()
    
    def _load_config(self) -> Dict[str, Any]:
        """加载数据库配置"""
        default_config = {
            "enabled": False,
            "type": "sqlite",  # sqlite, mysql, postgresql
            "host": "localhost",
            "port": 3306,
            "database": "excel_toolkit",
            "username": "",
            "password": "",
            "sqlite_path": "excel_toolkit.db",  # SQLite文件路径
            "charset": "utf8mb4"
        }
        
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    user_config = json.load(f)
                default_config.update(user_config)
            except Exception:
                pass
        
        return default_config
    
    def save_config(self, config: Dict[str, Any]):
        """保存数据库配置"""
        self._config.update(config)
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self._config, f, indent=2, ensure_ascii=False)
            return True
        except Exception as e:
            print(f"保存数据库配置失败: {e}")
            return False
    
    def get_config(self) -> Dict[str, Any]:
        """获取当前配置"""
        return self._config.copy()
    
    def is_enabled(self) -> bool:
        """检查数据库是否启用"""
        return self._config.get("enabled", False)
    
    def get_connection_string(self) -> Optional[str]:
        """获取数据库连接字符串"""
        if not self.is_enabled():
            return None
        
        db_type = self._config.get("type", "sqlite")
        
        if db_type == "sqlite":
            sqlite_path = self._config.get("sqlite_path", "excel_toolkit.db")
            # 如果是相对路径，转换为绝对路径
            if not os.path.isabs(sqlite_path):
                sqlite_path = os.path.join(os.path.dirname(self.config_file), sqlite_path)
            return f"sqlite:///{sqlite_path}"
        
        elif db_type == "mysql":
            host = self._config.get("host", "localhost")
            port = self._config.get("port", 3306)
            database = self._config.get("database", "excel_toolkit")
            username = self._config.get("username", "")
            password = self._config.get("password", "")
            charset = self._config.get("charset", "utf8mb4")
            return f"mysql+pymysql://{username}:{password}@{host}:{port}/{database}?charset={charset}"
        
        elif db_type == "postgresql":
            host = self._config.get("host", "localhost")
            port = self._config.get("port", 5432)
            database = self._config.get("database", "excel_toolkit")
            username = self._config.get("username", "")
            password = self._config.get("password", "")
            return f"postgresql+psycopg2://{username}:{password}@{host}:{port}/{database}"
        
        return None


class DatabaseManager:
    """数据库管理器"""
    
    def __init__(self):
        self.config = DatabaseConfig()
        self.engine = None
        self.Session = None
        self._connected = False
    
    def connect(self):
        """
        连接数据库
        返回: (成功标志, 错误信息)
        """
        if not self.config.is_enabled():
            return False, "数据库未启用"
        
        if not _SQLALCHEMY_AVAILABLE:
            return False, "SQLAlchemy未安装，请运行: pip install sqlalchemy"
        
        db_type = self.config.get_config().get("type", "sqlite")
        
        # 检查驱动
        if db_type == "mysql" and not _DB_DRIVERS.get("mysql"):
            return False, "MySQL驱动未安装，请运行: pip install pymysql"
        if db_type == "postgresql" and not _DB_DRIVERS.get("postgresql"):
            return False, "PostgreSQL驱动未安装，请运行: pip install psycopg2-binary"
        
        try:
            conn_str = self.config.get_connection_string()
            if not conn_str:
                return False, "无法生成连接字符串"
            
            # 创建引擎
            self.engine = create_engine(
                conn_str,
                poolclass=QueuePool,
                pool_size=5,
                max_overflow=10,
                pool_pre_ping=True,  # 自动重连
                echo=False
            )
            
            # 测试连接
            with self.engine.connect() as conn:
                conn.execute(text("SELECT 1"))
            
            # 创建Session类
            self.Session = sessionmaker(bind=self.engine)
            
            # 初始化表结构
            self._init_tables()
            
            self._connected = True
            return True, "连接成功"
        
        except Exception as e:
            self._connected = False
            return False, f"连接失败: {str(e)}"
    
    def _init_tables(self):
        """初始化数据库表结构"""
        if not self.engine:
            return
        
        try:
            # 创建所有表
            Base.metadata.create_all(self.engine)
        except Exception as e:
            print(f"初始化表结构失败: {e}")
    
    def disconnect(self):
        """断开数据库连接"""
        if self.engine:
            self.engine.dispose()
            self.engine = None
            self.Session = None
            self._connected = False
    
    def is_connected(self) -> bool:
        """检查是否已连接"""
        return self._connected
    
    def get_session(self):
        """获取数据库会话"""
        if not self.is_connected():
            success, msg = self.connect()
            if not success:
                raise Exception(f"无法连接数据库: {msg}")
        return self.Session()
    
    def test_connection(self):
        """测试数据库连接"""
        return self.connect()


# 全局数据库管理器实例
_db_manager = None

def get_db_manager() -> DatabaseManager:
    """获取全局数据库管理器实例"""
    global _db_manager
    if _db_manager is None:
        _db_manager = DatabaseManager()
    return _db_manager

