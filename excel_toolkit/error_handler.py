"""
错误处理工具模块 - 统一的错误处理和日志记录

提供装饰器和辅助函数,简化异常处理逻辑。
"""
import os
import traceback
import functools
from typing import Callable, Optional, Any
from datetime import datetime
from excel_toolkit.exceptions import (
    ExcelToolkitError,
    FileLockedError,
    FileNotFoundError as CustomFileNotFoundError,
    InvalidColumnError,
    SheetNotFoundError
)
import builtins


# 错误日志目录
ERROR_LOG_DIR = os.path.join(os.path.expanduser("~"), ".excel_toolkit", "logs")


def ensure_log_dir():
    """确保日志目录存在"""
    try:
        os.makedirs(ERROR_LOG_DIR, exist_ok=True)
    except Exception:
        pass  # 如果创建失败,使用当前目录


def log_error(error: Exception, context: str = ""):
    """
    记录错误到日志文件
    
    Args:
        error: 异常对象
        context: 错误上下文信息
    """
    ensure_log_dir()
    
    # 生成日志文件名(按日期)
    log_file = os.path.join(
        ERROR_LOG_DIR,
        f"error_{datetime.now().strftime('%Y%m%d')}.log"
    )
    
    try:
        with open(log_file, "a", encoding="utf-8") as f:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            f.write(f"\n{'='*60}\n")
            f.write(f"时间: {timestamp}\n")
            if context:
                f.write(f"上下文: {context}\n")
            f.write(f"错误类型: {type(error).__name__}\n")
            f.write(f"错误信息: {str(error)}\n")
            f.write(f"\n堆栈跟踪:\n")
            f.write(traceback.format_exc())
            f.write(f"{'='*60}\n")
    except Exception:
        pass  # 日志记录失败不应影响主流程


def handle_file_error(file_path: str, error: Exception):
    """
    处理文件相关错误,转换为友好的自定义异常
    
    Args:
        file_path: 文件路径
        error: 原始异常
        
    Raises:
        FileLockedError, CustomFileNotFoundError, 或原异常
    """
    if isinstance(error, PermissionError):
        raise FileLockedError(os.path.basename(file_path))
    elif isinstance(error, builtins.FileNotFoundError):
        raise CustomFileNotFoundError(file_path)
    else:
        raise error


def safe_execute(func: Callable, 
                 error_logger: Optional[Callable] = None,
                 context: str = "") -> tuple[bool, Any]:
    """
    安全执行函数,捕获并记录异常
    
    Args:
        func: 要执行的函数
        error_logger: 错误日志记录函数(可选)
        context: 执行上下文描述
        
    Returns:
        (成功标志, 结果或错误信息)
    """
    try:
        result = func()
        return True, result
    except ExcelToolkitError as e:
        # 自定义异常,已经包含友好信息
        log_error(e, context)
        if error_logger:
            error_logger(e.get_user_message())
        return False, e.get_user_message()
    except Exception as e:
        # 未预期的异常
        log_error(e, context)
        error_msg = f"发生未预期的错误: {str(e)}"
        if error_logger:
            error_logger(error_msg)
        return False, error_msg


def error_handler(context: str = "", 
                 logger: Optional[Callable] = None,
                 reraise: bool = False):
    """
    装饰器: 自动处理函数异常
    
    Args:
        context: 错误上下文描述
        logger: 日志函数
        reraise: 是否重新抛出异常
        
    使用示例:
        @error_handler(context="处理SKU数据", logger=print)
        def process_sku(file):
            ...
    """
    def decorator(func):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            try:
                return func(*args, **kwargs)
            except ExcelToolkitError as e:
                # 自定义异常
                log_error(e, context or func.__name__)
                if logger:
                    logger(e.get_user_message())
                if reraise:
                    raise
                return None
            except Exception as e:
                # 未预期的异常
                log_error(e, context or func.__name__)
                error_msg = f"❌ 发生错误: {str(e)}"
                if logger:
                    logger(error_msg)
                if reraise:
                    raise
                return None
        return wrapper
    return decorator


def validate_excel_file(file_path: str) -> bool:
    """
    验证Excel文件是否可访问
    
    Args:
        file_path: 文件路径
        
    Returns:
        True 如果文件可访问
        
    Raises:
        CustomFileNotFoundError, FileLockedError
    """
    if not os.path.exists(file_path):
        raise CustomFileNotFoundError(file_path)
    
    # 检查文件是否被锁定
    try:
        with open(file_path, 'a'):
            pass
        return True
    except (IOError, PermissionError):
        raise FileLockedError(os.path.basename(file_path))


def validate_column_letter(column: str) -> bool:
    """
    验证Excel列号格式
    
    Args:
        column: 列号(如 A, B, AA)
        
    Returns:
        True 如果有效
        
    Raises:
        InvalidColumnError
    """
    if not column or not isinstance(column, str):
        raise InvalidColumnError(str(column))
    
    if not column.strip().isalpha():
        raise InvalidColumnError(column)
    
    return True


def get_user_friendly_error(error: Exception) -> str:
    """
    将异常转换为用户友好的错误信息
    
    Args:
        error: 异常对象
        
    Returns:
        友好的错误信息
    """
    if isinstance(error, ExcelToolkitError):
        return error.get_user_message()
    elif isinstance(error, PermissionError):
        return FileLockedError("文件").get_user_message()
    elif isinstance(error, builtins.FileNotFoundError):
        return "❌ 文件未找到\n💡 请检查文件路径是否正确"
    elif isinstance(error, ValueError):
        return f"❌ 数据格式错误\n📋 {str(error)}\n💡 请检查输入数据格式"
    else:
        return f"❌ 发生错误: {str(error)}\n💡 请查看日志文件获取详细信息"


def create_error_report(errors: list[tuple[str, Exception]]) -> str:
    """
    创建错误汇总报告
    
    Args:
        errors: [(操作名称, 异常对象), ...]
        
    Returns:
        格式化的错误报告
    """
    if not errors:
        return "✅ 所有操作成功完成"
    
    report = [f"⚠️ 处理过程中发生 {len(errors)} 个错误:\n"]
    
    for i, (operation, error) in enumerate(errors, 1):
        report.append(f"{i}. {operation}")
        if isinstance(error, ExcelToolkitError):
            report.append(f"   {error.message}")
        else:
            report.append(f"   {str(error)}")
    
    report.append("\n💡 详细错误信息已记录到日志文件")
    report.append(f"   日志路径: {ERROR_LOG_DIR}")
    
    return "\n".join(report)
