"""
自定义异常类 - 统一的异常体系

定义了应用中所有可能的异常类型,便于精确捕获和处理。
"""


class ExcelToolkitError(Exception):
    """Excel工具箱基础异常类"""
    
    def __init__(self, message: str, details: str = None, solution: str = None):
        """
        Args:
            message: 错误简短描述
            details: 错误详细信息
            solution: 建议的解决方案
        """
        self.message = message
        self.details = details
        self.solution = solution
        super().__init__(self.message)
    
    def get_user_message(self) -> str:
        """获取用户友好的完整错误信息"""
        parts = [f"❌ {self.message}"]
        if self.details:
            parts.append(f"\n📋 详细信息: {self.details}")
        if self.solution:
            parts.append(f"\n💡 解决方案: {self.solution}")
        return "\n".join(parts)


class FileAccessError(ExcelToolkitError):
    """文件访问相关错误"""
    pass


class FileLockedError(FileAccessError):
    """文件被锁定(被Excel/WPS占用)"""
    
    def __init__(self, file_name: str):
        super().__init__(
            message=f"文件被占用: {file_name}",
            details="文件可能正在被 Excel、WPS 或其他程序打开",
            solution="请关闭所有打开此文件的程序,然后重试"
        )


class FileNotFoundError(FileAccessError):
    """文件不存在"""
    
    def __init__(self, file_path: str):
        super().__init__(
            message=f"文件未找到",
            details=f"路径: {file_path}",
            solution="请检查文件路径是否正确,或文件是否已被移动/删除"
        )


class DataValidationError(ExcelToolkitError):
    """数据验证错误"""
    pass


class InvalidColumnError(DataValidationError):
    """无效的列号"""
    
    def __init__(self, column: str):
        super().__init__(
            message=f"列号无效: {column}",
            details="Excel列号应为字母形式,如 A、B、AA、AB 等",
            solution="请输入有效的Excel列号(只包含字母)"
        )


class SheetNotFoundError(DataValidationError):
    """工作表不存在"""
    
    def __init__(self, sheet_name: str, available_sheets: list = None):
        details = f"工作表 '{sheet_name}' 不存在"
        if available_sheets:
            details += f"\n可用的工作表: {', '.join(available_sheets)}"
        
        super().__init__(
            message="工作表不存在",
            details=details,
            solution="请从下拉列表中选择正确的工作表名称"
        )


class EmptyDataError(DataValidationError):
    """数据为空"""
    
    def __init__(self, context: str = ""):
        super().__init__(
            message="数据为空",
            details=f"{context}没有找到任何有效数据" if context else "没有找到任何有效数据",
            solution="请检查文件内容是否正确,或选择的列/工作表是否包含数据"
        )


class DatabaseError(ExcelToolkitError):
    """数据库相关错误"""
    pass


class SKUNotFoundError(DataValidationError):
    """SKU未找到"""
    
    def __init__(self, sku: str):
        super().__init__(
            message=f"SKU未找到: {sku}",
            details="商品资料库中不存在此SKU",
            solution="请检查SKU是否正确,或更新商品资料库"
        )


class OCRError(ExcelToolkitError):
    """OCR识别错误"""
    pass


class DependencyMissingError(ExcelToolkitError):
    """缺少依赖"""
    
    def __init__(self, dependency: str, install_hint: str = None):
        solution = f"请安装 {dependency}"
        if install_hint:
            solution += f"\n安装方法: {install_hint}"
        
        super().__init__(
            message=f"缺少必需的依赖: {dependency}",
            details=f"此功能需要 {dependency} 支持",
            solution=solution
        )


class ConfigurationError(ExcelToolkitError):
    """配置错误"""
    pass


class InvalidMappingError(ConfigurationError):
    """映射配置错误"""
    
    def __init__(self, mapping_name: str, reason: str):
        super().__init__(
            message=f"映射配置错误: {mapping_name}",
            details=reason,
            solution="请检查配置文件或重新设置映射关系"
        )
