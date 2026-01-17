"""
UI 配置模块 - 统一的界面样式配置

定义所有界面元素的样式、间距、颜色等配置
"""

# ==================== 间距配置 ====================

SPACING = {
    # 外边距
    'outer_padding': 15,           # 主容器外边距
    'section_padding': 12,         # 区块内边距
    'group_padding': 10,           # 分组内边距
    
    # 内边距
    'control_padding_x': 8,        # 控件水平间距
    'control_padding_y': 6,        # 控件垂直间距
    'row_spacing': 8,              # 行间距
    'section_spacing': 15,         # 区块间距
    
    # 按钮间距
    'button_padding_x': 10,        # 按钮水平间距
    'button_padding_y': 8,         # 按钮垂直间距
    'button_spacing': 8,           # 按钮之间间距
}

# ==================== 字体配置 ====================

FONTS = {
    'title': ("Microsoft YaHei UI", 16, "bold"),      # 主标题
    'subtitle': ("Microsoft YaHei UI", 12, "bold"),   # 副标题
    'section': ("Microsoft YaHei UI", 11, "bold"),    # 区块标题
    'label': ("Microsoft YaHei UI", 10),              # 普通标签
    'button': ("Microsoft YaHei UI", 10),             # 按钮文字
    'status': ("Microsoft YaHei UI", 9),              # 状态栏
    'log': ("Consolas", 9),                           # 日志字体(等宽)
    'hint': ("Microsoft YaHei UI", 9),                # 提示文字
}

# ==================== 颜色配置 ====================

COLORS = {
    # 主色调
    'primary': '#3B82F6',          # 主要蓝色
    'primary_hover': '#2563EB',    # 主要蓝色(悬停)
    'primary_light': '#DBEAFE',    # 浅蓝色背景
    
    # 辅助色
    'secondary': '#6B7280',        # 次要灰色
    'success': '#10B981',          # 成功绿色
    'warning': '#F59E0B',          # 警告橙色
    'error': '#EF4444',            # 错误红色
    'info': '#3B82F6',             # 信息蓝色
    
    # 背景色
    'bg_light': '#F9FAFB',         # 浅色背景
    'bg_card': '#FFFFFF',          # 卡片背景
    'bg_hover': '#F3F4F6',         # 悬停背景
    
    # 文字色
    'text_primary': '#111827',     # 主要文字
    'text_secondary': '#6B7280',   # 次要文字
    'text_hint': '#9CA3AF',        # 提示文字
    
    # 边框色
    'border_light': '#E5E7EB',     # 浅边框
    'border_normal': '#D1D5DB',    # 普通边框
    'border_dark': '#9CA3AF',      # 深边框
}

# ==================== 组件尺寸配置 ====================

SIZES = {
    # 按钮尺寸
    'button_width_small': 8,       # 小按钮宽度
    'button_width_normal': 12,     # 普通按钮宽度
    'button_width_large': 16,      # 大按钮宽度
    
    # 输入框尺寸
    'entry_width_small': 6,        # 小输入框(列号)
    'entry_width_normal': 15,      # 普通输入框
    'entry_width_large': 30,       # 大输入框
    
    # 下拉框尺寸
    'combobox_width_small': 12,    # 小下拉框
    'combobox_width_normal': 20,   # 普通下拉框
    'combobox_width_large': 30,    # 大下拉框
    
    # 日志区域
    'log_height': 10,              # 日志区域高度(行数)
    
    # 标签宽度
    'label_width_small': 8,        # 小标签宽度
    'label_width_normal': 12,      # 普通标签宽度
    'label_width_large': 15,       # 大标签宽度
}

# ==================== 图标配置 ====================

ICONS = {
    # 功能图标
    'file': '📁',
    'folder': '📂',
    'excel': '📊',
    'pdf': '📄',
    'database': '🗄️',
    'warehouse': '🏭',
    'package': '📦',
    'sku': '🏷️',
    
    # 操作图标
    'play': '▶️',
    'stop': '⏹️',
    'refresh': '🔄',
    'delete': '🗑️',
    'add': '➕',
    'edit': '✏️',
    'save': '💾',
    'export': '📤',
    'import': '📥',
    
    # 状态图标
    'success': '✅',
    'error': '❌',
    'warning': '⚠️',
    'info': 'ℹ️',
    'loading': '⏳',
    
    # 其他图标
    'search': '🔍',
    'settings': '⚙️',
    'help': '❓',
    'about': 'ℹ️',
    'theme': '🎨',
    'pin': '📌',
    'log': '📝',
    'clear': '🧹',
}

# ==================== 布局模板 ====================

LAYOUT_TEMPLATES = {
    # 标准表单行布局
    'form_row': {
        'fill': 'x',
        'padx': SPACING['control_padding_x'],
        'pady': SPACING['row_spacing'] // 2,
    },
    
    # 按钮组布局
    'button_group': {
        'fill': 'x',
        'padx': SPACING['control_padding_x'],
        'pady': SPACING['section_spacing'],
    },
    
    # 区块容器布局
    'section_frame': {
        'fill': 'x',
        'padx': SPACING['section_padding'],
        'pady': SPACING['section_spacing'],
    },
    
    # 日志容器布局
    'log_frame': {
        'fill': 'both',
        'expand': True,
        'padx': SPACING['section_padding'],
        'pady': SPACING['section_spacing'],
    },
}

# ==================== 辅助函数 ====================

def get_card_style():
    """获取卡片样式配置"""
    return {
        'relief': 'flat',
        'borderwidth': 1,
        'background': COLORS['bg_card'],
    }

def get_button_padding():
    """获取按钮内边距"""
    return (SPACING['button_padding_x'], SPACING['button_padding_y'])

def get_section_padding():
    """获取区块内边距"""
    return SPACING['section_padding']

def apply_tooltip_style(widget, text):
    """应用统一的提示样式"""
    from excel_toolkit.tooltip import create_tooltip
    create_tooltip(widget, text)
