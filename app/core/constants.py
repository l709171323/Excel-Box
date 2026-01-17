"""
UI Constants Module - Constant definitions for the Flet-based Excel Toolkit

Defines all UI-related constants including spacing, dimensions, and static values.
"""

# ==================== Spacing Configuration ====================

class Spacing:
    """Spacing values for UI elements"""
    # Outer margins
    OUTER_PADDING = 15
    SECTION_PADDING = 12
    GROUP_PADDING = 10

    # Inner padding
    CONTROL_PADDING_X = 8
    CONTROL_PADDING_Y = 6
    ROW_SPACING = 8
    SECTION_SPACING = 15

    # Button spacing
    BUTTON_PADDING_X = 10
    BUTTON_PADDING_Y = 8
    BUTTON_SPACING = 8


# ==================== Font Configuration ====================

class FontSize:
    """Font size constants"""
    TITLE = 16
    SUBTITLE = 12
    SECTION = 11
    LABEL = 10
    BUTTON = 10
    STATUS = 9
    LOG = 9
    HINT = 9


class FontFamily:
    """Font family constants"""
    DEFAULT = "Microsoft YaHei UI"
    MONOSPACE = "Consolas"


# ==================== Size Configuration ====================

class Size:
    """Component size constants"""
    # Button widths (in characters)
    BUTTON_WIDTH_SMALL = 8
    BUTTON_WIDTH_NORMAL = 12
    BUTTON_WIDTH_LARGE = 16

    # Input widths (in characters)
    ENTRY_WIDTH_SMALL = 6
    ENTRY_WIDTH_NORMAL = 15
    ENTRY_WIDTH_LARGE = 30

    # Dropdown widths
    COMBOBOX_WIDTH_SMALL = 12
    COMBOBOX_WIDTH_NORMAL = 20
    COMBOBOX_WIDTH_LARGE = 30

    # Log area height (in lines)
    LOG_HEIGHT = 10

    # Window dimensions
    WINDOW_WIDTH = 1080
    WINDOW_HEIGHT = 780
    WINDOW_MIN_WIDTH = 960
    WINDOW_MIN_HEIGHT = 720


# ==================== Icon Configuration ====================

class Icon:
    """Icon constants using string identifiers for Flet icons"""
    # Function icons (using string identifiers instead of ft.icons)
    FILE = "FOLDER_OPEN"
    EXCEL = "TABLE_CHART"
    PDF = "PICTURE_AS_PDF"
    DATABASE = "STORAGE"
    WAREHOUSE = "WAREHOUSE"
    PACKAGE = "INVENTORY_2"
    SKU = "TAG"

    # Action icons
    PLAY = "PLAY_ARROW"
    STOP = "STOP"
    REFRESH = "REFRESH"
    DELETE = "DELETE"
    ADD = "ADD"
    EDIT = "EDIT"
    SAVE = "SAVE"
    EXPORT = "DOWNLOAD"
    IMPORT = "UPLOAD"
    CLEAR = "CLEAR"

    # Status icons
    SUCCESS = "CHECK_CIRCLE"
    ERROR = "ERROR"
    WARNING = "WARNING"
    INFO = "INFO"
    LOADING = "HOURGLASS_EMPTY"

    # Navigation icons
    HOME = "HOME"
    SETTINGS = "SETTINGS"
    HELP = "HELP"
    ABOUT = "INFO_OUTLINE"
    THEME = "PALETTE"
    PIN = "PUSH_PIN"

    # Tab icons
    TAB_STATES = "MAP_OUTLINED"
    TAB_SKU = "TAG_OUTLINED"
    TAB_HIGHLIGHT = "HIGHLIGHT_OUTLINED"
    TAB_INSERT = "ADD_ROW_OUTLINED"
    TAB_COMPARE = "COMPARE_OUTLINED"
    TAB_PDF = "PICTURE_AS_PDF_OUTLINED"
    TAB_PREFIX = "TEXT_FIELDS_OUTLINED"
    TAB_FOOTER = "INSERT_PAGE_BREAK_OUTLINED"
    TAB_ROUTER = "LOCATION_ON_OUTLINED"
    TAB_ENTRY = "EDIT_NOTE_OUTLINED"
    TAB_SHIPPING = "LOCAL_SHIPPING_OUTLINED"
    TAB_PPT = "SLIDESHOW_OUTLINED"
    TAB_IMAGE = "IMAGE_OUTLINED"
    TAB_DELETE = "DELETE_OUTLINED"


# ==================== Text Constants ====================

class Text:
    """Static text constants"""
    APP_TITLE = "Excel 工具箱"
    APP_VERSION = "2.3"
    APP_AUTHOR = "果汁梨"

    # Tab names
    TAB_NAMES = [
        "[1] 州名转换",
        "[2] SKU填充",
        "[3] 高亮重复",
        "[4] 插入行",
        "[5] 对比列",
        "[6] PDF拆分",
        "[7] 前缀填充",
        "[8] 面单页脚",
        "[9] 仓库推荐",
        "[10] 录入库存",
        "[11] 模板填充",
        "[12] PPT转PDF",
        "[13] 图片压缩",
        "[14] 删除列",
    ]

    # Status messages
    STATUS_READY = "就绪"
    STATUS_PROCESSING = "处理中..."
    STATUS_COMPLETED = "完成"
    STATUS_ERROR = "错误"

    # Theme modes
    THEME_LIGHT = "浅色"
    THEME_DARK = "深色"
    THEME_SYSTEM = "系统"

    # Button labels
    BTN_RUN = "执行"
    BTN_SELECT_FILE = "选择文件"
    BTN_SELECT_DIR = "选择目录"
    BTN_CLEAR = "清空"
    BTN_SAVE = "保存"
    BTN_EXPORT = "导出"
    BTN_IMPORT = "导入"

    # Default values
    FILE_NOT_SELECTED = "未选择文件"
    DIR_NOT_SELECTED = "未选择输出目录"


# ==================== Shortcut Keys ====================

class Shortcut:
    """Keyboard shortcut constants"""
    HELP = "F1"
    QUIT = "Ctrl+Q"
    NEXT_TAB = "Ctrl+Tab"
    PREV_TAB = "Ctrl+Shift+Tab"
    OPEN_FILE = "Ctrl+O"
    RUN = "Ctrl+R"
    CLEAR_LOG = "Ctrl+L"


# ==================== Tooltip Text ====================

class Tooltip:
    """Tooltip text constants"""
    ABOUT = "查看软件版本和作者信息"
    HELP = "打开帮助文档"
    THEME = "切换主题"
    TOPMOST = "窗口置顶"
    FILE_PICKER = "点击选择文件"
    DIR_PICKER = "点击选择输出目录"
