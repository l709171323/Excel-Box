"""
Excel 工具箱 - UI 模块

此目录包含所有Tab页面的UI定义，采用Mixin模式实现。
每个Tab模块包含一个Mixin类，提供create_tabX_xxx和run_toolX方法。

结构说明：
- mixins.py     - 共享的日志和文件选择功能
- tab1_states.py  - [1] 转换州名
- tab2_skus.py    - [2] 填充SKU信息
- tab3_highlight.py - [3] 高亮重复项
- tab4_insert.py  - [4] 插入缺失行
- tab5_compare.py - [5] 对比列数据
- tab6_pdf.py     - [6] 拆分订单PDF
- tab7_prefix.py  - [7] 前缀填充承运商
- tab9_router.py  - [9] 建议发货仓库
- tab10_entry.py  - [10] 录入发货信息

使用方式：
在主应用类中继承所有Mixin类即可获得完整功能。

示例：
    class ToolkitApp(
        LoggerMixin,
        FileSelectMixin,
        Tab1StatesMixin,
        Tab2SkusMixin,
        ...
    ):
        pass
"""

from excel_toolkit.ui.mixins import LoggerMixin, FileSelectMixin, get_sheet_names
from excel_toolkit.ui.tab1_states import Tab1StatesMixin
from excel_toolkit.ui.tab2_skus import Tab2SkusMixin
from excel_toolkit.ui.tab3_highlight import Tab3HighlightMixin
from excel_toolkit.ui.tab4_insert import Tab4InsertMixin
from excel_toolkit.ui.tab5_compare import Tab5CompareMixin
from excel_toolkit.ui.tab6_pdf import Tab6PdfMixin
from excel_toolkit.ui.tab7_prefix import Tab7PrefixMixin
from excel_toolkit.ui.tab8_pdf_footer import Tab8PdfFooterMixin
from excel_toolkit.ui.tab9_router import Tab9RouterMixin
from excel_toolkit.ui.tab10_entry import Tab10EntryMixin
from excel_toolkit.ui.tab11_shipping import Tab11ShippingMixin
from excel_toolkit.ui.tab12_ppt import Tab12PptMixin
from excel_toolkit.ui.tab13_image_compress import Tab13ImageCompressMixin
from excel_toolkit.ui.tab14_delete_cols import Tab14DeleteColsMixin

__all__ = [
    'LoggerMixin',
    'FileSelectMixin',
    'get_sheet_names',
    'Tab1StatesMixin',
    'Tab2SkusMixin',
    'Tab3HighlightMixin',
    'Tab4InsertMixin',
    'Tab5CompareMixin',
    'Tab6PdfMixin',
    'Tab7PrefixMixin',
    'Tab8PdfFooterMixin',
    'Tab9RouterMixin',
    'Tab10EntryMixin',
    'Tab11ShippingMixin',
    'Tab12PptMixin',
    'Tab13ImageCompressMixin',
    'Tab14DeleteColsMixin',
]
