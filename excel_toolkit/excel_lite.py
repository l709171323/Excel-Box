#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
轻量级Excel处理模块
替代openpyxl，避免numpy依赖，减小打包体积
"""
import os
import zipfile
import xml.etree.ElementTree as ET
from typing import List, Dict, Any, Optional, Union, Tuple
import xlrd
import xlsxwriter
from defusedxml import ElementTree as SafeET


class ExcelLiteError(Exception):
    """Excel处理异常"""
    pass


def column_index_from_string(column_str: str) -> int:
    """将列字母转换为索引（A=1, B=2, ...）"""
    result = 0
    for char in column_str.upper():
        result = result * 26 + (ord(char) - ord('A') + 1)
    return result


def get_column_letter(col_idx: int) -> str:
    """将列索引转换为字母（1=A, 2=B, ...）"""
    result = ""
    while col_idx > 0:
        col_idx -= 1
        result = chr(col_idx % 26 + ord('A')) + result
        col_idx //= 26
    return result


class ExcelWorksheet:
    """轻量级工作表类"""
    
    def __init__(self, reader, sheet_name: str):
        self.reader = reader
        self.title = sheet_name
        self._data = None
        self._max_row = None
        self._max_col = None
    
    @property
    def rows(self):
        """返回所有行的迭代器（兼容openpyxl）"""
        self._load_data()
        for row_data in self._data:
            yield [ExcelCell(value) for value in row_data]
    
    @property
    def max_row(self) -> int:
        """获取最大行数"""
        if self._max_row is None:
            self._load_data()
        return self._max_row or 0
    
    @property
    def max_column(self) -> int:
        """获取最大列数"""
        if self._max_col is None:
            self._load_data()
        return self._max_col or 0
    
    def _load_data(self):
        """加载工作表数据"""
        if self._data is None:
            self._data = self.reader.get_sheet_data(self.title)
            self._max_row = len(self._data)
            self._max_col = max(len(row) for row in self._data) if self._data else 0
    
    def __getitem__(self, row_index):
        """支持 worksheet[row_number] 语法访问行"""
        self._load_data()
        
        # 转换为0基索引
        if isinstance(row_index, int):
            if row_index < 1:
                raise IndexError("行索引必须从1开始")
            
            row_idx = row_index - 1
            if row_idx >= len(self._data):
                # 返回空行
                return [ExcelCell('') for _ in range(self._max_col or 0)]
            
            row_data = self._data[row_idx]
            # 确保行数据长度与最大列数一致
            while len(row_data) < (self._max_col or 0):
                row_data.append('')
            
            return [ExcelCell(value) for value in row_data]
        else:
            raise TypeError("行索引必须是整数")
    
    def iter_rows(self, min_row=1, max_row=None, min_col=1, max_col=None, values_only=False):
        """迭代行数据"""
        self._load_data()
        
        if not self._data:
            return  # 空数据，直接返回
        
        start_row = (min_row - 1) if min_row else 0
        end_row = (max_row - 1) if max_row else len(self._data) - 1
        start_col = (min_col - 1) if min_col else 0
        
        # 计算最大列数
        if max_col:
            end_col = max_col - 1
        else:
            # 找出所有行中的最大列数
            max_cols_in_data = max(len(row) for row in self._data) if self._data else 0
            end_col = max_cols_in_data - 1 if max_cols_in_data > 0 else 0
        
        for row_idx in range(start_row, min(end_row + 1, len(self._data))):
            row_data = self._data[row_idx]
            if values_only:
                # 返回值的元组
                row_values = []
                for col_idx in range(start_col, end_col + 1):
                    row_values.append(row_data[col_idx] if col_idx < len(row_data) else '')
                yield tuple(row_values)
            else:
                # 返回单元格对象列表
                cells = []
                for col_idx in range(start_col, end_col + 1):
                    value = row_data[col_idx] if col_idx < len(row_data) else ''
                    cells.append(ExcelCell(value))
                yield cells
    
    def cell(self, row: int, column: int):
        """获取单元格"""
        self._load_data()
        if 1 <= row <= len(self._data) and 1 <= column <= len(self._data[row-1]):
            return ExcelCell(self._data[row-1][column-1])
        return ExcelCell('')


class ExcelCell:
    """轻量级单元格类"""
    
    def __init__(self, value):
        self.value = value


class ExcelReader:
    """轻量级Excel读取器"""
    
    def __init__(self, file_path: str, read_only: bool = False, data_only: bool = False):
        self.file_path = file_path
        self.file_ext = os.path.splitext(file_path)[1].lower()
        self._workbook = None
        self._sheet_names = None
        self._worksheets = None
        # 兼容参数，暂时不使用
        self.read_only = read_only
        self.data_only = data_only
        
    def __enter__(self):
        return self
        
    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()
    
    def __getitem__(self, sheet_name: str):
        """支持 wb[sheet_name] 语法"""
        if sheet_name not in self.sheetnames:
            raise KeyError(f"工作表 '{sheet_name}' 不存在")
        return ExcelWorksheet(self, sheet_name)
    
    @property
    def worksheets(self):
        """获取所有工作表对象列表"""
        if self._worksheets is None:
            self._worksheets = [ExcelWorksheet(self, name) for name in self.sheetnames]
        return self._worksheets
    
    def close(self):
        """关闭工作簿"""
        if self._workbook and hasattr(self._workbook, 'release_resources'):
            self._workbook.release_resources()
        self._workbook = None
    
    @property
    def sheetnames(self) -> List[str]:
        """获取所有工作表名称"""
        if self._sheet_names is not None:
            return self._sheet_names
            
        if self.file_ext == '.xls':
            # 使用xlrd读取旧版Excel
            try:
                wb = xlrd.open_workbook(self.file_path)
                self._sheet_names = wb.sheet_names()
                return self._sheet_names
            except Exception as e:
                raise ExcelLiteError(f"无法读取.xls文件: {e}")
                
        elif self.file_ext in ['.xlsx', '.xlsm']:
            # 使用XML解析读取新版Excel
            try:
                self._sheet_names = self._get_xlsx_sheet_names()
                return self._sheet_names
            except Exception as e:
                raise ExcelLiteError(f"无法读取.xlsx文件: {e}")
        else:
            raise ExcelLiteError(f"不支持的文件格式: {self.file_ext}")
    
    def _get_xlsx_sheet_names(self) -> List[str]:
        """从xlsx文件中提取工作表名称"""
        try:
            with zipfile.ZipFile(self.file_path, 'r') as zip_file:
                # 读取workbook.xml获取工作表信息
                workbook_xml = zip_file.read('xl/workbook.xml')
                root = SafeET.fromstring(workbook_xml)
                
                # 查找所有工作表
                sheets = []
                for sheet in root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet'):
                    name = sheet.get('name')
                    if name:
                        sheets.append(name)
                
                return sheets
        except Exception as e:
            raise ExcelLiteError(f"解析xlsx文件失败: {e}")
    
    def get_sheet_data(self, sheet_name: str, max_rows: int = None) -> List[List[Any]]:
        """获取工作表数据"""
        if self.file_ext == '.xls':
            return self._get_xls_sheet_data(sheet_name, max_rows)
        elif self.file_ext in ['.xlsx', '.xlsm']:
            return self._get_xlsx_sheet_data(sheet_name, max_rows)
        else:
            raise ExcelLiteError(f"不支持的文件格式: {self.file_ext}")
    
    def _get_xls_sheet_data(self, sheet_name: str, max_rows: int = None) -> List[List[Any]]:
        """从xls文件读取数据"""
        try:
            wb = xlrd.open_workbook(self.file_path)
            if sheet_name not in wb.sheet_names():
                raise ExcelLiteError(f"工作表 '{sheet_name}' 不存在")
            
            sheet = wb.sheet_by_name(sheet_name)
            data = []
            
            end_row = min(sheet.nrows, max_rows) if max_rows else sheet.nrows
            
            for row_idx in range(end_row):
                row_data = []
                for col_idx in range(sheet.ncols):
                    cell = sheet.cell(row_idx, col_idx)
                    value = cell.value
                    
                    # 处理日期类型
                    if cell.ctype == xlrd.XL_CELL_DATE:
                        try:
                            import datetime
                            date_tuple = xlrd.xldate_as_tuple(value, wb.datemode)
                            value = datetime.datetime(*date_tuple)
                        except:
                            pass
                    
                    row_data.append(value)
                data.append(row_data)
            
            return data
        except Exception as e:
            raise ExcelLiteError(f"读取xls工作表失败: {e}")
    
    def _get_xlsx_sheet_data(self, sheet_name: str, max_rows: int = None) -> List[List[Any]]:
        """从xlsx文件读取数据（简化版本，只读取值）"""
        try:
            with zipfile.ZipFile(self.file_path, 'r') as zip_file:
                # 获取工作表ID
                sheet_id = self._get_sheet_id(zip_file, sheet_name)
                if not sheet_id:
                    raise ExcelLiteError(f"工作表 '{sheet_name}' 不存在")
                
                # 读取工作表数据
                sheet_xml_path = f'xl/worksheets/sheet{sheet_id}.xml'
                sheet_xml = zip_file.read(sheet_xml_path)
                
                # 读取共享字符串
                shared_strings = self._get_shared_strings(zip_file)
                
                # 解析工作表数据
                return self._parse_sheet_xml(sheet_xml, shared_strings, max_rows)
                
        except Exception as e:
            raise ExcelLiteError(f"读取xlsx工作表失败: {e}")
    
    def _get_sheet_id(self, zip_file: zipfile.ZipFile, sheet_name: str) -> Optional[str]:
        """获取工作表ID（基于在workbook.xml中的顺序）"""
        workbook_xml = zip_file.read('xl/workbook.xml')
        root = SafeET.fromstring(workbook_xml)
        
        sheet_index = 1  # 工作表文件从sheet1.xml开始
        for sheet in root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet'):
            if sheet.get('name') == sheet_name:
                return str(sheet_index)
            sheet_index += 1
        return None
    
    def _get_shared_strings(self, zip_file: zipfile.ZipFile) -> List[str]:
        """获取共享字符串表"""
        try:
            shared_strings_xml = zip_file.read('xl/sharedStrings.xml')
            root = SafeET.fromstring(shared_strings_xml)
            
            ns = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
            strings = []
            
            for si in root.findall(f'.//{ns}si'):
                # 尝试直接获取 <t> 元素
                text_elem = si.find(f'{ns}t')
                if text_elem is not None and text_elem.text:
                    strings.append(text_elem.text)
                else:
                    # 处理富文本格式：<si><r><t>text1</t></r><r><t>text2</t></r></si>
                    text_parts = []
                    for r_elem in si.findall(f'{ns}r'):
                        t_elem = r_elem.find(f'{ns}t')
                        if t_elem is not None and t_elem.text:
                            text_parts.append(t_elem.text)
                    
                    if text_parts:
                        strings.append(''.join(text_parts))
                    else:
                        # 最后尝试获取所有 <t> 元素
                        all_t = si.findall(f'.//{ns}t')
                        if all_t:
                            strings.append(''.join(t.text or '' for t in all_t))
                        else:
                            strings.append('')
            
            return strings
        except:
            return []
    
    def _parse_sheet_xml(self, sheet_xml: bytes, shared_strings: List[str], max_rows: int = None) -> List[List[Any]]:
        """解析工作表XML数据"""
        root = SafeET.fromstring(sheet_xml)
        data = []
        
        rows = root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row')
        
        for row_idx, row in enumerate(rows):
            if max_rows and row_idx >= max_rows:
                break
                
            row_data = []
            cells = row.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c')
            
            # 获取当前行的最大列数
            max_col = 0
            cell_dict = {}
            
            for cell in cells:
                cell_ref = cell.get('r', '')
                if cell_ref:
                    # 解析单元格引用（如A1, B2）
                    col_str = ''.join(c for c in cell_ref if c.isalpha())
                    col_idx = column_index_from_string(col_str) - 1
                    max_col = max(max_col, col_idx)
                    
                    # 获取单元格值
                    value = self._get_cell_value(cell, shared_strings)
                    cell_dict[col_idx] = value
            
            # 构建行数据
            for col_idx in range(max_col + 1):
                row_data.append(cell_dict.get(col_idx, ''))
            
            data.append(row_data)
        
        return data
    
    def _get_cell_value(self, cell_elem, shared_strings: List[str]) -> Any:
        """获取单元格值"""
        cell_type = cell_elem.get('t', '')
        value_elem = cell_elem.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
        
        if value_elem is None:
            return ''
        
        value = value_elem.text or ''
        
        if cell_type == 's':  # 共享字符串
            try:
                idx = int(value)
                if 0 <= idx < len(shared_strings):
                    return shared_strings[idx]
            except (ValueError, IndexError):
                pass
        elif cell_type == 'b':  # 布尔值
            return value == '1'
        elif cell_type in ('n', ''):  # 数字
            try:
                if '.' in value:
                    return float(value)
                else:
                    return int(value)
            except ValueError:
                pass
        
        return value


class ExcelWriter:
    """轻量级Excel写入器"""
    
    def __init__(self, file_path: str):
        self.file_path = file_path
        self.workbook = xlsxwriter.Workbook(file_path)
        self.worksheets = {}
    
    def __enter__(self):
        return self
        
    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()
    
    def close(self):
        """关闭工作簿"""
        if self.workbook:
            self.workbook.close()
            self.workbook = None
    
    def create_sheet(self, sheet_name: str):
        """创建工作表"""
        if sheet_name in self.worksheets:
            return self.worksheets[sheet_name]
        
        worksheet = self.workbook.add_worksheet(sheet_name)
        self.worksheets[sheet_name] = worksheet
        return worksheet
    
    def write_data(self, sheet_name: str, data: List[List[Any]], start_row: int = 0, start_col: int = 0):
        """写入数据到工作表"""
        worksheet = self.create_sheet(sheet_name)
        
        for row_idx, row_data in enumerate(data):
            for col_idx, value in enumerate(row_data):
                worksheet.write(start_row + row_idx, start_col + col_idx, value)
    
    def set_cell_value(self, sheet_name: str, row: int, col: int, value: Any):
        """设置单元格值"""
        worksheet = self.create_sheet(sheet_name)
        worksheet.write(row, col, value)
    
    def set_cell_color(self, sheet_name: str, row: int, col: int, color: str):
        """设置单元格背景色"""
        worksheet = self.create_sheet(sheet_name)
        format_obj = self.workbook.add_format({'bg_color': color})
        worksheet.write(row, col, '', format_obj)


# 兼容性函数，保持与原openpyxl代码的接口一致
def load_workbook(filename: str, read_only: bool = False, data_only: bool = False):
    """加载工作簿（兼容openpyxl接口）"""
    return ExcelReader(filename)


# 工具函数
def get_sheet_names(file_path: str) -> Optional[List[str]]:
    """获取Excel文件的所有工作表名称"""
    try:
        with ExcelReader(file_path) as reader:
            return reader.sheetnames
    except Exception as e:
        print(f"读取文件失败: {e}")
        return None