from excel_toolkit.excel_lite import ExcelReader, ExcelWriter
from excel_toolkit.excel_lite import column_index_from_string


def process_prefix_fill(file_name, src_col_letter, dst_col_letter, logger=print):
    try:
        wb = ExcelReader(file_name)
    except Exception as e:
        return f"加载文件失败: {e}"

    try:
        src_idx = column_index_from_string(str(src_col_letter).strip())
        dst_idx = column_index_from_string(str(dst_col_letter).strip())
    except Exception:
        wb.close(); return f"错误：列号无效。请检查源列 '{src_col_letter}' 与目标列 '{dst_col_letter}' 是否为有效Excel列号。"

    total_sheets_processed = 0
    total_cells_filled = 0
    logger(f"开始遍历所有工作表，读取 {src_col_letter} 列，根据前缀填充 {dst_col_letter} 列...")
    for ws in wb.worksheets:
        if ws.max_row <= 1:
            continue
        logger(f"  > 正在处理工作表: {ws.title}")
        sheet_count = 0
        for r in range(2, ws.max_row + 1):
            src_val = ws.cell(row=r, column=src_idx).value
            if src_val is None or src_val == "":
                continue
            s = str(src_val).strip()
            if not s:
                continue
            first = s[0]
            if first.isalpha():
                first = first.upper()
            if first == '9':
                ws.cell(row=r, column=dst_idx).value = 'usps'
                sheet_count += 1
            elif first == 'G':
                ws.cell(row=r, column=dst_idx).value = 'GOFO'
                sheet_count += 1
            elif first == 'U':
                ws.cell(row=r, column=dst_idx).value = 'UniUni'
                sheet_count += 1
        if sheet_count > 0:
            total_sheets_processed += 1
            total_cells_filled += sheet_count

    try:
        wb.save(file_name)
        wb.close()
        if total_sheets_processed > 0:
            return (
                "前缀填充完成！\n\n"
                f"总共在 {total_sheets_processed} 个工作表中填充了 {total_cells_filled} 个单元格。\n"
                f"源列: {src_col_letter}，目标列: {dst_col_letter}"
            )
        else:
            return (
                "前缀填充完成。\n"
                f"未在任何工作表中找到需要填充的行。源列: {src_col_letter}，目标列: {dst_col_letter}"
            )
    except PermissionError:
        return f"错误：保存失败！\n\n请关闭Excel文件 '{file_name}' 后再运行。"
    except Exception as e:
        return f"保存文件时发生未知错误: {e}"