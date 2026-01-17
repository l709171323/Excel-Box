from excel_toolkit.excel_lite import ExcelReader, ExcelWriter
from excel_toolkit.excel_lite import column_index_from_string


def process_compare_columns(file_x_path, sheets_x_names, col_x_letter, file_y_path, sheet_y_name, col_y_letter, logger=print, ignore_duplicates=True):
    try:
        x_col_index = column_index_from_string(str(col_x_letter).strip())
        y_col_index = column_index_from_string(str(col_y_letter).strip())
    except Exception:
        return f"错误：列号无效。请检查 X 列 '{col_x_letter}' 和 Y 列 '{col_y_letter}' 是否为有效的Excel列号。"

    try:
        logger(f"正在加载表格X: {file_x_path}...")
        wb_X = ExcelReader(file_x_path)
    except Exception as e:
        return f"加载表格X失败: {e}"

    # 收集X端
    x_values = set() if ignore_duplicates else {}
    valid_x_sheets = []
    for sname in sheets_x_names or []:
        if sname in wb_X.sheetnames:
            valid_x_sheets.append(sname)
        else:
            logger(f"  > 警告：表格X中未找到子表 '{sname}'，已跳过。")

    if not valid_x_sheets:
        wb_X.close(); return "错误：未选中任何有效的表格X子表。"

    for sname in valid_x_sheets:
        ws = wb_X[sname]
        logger(f"  > 读取X的子表: {sname}")
        for row in ws.iter_rows(min_col=x_col_index, max_col=x_col_index, min_row=2):
            cell = row[0]
            val = cell.value
            if val is None or val == "":
                continue
            sval = str(val).strip()
            if not sval:
                continue
            if ignore_duplicates:
                x_values.add(sval)
            else:
                x_values[sval] = x_values.get(sval, 0) + 1
    if ignore_duplicates:
        logger(f"X端合计采集 {len(x_values)} 个唯一值。")
    else:
        logger(f"X端合计采集 {sum(x_values.values())} 个值（含重复），唯一值 {len(x_values)} 个。")
    wb_X.close()

    # Y端
    try:
        logger(f"正在加载表格Y: {file_y_path}...")
        wb_Y = ExcelReader(file_y_path)
        if sheet_y_name not in wb_Y.sheetnames:
            wb_Y.close(); return f"错误：表格Y中未找到子表 '{sheet_y_name}'。"
        ws_Y = wb_Y[sheet_y_name]
    except Exception as e:
        return f"加载表格Y失败: {e}"

    y_values = set() if ignore_duplicates else {}
    logger(f"  > 读取Y的子表: {sheet_y_name}")
    for row in ws_Y.iter_rows(min_col=y_col_index, max_col=y_col_index, min_row=2):
        cell = row[0]
        val = cell.value
        if val is None or val == "":
            continue
        sval = str(val).strip()
        if not sval:
            continue
        if ignore_duplicates:
            y_values.add(sval)
        else:
            y_values[sval] = y_values.get(sval, 0) + 1
    if ignore_duplicates:
        logger(f"Y端合计采集 {len(y_values)} 个唯一值。")
    else:
        logger(f"Y端合计采集 {sum(y_values.values())} 个值（含重复），唯一值 {len(y_values)} 个。")
    wb_Y.close()

    if ignore_duplicates:
        missing_in_y = x_values - y_values
        extra_in_y = y_values - x_values

        def log_examples_set(title, items):
            logger(f"{title}：{len(items)} 个")
            count = 0
            for v in sorted(items):
                logger(f"  - {v}")
                count += 1
                if count >= 20:
                    logger("  ... 其余略")
                    break

        if missing_in_y:
            log_examples_set("Y中缺失的值 (存在于X)", missing_in_y)
        if extra_in_y:
            log_examples_set("Y中的多余值 (不在X)", extra_in_y)

        if not missing_in_y and not extra_in_y:
            return (
                "对比完成：一致。\n\n"
                f"X端唯一值: {len(x_values)}；Y端唯一值: {len(y_values)}。\n"
                "两个表的指定列数据集合完全相同。"
            )
        else:
            return (
                "对比完成：存在差异。\n\n"
                f"X端唯一值: {len(x_values)}；Y端唯一值: {len(y_values)}。\n"
                f"Y缺失: {len(missing_in_y)}；Y多余: {len(extra_in_y)}。\n"
                "详细差异已在日志中展示部分示例。"
            )
    else:
        y_counts = y_values
        x_counts = x_values
        missing_in_y_counts = {}
        extra_in_y_counts = {}
        for v, xc in x_counts.items():
            yc = y_counts.get(v, 0)
            if xc > yc:
                missing_in_y_counts[v] = xc - yc
        for v, yc in y_counts.items():
            xc = x_counts.get(v, 0)
            if yc > xc:
                extra_in_y_counts[v] = yc - xc

        def log_examples_counts(title, items_dict):
            logger(f"{title}：{len(items_dict)} 个值存在计数差异")
            count = 0
            for v in sorted(items_dict.keys()):
                xc = x_counts.get(v, 0)
                yc = y_counts.get(v, 0)
                diff = abs(xc - yc)
                logger(f"  - {v}: X={xc}, Y={yc}, 差={diff}")
                count += 1
                if count >= 20:
                    logger("  ... 其余略")
                    break

        if missing_in_y_counts:
            log_examples_counts("Y中缺失的值（考虑重复次数）", missing_in_y_counts)
        if extra_in_y_counts:
            log_examples_counts("Y中的多余值（考虑重复次数）", extra_in_y_counts)

        total_x = sum(x_counts.values())
        total_y = sum(y_counts.values())
        total_missing = sum(missing_in_y_counts.values())
        total_extra = sum(extra_in_y_counts.values())

        if total_missing == 0 and total_extra == 0:
            return (
                "对比完成：一致（考虑重复次数）。\n\n"
                f"X总值: {total_x}；Y总值: {total_y}。唯一值 X={len(x_counts)}，Y={len(y_counts)}。\n"
                "两个表的指定列数据在计数上完全一致。"
            )
        else:
            return (
                "对比完成：存在差异（考虑重复次数）。\n\n"
                f"X总值: {total_x}；Y总值: {total_y}。唯一值 X={len(x_counts)}，Y={len(y_counts)}。\n"
                f"Y缺失总数: {total_missing}；Y多余总数: {total_extra}。\n"
                "详细差异已在日志中展示部分示例。"
            )
