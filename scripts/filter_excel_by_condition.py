import os
import argparse
import json
import re
from pathlib import Path
from openpyxl import load_workbook, Workbook
from openpyxl.utils import column_index_from_string

SUPPORTED_OPERATORS = {
    "contains", "not_contains", "equals", "not_equals",
    "gt", "lt", "gte", "lte", "is_empty", "is_not_empty"
}


def is_numeric(val):
    if val is None:
        return False
    try:
        float(val)
        return True
    except (ValueError, TypeError):
        return False


def evaluate_condition(cell_value, op, target_value):
    """计算单个条件是否满足（支持逗号分割多值匹配）"""
    if cell_value is None:
        cv_str = ""
    else:
        cv_str = str(cell_value).strip()

    if op == "is_empty":
        return cv_str == ""
    if op == "is_not_empty":
        return cv_str != ""

    if isinstance(target_value, str):
        raw_targets = re.split(r'[,，]', target_value)
        targets = [t.strip() for t in raw_targets if t.strip()]
    else:
        targets = [str(target_value).strip()]

    if not targets:
        return op != "equals"

    if op == "contains":
        for t in targets:
            if t.lower() in cv_str.lower():
                return True
        return False

    if op == "not_contains":
        for t in targets:
            if t.lower() in cv_str.lower():
                return False
        return True

    if op == "equals":
        for t in targets:
            if cv_str.lower() == t.lower():
                return True
        return False

    if op == "not_equals":
        for t in targets:
            if cv_str.lower() == t.lower():
                return False
        return True

    if op in ("gt", "lt", "gte", "lte"):
        t0 = targets[0] if targets else "0"
        if is_numeric(cell_value) and is_numeric(t0):
            cv_num = float(cell_value)
            tv_num = float(t0)
            if op == "gt": return cv_num > tv_num
            if op == "lt": return cv_num < tv_num
            if op == "gte": return cv_num >= tv_num
            if op == "lte": return cv_num <= tv_num
        else:
            if op == "gt": return cv_str > t0
            if op == "lt": return cv_str < t0
            if op == "gte": return cv_str >= t0
            if op == "lte": return cv_str <= t0

    return False


def load_header(ws):
    """读取工作表首行作为表头列表"""
    for row in ws.iter_rows(min_row=1, max_row=1, values_only=True):
        return list(row)
    return []


def resolve_column_index(col_name, col_to_idx, header):
    """
    将列描述符解析为零基列索引。
    支持: 直接表头名匹配、"A (列名)" 格式、纯字母(A/B/C)匹配。
    返回索引或 None。
    """
    col_name = col_name.strip()
    if not col_name:
        return None

    if col_name in col_to_idx:
        return col_to_idx[col_name]

    match = re.match(r"^([A-Z]+)\s*\((.*)\)$", col_name, re.IGNORECASE)
    if match:
        letter, name = match.groups()
        if name in col_to_idx:
            return col_to_idx[name]
        try:
            idx = column_index_from_string(letter) - 1
            if 0 <= idx < len(header):
                return idx
        except (ValueError, IndexError):
            pass

    letter_match = re.match(r"^([A-Z]+)", col_name, re.IGNORECASE)
    if letter_match:
        try:
            idx = column_index_from_string(letter_match.group()) - 1
            if 0 <= idx < len(header):
                return idx
        except (ValueError, IndexError):
            pass

    return None


def resolve_conditions(conditions, header):
    """
    验证并解析条件列表，将列描述符解析为索引。
    返回 (valid_conditions, warnings)。
    """
    col_to_idx = {str(h).strip(): i for i, h in enumerate(header) if h is not None}
    valid = []
    warnings = []

    for cond in conditions:
        col_name = (cond.get("col") or "").strip()
        if not col_name:
            warnings.append(f"跳过空列名的条件")
            continue

        idx = resolve_column_index(col_name, col_to_idx, header)
        if idx is not None:
            cond["idx"] = idx
            valid.append(cond)
        else:
            warnings.append(f"无法识别列 '{col_name}'，已跳过")

    return valid, warnings


def match_row(row_values, valid_conditions):
    """检查数据行是否满足所有条件（AND 逻辑）"""
    for cond in valid_conditions:
        val = row_values[cond["idx"]]
        if not evaluate_condition(val, cond["op"], cond.get("val")):
            return False
    return True


def build_output_path(src_path, dest_path):
    """解析输出路径：目录则追加文件名，存在则追加编号后缀"""
    if dest_path.is_dir():
        dest_path = dest_path / f"Filtered_{src_path.name}"

    if dest_path.exists():
        stem = dest_path.stem
        ext = dest_path.suffix
        counter = 1
        while True:
            new_dest = dest_path.parent / f"{stem}_{counter}{ext}"
            if not new_dest.exists():
                return new_dest
            counter += 1

    return dest_path


def load_conditions_from_excel(cond_file_path):
    """
    从 Excel 条件文档读取筛选条件。
    文档格式：首个 Sheet，首行为表头（目标列 / 操作符 / 比对值），按表头名匹配列位置。
    返回条件列表: [{"col": "...", "op": "...", "val": "..."}, ...]
    """
    path = Path(cond_file_path).resolve()
    if not path.exists():
        raise FileNotFoundError(f"条件文档不存在: {path}")
    if not path.is_file():
        raise ValueError(f"条件文档路径不是文件: {path}")

    wb = load_workbook(path, data_only=True)
    ws = wb.active
    headers = load_header(ws)
    if not headers:
        wb.close()
        raise ValueError("条件文档为空或无法读取表头")

    header_map = {str(h).strip(): i for i, h in enumerate(headers) if h is not None}

    col_idx = None
    op_idx = None
    val_idx = None
    for key in header_map:
        k = key.lower()
        if k in ("目标列", "col", "column", "列名", "列"):
            col_idx = header_map[key]
        elif k in ("操作符", "op", "operator", "运算符", "条件"):
            op_idx = header_map[key]
        elif k in ("比对值", "val", "value", "目标值", "值"):
            val_idx = header_map[key]

    missing = []
    if col_idx is None:
        missing.append("目标列")
    if op_idx is None:
        missing.append("操作符")
    if val_idx is None:
        missing.append("比对值")
    if missing:
        wb.close()
        raise ValueError(f"条件文档缺少必要列: {', '.join(missing)}。请确保表头包含「目标列」「操作符」「比对值」")

    conditions = []
    warnings = []
    for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        col_val = str(row[col_idx]).strip() if len(row) > col_idx and row[col_idx] is not None else ""
        op_val = str(row[op_idx]).strip() if len(row) > op_idx and row[op_idx] is not None else ""
        target_val = str(row[val_idx]) if len(row) > val_idx and row[val_idx] is not None else ""

        if not col_val:
            warnings.append(f"第 {row_idx} 行: 目标列为空，已跳过")
            continue

        if op_val not in SUPPORTED_OPERATORS:
            warnings.append(f"第 {row_idx} 行: 不支持的操作符 '{op_val}'，已跳过。支持: {', '.join(sorted(SUPPORTED_OPERATORS))}")
            continue

        conditions.append({"col": col_val, "op": op_val, "val": target_val})

    wb.close()

    if warnings:
        for w in warnings:
            print(f"⚠️ {w}")

    if not conditions:
        raise ValueError("条件文档中未找到有效的筛选条件")

    print(f"✅ 从条件文档成功导入 {len(conditions)} 个筛选条件")
    return conditions


def filter_excel(src_path_str, dest_path_str, conditions_json, sheet_name=None):
    """执行多条件筛选"""
    src_path = Path(src_path_str).resolve()
    dest_path = Path(dest_path_str).resolve()

    try:
        conditions = json.loads(conditions_json)
    except (json.JSONDecodeError, TypeError) as e:
        print(f"❌ 条件解析失败: {e}")
        return

    if not src_path.exists():
        print(f"❌ 源文件不存在: {src_path}")
        return

    print(f"📄 正在读取源文件: {src_path.name}")
    wb = load_workbook(src_path, data_only=True, read_only=True)

    # 根据 sheet_name 选择工作表
    if sheet_name:
        if sheet_name not in wb.sheetnames:
            print(f"❌ Sheet '{sheet_name}' 不存在，可用 Sheet：{wb.sheetnames}")
            wb.close()
            return
        ws = wb[sheet_name]
        print(f"📑 已选择 Sheet：{sheet_name}")
    else:
        ws = wb.active
        print(f"📑 使用默认 Sheet：{ws.title}")

    header = load_header(ws)
    if not header:
        print("❌ 无法识别表头或文件为空。")
        wb.close()
        return

    valid_conditions, warnings = resolve_conditions(conditions, header)
    for w in warnings:
        print(f"⚠️ {w}")

    if not valid_conditions:
        print("❌ 没有有效的过滤条件。")
        wb.close()
        return

    print(f"🔍 开始过滤，有效条件数: {len(valid_conditions)}")

    wb_out = Workbook(write_only=True)
    ws_out = wb_out.create_sheet("Filtered Result")
    ws_out.append(header)

    match_count = 0
    total_count = 0

    for row_values in ws.iter_rows(min_row=2, values_only=True):
        total_count += 1
        if match_row(row_values, valid_conditions):
            ws_out.append(row_values)
            match_count += 1
        if total_count % 1000 == 0:
            print(f"🚀 已扫描 {total_count} 行，匹配 {match_count} 行...")

    wb.close()

    dest_path = build_output_path(src_path, dest_path)
    print(f"💾 正在保存结果至: {dest_path.name} ...")
    wb_out.save(dest_path)
    print(f"✅ 处理完成！扫描 {total_count} 行，导出 {match_count} 行记录。")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Excel 多条件筛选工具")
    parser.add_argument("-s", "--src", required=True, help="源文件路径")
    parser.add_argument("-o", "--out", required=True, help="输出路径")
    parser.add_argument("-c", "--conditions", default=None, help="条件 JSON 字符串")
    parser.add_argument("--cond-file", default=None, help="条件 Excel 文档路径（与 -c 互斥）")
    parser.add_argument("--sheet", default=None, help="指定读取的 Sheet 名称（默认取第一个 Sheet）")

    args = parser.parse_args()

    if args.cond_file and args.conditions:
        print("❌ --cond-file 和 -c 不能同时使用，请选择其中一个。")
        exit(1)
    if not args.cond_file and not args.conditions:
        print("❌ 必须提供 -c 或 --cond-file 参数。")
        exit(1)

    if args.cond_file:
        conds = load_conditions_from_excel(args.cond_file)
        conditions_str = json.dumps(conds, ensure_ascii=False)
    else:
        conditions_str = args.conditions

    filter_excel(args.src, args.out, conditions_str, sheet_name=args.sheet)
