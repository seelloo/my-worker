"""
cross_sheet_merge.py
--------------------
从同一 Excel 文件的多个 Sheet 中，按用户定义的规则，
将不同 Sheet 的不同列垂直拼接，生成新 Sheet 的各列。

输入 JSON 结构示例：
{
  "src_file": "C:/data/source.xlsx",
  "out_file": "C:/data/output.xlsx",
  "out_sheet": "合并结果",
  "skip_header": true,
  "columns": [
    {
      "header": "新列名1",
      "segments": [
        {"sheet": "Sheet1", "col": "A"},
        {"sheet": "Sheet2", "col": "B"}
      ]
    },
    {
      "header": "新列名2",
      "segments": [
        {"sheet": "Sheet1", "col": "C"},
        {"sheet": "Sheet3", "col": "A"}
      ]
    }
  ]
}
"""

import argparse
import json
import sys
from pathlib import Path

import openpyxl
from openpyxl.utils import column_index_from_string


def get_sheet_names(src_file: str) -> list:
    """返回 Excel 文件中所有 Sheet 名称列表。"""
    wb = openpyxl.load_workbook(src_file, read_only=True, data_only=True)
    names = wb.sheetnames
    wb.close()
    return names


def get_sheet_columns(src_file: str, sheet_name: str) -> list:
    """
    返回指定 Sheet 第一行（表头）的列信息列表。
    格式：[{"letter": "A", "header": "列名1"}, ...]
    """
    wb = openpyxl.load_workbook(src_file, read_only=True, data_only=True)
    if sheet_name not in wb.sheetnames:
        wb.close()
        return []

    ws = wb[sheet_name]
    result = []
    for i, cell in enumerate(next(ws.iter_rows(min_row=1, max_row=1), []), start=1):
        letter = _col_index_to_letter(i)
        header = str(cell.value) if cell.value is not None else ""
        result.append({"letter": letter, "header": header})
    wb.close()
    return result


def _col_index_to_letter(col_index: int) -> str:
    """将列索引（1-based）转换为 Excel 列字母（A, B, ..., AA, ...）。"""
    name = ""
    while col_index > 0:
        mod = (col_index - 1) % 26
        name = chr(65 + mod) + name
        col_index = (col_index - mod) // 26
    return name


def _read_column_values(ws, col_letter: str, skip_header: bool) -> list:
    """
    从工作表中读取指定列的所有值（排除空行末尾）。
    skip_header=True 时跳过第一行。
    """
    try:
        col_idx = column_index_from_string(col_letter.upper())
    except Exception:
        print(f"  ⚠️  无效列标识: '{col_letter}'，跳过该片段。")
        return []

    values = []
    for row in ws.iter_rows(min_col=col_idx, max_col=col_idx, values_only=True):
        values.append(row[0])

    if skip_header and values:
        values = values[1:]

    return values


def cross_sheet_merge(config: dict) -> bool:
    """
    按照 config 配置执行跨 Sheet 列合并。
    返回 True 表示成功。
    """
    src_file = config.get("src_file", "").strip()
    out_file = config.get("out_file", "").strip()
    out_sheet = config.get("out_sheet", "合并结果").strip() or "合并结果"
    skip_header = config.get("skip_header", True)
    columns_cfg = config.get("columns", [])

    # ── 参数校验 ──────────────────────────────────────────────
    if not src_file:
        print("❌ 错误：缺少 src_file 参数。")
        return False
    src_path = Path(src_file).resolve()
    if not src_path.exists() or not src_path.is_file():
        print(f"❌ 错误：源文件不存在：{src_path}")
        return False
    if not out_file:
        print("❌ 错误：缺少 out_file 参数。")
        return False
    if not columns_cfg:
        print("❌ 错误：columns 配置为空，无需合并。")
        return False

    out_path = Path(out_file).resolve()
    out_path.parent.mkdir(parents=True, exist_ok=True)

    print(f"📂 源文件：{src_path.name}")
    print(f"📝 输出文件：{out_path}")
    print(f"📋 输出 Sheet：{out_sheet}")
    print(f"{'⏭️  跳过' if skip_header else '📌 保留'}各 Sheet 原始表头")
    print(f"🔢 共需生成 {len(columns_cfg)} 个输出列\n")

    # ── 读取源文件 ────────────────────────────────────────────
    try:
        src_wb = openpyxl.load_workbook(src_path, data_only=True)
    except Exception as e:
        print(f"❌ 无法读取源文件：{e}")
        return False

    # ── 按列配置逐列组装数据 ───────────────────────────────────
    headers = []
    col_data = []  # list of list，每个子 list 是一个输出列的所有行值

    for col_cfg in columns_cfg:
        col_header = col_cfg.get("header", "")
        segments = col_cfg.get("segments", [])
        merged_values = []

        print(f"  🔧 处理输出列「{col_header}」，共 {len(segments)} 个片段：")

        for seg in segments:
            sheet_name = seg.get("sheet", "").strip()
            col_letter = seg.get("col", "").strip()

            if not sheet_name or not col_letter:
                print(f"    ⚠️  跳过无效片段（sheet 或 col 为空）")
                continue

            if sheet_name not in src_wb.sheetnames:
                print(f"    ⚠️  Sheet「{sheet_name}」不存在，跳过该片段。")
                continue

            ws = src_wb[sheet_name]
            values = _read_column_values(ws, col_letter, skip_header)
            print(f"    ✅ Sheet「{sheet_name}」列「{col_letter}」→ 读取 {len(values)} 行")
            merged_values.extend(values)

        headers.append(col_header)
        col_data.append(merged_values)
        print(f"  📊 列「{col_header}」合并完毕，共 {len(merged_values)} 行数据\n")

    src_wb.close()

    # ── 写入输出文件 ──────────────────────────────────────────
    # 计算最大行数（各列数据量可能不一致，用空值补齐）
    max_rows = max((len(d) for d in col_data), default=0)
    if max_rows == 0:
        print("❌ 所有列均无有效数据，取消写入。")
        return False

    # 如果目标文件已存在，加载并追加 Sheet；否则新建
    if out_path.exists():
        out_wb = openpyxl.load_workbook(out_path)
        # 若同名 Sheet 已存在则加编号避免冲突
        final_sheet_name = _unique_sheet_name(out_wb.sheetnames, out_sheet)
        out_ws = out_wb.create_sheet(title=final_sheet_name)
    else:
        out_wb = openpyxl.Workbook()
        out_ws = out_wb.active
        out_ws.title = out_sheet
        final_sheet_name = out_sheet

    # 写入表头
    out_ws.append(headers)

    # 写入数据行（列对齐，不足的列用 None 补齐）
    for row_idx in range(max_rows):
        row = [
            col_data[c][row_idx] if row_idx < len(col_data[c]) else None
            for c in range(len(col_data))
        ]
        out_ws.append(row)

    try:
        out_wb.save(out_path)
        out_wb.close()
        print(f"🎉 完成！输出 Sheet「{final_sheet_name}」共 {max_rows} 行数据，已保存至：")
        print(f"   {out_path}")
        return True
    except Exception as e:
        print(f"❌ 保存失败：{e}")
        return False


def _unique_sheet_name(existing: list, name: str) -> str:
    """确保 Sheet 名称在现有列表中唯一，超长截断，冲突加编号。"""
    base = name[:28].strip()
    candidate = base
    i = 1
    while candidate in existing:
        candidate = f"{base[:25]}_{i}"
        i += 1
    return candidate


def main():
    parser = argparse.ArgumentParser(
        description="跨 Sheet 多列自由组合合并工具"
    )
    parser.add_argument("-s", "--src", required=True, help="源 Excel 文件路径")
    parser.add_argument("-o", "--out", required=True, help="输出 Excel 文件路径")
    parser.add_argument(
        "-c", "--config", required=True,
        help="列配置 JSON 字符串，或 JSON 文件路径"
    )
    parser.add_argument(
        "--out_sheet", default="合并结果",
        help="输出 Sheet 名称（默认：合并结果）"
    )
    parser.add_argument(
        "--keep_header", action="store_true",
        help="保留各 Sheet 原始表头（默认跳过）"
    )
    args = parser.parse_args()

    # 解析 columns 配置（可以是 JSON 字符串或 JSON 文件路径）
    cfg_str = args.config.strip()
    if cfg_str.startswith("[") or cfg_str.startswith("{"):
        try:
            columns_cfg = json.loads(cfg_str)
        except json.JSONDecodeError as e:
            print(f"❌ JSON 解析失败：{e}")
            sys.exit(1)
    else:
        cfg_path = Path(cfg_str)
        if not cfg_path.exists():
            print(f"❌ 配置文件不存在：{cfg_path}")
            sys.exit(1)
        with open(cfg_path, encoding="utf-8") as f:
            columns_cfg = json.load(f)

    config = {
        "src_file": args.src,
        "out_file": args.out,
        "out_sheet": args.out_sheet,
        "skip_header": not args.keep_header,
        "columns": columns_cfg if isinstance(columns_cfg, list) else columns_cfg.get("columns", []),
    }

    success = cross_sheet_merge(config)
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
