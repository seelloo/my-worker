"""
excel_column_reorder.py
从源 Excel 文件中按指定字段顺序提取列，输出到新 Excel 文件。
- 源文件缺失的字段：跳过警告，以空列代替（标题保留）
- 所有输出单元格格式强制为文本型
"""
from __future__ import annotations

# The list of your imports...
import argparse
import json
import sys
from pathlib import Path

from openpyxl import load_workbook, Workbook
from openpyxl.styles import numbers


def reorder_columns(src_path_str: str, dest_path_str: str, fields: list[str], sheet_name: str = None):
    """
    按 fields 顺序从源文件提取数据写入目标文件。

    Args:
        src_path_str: 源 Excel 文件路径
        dest_path_str: 输出 Excel 文件路径
        fields: 有序字段名列表（可含源文件中不存在的字段）
        sheet_name: 指定读取的 Sheet 名；None 则取 active sheet
    """
    src_path = Path(src_path_str).resolve()
    dest_path = Path(dest_path_str).resolve()

    if not src_path.exists():
        print(f"❌ 源文件不存在: {src_path}")
        sys.exit(1)

    if not fields:
        print("❌ 字段列表为空，无法执行。")
        sys.exit(1)

    print(f"📄 正在读取源文件: {src_path.name}")
    wb_src = load_workbook(src_path, data_only=True, read_only=True)

    # 选取目标 Sheet
    if sheet_name and sheet_name in wb_src.sheetnames:
        ws_src = wb_src[sheet_name]
    else:
        if sheet_name:
            print(f"⚠️ Sheet 「{sheet_name}」不存在，已自动使用第一个 Sheet。")
        ws_src = wb_src.active

    print(f"📋 当前 Sheet: {ws_src.title}")

    # 读取表头（第一行）
    header_row = next(ws_src.iter_rows(min_row=1, max_row=1, values_only=True), None)
    if not header_row:
        print("❌ 源文件表头为空。")
        wb_src.close()
        sys.exit(1)

    # 构建列名 → 列索引映射（大小写不敏感降级）
    header = [str(h).strip() if h is not None else "" for h in header_row]
    col_map_exact = {h: i for i, h in enumerate(header) if h}
    col_map_lower = {h.lower(): i for i, h in enumerate(header) if h}

    # 解析每个目标字段对应的源列索引（None 表示不存在，输出空列）
    field_indices = []
    for field in fields:
        if field in col_map_exact:
            field_indices.append(col_map_exact[field])
        elif field.lower() in col_map_lower:
            idx = col_map_lower[field.lower()]
            print(f"⚠️ 字段「{field}」已通过大小写不敏感方式匹配到「{header[idx]}」")
            field_indices.append(idx)
        else:
            print(f"⚠️ 字段「{field}」在源文件中不存在，将以空数据代替。")
            field_indices.append(None)

    print(f"✅ 共 {len(fields)} 个目标字段，其中 {field_indices.count(None)} 个为空替代列。")

    # 创建输出工作簿
    wb_out = Workbook(write_only=True)
    ws_out = wb_out.create_sheet("重组结果")

    TEXT_FORMAT = numbers.FORMAT_TEXT  # '@'

    def make_text_cell(value):
        """创建强制文本格式的单元格写入值"""
        from openpyxl.cell import WriteOnlyCell
        cell = WriteOnlyCell(ws_out, value=str(value) if value is not None else "")
        cell.number_format = TEXT_FORMAT
        return cell

    # 写入表头行
    header_cells = [make_text_cell(fields[i]) for i in range(len(fields))]
    ws_out.append(header_cells)

    # 逐行写入数据
    row_count = 0
    for row_idx, row_values in enumerate(ws_src.iter_rows(min_row=2, values_only=True), start=2):
        row_count += 1
        out_row = []
        for src_idx in field_indices:
            if src_idx is None:
                # 字段不存在：以空文本代替
                cell = make_text_cell("")
            else:
                val = row_values[src_idx] if src_idx < len(row_values) else None
                cell = make_text_cell(val)
            out_row.append(cell)
        ws_out.append(out_row)

        if row_count % 5000 == 0:
            print(f"🚀 已处理 {row_count} 行...")

    wb_src.close()

    # 确保输出目录存在
    dest_path.parent.mkdir(parents=True, exist_ok=True)

    print(f"💾 正在保存至: {dest_path.name} ...")
    wb_out.save(dest_path)
    print(f"✅ 处理完成！共输出 {row_count} 行，{len(fields)} 列。文件已保存至: {dest_path}")


def parse_txt_fields(txt_path_str: str) -> list[str]:
    """
    解析 TXT 文件，每行一个字段名，空行跳过。
    自动探测编码：utf-8-sig → gbk → gb18030 → latin-1（兜底）。
    返回有序字段名列表。
    """
    txt_path = Path(txt_path_str).resolve()
    if not txt_path.exists():
        raise FileNotFoundError(f"TXT 文件不存在: {txt_path}")
    if not txt_path.is_file():
        raise ValueError(f"路径不是文件: {txt_path}")

    # 按优先级尝试编码（覆盖 UTF-8 BOM / 无 BOM / Windows GBK / 兜底 latin-1）
    _ENCODINGS = ["utf-8-sig", "gbk", "gb18030", "latin-1"]
    content = None
    detected_enc = None
    raw = txt_path.read_bytes()
    for enc in _ENCODINGS:
        try:
            content = raw.decode(enc)
            detected_enc = enc
            break
        except (UnicodeDecodeError, LookupError):
            continue

    if content is None:
        raise ValueError("TXT 文件编码无法识别，请将文件另存为 UTF-8 格式后重试")

    print(f"[INFO] TXT encoding detected: {detected_enc}")
    fields = [line.strip() for line in content.splitlines() if line.strip()]

    if not fields:
        raise ValueError("TXT 文件中未找到有效字段名（所有行均为空）")

    return fields


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Excel 字段重组导出工具")
    parser.add_argument("-s", "--src", required=True, help="源 Excel 文件路径")
    parser.add_argument("-o", "--out", required=True, help="输出 Excel 文件路径")
    parser.add_argument(
        "-f", "--fields", default=None,
        help="有序字段名 JSON 数组，如 '[\"姓名\",\"部门\"]'"
    )
    parser.add_argument(
        "--txt", default=None,
        help="字段名 TXT 文件路径（每行一个字段名，与 -f 互斥）"
    )
    parser.add_argument("--sheet", default=None, help="指定 Sheet 名称（可选）")

    args = parser.parse_args()

    if args.txt and args.fields:
        print("❌ --txt 和 -f 不能同时使用，请选择其中一个。")
        sys.exit(1)
    if not args.txt and not args.fields:
        print("❌ 必须提供 -f 或 --txt 参数。")
        sys.exit(1)

    if args.txt:
        try:
            target_fields = parse_txt_fields(args.txt)
            print(f"📥 从 TXT 文件读取到 {len(target_fields)} 个字段。")
        except (FileNotFoundError, ValueError) as e:
            print(f"❌ {e}")
            sys.exit(1)
    else:
        try:
            target_fields = json.loads(args.fields)
        except json.JSONDecodeError as e:
            print(f"❌ 字段 JSON 解析失败: {e}")
            sys.exit(1)

    reorder_columns(args.src, args.out, target_fields, sheet_name=args.sheet)
