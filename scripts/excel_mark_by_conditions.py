"""
excel_mark_by_conditions.py
───────────────────────────────────────────────────────────────
功能：读取条件文件，对数据 Excel 进行逐行匹配，满足某一行的
      全部条件（AND）则在新增列中打上对应"标识"。

条件文件格式（Excel）：
    列1: 标识    ← 命中时写入的文字
    列2..N: 条件  ← 格式  列字母:列字母=值  或  列字母:列字母>=值
                   例如  AG:AG=是  /  DM:DM>=129

支持的比较操作符：= != > >= < <=
若某行中条件列为空则忽略该条件列。

用法：
    python scripts/excel_mark_by_conditions.py \
        -s 源文件.xlsx \
        -c 条件文件.xlsx \
        -o 输出文件.xlsx \
        --mark_col 高值标识   # 新增列的名称，默认"标识"
"""
import argparse
import re
import sys
import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.styles import PatternFill, Font


# ── 条件解析 ─────────────────────────────────────────────────────────────────
_OP_RE = re.compile(r'^([A-Z]+):([A-Z]+)\s*(>=|<=|!=|>|<|=)\s*(.+)$', re.IGNORECASE)


def parse_condition(cond_str: str):
    """
    解析单个条件字符串。
    返回 (col_letter, op, value) 或 None（若格式不合法）。
    例: 'AG:AG=是'  ->  ('AG', '=', '是')
         'DM:DM>=129' -> ('DM', '>=', '129')
    """
    cond_str = str(cond_str).strip()
    if not cond_str or cond_str.lower() == 'nan':
        return None
    m = _OP_RE.match(cond_str)
    if not m:
        print(f"  ⚠️  无法解析条件：{cond_str!r}，已跳过", file=sys.stderr)
        return None
    col_letter = m.group(2).upper()
    op = m.group(3)
    value = m.group(4).strip()
    return (col_letter, op, value)


def match_value(cell_val, op: str, target: str) -> bool:
    """比较单元格值与目标值，支持数值与字符串双模式。"""
    if cell_val is None:
        cell_val = ''
    cell_str = str(cell_val).strip()
    target_str = target.strip().strip('"').strip("'")

    # 尝试数值比较
    try:
        cell_num = float(cell_str)
        target_num = float(target_str)
        return {
            '=': cell_num == target_num,
            '!=': cell_num != target_num,
            '>': cell_num > target_num,
            '>=': cell_num >= target_num,
            '<': cell_num < target_num,
            '<=': cell_num <= target_num,
        }[op]
    except (ValueError, KeyError):
        pass

    # 字符串比较
    return {
        '=': cell_str == target_str,
        '!=': cell_str != target_str,
        '>': cell_str > target_str,
        '>=': cell_str >= target_str,
        '<': cell_str < target_str,
        '<=': cell_str <= target_str,
    }.get(op, False)


# ── 主逻辑 ───────────────────────────────────────────────────────────────────
def run(src_file: str, cond_file: str, out_file: str, mark_col: str, sheet_name: str = ''):
    print(f"📖 读取源文件：{src_file}")
    wb = load_workbook(src_file)
    ws = wb.active if not sheet_name else wb[sheet_name]
    max_col = ws.max_column
    max_row = ws.max_row

    # 读表头（第1行）
    header_row = [ws.cell(row=1, column=c).value for c in range(1, max_col + 1)]

    # ── 读取并解析条件文件 ──────────────────────────────────────────────────
    print(f"📋 读取条件文件：{cond_file}")
    df_cond = pd.read_excel(cond_file, dtype=str)
    if df_cond.empty:
        print("❌ 条件文件为空，退出。", file=sys.stderr)
        sys.exit(1)

    # 第一列是标识，其余列是条件
    label_col = df_cond.columns[0]
    cond_cols = df_cond.columns[1:]

    # 构建规则列表: [(标识, [(col_letter, op, value), ...]), ...]
    rules = []
    for _, row in df_cond.iterrows():
        label = str(row[label_col]).strip()
        conditions = []
        for cc in cond_cols:
            raw = row.get(cc, '')
            if pd.isna(raw) or str(raw).strip() == '':
                continue
            parsed = parse_condition(str(raw))
            if parsed:
                conditions.append(parsed)
        if label and conditions:
            rules.append((label, conditions))
            print(f"  ✅ 规则「{label}」: {len(conditions)} 个条件")

    if not rules:
        print("❌ 没有解析到有效规则，退出。", file=sys.stderr)
        sys.exit(1)

    # ── 预处理：验证列字母是否在范围内 ────────────────────────────────────
    valid_rules = []
    for label, conditions in rules:
        valid_conds = []
        for col_letter, op, value in conditions:
            try:
                col_idx = column_index_from_string(col_letter)
                if col_idx > max_col:
                    print(f"  ⚠️  规则「{label}」中列 {col_letter}(第{col_idx}列) 超出源文件列数({max_col})，忽略此条件")
                    continue
                valid_conds.append((col_letter, col_idx, op, value))
            except Exception:
                print(f"  ⚠️  无法识别列字母：{col_letter}，跳过")
        if valid_conds:
            valid_rules.append((label, valid_conds))

    # ── 逐行匹配 ───────────────────────────────────────────────────────────
    # 新列追加在最后
    new_col_idx = max_col + 1
    ws.cell(row=1, column=new_col_idx).value = mark_col

    hit_count = 0
    for data_row in range(2, max_row + 1):
        matched_labels = []
        for label, conditions in valid_rules:
            all_match = True
            for col_letter, col_idx, op, target in conditions:
                cell_val = ws.cell(row=data_row, column=col_idx).value
                if not match_value(cell_val, op, target):
                    all_match = False
                    break
            if all_match:
                matched_labels.append(label)

        if matched_labels:
            ws.cell(row=data_row, column=new_col_idx).value = '|'.join(matched_labels)
            hit_count += 1

    # ── 新增列表头样式（橙色底色与其他表头保持一致）────────────────────
    header_cell = ws.cell(row=1, column=new_col_idx)
    header_cell.fill = PatternFill('solid', start_color='E36209', end_color='E36209')
    header_cell.font = Font(name='Arial', bold=True, color='FFFFFF', size=10)

    # ── 保存 ───────────────────────────────────────────────────────────────
    # 若输出路径与源文件相同则覆盖，否则另存
    import shutil
    if os.path.abspath(src_file) != os.path.abspath(out_file):
        shutil.copy2(src_file, out_file)
        # 需要在 copy 的文件上重新操作
        # （已经在 wb 上修改，直接 save 即可）

    wb.save(out_file)
    print(f"\n✅ 完成！共扫描 {max_row - 1} 行，命中 {hit_count} 行 → {out_file}")
    print(f"   新增列「{mark_col}」位于第 {new_col_idx} 列（{get_column_letter(new_col_idx)}）")


# ── CLI ───────────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(description="按条件文件对 Excel 打标识")
    parser.add_argument('-s', '--src',      required=True, help='源 Excel 文件路径')
    parser.add_argument('-c', '--cond',     required=True, help='条件 Excel 文件路径')
    parser.add_argument('-o', '--out',      required=True, help='输出 Excel 文件路径')
    parser.add_argument('--mark_col',       default='高值标识', help='新增标识列的列名（默认：高值标识）')
    parser.add_argument('--sheet',          default='',   help='源文件 Sheet 名（默认读第一个）')
    args = parser.parse_args()

    if not os.path.isfile(args.src):
        print(f"❌ 源文件不存在：{args.src}", file=sys.stderr)
        sys.exit(1)
    if not os.path.isfile(args.cond):
        print(f"❌ 条件文件不存在：{args.cond}", file=sys.stderr)
        sys.exit(1)

    run(args.src, args.cond, args.out, args.mark_col, args.sheet)


if __name__ == '__main__':
    main()
