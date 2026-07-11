"""
excel_sheet_splitter.py
========================
Excel 多 Sheet 拆分与重组核心逻辑

拆分（split_by_column）：
  - 读取一个多 Sheet 的 Excel 文件
  - 按每个 Sheet 中指定列的唯一值分组
  - 每个唯一值输出一个 Excel 文件，文件内保留原来所有 Sheet
    （但每 Sheet 只包含该值对应的行）

重组（merge_by_filename）：
  - 扫描一个目录内所有 .xlsx 文件
  - 按文件名分组，文件名相同的文件合并为一个新文件
  - 每个源文件的 Sheet 数据叠加追加到目标文件的对应 Sheet 中
"""

import os
import sys

try:
    import pandas as pd
    from openpyxl import load_workbook
except ImportError:
    print("❌ 缺少依赖：请执行 pip install pandas openpyxl")
    sys.exit(1)


# ─── 拆分 ────────────────────────────────────────────────────────────────────

def split_by_column(input_path: str, split_col: str, output_dir: str) -> dict:
    """
    将一个含多个 Sheet 的 Excel 文件，按指定列的值拆分成多个文件。

    Args:
        input_path:  源 Excel 文件的绝对路径
        split_col:   用于分组拆分的列名
        output_dir:  输出目录（不存在时自动创建）

    Returns:
        dict: {
            "output_files": [str],   # 生成的文件路径列表
            "split_values": [str],   # 所有唯一值列表
            "sheet_names": [str],    # Sheet 名列表
            "skipped_sheets": [str]  # 不含指定列的 Sheet（跳过）
        }
    """
    if not os.path.isfile(input_path):
        raise FileNotFoundError(f"源文件不存在：{input_path}")

    os.makedirs(output_dir, exist_ok=True)

    print(f"📂 读取文件：{input_path}")
    xl = pd.ExcelFile(input_path, engine='openpyxl')
    sheet_names = xl.sheet_names
    print(f"📋 共 {len(sheet_names)} 个 Sheet：{sheet_names}")

    # 读取所有 Sheet
    sheets_data: dict[str, pd.DataFrame] = {}
    skipped_sheets: list[str] = []

    for sheet in sheet_names:
        df = xl.parse(sheet, dtype=str)
        if split_col not in df.columns:
            print(f"⚠️  Sheet「{sheet}」中未找到列「{split_col}」，跳过。")
            skipped_sheets.append(sheet)
            continue
        sheets_data[sheet] = df
        print(f"  ✅ Sheet「{sheet}」：{len(df)} 行，含目标列「{split_col}」")

    if not sheets_data:
        raise ValueError(f"所有 Sheet 中均未找到列「{split_col}」，请检查列名。")

    # 收集所有 Sheet 中该列的唯一值（全集）
    all_values: set = set()
    for df in sheets_data.values():
        all_values.update(df[split_col].dropna().astype(str).unique())

    all_values = sorted(all_values)
    print(f"\n🔑 共发现 {len(all_values)} 个唯一值：{all_values[:10]}{'...' if len(all_values) > 10 else ''}")

    output_files: list[str] = []
    total = len(all_values)

    for idx, val in enumerate(all_values, 1):
        # 清理文件名中的非法字符
        safe_val = _safe_filename(val)
        out_path = os.path.join(output_dir, f"{safe_val}.xlsx")

        print(f"\n[PROGRESS: {int(idx / total * 85)}%]")
        print(f"  📝 正在生成：{safe_val}.xlsx（值 = 「{val}」）")

        with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
            written = False
            for sheet_name, df in sheets_data.items():
                subset = df[df[split_col].astype(str) == val].copy()
                if subset.empty:
                    # 即使该 Sheet 下没有该值的行，也写入空 Sheet（仅含表头），保持结构一致
                    subset = df.head(0)
                subset.to_excel(writer, sheet_name=sheet_name, index=False)
                written = True

            # 若有被跳过的 Sheet（不含目标列），原样写入
            for skipped in skipped_sheets:
                df_skipped = xl.parse(skipped, dtype=str)
                df_skipped.to_excel(writer, sheet_name=skipped, index=False)
                written = True

        if written:
            output_files.append(out_path)

    print(f"\n[PROGRESS: 100%]")
    print(f"\n✅ 拆分完成！共生成 {len(output_files)} 个文件，输出至：{output_dir}")

    return {
        "output_files": output_files,
        "split_values": all_values,
        "sheet_names": list(sheets_data.keys()),
        "skipped_sheets": skipped_sheets,
    }


# ─── 重组 ────────────────────────────────────────────────────────────────────

def merge_by_filename(input_dir: str, output_dir: str) -> dict:
    """
    扫描目录内所有 .xlsx 文件，按文件名分组合并为新文件。
    文件名相同的多个文件，其各 Sheet 数据叠加追加，合并到一个文件中。

    典型使用场景：将多次拆分输出的同名文件（来自不同批次/子目录）重组还原。

    Args:
        input_dir:   来源目录（含多个 .xlsx 文件，文件名即为分组键）
        output_dir:  输出目录（不存在时自动创建）

    Returns:
        dict: {
            "output_files": [str],    # 生成的文件路径列表
            "merged_groups": int,     # 合并的分组数（即唯一文件名数）
            "source_count": int       # 来源文件总数
        }
    """
    if not os.path.isdir(input_dir):
        raise FileNotFoundError(f"来源目录不存在：{input_dir}")

    os.makedirs(output_dir, exist_ok=True)

    # 收集目录下所有 xlsx（不递归）
    xlsx_files = [
        f for f in os.listdir(input_dir)
        if f.lower().endswith('.xlsx') and os.path.isfile(os.path.join(input_dir, f))
    ]

    if not xlsx_files:
        raise ValueError(f"来源目录中没有找到任何 .xlsx 文件：{input_dir}")

    print(f"📂 来源目录：{input_dir}")
    print(f"📋 发现 {len(xlsx_files)} 个 .xlsx 文件")

    # 按文件名（去掉 .xlsx 后缀）分组（此版本文件名即为组名）
    groups: dict[str, list[str]] = {}
    for fname in xlsx_files:
        groups.setdefault(fname, []).append(os.path.join(input_dir, fname))

    output_files: list[str] = []
    total = len(groups)

    for idx, (fname, src_paths) in enumerate(sorted(groups.items()), 1):
        print(f"\n[PROGRESS: {int(idx / total * 85)}%]")
        out_path = os.path.join(output_dir, fname)
        print(f"  🔗 合并 → {fname}（{len(src_paths)} 个来源）")

        # 按 Sheet 合并数据
        # 结构：{ sheet_name: [df1, df2, ...] }
        merged: dict[str, list[pd.DataFrame]] = {}

        for src_path in src_paths:
            xl = pd.ExcelFile(src_path, engine='openpyxl')
            for sheet in xl.sheet_names:
                df = xl.parse(sheet, dtype=str)
                merged.setdefault(sheet, []).append(df)

        with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
            for sheet_name, dfs in merged.items():
                # 纵向堆叠，去重表头（pandas concat 自动处理）
                combined = pd.concat(dfs, ignore_index=True)
                # 去除完全重复行（可选，默认保留所有行）
                combined.to_excel(writer, sheet_name=sheet_name, index=False)

        output_files.append(out_path)

    print(f"\n[PROGRESS: 100%]")
    print(f"\n✅ 重组完成！共合并 {total} 组，输出至：{output_dir}")

    return {
        "output_files": output_files,
        "merged_groups": total,
        "source_count": len(xlsx_files),
    }


# ─── 工具函数 ─────────────────────────────────────────────────────────────────

def _safe_filename(val: str) -> str:
    """将字符串转为安全的文件名（替换 Windows/Linux 非法字符）"""
    illegal = r'\/:*?"<>|'
    result = val
    for ch in illegal:
        result = result.replace(ch, '_')
    # 限制长度防止路径过长
    return result[:80].strip()


# ─── CLI 入口 ─────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    import argparse

    parser = argparse.ArgumentParser(description='Excel 多 Sheet 拆分与重组工具')
    subparsers = parser.add_subparsers(dest='action', required=True)

    # split 子命令
    sp = subparsers.add_parser('split', help='按列值拆分 Excel')
    sp.add_argument('--input',      required=True, help='源 Excel 文件路径')
    sp.add_argument('--col',        required=True, help='用于拆分的列名')
    sp.add_argument('--output_dir', required=True, help='输出目录')

    # merge 子命令
    mp = subparsers.add_parser('merge', help='按文件名重组 Excel')
    mp.add_argument('--input_dir',  required=True, help='来源目录')
    mp.add_argument('--output_dir', required=True, help='输出目录')

    args = parser.parse_args()

    if args.action == 'split':
        result = split_by_column(args.input, args.col, args.output_dir)
        print(f"\n📊 生成文件数：{len(result['output_files'])}")
        print(f"📋 唯一值列表：{result['split_values']}")
        print(f"⚠️  跳过 Sheet：{result['skipped_sheets']}")

    elif args.action == 'merge':
        result = merge_by_filename(args.input_dir, args.output_dir)
        print(f"\n📊 合并组数：{result['merged_groups']}")
        print(f"📁 来源文件数：{result['source_count']}")
