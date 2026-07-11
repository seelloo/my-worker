"""
excel_proportional_alloc.py
===========================
按比例分配引擎：
- 读取文件 A 中的关联键（key_a）和权重列（weight_a）
- 读取文件 B 中的关联键（key_b）和待分配值列（value_b）
- 按 key_a == key_b 进行匹配，将 B 文件的 value 值按 A 文件的权重比例分配
- 保证最终分配结果的精确合计值等于原始 B 文件的值（无缝处理尾差，甚至 A 文件关联键乱序也支持）。
- 支持自定义输出列（可来自 A、B 或 结果列）
"""

import os
import argparse
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from openpyxl import load_workbook, Workbook


def get_unique_path(base_path: Path) -> Path:
    if not base_path.exists():
        return base_path
    name = base_path.stem
    ext = base_path.suffix
    dir_path = base_path.parent
    counter = 1
    while True:
        new_path = dir_path / f"{name}_{counter}{ext}"
        if not new_path.exists():
            return new_path
        counter += 1


def excel_proportional_alloc(
    file_a: str,
    file_b: str,
    output_path: str,
    key_a_name: str,
    weight_a_name: str,
    key_b_name: str,
    value_b_name: str,
    result_col: str = "分配结果",
    output_columns: list = None,
    decimal_places: int = 2,
):
    """
    核心逻辑（动态余额分配法）：
    1. 索引 B 表，将 key -> sum(value) 保存至字典。
    2. 第一遍扫描 A 表，将 key -> sum(weight) 和 count() 统计保存。
    3. 组建输出列映射关系。
    4. 第二遍扫描 A 表：
       为了彻底解决乱序与尾差问题，维护一张字典记录 "该 key 还剩多少钱可分，还剩多少权重未处理"。
       对于该 key 下的每一行：
       - 若已经是最后一行（剩余权重 == 当前权重），直接将 "剩余的钱" 作为该行分得的结果。（完美吸收所有舍入尾差）
       - 否则，该行结果 = 剩余的钱 * ( 当前权重 / 剩余权重 )，进行精确 round(decimal_places)
       - 将分配出的钱从 "剩余的钱" 扣除，当前权重从 "剩余权重" 扣除
       * 如果全组权重都是 0，则转为按剩余记录数平均分法。
    """
    path_a = Path(file_a)
    path_b = Path(file_b)
    out = Path(output_path)

    if not path_a.exists():
        print(f"❌ 找不到文件 A: {path_a}")
        return
    if not path_b.exists():
        print(f"❌ 找不到文件 B: {path_b}")
        return

    # ── 1. 读取文件 B，建立 key -> value_b 映射 ──────────────────────────────
    print("📄 正在索引文件 B...")
    wb_b = load_workbook(path_b, data_only=True, read_only=True)
    ws_b = wb_b.active
    header_b = []
    for row in ws_b.iter_rows(min_row=1, max_row=1, values_only=True):
        header_b = list(row)
        break

    try:
        idx_key_b = header_b.index(key_b_name)
        idx_val_b = header_b.index(value_b_name)
    except ValueError as e:
        print(f"❌ 文件 B 中缺少关键字段: {e}")
        wb_b.close()
        return

    map_b_value = {}   # key (str) -> sum of value (Decimal)
    map_b_row = {}     # key (str) -> first matched row tuple 
    b_count = 0
    
    for row in ws_b.iter_rows(min_row=2, values_only=True):
        kb = row[idx_key_b]
        if kb is None:
            continue
        kb = str(kb)
        vb = row[idx_val_b]
        try:
            vb_val = Decimal(str(vb)) if vb is not None else Decimal("0")
        except:
            vb_val = Decimal("0")
            
        map_b_value[kb] = map_b_value.get(kb, Decimal("0")) + vb_val
        if kb not in map_b_row:
            map_b_row[kb] = row
        b_count += 1
    wb_b.close()
    print(f"✅ 文件 B 索引完成，共 {b_count} 行有效数据，{len(map_b_value)} 个唯一关联键。")

    # ── 2. 第一遍扫描 A 文件：统计总权重和行数 ─────────────────────────────────
    print("📄 第一遍扫描文件 A（统计各组权重总和）...")
    wb_a1 = load_workbook(path_a, data_only=True, read_only=True)
    ws_a1 = wb_a1.active
    header_a = []
    for row in ws_a1.iter_rows(min_row=1, max_row=1, values_only=True):
        header_a = list(row)
        break

    try:
        idx_key_a = header_a.index(key_a_name)
        idx_wt_a = header_a.index(weight_a_name)
    except ValueError as e:
        print(f"❌ 文件 A 中缺少关键字段: {e}")
        wb_a1.close()
        return

    group_stats = {}  # key -> {"sum_wt": Decimal, "count": int}
    for row in ws_a1.iter_rows(min_row=2, values_only=True):
        ka = row[idx_key_a]
        if ka is None:
            continue
        ka = str(ka)
        if ka not in map_b_value:
            continue
            
        wt = row[idx_wt_a]
        try:
            wt_val = Decimal(str(wt)) if wt is not None else Decimal("0")
        except:
            wt_val = Decimal("0")
            
        if ka not in group_stats:
            group_stats[ka] = {"sum_wt": Decimal("0"), "count": 0}
            
        group_stats[ka]["sum_wt"] += wt_val
        group_stats[ka]["count"] += 1
    wb_a1.close()
    print(f"✅ 权重统计完成，涉及 {len(group_stats)} 个匹配关联键。")

    # ── 3. 构建输出字段映射 ────────────────────────────────────────────────────
    field_map = {}
    for i, h in enumerate(header_a):
        if h is not None:
            field_map[str(h)] = ("A", i)
    for i, h in enumerate(header_b):
        if h is not None:
            field_map[str(h)] = ("B", i)
    field_map[result_col] = ("RES", None)

    if output_columns:
        final_header = [c for c in output_columns if c in field_map]
        if not final_header:
            print("⚠️ 指定的输出列均无效，将使用默认输出（A 所有列 + 结果列）。")
            final_header = [h for h in header_a if h is not None] + [result_col]
    else:
        final_header = [h for h in header_a if h is not None] + [result_col]

    # ── 4. 第二遍扫描 A 文件：流式分配并拦截尾差 ──────────────────────────────────
    if out.is_dir():
        out = out / f"Alloc_{path_a.name}"
    final_out_path = get_unique_path(out)

    # 状态数据
    rem_b = {k: v for k, v in map_b_value.items()}
    rem_wt = {k: stat["sum_wt"] for k, stat in group_stats.items()}
    rem_count = {k: stat["count"] for k, stat in group_stats.items()}
    
    unit = Decimal(10) ** -decimal_places

    print("📄 第二遍扫描文件 A（动态余额分配写出）...")
    wb_a2 = load_workbook(path_a, data_only=True, read_only=True)
    ws_a2 = wb_a2.active

    processed = 0
    skipped = 0

    try:
        wb_out = Workbook(write_only=True)
        ws_out = wb_out.create_sheet("Result")
        ws_out.append(final_header)

        for row in ws_a2.iter_rows(min_row=2, values_only=True):
            ka = row[idx_key_a]
            if ka is None:
                skipped += 1
                continue
            ka = str(ka)

            if ka not in map_b_value:
                # 以 B 表为基准，如果 B 表没有记录，就不输出该行
                skipped += 1
                continue

            row_b = map_b_row.get(ka)
            # 获取该行独立权重
            wt = row[idx_wt_a]
            try:
                wt_val = Decimal(str(wt)) if wt is not None else Decimal("0")
            except:
                wt_val = Decimal("0")
            
            sum_wt_orig = group_stats[ka]["sum_wt"]
            b_now = rem_b[ka]
            w_now = rem_wt[ka]
            c_now = rem_count[ka]

            dec_res = None
            if sum_wt_orig != 0:
                # 如果有非 0 权重：按权重比例扣除
                if w_now == wt_val:
                    # 这是记录中的最后一条符合条件的权重！吸收全部剩余余额防尾差
                    dec_res = b_now
                else:
                    if w_now != 0:
                        raw = b_now * wt_val / w_now
                        dec_res = raw.quantize(unit, rounding=ROUND_HALF_UP)
                    else:
                        dec_res = Decimal("0")
                rem_wt[ka] -= wt_val
            else:
                # 全组所有权重求和都是 0：按总行数平均划分
                if c_now == 1:
                    # 该组只剩最终 1 个记录
                    dec_res = b_now
                else:
                    if c_now != 0:
                        raw = b_now / Decimal(str(c_now))
                        dec_res = raw.quantize(unit, rounding=ROUND_HALF_UP)
                    else:
                        dec_res = Decimal("0")
            
            rem_b[ka] -= dec_res if dec_res is not None else Decimal("0")
            rem_count[ka] -= 1
            
            # 转换 Decimal 为 float 写入 Excel（保持可计算性）
            res_val = float(dec_res) if dec_res is not None else 0.0

            # 构建并追加行组合
            out_row = []
            for col_name in final_header:
                source, idx = field_map[col_name]
                if source == "A":
                    out_row.append(row[idx] if idx < len(row) else None)
                elif source == "B":
                    out_row.append(row_b[idx] if (row_b and idx < len(row_b)) else None)
                elif source == "RES":
                    out_row.append(res_val)
                    
            ws_out.append(out_row)
            processed += 1

            if processed % 1000 == 0:
                print(f"🚀 已处理 {processed} 行...")

        wb_a2.close()
        print(f"💾 正在写出结果至: {final_out_path.name}...")
        wb_out.save(final_out_path)
        print(
            f"✅ 处理完成！共写出 {processed} 行数据"
            + (f"，跳过 {skipped} 行（关联键为空）" if skipped else "")
            + f"。\n📁 输出文件: {final_out_path}"
        )
    finally:
        if "wb_out" in locals():
            wb_out.close()
        if "wb_a2" in locals():
            wb_a2.close()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="按比例分配引擎：将 B 文件的值按 A 文件权重比例分配到 A 的每一行。"
    )
    parser.add_argument("--file_a", required=True, help="Excel A 文件路径")
    parser.add_argument("--file_b", required=True, help="Excel B 文件路径")
    parser.add_argument("--out", required=True, help="输出 Excel 路径")
    parser.add_argument("--ka", required=True, help="A 文件关联键列名")
    parser.add_argument("--wa", required=True, help="A 文件权重列名（用于计算比例）")
    parser.add_argument("--kb", required=True, help="B 文件关联键列名")
    parser.add_argument("--vb", required=True, help="B 文件待分配值列名")
    parser.add_argument("--res", default="分配结果", help="输出结果列名")
    parser.add_argument("--cols", default=None, help="输出列名，逗号分隔（可任意混合 A/B 字段）")
    parser.add_argument("--decimals", type=int, default=2, help="结果保留小数位数（默认 2）")

    args = parser.parse_args()

    out_cols = None
    if args.cols:
        out_cols = [c.strip() for c in args.cols.split(",") if c.strip()]

    excel_proportional_alloc(
        file_a=args.file_a,
        file_b=args.file_b,
        output_path=args.out,
        key_a_name=args.ka,
        weight_a_name=args.wa,
        key_b_name=args.kb,
        value_b_name=args.vb,
        result_col=args.res,
        output_columns=out_cols,
        decimal_places=args.decimals,
    )
