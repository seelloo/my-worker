"""
Excel 数值列条件分组工具

根据用户定义的多条分组规则，对指定数值列进行分组编号。
每条规则包含 1~3 个筛选条件（AND 逻辑）和每组记录数。
系统按规则顺序处理，全局组号递增，未匹配行标记为「未分组」。
"""

import os
import argparse
import json
import random
from pathlib import Path
from openpyxl import load_workbook, Workbook
from openpyxl.utils import column_index_from_string
import re


def resolve_column_index(col_name, header):
    """
    将用户输入的列标识解析为 0-based 列索引。
    支持：表头名称 / Excel 列字母 (A, B, AA) / 格式 "B (金额)"
    """
    col_name = col_name.strip()
    if not col_name:
        return None

    # 建立表头名称到索引的映射
    col_to_idx = {str(h).strip(): i for i, h in enumerate(header) if h is not None}

    # 1. 直接匹配表头名称
    if col_name in col_to_idx:
        return col_to_idx[col_name]

    # 2. 尝试解析 "B (金额)" 格式
    match = re.match(r"^([A-Z]+)\s*\((.+)\)$", col_name, re.I)
    if match:
        letter, name = match.groups()
        if name.strip() in col_to_idx:
            return col_to_idx[name.strip()]
        try:
            idx = column_index_from_string(letter.upper()) - 1
            if 0 <= idx < len(header):
                return idx
        except Exception:
            pass

    # 3. 纯字母列标识 (A, B, AA...)
    if re.match(r"^[A-Za-z]+$", col_name):
        try:
            idx = column_index_from_string(col_name.upper()) - 1
            if 0 <= idx < len(header):
                return idx
        except Exception:
            pass

    return None


def evaluate_condition(cell_value, op, target_value):
    """
    判断单元格值是否满足单个条件。
    cell_value: 单元格原始值
    op: 运算符 (gt, gte, lt, lte, eq, neq)
    target_value: 目标数值
    """
    if cell_value is None:
        return False
    try:
        num = float(cell_value)
        target = float(target_value)
    except (ValueError, TypeError):
        return False

    if op == "gt":
        return num > target
    elif op == "gte":
        return num >= target
    elif op == "lt":
        return num < target
    elif op == "lte":
        return num <= target
    elif op == "eq":
        return num == target
    elif op == "neq":
        return num != target
    return False


def check_rule_match(cell_value, conditions):
    """
    检查单元格值是否满足一条规则的所有条件（AND 逻辑）。
    conditions: [{op, value}, ...] 长度 1~3
    """
    for cond in conditions:
        op = cond.get("op", "").strip()
        val = cond.get("value")
        if not op or val is None or str(val).strip() == "":
            continue  # 跳过空条件（可选条件未填）
        if not evaluate_condition(cell_value, op, val):
            return False
    return True


def group_assign(src_path_str, out_path_str, target_col, rules_json, group_col_name="组编码", concat_col=None):
    """
    核心分组逻辑。

    :param src_path_str: 源 Excel 路径
    :param out_path_str: 输出 Excel 路径
    :param target_col: 目标数值列标识
    :param rules_json: 规则 JSON 字符串或已解析列表
    :param group_col_name: 新增的组编码列名
    :param concat_col: 需要在第二个 Sheet 中汇总拼接的列标识
    """
    src_path = Path(src_path_str).resolve()
    out_path = Path(out_path_str).resolve()

    # 解析规则
    if isinstance(rules_json, str):
        try:
            rules = json.loads(rules_json)
        except Exception as e:
            print(f"❌ 规则 JSON 解析失败: {e}")
            return
    else:
        rules = rules_json

    if not rules:
        print("❌ 未提供分组规则。")
        return

    if not src_path.exists():
        print(f"❌ 源文件不存在: {src_path}")
        return

    print(f"📄 正在读取源文件: {src_path.name}")
    wb = load_workbook(src_path, data_only=True, read_only=True)
    ws = wb.active

    # 读取表头
    header = []
    for row in ws.iter_rows(min_row=1, max_row=1, values_only=True):
        header = list(row)
        break

    if not header:
        print("❌ 无法识别表头或文件为空。")
        wb.close()
        return

    # 解析目标列索引
    col_idx = resolve_column_index(target_col, header)
    if col_idx is None:
        print(f"❌ 无法识别目标列「{target_col}」。可用列：{[str(h) for h in header if h]}")
        wb.close()
        return

    print(f"🎯 目标列: [{target_col}] → 第 {col_idx + 1} 列 (表头: {header[col_idx]})")
    print(f"📋 共 {len(rules)} 条分组规则")

    # 读取所有数据行
    all_rows = []
    for row_values in ws.iter_rows(min_row=2, values_only=True):
        all_rows.append(list(row_values))
    wb.close()

    total = len(all_rows)
    if total == 0:
        print("⚠️ 源文件没有数据行。")
        return

    print(f"📊 共读取 {total} 行数据")

    # 全局可用行索引（0-based）
    available_indices = set(range(total))
    
    # 分配组号
    group_codes = [None] * total
    current_group = 1

    for rule_idx, rule in enumerate(rules):
        conditions = rule.get("conditions", [])
        group_size = max(1, int(rule.get("group_size", 1)))
        sum_threshold = rule.get("sum_threshold")
        sum_op = rule.get("sum_op", "gt")
        
        # 过滤掉空条件
        valid_conditions = [c for c in conditions if c.get("op") and c.get("value") is not None and str(c.get("value")).strip() != ""]
        
        if conditions and not valid_conditions:
            print(f"⚠️ 规则 {rule_idx + 1} 没有有效条件，已跳过。")
            continue

        # 候选行为全部尚未分组的行
        rule_candidates = sorted(list(available_indices))
                
        matched_groups = 0
        matched_rows_count = 0
        
        if valid_conditions:
            # 随机总和搜索模式：条件作用于组的数值总和
            max_attempts = 5000
            attempts = 0
            
            while len(rule_candidates) >= group_size:
                # 随机抽取候选
                sampled = random.sample(rule_candidates, group_size)
                
                # 计算总和
                total_sum = 0
                valid_sum = True
                for idx in sampled:
                    val = all_rows[idx][col_idx] if col_idx < len(all_rows[idx]) else None
                    if val is None:
                        valid_sum = False
                        break
                    try:
                        total_sum += float(val)
                    except (ValueError, TypeError):
                        valid_sum = False
                        break
                        
                # 判断总和是否满足条件
                if valid_sum and check_rule_match(total_sum, valid_conditions):
                    # 匹配成功，分配组号
                    for idx in sampled:
                        group_codes[idx] = current_group
                        available_indices.remove(idx)
                        rule_candidates.remove(idx)
                    current_group += 1
                    matched_groups += 1
                    matched_rows_count += group_size
                    attempts = 0 # 重置尝试次数
                else:
                    attempts += 1
                    if attempts >= max_attempts:
                        print(f"  ⚠️ 规则 {rule_idx + 1}: 尝试 {max_attempts} 次后无法再找到满足总和条件的组合。")
                        break
            
            print(f"  规则 {rule_idx + 1}: 匹配 {matched_rows_count} 行 → 生成了 {matched_groups} 个组 (每组 {group_size} 条, 总和需满足条件)")
        else:
            # 原有的顺序分配模式 (无条件时纯按 size 分组)
            for chunk_start in range(0, len(rule_candidates), group_size):
                chunk = rule_candidates[chunk_start:chunk_start + group_size]
                for idx in chunk:
                    group_codes[idx] = current_group
                    available_indices.remove(idx)
                current_group += 1
                matched_groups += 1
                matched_rows_count += len(chunk)
            print(f"  规则 {rule_idx + 1}: 匹配 {matched_rows_count} 行 → 生成了 {matched_groups} 个组 (每组 {group_size} 条)")

    # 未匹配行标记为「未分组」
    unmatched = sum(1 for g in group_codes if g is None)
    matched_total = total - unmatched
    total_groups = current_group - 1

    for i in range(total):
        if group_codes[i] is None:
            group_codes[i] = "未分组"

    print(f"\n📊 分组统计:")
    print(f"   已分组: {matched_total} 行 → {total_groups} 个组")
    print(f"   未分组: {unmatched} 行")

    # 写出结果
    wb_out = Workbook(write_only=True)
    ws_out = wb_out.create_sheet("分组结果")

    # 写表头（原始表头 + 组编码列）
    out_header = list(header) + [group_col_name]
    ws_out.append(out_header)

    def sort_key(i):
        code = group_codes[i]
        if code == "未分组":
            return (float('inf'), i)
        return (code, i)

    sorted_indices = sorted(range(total), key=sort_key)

    for i in sorted_indices:
        row = all_rows[i]
        # 确保行长度与表头对齐
        padded = list(row) + [None] * max(0, len(header) - len(row))
        padded.append(group_codes[i])
        ws_out.append(padded)

    # 确保输出目录存在
    out_path.parent.mkdir(parents=True, exist_ok=True)

    # 防重名
    if out_path.exists():
        stem = out_path.stem
        ext = out_path.suffix
        counter = 1
        while True:
            new_path = out_path.parent / f"{stem}_{counter}{ext}"
            if not new_path.exists():
                out_path = new_path
                break
            counter += 1

    # 如果指定了 --concat_col，则新增一个汇总 Sheet
    if concat_col:
        concat_col_idx = resolve_column_index(concat_col, header)
        if concat_col_idx is not None:
            ws_summary = wb_out.create_sheet("分组汇总")
            ws_summary.append(["组号", f"{header[concat_col_idx]}汇总"])
            
            # 按组别收集该列的值
            group_values = {}
            for i, row in enumerate(all_rows):
                g = group_codes[i]
                if g != "未分组":
                    val = row[concat_col_idx] if concat_col_idx < len(row) else ""
                    if val is not None:
                        val_str = str(val).strip()
                        if val_str:
                            if g not in group_values:
                                group_values[g] = []
                            group_values[g].append(val_str)
            
            # 按组号排序写入
            for g in sorted(group_values.keys()):
                vals = group_values[g]
                # 单引号包裹，逗号隔开
                joined_str = ",".join(f"'{v}'" for v in vals)
                ws_summary.append([g, joined_str])
            print(f"📄 已生成「分组汇总」Sheet，拼接列：{header[concat_col_idx]}")
        else:
            print(f"⚠️ 无法识别拼接列「{concat_col}」，跳过生成汇总 Sheet。")

    print(f"\n💾 正在保存结果至: {out_path.name} ...")
    wb_out.save(out_path)
    print(f"✅ 分组完成！共 {total} 行数据，{total_groups} 个组，输出至 {out_path}")


def main():
    parser = argparse.ArgumentParser(description="Excel 数值列条件分组工具")
    parser.add_argument("-s", "--src", required=True, help="源 Excel 文件路径")
    parser.add_argument("-o", "--out", required=True, help="输出 Excel 文件路径")
    parser.add_argument("-c", "--col", required=True, help="目标数值列（列名或字母）")
    parser.add_argument("-r", "--rules", required=True, help="规则 JSON 文件路径或 JSON 字符串")
    parser.add_argument("--group_col_name", default="组编码", help="新增的组编码列名")
    parser.add_argument("--concat_col", help="用于生成汇总Sheet的拼接列（列名或字母）")

    args = parser.parse_args()

    # 支持从文件读取规则
    rules_input = args.rules
    if os.path.isfile(rules_input):
        with open(rules_input, "r", encoding="utf-8") as f:
            rules_input = f.read()

    group_assign(args.src, args.out, args.col, rules_input, args.group_col_name, args.concat_col)

if __name__ == "__main__":
    main()
