"""
佣金电子签日报表生成器
======================
重构自老系统：
  - Fbasics.py      (入口逻辑)
  - DataBasics.py   (工具函数 is_nested_list)
  - Xlsopera.py     (核心函数 accounts_daily_report / copy_excel_file)

业务流程：
  1. 读取报账单台账 Excel，按"是否已报账"列分组
  2. 提取每组的"财辅报账单号"列表
  3. 遍历每组，从佣金支付日报表中筛选对应数据
  4. 累加"实付金额(元)"，以总金额命名输出文件
  5. 拷贝模板文件 → 追加写入筛选数据（从指定行开始，无表头）

用法示例：
  python scripts/accounts_daily_report.py \
      -b "D:/数据/bzlist202603.xlsx" \
      -s "D:/数据/" \
      -d "佣金支付日报表.xlsx" \
      -t "GX-SF006佣金外包费支付日报表202604.xlsx"
"""

import argparse
import os
import shutil
import time
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook


# ---------------------------------------------------------------------------
# 工具函数（内联自 DataBasics.zhcDataOpera）
# ---------------------------------------------------------------------------

def is_nested_list(lst: list) -> bool:
    """
    判断列表是否为嵌套列表（即列表中包含至少一个子列表）。
    内联自 DataBasics.zhcDataOpera.is_nested_list()。
    """
    return any(isinstance(item, list) for item in lst)


# ---------------------------------------------------------------------------
# Excel 操作函数（内联自 Xlsopera.zhcXlsOpera）
# ---------------------------------------------------------------------------

def copy_excel_file(source_path: str, target_path: str) -> None:
    """
    用 openpyxl 拷贝 Excel 文件（保留格式/样式）。
    内联自 Xlsopera.zhcXlsOpera.copy_excel_file()。
    """
    workbook = load_workbook(source_path)
    workbook.save(target_path)


def accounts_daily_report(
    bzbm_list: list,
    base_path: str,
    accounts_daily_file: str,
    sign_report_file: str,
    group_col: str = "是否已报账",
    id_col: str = "财辅报账单号",
    usecols: list = None,
    amount_col: str = "实付金额(元)",
    startrow: int = 3,
    skiprows: int = 1,
    sheet_name_to_read: str | int = 0,
) -> dict:
    """
    按报账单号列表从日报表中筛选数据，并写入以总金额命名的电子签报表。

    重构自 Xlsopera.zhcXlsOpera.accounts_daily_report()。
    修复原始 Bug：原代码用相对路径 file_output_name 判断文件是否存在，
    应使用完整路径 file_path，否则在非当前目录运行时始终走"不存在"分支。

    参数：
        bzbm_list         : 本批次报账单号列表
        base_path         : 日报表所在目录（绝对路径）
        accounts_daily_file : 佣金支付日报表文件名
        sign_report_file  : 电子签模板文件名
        group_col         : 报账单台账中的分组列名（默认"是否已报账"）
        id_col            : 报账单号列名（默认"财辅报账单号"）
        usecols           : 从日报表中读取的列名列表，None 则读取全部列
        amount_col        : 累加金额的列名（默认"实付金额(元)"）
        startrow          : 数据写入目标 Sheet 的起始行（0-indexed，默认 3）
        skiprows          : 读取日报表时跳过的行数（默认 1，跳过脏数据首行）
        sheet_name_to_read: 读取日报表时的Sheet名称或索引（默认 0）

    返回：
        包含执行结果信息的字典
    """
    start_time = time.process_time()

    # 1. 读取日报表
    source_daily_path = os.path.join(base_path, accounts_daily_file)
    xsd_info = pd.read_excel(
        source_daily_path,
        sheet_name=sheet_name_to_read,
        engine="openpyxl",
        skiprows=skiprows,
        usecols=usecols,  # None 时读全部列
    )

    # 2. 按报账单号筛选
    xsd_read = xsd_info[xsd_info[id_col].isin(bzbm_list)]

    # 3. 累加实付金额
    a_sum = int(xsd_read[amount_col].sum())

    # 4. 构造输出文件名（模板名去扩展名 + 总金额 + .xlsx）
    sign_report_stem = os.path.splitext(sign_report_file)[0]
    output_filename = f"{sign_report_stem}-{a_sum}.xlsx"
    file_path = os.path.join(base_path, output_filename)  # ✅ 完整路径（修复原始 bug）

    # 5. 如输出文件不存在则先拷贝模板，再追加写入数据
    sheet_name = "Sheet1"
    if not os.path.exists(file_path):  # ✅ 用完整路径判断（原始 bug 修复点）
        source_template_path = os.path.join(base_path, sign_report_file)
        copy_excel_file(source_template_path, file_path)
        is_exist = False
    else:
        is_exist = True

    with pd.ExcelWriter(file_path, mode="a", if_sheet_exists="overlay", engine="openpyxl") as writer:
        xsd_read.to_excel(writer, sheet_name=sheet_name, startrow=startrow, index=False, header=None)

    stop_time = time.process_time()

    return {
        "output_file": file_path,
        "total_amount": a_sum,
        "matched_rows": len(xsd_read),
        "file_existed": is_exist,
        "elapsed_sec": round(stop_time - start_time, 4),
    }


# ---------------------------------------------------------------------------
# 主入口（重构自 Fbasics.py）
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="佣金电子签日报表生成器：按报账单台账分组，从日报表筛选数据并写入电子签模板。"
    )
    parser.add_argument(
        "-b", "--bzlist",
        required=True,
        help="报账单台账 Excel 文件路径（如 bzlist202603.xlsx）",
    )
    parser.add_argument(
        "-s", "--src_dir",
        required=True,
        help="日报表和模板文件所在目录路径",
    )
    parser.add_argument(
        "-d", "--daily",
        required=True,
        help="佣金支付日报表文件名（如 佣金支付日报表.xlsx）",
    )
    parser.add_argument(
        "-t", "--template",
        required=True,
        help="电子签模板文件名（如 GX-SF006佣金外包费支付日报表202604.xlsx）",
    )
    parser.add_argument(
        "-g", "--group_col",
        default="是否已报账",
        help="报账单台账中的分组列名（默认：是否已报账）",
    )
    parser.add_argument(
        "-i", "--id_col",
        default="财辅报账单号",
        help="报账单号列名（默认：财辅报账单号）",
    )
    parser.add_argument(
        "--amount_col",
        default="实付金额(元)",
        help="累加金额的列名（默认：实付金额(元)）",
    )
    parser.add_argument(
        "--startrow",
        type=int,
        default=3,
        help="数据写入目标 Sheet 的起始行，0-indexed（默认：3，即第4行）",
    )
    parser.add_argument(
        "--skiprows",
        type=int,
        default=1,
        help="读取日报表时跳过的行数（默认：1，跳过脏数据首行）",
    )

    parser.add_argument(
        "--sheet_name",
        default=0,
        help="读取日报表时的Sheet名称或索引（默认：0，即第一个Sheet）",
    )

    args = parser.parse_args()

    # 将 sheet_name 尝试转换为 int（如果是数字索引的话）
    sheet_name = args.sheet_name
    if isinstance(sheet_name, str) and sheet_name.isdigit():
        sheet_name = int(sheet_name)

    # 校验输入文件
    if not os.path.isfile(args.bzlist):
        print(f"❌ 报账单台账文件不存在：{args.bzlist}")
        return
    if not os.path.isdir(args.src_dir):
        print(f"❌ 源目录不存在：{args.src_dir}")
        return

    # 读取报账单台账并按分组列分组
    print(f"📖 读取报账单台账：{args.bzlist}")
    bzlist_info = pd.read_excel(args.bzlist)

    if args.group_col not in bzlist_info.columns:
        print(f"❌ 报账单台账中不存在分组列：'{args.group_col}'")
        print(f"   现有列：{list(bzlist_info.columns)}")
        return

    if args.id_col not in bzlist_info.columns:
        print(f"❌ 报账单台账中不存在单号列：'{args.id_col}'")
        print(f"   现有列：{list(bzlist_info.columns)}")
        return

    # 提取各分组的报账单号列表（list of list）
    bzbm_list = [
        list(group[args.id_col])
        for _, group in bzlist_info.groupby(args.group_col)
    ]

    print(f"📊 共发现 {len(bzbm_list)} 个分组")

    # 遍历每组，生成报表
    if is_nested_list(bzbm_list):
        for idx, item in enumerate(bzbm_list):
            print(f"\n▶ 处理第 {idx} 组，共 {len(item)} 条报账单号...")
            result = accounts_daily_report(
                bzbm_list=item,
                base_path=args.src_dir,
                accounts_daily_file=args.daily,
                sign_report_file=args.template,
                id_col=args.id_col,
                amount_col=args.amount_col,
                startrow=args.startrow,
                skiprows=args.skiprows,
                sheet_name_to_read=sheet_name,
            )
            _print_result(result)
    else:
        # 非嵌套时作为单组处理
        flat_list = [item for sublist in bzbm_list for item in sublist]
        print(f"\n▶ 单组模式，共 {len(flat_list)} 条报账单号...")
        result = accounts_daily_report(
            bzbm_list=flat_list,
            base_path=args.src_dir,
            accounts_daily_file=args.daily,
            sign_report_file=args.template,
            id_col=args.id_col,
            amount_col=args.amount_col,
            startrow=args.startrow,
            skiprows=args.skiprows,
            sheet_name_to_read=sheet_name,
        )
        _print_result(result)

    print("\n✅ 全部分组处理完毕！")


def _print_result(result: dict) -> None:
    """格式化打印单次处理结果。"""
    status = "（已存在，追加写入）" if result["file_existed"] else "（新建自模板）"
    print(f"   💰 实付总金额：{result['total_amount']:,} 元")
    print(f"   📄 匹配数据行：{result['matched_rows']} 行")
    print(f"   📁 输出文件  ：{result['output_file']} {status}")
    print(f"   ⏱ 耗时       ：{result['elapsed_sec']} 秒")


if __name__ == "__main__":
    main()
