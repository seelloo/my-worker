"""
split_excel_to_single_sheets.py
================================
将一个 Excel 文件的多个 Sheet 拆分成一个个独立的 Excel 文件。
每个文件只包含一个 Sheet 的数据，文件名以原始 Sheet 名称命名。
"""

import os
import sys
import argparse
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    print("❌ 缺少依赖：请执行 pip install pandas openpyxl")
    sys.exit(1)

def split_to_single_sheets(input_file: str, output_dir: str):
    input_path = Path(input_file)
    out_dir = Path(output_dir)

    if not input_path.exists() or not input_path.is_file():
        print(f"❌ 找不到源文件: {input_path}")
        sys.exit(1)

    out_dir.mkdir(parents=True, exist_ok=True)
    
    print(f"📂 开始读取文件: {input_path.name}")
    try:
        xl = pd.ExcelFile(input_path, engine='openpyxl')
        sheet_names = xl.sheet_names
    except Exception as e:
        print(f"❌ 读取 Excel 文件失败: {e}")
        sys.exit(1)

    print(f"📋 发现 {len(sheet_names)} 个 Sheet: {sheet_names}")
    
    success_count = 0
    total = len(sheet_names)
    
    for idx, sheet_name in enumerate(sheet_names, 1):
        print(f"\n[PROGRESS: {int(idx / total * 90)}%]")
        
        # 清理 Sheet 名中不能作为文件名的字符
        safe_name = "".join([c for c in sheet_name if c not in r'\/:*?"<>|'])
        if not safe_name:
            safe_name = f"Sheet_{idx}"
            
        out_file = out_dir / f"{safe_name}.xlsx"
        
        print(f"  📝 正在提取 Sheet「{sheet_name}」 → {out_file.name}")
        try:
            df = xl.parse(sheet_name, dtype=str)  # 强制作为字符串读取，防止数据格式丢失
            df.to_excel(out_file, index=False, engine='openpyxl')
            success_count += 1
        except Exception as e:
            print(f"  ❌ 提取失败: {e}")

    print(f"\n[PROGRESS: 100%]")
    print(f"\n✅ 拆分完成！成功生成 {success_count}/{total} 个文件，输出至：{out_dir}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="将一个 Excel 文件的多个 Sheet 拆分成一个个独立的 Excel 文件")
    parser.add_argument("-i", "--input", required=True, help="源 Excel 文件路径")
    parser.add_argument("-o", "--output_dir", required=True, help="输出目录路径")
    
    args = parser.parse_args()
    split_to_single_sheets(args.input, args.output_dir)
