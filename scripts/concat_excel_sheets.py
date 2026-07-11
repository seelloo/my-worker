import os
import argparse
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

def sanitize_sheet_name(name: str) -> str:
    invalid_chars = ['*', ':', '?', '/', '\\', '[', ']']
    for char in invalid_chars:
        name = name.replace(char, '')
    return name[:31]

def get_unique_sheet_name(wb: Workbook, base_name: str) -> str:
    base_name = sanitize_sheet_name(base_name)
    if not base_name:
        base_name = "Sheet"
    
    existing_names = wb.sheetnames
    if base_name not in existing_names:
        return base_name
        
    counter = 1
    while True:
        # 确保加上后缀后不超过31个字符
        suffix = f"_{counter}"
        max_base_len = 31 - len(suffix)
        new_name = f"{base_name[:max_base_len]}{suffix}"
        if new_name not in existing_names:
            return new_name
        counter += 1

def concat_excel_folder(src_path_str: str, output_path: str, start_row: int = 1, mode: str = "single", target_sheet: str = ""):
    """
    性能隔离优化版：将一个文件夹内所有 Excel 的数据（忽略表头，按行追加）合并成一个超级大表。
    使用 read_only = True 和 write_only = True。
    """
    src_dir = Path(src_path_str)
    out_file = Path(output_path)
    
    if not src_dir.exists():
        print(f"❌ 找不到源目录: {src_dir}")
        return
        
    if out_file.is_dir():
        out_file = out_file / "Merged_Data.xlsx"
    final_out_path = get_unique_path(out_file)

    print(f"🚀 开始全量扫荡目录: {src_dir.name} ...")
    
    # 获取所有待合并文件
    files = [f for f in src_dir.rglob("*.xlsx") if f.is_file() and f != final_out_path]
    if not files:
        print("❌ 未发现任何有效的 .xlsx 文件。")
        return

    # 初始化流式输出
    wb_out = Workbook(write_only=True)
    ws_out = None
    
    if mode != "分Sheet保留原名":
        ws_out = wb_out.create_sheet("Merged")
    
    header_captured = False
    total_rows = 0
    file_count = 0

    for f_path in files:
        wb_src = None
        try:
            print(f"📄 正在吸入文件: {f_path.name} ...")
            wb_src = load_workbook(f_path, data_only=True, read_only=True)
            if mode == "分Sheet保留原名":
                sheet_name = get_unique_sheet_name(wb_out, f_path.stem)
                ws_out = wb_out.create_sheet(sheet_name)
                header_captured = False # 每个文件独立判断表头

            for ws_src in wb_src.worksheets:
                # 若指定了 target_sheet，跳过名称不匹配的 Sheet
                if target_sheet and ws_src.title != target_sheet:
                    continue
                row_idx = 0
                for row_data in ws_src.iter_rows(values_only=True):
                    row_idx += 1
                    
                    # 生成器捕获统一表头
                    if start_row > 1:
                        if row_idx < start_row:
                            if not header_captured:
                                ws_out.append(row_data)
                            continue
                    
                    # 数据行写入
                    if any(v is not None and str(v).strip() != "" for v in row_data):
                        ws_out.append(row_data)
                        total_rows += 1
                        
                header_captured = True # 只要处理过哪怕一个非空的 start_row 以前的部分，就认为表格头已固定
            file_count += 1
        except Exception as e:
            print(f"❌ 读取 '{f_path.name}' 时发生抖动: {e}")
        finally:
            if wb_src:
                wb_src.close()

    if total_rows == 0:
        print("⚠️ 未发现可合并的有效数据。")
        return

    print(f"💾 正在封存合并结果至: {final_out_path.name} ...")
    wb_out.save(final_out_path)
    wb_out.close()
    print(f"✅ 完成！统共扫描了 {file_count} 个文件，累计向下拼接了 {total_rows} 行核心数据。")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="性能隔离版：全文件夹 Excel 物理拼接引擎。")
    parser.add_argument("-s", "--src", required=True, help="源文件夹路径")
    parser.add_argument("-o", "--out", required=True, help="输出 Excel 文件路径")
    parser.add_argument("-r", "--start_row", type=int, default=1, help="有效数据起始行 (从1开始)")
    parser.add_argument("-m", "--mode", default="single", help="合并模式: 'single' 或 '分Sheet保留原名'")
    
    parser.add_argument("-t", "--sheet", default="", help="仅读取指定名称的 Sheet（留空则读所有 Sheet）")
    
    args = parser.parse_args()
    concat_excel_folder(args.src, args.out, args.start_row, args.mode, args.sheet)
