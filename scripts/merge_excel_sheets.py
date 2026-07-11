from utils.file_utils import FileUtils
from utils.excel_utils import ExcelUtils
from utils.error_handler import ErrorHandler
from utils.cli_parser import CLIParser
import argparse
import sys
from pathlib import Path
import openpyxl
import xlrd

def merge_sheets(src_dir: str, out_file: str, target_sheet: str):
    src_path = FileUtils.resolve_path(src_dir)
    out_path = FileUtils.resolve_path(out_file)

    if not FileUtils.validate_directory_exists(src_path):
        ErrorHandler.handle_directory_not_found(src_dir)
        return

    out_path = FileUtils.get_output_path(src_path, out_path)
    out_path = FileUtils.handle_existing_file(out_path)

    # 使用 write_only=True 极致减少内存占用
    out_wb = ExcelUtils.create_output_workbook()

    existing_sheet_names = set()

    ErrorHandler.print_info(f"🚀 开始在大数据量模式下合并 Sheet: '{target_sheet}'...")

    processed_count = 0
    skipped_count = 0
    excel_files = list(src_path.rglob("*.xls")) + list(src_path.rglob("*.xlsx"))

    for file_path in excel_files:
        if file_path == out_path:
            continue

        try:
            if file_path.suffix.lower() == '.xlsx':
                row_gen = ExcelUtils.process_xlsx_gen(file_path, target_sheet)
                if row_gen is None:
                    ErrorHandler.print_skip(f"'{file_path.name}' (未找到目标 Sheet)")
                    skipped_count += 1
                    continue

                valid_name = ExcelUtils.get_valid_sheet_name(file_path.stem, existing_sheet_names)
                existing_sheet_names.add(valid_name)

                ws = ExcelUtils.create_output_sheet(out_wb, valid_name)
                # 使用生成器流式写入，避免内存积压
                for row_data in row_gen:
                    ws.append(row_data)

            elif file_path.suffix.lower() == '.xls':
                rows = ExcelUtils.process_xls(file_path, target_sheet)
                if rows is None:
                    ErrorHandler.print_skip(f"'{file_path.name}' (未找到目标 Sheet)")
                    skipped_count += 1
                    continue

                valid_name = ExcelUtils.get_valid_sheet_name(file_path.stem, existing_sheet_names)
                existing_sheet_names.add(valid_name)
                ws = ExcelUtils.create_output_sheet(out_wb, valid_name)
                for row_data in rows:
                    ws.append(row_data)

            ErrorHandler.print_success(f"成功合并 [Streaming]: '{file_path.name}'")
            processed_count += 1

        except Exception as e:
            ErrorHandler.handle_excel_processing_error(file_path.name, e)
            skipped_count += 1

    if processed_count == 0:
        ErrorHandler.print_error("\n未能成功提取任何数据。")
        return

    ErrorHandler.print_info(f"\n💾 正在保存至: {out_path.name} ...")
    try:
        ExcelUtils.save_workbook(out_wb, out_path)
        ErrorHandler.print_success(f"并行合并圆满成功！共计包含 {processed_count} 个分离子表。")
    except Exception as e:
        ErrorHandler.handle_file_save_error(out_path.name, e)

if __name__ == "__main__":
    parser = CLIParser.create_parser("性能隔离优化版：多文档 Sheet 提取与物理合并。")
    parser.add_argument("-s", "--src", required=True, help="源文件夹路径")
    parser.add_argument("-o", "--out", required=True, help="输出文件路径")
    CLIParser.add_sheet_arguments(parser)
    args = CLIParser.parse_arguments(parser)
    merge_sheets(args.src, args.out, args.sheet)