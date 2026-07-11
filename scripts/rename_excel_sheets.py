from utils.file_utils import FileUtils
from utils.excel_utils import ExcelUtils
from utils.error_handler import ErrorHandler
from utils.cli_parser import CLIParser
import argparse
import sys
from pathlib import Path
import openpyxl

def rename_excel_sheets(src_file: str, dest_path_str: str, remove_str: str = None, add_str: str = None, add_position: str = "back"):
    src_path = FileUtils.resolve_path(src_file)
    dest_path = FileUtils.resolve_path(dest_path_str)

    if not FileUtils.validate_file_exists(src_path):
        ErrorHandler.handle_file_not_found(src_file)
        return

    out_file_path = FileUtils.get_output_path(src_path, dest_path)
    out_file_path = FileUtils.handle_existing_file(out_file_path)

    try:
        # 读取 .xlsx 文件
        wb = openpyxl.load_workbook(src_path)
    except Exception as e:
        ErrorHandler.handle_file_read_error(src_path.name, e)
        return

    existing_names = set()
    renamed_count_in_file = 0

    ErrorHandler.print_info(f"正在处理: '{src_path.name}'")
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]

        # 即使没有任何改变的内容也要加入 existing_names 中占坑，防重名
        new_sheet_name = ExcelUtils.generate_new_sheet_name(
            sheet_name, remove_str, add_str, add_position, existing_names
        )

        if new_sheet_name != sheet_name:
            ws.title = new_sheet_name
            ErrorHandler.print_success(f"   '{sheet_name}' -> '{new_sheet_name}'")
            renamed_count_in_file += 1

        existing_names.add(new_sheet_name)

    try:
        ExcelUtils.save_workbook(wb, out_file_path)
        if renamed_count_in_file > 0:
            ErrorHandler.print_success(f"💾 另存为 -> '{out_file_path.name}' (修改了 {renamed_count_in_file} 个 Sheet)\n")
        else:
            ErrorHandler.print_info(f"👉 另存为 -> '{out_file_path.name}' (无更改)\n")
    except Exception as e:
        ErrorHandler.handle_file_save_error(out_file_path.name, e)

    ErrorHandler.print_success(f"🎉 处理完成！累计重命名了 {renamed_count_in_file} 个 Sheet。")

def main():
    parser = CLIParser.create_parser(
        description="修改指定 Excel (.xlsx) 文件中的所有 Sheet 名称，并将修改后的文件另存。"
    )
    CLIParser.add_common_arguments(parser)
    CLIParser.add_rename_arguments(parser)
    args = CLIParser.parse_arguments(parser)

    CLIParser.validate_rename_arguments(args.remove, args.add)

    rename_excel_sheets(args.file, args.out, args.remove, args.add, args.pos)

if __name__ == "__main__":
    main()
