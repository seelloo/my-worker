import os
import csv
import argparse
import logging
from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell

logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

def detect_encoding(file_path):
    encodings = ['utf-8', 'gbk', 'gb2312', 'utf-8-sig']
    for enc in encodings:
        try:
            with open(file_path, 'r', encoding=enc) as f:
                f.read(1024)
            return enc
        except UnicodeDecodeError:
            continue
    return 'utf-8'

def convert_csvs_to_excel(src_dir, out_file):
    if not os.path.isdir(src_dir):
        logging.error(f"源目录不存在: {src_dir}")
        return False
        
    csv_files = [f for f in os.listdir(src_dir) if f.lower().endswith('.csv')]
    if not csv_files:
        logging.warning(f"目录中未找到 CSV 文件: {src_dir}")
        return False
        
    # Use write_only=True for memory efficiency
    wb = Workbook(write_only=True)
    
    for filename in csv_files:
        filepath = os.path.join(src_dir, filename)
        # Sheetnames can't be > 31 characters
        sheet_name = os.path.splitext(filename)[0][:31]
        
        # Some characters are not allowed in Excel sheet names
        for invalid_char in ['\\', '/', '?', '*', '[', ']', ':']:
            sheet_name = sheet_name.replace(invalid_char, '_')
            
        if not sheet_name:
            sheet_name = "Sheet_" + filename[:10]
            
        ws = wb.create_sheet(title=sheet_name)
        
        encoding = detect_encoding(filepath)
        logging.info(f"正在处理 {filename} (使用编码 {encoding})")
        
        try:
            with open(filepath, 'r', encoding=encoding, newline='') as f:
                reader = csv.reader(f)
                for row_data in reader:
                    # Create formatted cells for each row
                    row_cells = []
                    for val in row_data:
                        cell = WriteOnlyCell(ws, value=str(val))
                        cell.number_format = '@' # Force text format
                        row_cells.append(cell)
                    ws.append(row_cells)
        except Exception as e:
            logging.error(f"处理文件 {filename} 时出错: {e}")
            
    try:
        logging.info(f"正在保存 Excel 文件: {out_file}")
        wb.save(out_file)
        wb.close()
        logging.info(f"成功将 {len(csv_files)} 个 CSV 文件导出到: {out_file}")
        return True
    except Exception as e:
        logging.error(f"保存 Excel 文件失败: {e}")
        return False
    finally:
        wb.close()

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="批量合并 CSV 到 Excel 并设置文本格式")
    parser.add_argument("-s", "--src_dir", required=True, help="包含 CSV 文件的源目录")
    parser.add_argument("-o", "--out_file", required=True, help="输出 Excel 路径")
    args = parser.parse_args()
    
    convert_csvs_to_excel(args.src_dir, args.out_file)
