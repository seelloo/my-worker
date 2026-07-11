import os
import argparse
from pathlib import Path
from openpyxl import load_workbook, Workbook

def get_unique_path(base_path: Path) -> Path:
    """
    检查提供路径是否已存在。如果存在，自动在文件名后追加 _1, _2 等后缀以防覆盖。
    """
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

def get_excel_header(file_path: str) -> list:
    """
    性能优化版：使用 read_only 获取表头
    """
    path = Path(file_path)
    if not path.exists():
        return []
    # 使用 read_only 模式极大减少获取表头时的内存占用
    wb = None
    try:
        wb = load_workbook(path, data_only=True, read_only=True)
        ws = wb.active
        header = []
        # 只取第一行
        for row in ws.iter_rows(min_row=1, max_row=1, values_only=True):
            header = [v for v in row if v is not None]
            break
        return header
    except:
        return []
    finally:
        if wb:
            wb.close()

def excel_join_multiply(file_a: str, file_b: str, output_path: str, 
                        key_a_name: str = 'a', key_b_name: str = 'b', 
                        col_a_name: str = 'c', col_b_name: str = 'c',
                        result_col: str = '乘积结果',
                        output_columns: list = None):
    """
    大数据量优化版：
    1. 使用 read_only=True 加载 A、B 两表，支持千万级数据。
    2. 使用 write_only=True 进行结果流式写入，内存占用恒定。
    3. 哈希索引加速关联计算。
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

    print(f"📄 正在索引文件 B (Hash Indexing)...")
    # B 表通常建议作为较小的表（虽然 read_only 能支持大表，但 map_b 仍占内存）
    wb_b = load_workbook(path_b, data_only=True, read_only=True)
    ws_b = wb_b.active
    
    # 获取表头
    header_b = []
    for row in ws_b.iter_rows(min_row=1, max_row=1, values_only=True):
        header_b = list(row)
        break
        
    try:
        idx_key_b = header_b.index(key_b_name)
        idx_col_b = header_b.index(col_b_name)
    except ValueError as e:
        print(f"❌ 文件 B 中缺少关键字段: {e}")
        wb_b.close()
        return

    # map_b: key -> row_tuple (仅保存计算所需的部分以进一步节省内存)
    map_b = {}
    for row in ws_b.iter_rows(min_row=2, values_only=True):
        kb_val = row[idx_key_b]
        if kb_val is not None:
            # 存储整行 (tuple) 供后续字段挑选
            map_b[str(kb_val)] = row
    
    wb_b.close()
    print(f"✅ 文件 B 索引构建完成，共载入 {len(map_b)} 个映射项。")

    print(f"📄 正在处理文件 A (Streaming Mode)...")
    wb_a = load_workbook(path_a, data_only=True, read_only=True)
    ws_a = wb_a.active
    
    header_a = []
    for row in ws_a.iter_rows(min_row=1, max_row=1, values_only=True):
        header_a = list(row)
        break

    try:
        idx_key_a = header_a.index(key_a_name)
        idx_col_a = header_a.index(col_a_name)
    except ValueError as e:
        print(f"❌ 文件 A 中缺少关键字段: {e}")
        wb_a.close()
        return

    # 结果字段映射索引
    field_map = {}
    for i, h in enumerate(header_a): field_map[h] = ('A', i)
    for i, h in enumerate(header_b): field_map[h] = ('B', i)
    field_map[result_col] = ('RES', None)

    if output_columns:
        final_header = [c for c in output_columns if c in field_map]
    else:
        final_header = header_a + [result_col]

    # 确定输出路径
    if out.is_dir():
        out = out / f"Result_{path_a.name}"
    final_out_path = get_unique_path(out)

    try:
        # 极简模式：流式写入输出
        wb_out = Workbook(write_only=True)
        ws_out = wb_out.create_sheet("Result")
        ws_out.append(final_header)
        
        processed_count = 0
        # 物理隔离环境无法通过 UI 获取 max_row（耗时且在 read_only 下不可靠），我们采用行数计数
        for row_a in ws_a.iter_rows(min_row=2, values_only=True):
            ka_val = str(row_a[idx_key_a]) if row_a[idx_key_a] is not None else None
            ca_val = None
            if row_a[idx_col_a] is not None:
                try: ca_val = float(row_a[idx_col_a])
                except: ca_val = None
            
            row_b = map_b.get(ka_val)
            res_val = None
            
            if row_b:
                cb_val = None
                if row_b[idx_col_b] is not None:
                    try: cb_val = float(row_b[idx_col_b])
                    except: cb_val = None
                
                if ca_val is not None and cb_val is not None:
                    res_val = round(ca_val * cb_val, 2)
            
            # 组装行数据
            out_row = []
            for col_name in final_header:
                source, idx = field_map[col_name]
                if source == 'A': out_row.append(row_a[idx])
                elif source == 'B': out_row.append(row_b[idx] if row_b else None)
                elif source == 'RES': out_row.append(res_val)
            
            ws_out.append(out_row)
            processed_count += 1
            
            if processed_count % 1000 == 0:
                print(f"🚀 已处理 {processed_count} 行...")

        wb_a.close()
        
        print(f"💾 正在封存结果至: {final_out_path.name} ...")
        wb_out.save(final_out_path)
        print(f"✅ 处理圆满完成！累计导出 {processed_count} 行深度关联数据。")
    finally:
        if 'wb_out' in locals():
            wb_out.close()
        if 'wb_a' in locals():
            wb_a.close()

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="性能隔离优化版：多表关联乘积引擎。")
    parser.add_argument("--file_a", required=True)
    parser.add_argument("--file_b", required=True)
    parser.add_argument("--out", required=True)
    parser.add_argument("--ka", default="a")
    parser.add_argument("--kb", default="b")
    parser.add_argument("--ca", default="c")
    parser.add_argument("--cb", default="c")
    parser.add_argument("--res", default="乘积结果")
    parser.add_argument("--cols", default=None, help="输出列名，逗号分隔")
    
    args = parser.parse_args()
    
    # 解析输出列
    out_cols = None
    if args.cols:
        out_cols = [c.strip() for c in args.cols.split(',') if c.strip()]

    excel_join_multiply(args.file_a, args.file_b, args.out, 
                        args.ka, args.kb, args.ca, args.cb, args.res,
                        output_columns=out_cols)
