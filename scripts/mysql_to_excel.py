"""
MySQL 数据库表导出为 Excel（强制文本格式）
支持：多表配置、WHERE 条件筛选、单元格强制文本格式
引用已保存的数据库连接 (--connection_name)
"""

import argparse
import logging
import json
import sys
import os
import sqlite3

logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')


def _find_db_path():
    """查找本地保存连接的 SQLite 数据库文件"""
    candidates = [
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "data", "workflows.db"),
        os.path.join(os.getcwd(), "data", "workflows.db"),
    ]
    for p in candidates:
        if os.path.isfile(p):
            return os.path.abspath(p)
    return None


def _parse_excel_tables_config(df):
    """
    从 DataFrame 解析表配置。
    支持的列名（中英文均可）：
      sheet_name/Sheet名, table/表名, columns/列, where/条件, limit/限制行数
    """
    col_map = {}
    for i, col_name in enumerate(df.columns):
        n = str(col_name).strip().lower()
        if n in ('sheet_name', 'sheet名', 'sheet'):
            col_map['sheet_name'] = i
        elif n in ('table', '表名', 'table_name', 'table_name'):
            col_map['table'] = i
        elif n in ('columns', '列', 'cols', 'col'):
            col_map['columns'] = i
        elif n in ('where', '条件', 'filter'):
            col_map['where'] = i
        elif n in ('limit', '限制', '行数', 'limit_rows'):
            col_map['limit'] = i

    if 'table' not in col_map:
        logging.error("表配置 Excel 中未找到「table/表名」列")
        return []

    configs = []
    for _, row in df.iterrows():
        cfg = {"table": str(row.iloc[col_map['table']]).strip()}
        for key, idx in col_map.items():
            if key == 'table':
                continue
            val = row.iloc[idx]
            if pd.notna(val) and str(val).strip():
                cfg[key] = str(val).strip()
        if not cfg.get("sheet_name"):
            cfg["sheet_name"] = cfg["table"]
        configs.append(cfg)

    return configs


def lookup_saved_connection(connection_name: str):
    """从本地 SQLite 数据库中查询已保存的连接信息"""
    db_path = _find_db_path()
    if not db_path:
        return None

    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    try:
        cursor.execute(
            "SELECT * FROM database_connections WHERE name = ?", (connection_name,)
        )
        row = cursor.fetchone()
        if row:
            keys = row.keys()
            return {
                "host": row["host"],
                "port": row["port"],
                "user": row["user"],
                "password": row["password"],
                "database": row["database_name"],
                "charset": row["charset"] if "charset" in keys else "utf8mb4",
                "timeout": row["timeout"] if "timeout" in keys else 30,
            }
        return None
    finally:
        conn.close()

try:
    import pymysql
except ImportError:
    print("❌ 缺少 pymysql 依赖，请运行: pip install pymysql")
    sys.exit(1)

try:
    from openpyxl import Workbook
    from openpyxl.cell import WriteOnlyCell
except ImportError:
    print("❌ 缺少 openpyxl 依赖，请运行: pip install openpyxl")
    sys.exit(1)


def get_connection(db_config):
    """建立 MySQL 连接"""
    return pymysql.connect(
        host=db_config["host"],
        port=db_config.get("port", 3306),
        user=db_config["user"],
        password=db_config["password"],
        database=db_config["database"],
        charset=db_config.get("charset", "utf8mb4"),
        cursorclass=pymysql.cursors.DictCursor,
        connect_timeout=db_config.get("timeout", 30),
    )


def fetch_table_data(conn, table_cfg):
    """从单个表中查询数据"""
    table_name = table_cfg["table"]
    columns = table_cfg.get("columns", "*")
    where = table_cfg.get("where", "").strip()

    if columns == "*" or not columns:
        col_clause = "*"
    else:
        cols = [c.strip() for c in columns.split(",") if c.strip()]
        col_clause = ", ".join(cols)

    sql = f"SELECT {col_clause} FROM `{table_name}`"
    params = []

    if where:
        sql += f" WHERE {where}"

    limit = table_cfg.get("limit")
    if limit and str(limit).strip():
        sql += f" LIMIT {int(limit)}"

    logging.info(f"查询表 [{table_cfg.get('sheet_name', table_name)}]: {sql}")

    with conn.cursor() as cursor:
        cursor.execute(sql, tuple(params) if params else None)
        rows = cursor.fetchall()

    if not rows:
        logging.warning(f"表 [{table_cfg.get('sheet_name', table_name)}] 无数据")
        return [], []

    fieldnames = list(rows[0].keys())
    return fieldnames, rows


def write_to_excel(fieldnames, rows, sheet_name, ws):
    """将数据写入 Excel Sheet（强制文本格式）"""
    max_col = len(fieldnames)

    # 写入表头（也强制文本格式）
    header_cells = []
    for val in fieldnames:
        cell = WriteOnlyCell(ws, value=str(val) if val is not None else "")
        cell.number_format = "@"
        header_cells.append(cell)
    ws.append(header_cells)

    # 写入数据行
    for row_idx, row in enumerate(rows, start=2):
        row_cells = []
        for col_idx, field in enumerate(fieldnames, start=1):
            val = row.get(field)
            if val is None:
                str_val = ""
            elif isinstance(val, (list, dict)):
                str_val = json.dumps(val, ensure_ascii=False)
            else:
                str_val = str(val)
            cell = WriteOnlyCell(ws, value=str_val)
            cell.number_format = "@"
            row_cells.append(cell)
        ws.append(row_cells)

    logging.info(f"Sheet [{sheet_name}] 已写入 {len(rows)} 行 × {max_col} 列")


def export_mysql_to_excel(db_config, tables_cfg, output_file):
    """主流程：连接 MySQL，查询多表数据，导出到 Excel"""
    wb = Workbook(write_only=True)

    conn = get_connection(db_config)
    try:
        for i, table_cfg in enumerate(tables_cfg):
            sheet_name = table_cfg.get("sheet_name", table_cfg["table"])[:31]
            for ch in ['\\', '/', '?', '*', '[', ']', ':']:
                sheet_name = sheet_name.replace(ch, '_')
            if not sheet_name.strip():
                sheet_name = f"Table_{i+1}"

            ws = wb.create_sheet(title=sheet_name)
            fieldnames, rows = fetch_table_data(conn, table_cfg)
            if fieldnames:
                write_to_excel(fieldnames, rows, sheet_name, ws)
            else:
                logging.warning(f"跳过空表: {sheet_name}")

        os.makedirs(os.path.dirname(os.path.abspath(output_file)), exist_ok=True)
        wb.save(output_file)
        logging.info(f"成功导出 {len(tables_cfg)} 张表到: {output_file}")
    finally:
        conn.close()
        wb.close()


def main():
    parser = argparse.ArgumentParser(
        description="MySQL 数据库表导出为 Excel（强制文本格式），支持多表配置与条件筛选"
    )
    # 方式一：引用已保存的连接（优先）
    parser.add_argument(
        "--connection_name",
        help="使用已保存的数据库连接名称（密码从本地 SQLite 读取，不通过命令行传递）"
    )

    # 方式二：直接传参
    parser.add_argument("-H", "--host", default=None, help="MySQL 主机地址")
    parser.add_argument("-P", "--port", type=int, default=3306, help="MySQL 端口 (默认 3306)")
    parser.add_argument("-u", "--user", default=None, help="MySQL 用户名")
    parser.add_argument("-p", "--password", default=None, help="MySQL 密码")
    parser.add_argument("-d", "--database", default=None, help="MySQL 数据库名")
    parser.add_argument("-o", "--out", required=True, help="输出 Excel 文件路径")
    parser.add_argument(
        "-t", "--tables",
        default=None,
        help="表配置 JSON 数组，每项包含: sheet_name/table/columns/where/limit"
    )
    parser.add_argument(
        "--tables_file",
        default=None,
        help="从 Excel 或 JSON 文件导入表配置（与 -t 二选一）"
    )
    parser.add_argument("--charset", default="utf8mb4", help="字符集 (默认 utf8mb4)")
    parser.add_argument("--timeout", type=int, default=30, help="连接超时秒数 (默认 30)")

    args = parser.parse_args()

    if not args.connection_name and not all([args.host, args.user, args.password, args.database]):
        logging.error("必须提供 --connection_name 或同时提供 -H/-u/-p/-d")
        sys.exit(1)

    if args.connection_name:
        saved = lookup_saved_connection(args.connection_name)
        if not saved:
            logging.error(f"未找到名为「{args.connection_name}」的已保存连接")
            sys.exit(1)
        db_config = {
            "host": args.host or saved["host"],
            "port": args.port or saved["port"],
            "user": args.user or saved["user"],
            "password": args.password or saved["password"],
            "database": args.database or saved["database"],
            "charset": args.charset or saved["charset"],
            "timeout": args.timeout or saved["timeout"],
        }
        logging.info(f"使用已保存连接: {args.connection_name}")
    else:
        db_config = {
            "host": args.host,
            "port": args.port,
            "user": args.user,
            "password": args.password,
            "database": args.database,
            "charset": args.charset,
            "timeout": args.timeout,
        }

    # 解析表配置：优先从文件读取，其次从 -t JSON 参数
    if args.tables_file:
        if not os.path.isfile(args.tables_file):
            logging.error(f"表配置文件不存在: {args.tables_file}")
            sys.exit(1)
        ext = os.path.splitext(args.tables_file)[1].lower()
        if ext == '.json':
            with open(args.tables_file, 'r', encoding='utf-8') as f:
                tables_cfg = json.load(f)
        elif ext in ('.xlsx', '.xls'):
            import pandas as pd
            df = pd.read_excel(args.tables_file, dtype=str)
            tables_cfg = _parse_excel_tables_config(df)
        else:
            logging.error(f"不支持的表配置文件格式: {ext}，请使用 .json 或 .xlsx")
            sys.exit(1)
    elif args.tables:
        try:
            tables_cfg = json.loads(args.tables)
        except json.JSONDecodeError as e:
            logging.error(f"表配置 JSON 解析失败: {e}")
            sys.exit(1)
    else:
        logging.error("必须提供 -t/--tables 或 --tables_file 其中之一")
        sys.exit(1)

    if not isinstance(tables_cfg, list) or len(tables_cfg) == 0:
        logging.error("表配置必须是非空 JSON 数组")
        sys.exit(1)

    for idx, t in enumerate(tables_cfg):
        if "table" not in t:
            logging.error(f"表配置[{idx}] 缺少 'table' 字段")
            sys.exit(1)

    export_mysql_to_excel(db_config, tables_cfg, args.out)


if __name__ == "__main__":
    main()
