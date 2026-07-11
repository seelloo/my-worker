TASKS = {
    "copy_excel": {
        "name": "📂 复制并筛选 Excel",
        "category": "archive",
        "script": "scripts/copy_excel_files.py",
        "description": "搜索源目录中包含或排除特定关键字的 Excel 文件，并批量复制到目标目录。",
        "params": [
            {"id": "src_dir", "label": "源目录路径", "type": "text", "required": True, "rules": {"exists": True, "isdir": True}},
            {"id": "dest_dir", "label": "目标目录路径", "type": "text", "required": True},
            {"id": "keyword", "label": "包含关键字", "type": "text", "required": True},
            {"id": "exclude_keyword", "label": "排除关键字", "type": "text", "required": False}
        ],
        "cmd_builder": lambda p: ["python", "scripts/copy_excel_files.py", "-s", p["src_dir"], "-d", p["dest_dir"], "-k", p["keyword"]] + (["-e", p["exclude_keyword"]] if p.get("exclude_keyword") else [])
    },
    "merge_sheets": {
        "name": "📑 多文件合并为单文件多Sheet",
        "category": "archive",
        "script": "scripts/merge_excel_sheets.py",
        "description": "读取一个或多个文件中的 Sheet 并合并到一个指定文件中。",
        "params": [
            {"id": "src_dir", "label": "源文件夹路径", "type": "text", "required": True, "rules": {"exists": True, "isdir": True}},
            {"id": "out_file", "label": "输出 Excel 路径", "type": "text", "required": True, "rules": {"ext": [".xlsx"]}},
            {"id": "target_sheet", "label": "目标 Sheet 名 (选填)", "type": "text", "required": False}
        ],
        "cmd_builder": lambda p: ["python", "scripts/merge_excel_sheets.py", "-s", p["src_dir"], "-o", p["out_file"]] + (["-t", p["target_sheet"]] if p.get("target_sheet") else [])
    },
    "concat_excel": {
        "name": "🔗 多文件合并为单文件1/N-Sheet",
        "category": "archive",
        "script": "scripts/concat_excel_sheets.py",
        "description": "将一个文件夹内所有 Excel 的数据（忽略表头，按行追加）合并成一个超级大表。",
        "params": [
            {"id": "src_dir", "label": "源文件夹路径", "type": "text", "required": True, "rules": {"exists": True, "isdir": True}},
            {"id": "out_file", "label": "输出 Excel 路径", "type": "text", "required": True, "rules": {"ext": [".xlsx"]}},
            {"id": "merge_mode", "label": "合并模式", "type": "select", "options": ["合并到一个Sheet", "分Sheet保留原名"], "default": "合并到一个Sheet"},
            {"id": "start_row", "label": "数据起始行 (数字)", "type": "number", "default": 1},
            {"id": "target_sheet", "label": "指定 Sheet 名称（选填，不填则读所有 Sheet）", "type": "text", "required": False}
        ],
        "cmd_builder": lambda p: [
            "python", "scripts/concat_excel_sheets.py",
            "-s", p["src_dir"], "-o", p["out_file"],
            "-r", str(p.get("start_row", 1)),
            "-m", p.get("merge_mode", "合并到一个Sheet")
        ] + (["-t", p["target_sheet"]] if p.get("target_sheet") else [])
    },
    "excel_join_multiply": {
        "name": "🔗 Excel 关联计算 (与乘积结果)",
        "category": "processing",
        "script": "scripts/excel_join_multiply.py",
        "description": "将 A 文件的 a 字段关联 B 文件的 b 字段，计算乘积并支持自定义列排序输出。",
        "params": [
            {"id": "file_a", "label": "Excel A 路径", "type": "text", "rules": {"exists": True, "isfile": True}},
            {"id": "file_b", "label": "Excel B 路径", "type": "text", "rules": {"exists": True, "isfile": True}},
            {"id": "out", "label": "输出路径", "type": "text", "rules": {"ext": [".xlsx"]}},
            {"id": "ka", "label": "A 关联键", "type": "text", "rules": {"not_empty": True}},
            {"id": "kb", "label": "B 关联键", "type": "text", "rules": {"not_empty": True}},
            {"id": "ca", "label": "A 乘数项", "type": "text", "rules": {"not_empty": True}},
            {"id": "cb", "label": "B 乘数项", "type": "text", "rules": {"not_empty": True}},
            {"id": "res", "label": "结果列名", "type": "text", "rules": {"not_empty": True}}
        ],
        "is_core": True,
        "cmd_builder": lambda p: [
            "python", "scripts/excel_join_multiply.py", 
            "--file_a", p["file_a"], "--file_b", p["file_b"], "--out", p["out"], 
            "--ka", p["ka"], "--kb", p["kb"], "--ca", p["ca"], "--cb", p["cb"], "--res", p["res"]
        ] + (["--cols", ",".join(p["output_columns"])] if p.get("output_columns") else [])
    },
    "excel_xlookup": {
        "name": "🔎 Excel XLOOKUP 跨表查找",
        "category": "processing",
        "script": "scripts/excel_xlookup.py",
        "description": "从副文件按指定列匹配主文件的查找列，将副文件的查询列和匹配记录数追加到主文件末尾。",
        "params": [
            {"id": "main_file", "label": "主文件路径", "type": "text", "required": True, "rules": {"exists": True, "isfile": True}},
            {"id": "side_file", "label": "副文件路径", "type": "text", "required": True, "rules": {"exists": True, "isfile": True}},
            {"id": "out", "label": "输出路径", "type": "text", "required": True, "rules": {"ext": [".xlsx"]}},
            {"id": "lookup_col", "label": "主文件查找列", "type": "text", "required": True, "rules": {"not_empty": True}},
            {"id": "match_col", "label": "副文件匹配列", "type": "text", "required": True, "rules": {"not_empty": True}},
            {"id": "result_cols", "label": "副文件查询列(逗号分隔)", "type": "text", "required": True, "rules": {"not_empty": True}},
            {"id": "main_sheet", "label": "主文件Sheet名(可选)", "type": "text"},
            {"id": "side_sheet", "label": "副文件Sheet名(可选)", "type": "text"}
        ],
        "is_core": True,
        "cmd_builder": lambda p: [
            "python", "scripts/excel_xlookup.py",
            "--main", p["main_file"], "--side", p["side_file"], "--out", p["out"],
            "--lookup_col", p["lookup_col"], "--match_col", p["match_col"],
            "--result_cols", p["result_cols"]
        ] + (["--main_sheet", p["main_sheet"]] if p.get("main_sheet") else [])
        + (["--side_sheet", p["side_sheet"]] if p.get("side_sheet") else [])
    },
    "rename_sheets": {
        "name": "✏️ 批量重命名 Sheet",
        "category": "processing",
        "script": "scripts/rename_excel_sheets.py",
        "description": "对 Excel 文件中的所有工作表名称进行批量修改（删除关键词或加前后缀）。",
        "params": [
            {"id": "src_file", "label": "源 Excel 文件路径", "type": "text", "required": True, "rules": {"exists": True, "isfile": True}},
            {"id": "out_path", "label": "输出 Excel 路径", "type": "text", "required": True, "rules": {"ext": [".xlsx"]}},
            {"id": "remove_str", "label": "删除字符串", "type": "text", "required": False},
            {"id": "add_str", "label": "增加字符串", "type": "text", "required": False},
            {"id": "add_position", "label": "位置", "type": "select", "options": ["back", "front"], "default": "back"}
        ],
        "cmd_builder": lambda p: ["python", "scripts/rename_excel_sheets.py", "-s", p["src_file"], "-o", p["out_path"]] + (["-r", p["remove_str"]] if p.get("remove_str") else []) + (["-a", p["add_str"]] if p.get("add_str") else []) + (["-p", p.get("add_position", "back")])
    },
    "collect_column": {
        "name": "📍 提取特定列到新表",
        "category": "processing",
        "script": "scripts/collect_excel_column.py",
        "description": "从一个 Excel 中提取指定的列并保存到新表。",
        "params": [
            {"id": "src_file", "label": "源 Excel 路径", "type": "text", "required": True, "rules": {"exists": True, "isfile": True}},
            {"id": "out_path", "label": "输出 Excel 路径", "type": "text", "required": True},
            {"id": "target_column", "label": "目标列名 (如 A)", "type": "text", "default": "A", "rules": {"not_empty": True}},
            {"id": "out_sheet_name", "label": "新 Sheet 名称", "type": "text", "default": "Collected_Data"}
        ],
        "cmd_builder": lambda p: ["python", "scripts/collect_excel_column.py", "-s", p["src_file"], "-o", p["out_path"], "-c", p["target_column"], "-n", p.get("out_sheet_name", "Collected_Data")]
    },
    "clean_filenames": {
        "name": "🧹 文件名降噪清洗",
        "category": "processing",
        "script": "scripts/clean_filenames.py",
        "description": "去除文件名中的符号、英文、数字，仅保留中文。",
        "params": [
            {"id": "target_dir", "label": "目标目录", "type": "text", "required": True, "rules": {"exists": True, "isdir": True}},
            {"id": "rm_str1", "label": "删除词 1", "type": "text"},
            {"id": "rm_str2", "label": "删除词 2", "type": "text"},
            {"id": "rm_str3", "label": "删除词 3", "type": "text"},
            {"id": "rm_str4", "label": "删除词 4", "type": "text"},
            {"id": "add_str", "label": "追加字符", "type": "text"},
            {"id": "add_position", "label": "位置", "type": "select", "options": ["suffix", "prefix"], "default": "suffix"}
        ],
        "cmd_builder": lambda p: ["python", "scripts/clean_filenames.py", "-d", p["target_dir"]] + (["--remove", p.get("rm_str1", ""), p.get("rm_str2", ""), p.get("rm_str3", ""), p.get("rm_str4", "")]) + (["--add", p["add_str"]] if p.get("add_str") else []) + (["--pos", p.get("add_position", "suffix")])
    },
    "rename_to_dir": {
        "name": "🏷️ 文件重命名为目录名",
        "category": "processing",
        "script": "scripts/rename_files_to_dirname.py",
        "description": "将文件重命名为其所属文件夹名称。",
        "params": [
            {"id": "target_dir", "label": "目标目录路径", "type": "text", "required": True, "rules": {"exists": True, "isdir": True}}
        ],
        "cmd_builder": lambda p: ["python", "scripts/rename_files_to_dirname.py", "-d", p["target_dir"]]
    },
    "ofd_to_image": {
        "name": "🖼️ OFD 转图片",
        "category": "processing",
        "script": "scripts/ofd_to_image.py",
        "description": "将 OFD 国家标准版式文件转换为图片。",
        "params": [
            {"id": "src_dir", "label": "OFD 源目录", "type": "text", "required": True, "rules": {"exists": True, "isdir": True}},
            {"id": "out_dir", "label": "图片输出目录", "type": "text", "required": True},
            {"id": "format", "label": "输出格式", "type": "select", "options": ["jpg", "png"], "default": "jpg"}
        ],
        "cmd_builder": lambda p: ["python", "scripts/ofd_to_image.py", "-s", p["src_dir"], "-o", p["out_dir"], "-f", p.get("format", "jpg")]
    },
    "ofd_extract": {
        "name": "📰 OFD 发票全自动提取",
        "category": "extraction",
        "script": "scripts/extract_ofd_invoice.py",
        "description": "自动识别 OFD 增值税发票内容并导出到 Excel 汇总表。",
        "params": [
            {"id": "src_dir", "label": "OFD 票据源文件夹", "type": "text", "required": True, "rules": {"exists": True, "isdir": True}},
            {"id": "out_file", "label": "汇总 Excel 路径", "type": "text", "required": True, "rules": {"ext": [".xlsx"]}}
        ],
        "cmd_builder": lambda p: ["python", "scripts/extract_ofd_invoice.py", "-s", p["src_dir"], "-o", p["out_file"]]
    },
    "csv_to_excel": {
        "name": "批量 CSV 转 Excel (强制文本格式)",
        "category": "processing",
        "script": "scripts/csv_to_excel_text.py",
        "description": "将目录下的所有 CSV 文件合并到一个 Excel 的多个 Sheet 中，并强制所有单元格为纯文本格式（防止 0 丢失或科学计数法）。",
        "params": [
            {"id": "src_dir", "label": "CSV 源文件夹路径", "type": "text", "required": True, "rules": {"exists": True, "isdir": True}},
            {"id": "out_file", "label": "输出 Excel 路径", "type": "text", "required": True, "rules": {"ext": [".xlsx"]}}
        ],
        "cmd_builder": lambda p: ["python", "scripts/csv_to_excel_text.py", "-s", p["src_dir"], "-o", p["out_file"]]
    },
    "rename_excel_headers": {
        "name": "🏷️ 批量重命名 Excel 表头",
        "category": "processing",
        "script": "scripts/rename_excel_headers.py",
        "description": "根据映射表 (Excel) 批量替换目标 Excel 文件中的表头字段名。",
        "params": [
            {"id": "src", "label": "源 Excel 文件或文件夹", "type": "text", "required": True, "rules": {"exists": True}},
            {"id": "mapping", "label": "映射关系 Excel 文件", "type": "text", "required": True, "rules": {"exists": True, "isfile": True}},
            {"id": "out", "label": "输出路径 (文件夹或文件)", "type": "text", "required": True},
            {"id": "old_idx", "label": "旧名称列索引 (从 0 开始)", "type": "number", "default": 0},
            {"id": "new_idx", "label": "新名称列索引 (从 0 开始)", "type": "number", "default": 1}
        ],
        "cmd_builder": lambda p: ["python", "scripts/rename_excel_headers.py", "-s", p["src"], "-m", p["mapping"], "-o", p["out"], "--old_idx", str(p.get("old_idx", 0)), "--new_idx", str(p.get("new_idx", 1))]
    },
    "excel_proportional_alloc": {
        "name": "⚖️ Excel 按比例分配 (A→B 值拆分)",
        "category": "processing",
        "script": "scripts/excel_proportional_alloc.py",
        "description": "按关联键匹配 A/B 两个 Excel，将 B 文件的值按 A 文件权重列的比例逐行分配，并支持自定义输出字段与排序。",
        "is_core": True,
        "params": [
            {"id": "file_a", "label": "Excel A 路径", "type": "text", "rules": {"exists": True, "isfile": True}},
            {"id": "file_b", "label": "Excel B 路径", "type": "text", "rules": {"exists": True, "isfile": True}},
            {"id": "out", "label": "输出路径", "type": "text", "rules": {"ext": [".xlsx"]}},
            {"id": "ka", "label": "A 关联键列名", "type": "text", "rules": {"not_empty": True}},
            {"id": "wa", "label": "A 权重列名（用于计算比例）", "type": "text", "rules": {"not_empty": True}},
            {"id": "kb", "label": "B 关联键列名", "type": "text", "rules": {"not_empty": True}},
            {"id": "vb", "label": "B 待分配值列名", "type": "text", "rules": {"not_empty": True}},
            {"id": "res", "label": "结果列名", "type": "text", "rules": {"not_empty": True}},
            {"id": "decimals", "label": "结果小数位数", "type": "number", "default": 2}
        ],
        "cmd_builder": lambda p: [
            "python", "scripts/excel_proportional_alloc.py",
            "--file_a", p["file_a"], "--file_b", p["file_b"], "--out", p["out"],
            "--ka", p["ka"], "--wa", p["wa"],
            "--kb", p["kb"], "--vb", p["vb"],
            "--res", p.get("res", "分配结果"),
            "--decimals", str(p.get("decimals", 2))
        ] + (["--cols", ",".join(p["output_columns"])] if p.get("output_columns") else [])
    },
    "filter_excel": {
        "name": "🔍 多条件筛选 Excel 记录",
        "category": "processing",
        "script": "scripts/filter_excel_by_condition.py",
        "description": "基于一个或多个列的条件筛选记录（支持包含、等于、大于、小于等），满足所有条件的记录将导出到新表。可选从 Excel 条件文档导入筛选条件。",
        "params": [
            {"id": "src_file", "label": "源 Excel 文件路径", "type": "text", "required": True, "rules": {"exists": True, "isfile": True}},
            {"id": "out_path", "label": "输出结果保存路径", "type": "text", "required": True},
            {"id": "sheet",    "label": "指定 Sheet 名称（选填，默认读取第一个 Sheet）", "type": "sheet_select", "required": False, "src_ref": "src_file"},
            {"id": "conditions", "label": "过滤条件配置", "type": "condition_builder", "required": False},
            {"id": "cond_file", "label": "条件 Excel 文档路径（可选，与条件构造器二选一）", "type": "text", "required": False, "rules": {"exists": True, "isfile": True}},
            {
                "id": "_cond_hint",
                "type": "info",
                "label": "📋 条件文档格式说明",
                "col_hint": [
                    ["目标列", "目标列 / col / column / 列名 / 列"],
                    ["操作符", "操作符 / op / operator / 运算符 / 条件"],
                    ["比对值", "比对值 / val / value / 目标值 / 值"],
                ],
                "operators": [
                    ["contains",     "包含",       "多值用英文/中文逗号分隔，如：A,B,C"],
                    ["not_contains", "不包含",     "所有目标值均不包含时才通过"],
                    ["equals",       "等于",       "多值用逗号分隔，满足其一即通过"],
                    ["not_equals",   "不等于",     "所有目标值均不等于时才通过"],
                    ["gt",           "大于",       "数值优先比较，不是数值则字符串比较"],
                    ["lt",           "小于",       ""],
                    ["gte",          "大于等于",   ""],
                    ["lte",          "小于等于",   ""],
                    ["is_empty",     "为空",       "比对值列留空即可"],
                    ["is_not_empty", "不为空",     "比对值列留空即可"],
                ],
                "example": "状态 | equals | 已完成,处理中    /    金额 | gt | 1000    /    备注 | is_empty |",
            },
        ],

        "cmd_builder": lambda p: (
            ["python", "scripts/filter_excel_by_condition.py", "-s", p["src_file"], "-o", p["out_path"]]
            + (["--sheet", p["sheet"]] if p.get("sheet") else [])
            + (["--cond-file", p["cond_file"]] if p.get("cond_file") else ["-c", p.get("conditions", "[]")])
        )
    },

    "accounts_daily_report": {
        "name": "💰 佣金电子签日报表生成",
        "category": "processing",
        "script": "scripts/accounts_daily_report.py",
        "description": "读取报账单台账，按分组列拆分后从佣金日报表中筛选数据，按实付总金额命名并写入电子签模板文件。",
        "params": [
            {"id": "bzlist",    "label": "报账单台账 Excel 路径",   "type": "text", "required": True,  "rules": {"exists": True, "isfile": True}},
            {"id": "src_dir",   "label": "日报表 & 模板所在目录",   "type": "text", "required": True,  "rules": {"exists": True, "isdir": True}},
            {"id": "daily",     "label": "佣金支付日报表文件名",    "type": "text", "required": True,  "rules": {"not_empty": True}},
            {"id": "template",  "label": "电子签模板文件名",        "type": "text", "required": True,  "rules": {"not_empty": True}},
            {"id": "group_col", "label": "分组列名",                "type": "text", "required": False, "default": "是否已报账"},
            {"id": "id_col",    "label": "报账单号列名",            "type": "text", "required": False, "default": "财辅报账单号"},
            {"id": "amount_col","label": "累加金额列名",            "type": "text", "required": False, "default": "实付金额(元)"},
            {"id": "startrow",  "label": "写入起始行 (0-indexed)", "type": "number","required": False, "default": 3},
            {"id": "skiprows",  "label": "日报表跳过行数",          "type": "number","required": False, "default": 1}
        ],
        "cmd_builder": lambda p: [
            "python", "scripts/accounts_daily_report.py",
            "-b", p["bzlist"],
            "-s", p["src_dir"],
            "-d", p["daily"],
            "-t", p["template"],
            "-g", p.get("group_col", "是否已报账"),
            "-i", p.get("id_col", "财辅报账单号"),
            "--amount_col", p.get("amount_col", "实付金额(元)"),
            "--startrow",   str(p.get("startrow", 3)),
            "--skiprows",   str(p.get("skiprows", 1)),
        ]
    },
    "excel_group_assign": {
        "name": "🔢 Excel 数值列条件分组",
        "category": "processing",
        "script": "scripts/excel_group_assign.py",
        "description": "根据多条分组规则对指定数值列进行分组编号，支持条件匹配和随机总和搜索。",
        "params": [
            {"id": "src", "label": "源 Excel 路径", "type": "text", "required": True, "rules": {"exists": True, "isfile": True}},
            {"id": "out", "label": "输出 Excel 路径", "type": "text", "required": True, "rules": {"ext": [".xlsx"]}},
            {"id": "col", "label": "目标数值列（列名或字母）", "type": "text", "required": True, "rules": {"not_empty": True}},
            {"id": "rules", "label": "分组规则 JSON", "type": "textarea", "required": True, "rules": {"not_empty": True}},
            {"id": "group_col_name", "label": "组编码列名", "type": "text", "default": "组编码"},
            {"id": "concat_col", "label": "汇总拼接列（可选）", "type": "text"}
        ],
        "cmd_builder": lambda p: [
            "python", "scripts/excel_group_assign.py",
            "-s", p["src"], "-o", p["out"], "-c", p["col"], "-r", p["rules"],
            "--group_col_name", p.get("group_col_name", "组编码")
        ] + (["--concat_col", p["concat_col"]] if p.get("concat_col") else [])
    },
    "excel_column_reorder": {
        "name": "📐 Excel 列顺序重组",
        "category": "processing",
        "script": "scripts/excel_column_reorder.py",
        "description": "按指定字段顺序从源 Excel 中提取列，输出到新文件（缺失字段以空列代替，强制文本格式）。",
        "params": [
            {"id": "src", "label": "源 Excel 路径", "type": "text", "required": True, "rules": {"exists": True, "isfile": True}},
            {"id": "out", "label": "输出 Excel 路径", "type": "text", "required": True, "rules": {"ext": [".xlsx"]}},
            {"id": "fields", "label": "有序字段名 JSON 数组", "type": "textarea", "required": True, "rules": {"not_empty": True}},
            {"id": "sheet", "label": "指定 Sheet 名称（可选）", "type": "text"}
        ],
        "cmd_builder": lambda p: [
            "python", "scripts/excel_column_reorder.py",
            "-s", p["src"], "-o", p["out"], "-f", p["fields"]
        ] + (["--sheet", p["sheet"]] if p.get("sheet") else [])
    },
    "excel_sheet_split": {
        "name": "✂️ Excel 多 Sheet 按列拆分",
        "category": "archive",
        "script": "scripts/excel_sheet_splitter.py",
        "description": "将多 Sheet 的 Excel 按指定列的唯一值拆分成多个文件，每个文件保留所有 Sheet 结构。",
        "params": [
            {"id": "input", "label": "源 Excel 路径", "type": "text", "required": True, "rules": {"exists": True, "isfile": True}},
            {"id": "col", "label": "拆分列名", "type": "text", "required": True, "rules": {"not_empty": True}},
            {"id": "output_dir", "label": "输出目录", "type": "text", "required": True}
        ],
        "cmd_builder": lambda p: [
            "python", "scripts/excel_sheet_splitter.py", "split",
            "--input", p["input"], "--col", p["col"], "--output_dir", p["output_dir"]
        ]
    },
    "split_to_single_sheets": {
        "name": "✂️ 多Sheet拆分为单Sheet文件",
        "category": "archive",
        "script": "scripts/split_excel_to_single_sheets.py",
        "description": "将一个包含多个 Sheet 的 Excel 文件，按每个 Sheet 拆分成一个个独立的 Excel 文件（每个文件仅保留一个 Sheet 的数据）。",
        "params": [
            {"id": "input", "label": "源 Excel 文件路径", "type": "text", "required": True, "rules": {"exists": True, "isfile": True}},
            {"id": "output_dir", "label": "输出目录路径", "type": "text", "required": True}
        ],
        "cmd_builder": lambda p: ["python", "scripts/split_excel_to_single_sheets.py", "-i", p["input"], "-o", p["output_dir"]]
    },
    "excel_sheet_merge": {
        "name": "🔗 Excel 按文件名重组",
        "category": "archive",
        "script": "scripts/excel_sheet_splitter.py",
        "description": "扫描目录内所有 .xlsx 文件，按文件名分组合并为新文件（Sheet 数据叠加追加）。",
        "params": [
            {"id": "input_dir", "label": "来源目录", "type": "text", "required": True, "rules": {"exists": True, "isdir": True}},
            {"id": "output_dir", "label": "输出目录", "type": "text", "required": True}
        ],
        "cmd_builder": lambda p: [
            "python", "scripts/excel_sheet_splitter.py", "merge",
            "--input_dir", p["input_dir"], "--output_dir", p["output_dir"]
        ]
    },
    "cross_sheet_merge": {
        "name": "🧩 跨 Sheet 多列自由组合",
        "category": "processing",
        "script": "scripts/cross_sheet_merge.py",
        "description": "从同一 Excel 多个 Sheet 中按规则将不同列垂直拼接，生成新 Sheet。",
        "params": [
            {"id": "src", "label": "源 Excel 路径", "type": "text", "required": True, "rules": {"exists": True, "isfile": True}},
            {"id": "out", "label": "输出 Excel 路径", "type": "text", "required": True, "rules": {"ext": [".xlsx"]}},
            {"id": "config", "label": "列配置 JSON", "type": "textarea", "required": True, "rules": {"not_empty": True}},
            {"id": "out_sheet", "label": "输出 Sheet 名", "type": "text", "default": "合并结果"},
            {"id": "keep_header", "label": "保留原始表头", "type": "checkbox", "default": False}
        ],
        "cmd_builder": lambda p: [
            "python", "scripts/cross_sheet_merge.py",
            "-s", p["src"], "-o", p["out"], "-c", p["config"],
            "--out_sheet", p.get("out_sheet", "合并结果")
        ] + (["--keep_header"] if p.get("keep_header") else [])
    },
    "pdf_convert_word": {
        "name": "📄 PDF 转 Word",
        "category": "processing",
        "script": "scripts/pdf_manager.py",
        "description": "将已上传的 PDF 文件转换为 Word (.docx) 文档。",
        "params": [
            {"id": "filename", "label": "PDF 文件名", "type": "text", "required": True, "rules": {"not_empty": True}}
        ],
        "cmd_builder": lambda p: [
            "python", "-c",
            f"from scripts.pdf_manager import convert_to_word; print(convert_to_word('{p['filename']}'))"
        ]
    },
    "pdf_add_password": {
        "name": "🔒 PDF 添加密码保护",
        "category": "processing",
        "script": "scripts/pdf_manager.py",
        "description": "为 PDF 文件添加密码保护，生成新的加密文件。",
        "params": [
            {"id": "filename", "label": "PDF 文件名", "type": "text", "required": True, "rules": {"not_empty": True}},
            {"id": "user_password", "label": "打开密码", "type": "text", "required": True, "rules": {"not_empty": True}},
            {"id": "owner_password", "label": "权限密码（可选）", "type": "text"}
        ],
        "cmd_builder": lambda p: [
            "python", "-c",
            f"from scripts.pdf_manager import add_password; print(add_password('{p['filename']}', '{p['user_password']}', '{p.get('owner_password', '')}'))"
        ]
    },
    "workflow_engine": {
        "name": "🔄 执行工作流",
        "category": "workflow",
        "script": "scripts/workflow_engine.py",
        "description": "根据工作流配置执行一系列任务，支持顺序、并行、条件和循环执行。",
        "params": [
            {"id": "workflow_id", "label": "工作流ID", "type": "text", "required": True},
            {"id": "config", "label": "工作流配置(JSON)", "type": "textarea", "required": True}
        ],
        "cmd_builder": lambda p: ["python", "scripts/workflow_engine.py", "-w", p["workflow_id"], "-c", p["config"]]
    },
    "zhibiao_output": {
        "name": "📊 指标输出处理",
        "category": "processing",
        "script": "scripts/zhibiao_output.py",
        "description": "读取厅店和指标表，整合数据并输出。支持默认长表格式，或套用模板生成宽表格式。",
        "params": [
            {"id": "src_file", "label": "源 Excel 路径", "type": "text", "required": True, "rules": {"exists": True, "isfile": True}},
            {"id": "out_path", "label": "输出 Excel 路径", "type": "text", "required": True, "rules": {"ext": [".xlsx"]}},
            {"id": "out_format", "label": "输出格式", "type": "select", "options": ["默认长表", "套用模板宽表"], "default": "默认长表"},
            {"id": "template_file", "label": "宽表模板路径 (选宽表时有效)", "type": "text", "default": r"d:\my-project\my-worker\输出格式.xlsx"}
        ],
        "cmd_builder": lambda p: ["python", "scripts/zhibiao_output.py", "-i", p["src_file"], "-o", p["out_path"], "-f", p.get("out_format", "默认长表")] + (["-t", p["template_file"]] if p.get("template_file") else [])
    },
    "excel_mark_conditions": {
        "name": "🏷️ 按条件打标识",
        "category": "processing",
        "script": "scripts/excel_mark_by_conditions.py",
        "description": "读取条件文件（支持多行规则、多个 AND 条件），对源 Excel 逐行匹配，命中规则的行在新增列中写入对应标识文字。条件格式：列字母:列字母=值，支持 = != > >= < <= 操作符。",
        "params": [
            {"id": "src_file",  "label": "源 Excel 文件路径",  "type": "text", "required": True,  "rules": {"exists": True, "isfile": True}},
            {"id": "cond_file", "label": "条件 Excel 文件路径", "type": "text", "required": True,  "rules": {"exists": True, "isfile": True}},
            {"id": "out_file",  "label": "输出 Excel 路径",     "type": "text", "required": True,  "rules": {"ext": [".xlsx"]}},
            {"id": "mark_col",  "label": "新增标识列名",         "type": "text", "default": "高值标识"},
            {"id": "sheet",     "label": "源文件 Sheet 名（可选，默认第一个）", "type": "text"},
            {
                "id": "_hint",
                "type": "info",
                "label": "📋 条件文件格式说明",
                "col_hint": [
                    ["标识",  "命中后写入新列的文字，如：融光、高值"],
                    ["条件1", "格式：列字母:列字母=值，如：AG:AG=是"],
                    ["条件2", "格式：列字母:列字母>=数字，如：DM:DM>=129"],
                ],
                "operators": [
                    ["=",  "等于",   "字符串/数值均支持"],
                    ["!=", "不等于", ""],
                    [">",  "大于",   "优先按数值比较"],
                    [">=", "大于等于", ""],
                    ["<",  "小于",   ""],
                    ["<=", "小于等于", ""],
                ],
                "example": "AG:AG=是  /  DM:DM>=129  /  BI:BI=正常",
            },
        ],
        "cmd_builder": lambda p: [
            "python", "scripts/excel_mark_by_conditions.py",
            "-s", p["src_file"],
            "-c", p["cond_file"],
            "-o", p["out_file"],
            "--mark_col", p.get("mark_col", "高值标识"),
        ] + (["--sheet", p["sheet"]] if p.get("sheet") else [])
    },
    "kpi_extract_output": {
        "name": "📊 门店KPI指标提取输出",
        "category": "processing",
        "script": "scripts/kpi_extract_output.py",
        "description": "从门店KPI宽表（指标Sheet）提取数据，按账期月份和地市整合，输出为长表格式（每行=一个门店×一个指标），包含月度目标值、权重、封顶值、保底值。",
        "params": [
            {"id": "src_file",    "label": "源 Excel 文件路径（含指标Sheet）", "type": "text", "required": True, "rules": {"exists": True, "isfile": True}},
            {"id": "out_path",    "label": "输出 Excel 路径",                   "type": "text", "required": True, "rules": {"ext": [".xlsx"]}},
            {"id": "month",       "label": "账期月份（如 202606）",              "type": "text", "required": True, "rules": {"not_empty": True}},
            {"id": "city",        "label": "地市（如 南宁）",                    "type": "text", "required": True, "rules": {"not_empty": True}},
            {"id": "data_sheet",  "label": "数据Sheet名（默认：指标）",          "type": "text", "default": "指标"},
            {"id": "id_col",      "label": "编码列名（默认：店中商编码）",        "type": "text", "default": "店中商编码"},
            {"id": "name_col",    "label": "名称列名（默认：店中商名称）",        "type": "text", "default": "店中商名称"},
        ],
        "cmd_builder": lambda p: [
            "python", "scripts/kpi_extract_output.py",
            "-i", p["src_file"],
            "-o", p["out_path"],
            "-m", p["month"],
            "-c", p["city"],
            "--data_sheet", p.get("data_sheet", "指标"),
            "--id_col",     p.get("id_col",    "店中商编码"),
            "--name_col",   p.get("name_col",  "店中商名称"),
        ]
    },
}


def get_task_metadata():
    meta = []
    for tid, info in TASKS.items():
        meta.append({
            "id": tid, 
            "name": info["name"], 
            "category": info["category"],
            "description": info["description"], 
            "params": info["params"],
            "is_core": info.get("is_core", False)
        })
    return meta
