# 📊 Excel / 文件批处理工具平台

> 一个以 **Flask** 驱动的 Web 任务引擎（v3.0 · Blueprint 模式），集成 **20+ 个 Excel 与文件处理工具**。
> 所有任务通过独立子进程沙箱化执行，SSE 实时流式返回日志。
> 内置**工作流编排**、**定时调度**、**任务队列**与**自然语言路由 Agent**。

---

## ✨ 功能全览

### 🗂️ 归档类

| 任务 | 描述 |
|---|---|
| **复制并筛选 Excel** | 按关键字（包含 / 排除）在源目录搜索 Excel 文件，批量复制到目标目录 |
| **多个 Sheet 合并为一** | 将一个或多个文件的指定 Sheet 合并写入同一个 Excel |
| **文件夹 Excel 内容全合并** | 将目录下所有 Excel 数据（忽略表头）按行追加合并为超级大表 |
| **Excel 多 Sheet 按列拆分** | 将多 Sheet 的 Excel 按指定列的唯一值拆分为多个独立文件 |
| **Excel 按文件名重组** | 扫描目录内所有 .xlsx，按文件名分组合并（Sheet 数据叠加追加） |

### ⚙️ 处理类

| 任务 | 描述 |
|---|---|
| **Excel 关联计算（与乘积结果）** | 按关联键匹配两个 Excel，计算两列的乘积并支持自定义输出列排序 |
| **Excel XLOOKUP 跨表查找** | 从副文件按列匹配主文件，将查询结果追加到主文件末尾 |
| **Excel 按比例分配（A→B 值拆分）** | 按权重列比例将 B 文件的值逐行分配到 A 文件的每条记录 |
| **Excel 数值列条件分组** | 根据多条分组规则对指定数值列进行分组编号，支持随机总和搜索 |
| **Excel 列顺序重组** | 按指定字段顺序重新提取列，缺失字段以空列代替，强制文本格式 |
| **跨 Sheet 多列自由组合** | 从同一 Excel 多个 Sheet 中按规则将不同列垂直拼接为新 Sheet |
| **多条件筛选 Excel 记录** | 基于多列条件（包含、等于、大于、小于等）筛选记录并导出 |
| **批量重命名 Sheet** | 批量删除 Sheet 名中的关键词，或为 Sheet 名添加前缀 / 后缀 |
| **批量重命名 Excel 表头** | 根据映射 Excel 批量替换目标文件中的表头字段名 |
| **提取特定列到新表** | 从 Excel 中提取指定列并保存为新工作表 |
| **文件名降噪清洗** | 去除文件名中的符号、英文、数字，仅保留中文，支持批量追加字符 |
| **文件重命名为目录名** | 将文件批量重命名为其所属文件夹名称 |
| **批量 CSV 转 Excel（强制文本格式）** | 将目录下所有 CSV 合并到 Excel 多 Sheet，强制纯文本（防止前导零丢失） |
| **OFD 转图片** | 将国标 OFD 版式文件批量转换为 JPG / PNG 图片 |
| **佣金电子签日报表生成** | 按报账单台账分组，从佣金日报表筛选数据，以实付总金额命名并写入电子签模板 |
| **PDF 转 Word** | 将已上传的 PDF 文件转换为 Word (.docx) 文档 |
| **PDF 添加密码保护** | 为 PDF 文件添加打开密码与权限密码 |

### 📰 提取类

| 任务 | 描述 |
|---|---|
| **OFD 发票全自动提取** | 自动识别 OFD 增值税发票字段（开票方、金额、税额等），汇总导出到 Excel |

### 🔄 工作流类

| 功能 | 描述 |
|---|---|
| **执行工作流** | 根据工作流配置执行一系列任务，支持顺序、并行、条件和循环执行 |

---

## 🏗️ 项目结构

```
my-todo/
│
├── app.py                      # Flask 主服务（v3.0 Blueprint 模式）
├── config.py                   # 全局配置（Host / Port / 限流 / 日志）
│
├── routes/                     # Flask 路由蓝图
│   ├── tasks.py                # 任务触发 & SSE 实时日志流
│   ├── excel.py                # Excel 专项操作 API
│   ├── pdf.py                  # PDF 上传 / 管理 / 转换
│   ├── invoice.py              # 发票大盘（搜索 + 导出）
│   ├── schedules.py            # 定时任务 CRUD
│   ├── queue.py                # 任务队列管理
│   ├── logs.py                 # 执行日志查看
│   └── preview.py              # 文件预览
│
├── scripts/                    # 业务核心（任务脚本 + 基础设施）
│   ├── task_registry.py        # 任务注册中心（20+ 任务的元数据 & 命令构造器）
│   ├── workflow_engine.py      # 工作流编排引擎（顺序/并行/条件/循环）
│   ├── scheduler.py            # APScheduler 定时调度（Cron / Interval / Date）
│   ├── task_queue.py           # SQLite 任务队列（优先级 FIFO + 后台 Worker）
│   ├── agent.py                # 自然语言路由 Agent（关键词匹配任务 & 参数提取）
│   ├── job_logger.py           # 任务执行日志记录（SQLite 持久化）
│   ├── validator.py            # 后端参数校验器
│   │
│   ├── accounts_daily_report.py      # 佣金电子签日报表生成
│   ├── clean_filenames.py            # 文件名降噪清洗
│   ├── collect_excel_column.py       # 提取特定列
│   ├── concat_excel_sheets.py        # 文件夹 Excel 全合并
│   ├── copy_excel_files.py           # 复制并筛选 Excel
│   ├── cross_sheet_merge.py          # 跨 Sheet 多列自由组合
│   ├── csv_to_excel_text.py          # CSV 转 Excel（强制文本）
│   ├── excel_column_reorder.py       # 列顺序重组
│   ├── excel_group_assign.py         # 数值列条件分组
│   ├── excel_join_multiply.py        # Excel 关联计算（与乘积）
│   ├── excel_proportional_alloc.py   # Excel 按比例分配
│   ├── excel_sheet_splitter.py       # 多 Sheet 按列拆分 / 按文件名重组
│   ├── excel_xlookup.py              # Excel XLOOKUP 跨表查找
│   ├── extract_ofd_invoice.py        # OFD 发票自动提取
│   ├── filter_excel_by_condition.py  # 多条件筛选
│   ├── merge_excel_sheets.py         # 多 Sheet 合并
│   ├── ofd_to_image.py               # OFD 转图片
│   ├── pdf_manager.py                # PDF 转 Word / 加密
│   ├── rename_excel_headers.py       # 批量重命名表头
│   ├── rename_excel_sheets.py        # 批量重命名 Sheet
│   └── rename_files_to_dirname.py    # 文件重命名为目录名
│
├── workflows/                  # 工作流系统
│   ├── api/
│   │   └── workflow_api.py     # 工作流 REST API（Blueprint）
│   └── models/
│       ├── workflow.py         # 工作流数据模型（Workflow / Task / Execution）
│       └── database.py         # 工作流 SQLite 持久化
│
├── utils/                      # 通用工具库
│   ├── error_handler.py        # 统一错误处理器
│   ├── rate_limiter.py         # 请求限流（默认 100次/分钟）
│   ├── security.py             # 安全工具（路径校验等）
│   ├── excel_utils.py          # Excel 通用工具函数
│   ├── file_utils.py           # 文件操作工具函数
│   ├── flask_utils.py          # Flask 辅助函数
│   └── workflow_utils.py       # 工作流辅助函数
│
├── templates/
│   ├── index.html              # 主任务执行界面
│   ├── workflow_editor.html    # 工作流可视化编辑器
│   └── workflow_list.html      # 工作流管理列表
│
├── static/                     # CSS / JS 静态资源
│
├── data/                       # 运行时数据（自动创建）
│   ├── schedules.db            # 定时任务持久化
│   └── task_queue.db           # 任务队列持久化
│
├── invoices_archive.db         # 发票数据库（SQLite）
├── uploads/                    # 文件上传存储目录
├── app.log                     # 运行日志
└── requirements.txt            # Python 依赖清单
```

---

## 🚀 快速开始

### 1. 安装依赖

```bash
# 创建并激活虚拟环境（推荐）
python -m venv .venv
.venv\Scripts\Activate.ps1   # Windows PowerShell

# 安装核心依赖
pip install -r requirements.txt

# 可选：批量 PDF 转 Word 功能
pip install pdf2docx

# 可选：PDF 转图片功能
pip install PyMuPDF
```

### 2. 启动服务

```bash
python app.py
```

服务启动后访问：**http://127.0.0.1:5000**

---

## 🖥️ Web 界面操作

1. 打开浏览器访问 `http://127.0.0.1:5000`
2. 左侧任务列表选择需要的工具
3. 填写参数后点击 **执行**
4. 右侧实时日志窗口流式展示任务进度

所有任务均在独立子进程中运行（沙箱模式），任务崩溃不影响主服务。

---

## ⌨️ 命令行直接调用

所有脚本均支持命令行独立运行，使用 `-h` 查看任意脚本的完整帮助：

```bash
python scripts/<任意脚本>.py -h
```

---

### 📂 复制并筛选 Excel

```bash
python scripts/copy_excel_files.py \
    -s "D:/源文件夹" \
    -d "D:/目标文件夹" \
    -k "关键字"

# 同时排除包含某关键字的文件
python scripts/copy_excel_files.py \
    -s "D:/源文件夹" \
    -d "D:/目标文件夹" \
    -k "2024" \
    -e "草稿"
```

| 参数 | 说明 | 必填 |
|---|---|---|
| `-s` / `--src` | 源目录路径 | ✅ |
| `-d` / `--dest` | 目标目录路径 | ✅ |
| `-k` / `--keyword` | 文件名包含关键字 | ✅ |
| `-e` / `--exclude` | 文件名排除关键字 | ❌ |

---

### 📑 多个 Sheet 合并为一个文件

```bash
python scripts/merge_excel_sheets.py \
    -s "D:/源文件夹" \
    -o "D:/输出/合并.xlsx" \
    -t "运营数据"
```

| 参数 | 说明 | 必填 |
|---|---|---|
| `-s` / `--src` | 源目录路径 | ✅ |
| `-o` / `--out` | 输出 Excel 路径 | ✅ |
| `-t` / `--sheet` | 要提取的目标 Sheet 名 | ❌ |

---

### 🔗 文件夹 Excel 内容全合并（超级大表）

```bash
# 从第1行开始（含表头）
python scripts/concat_excel_sheets.py \
    -s "D:/源文件夹" \
    -o "D:/输出/合并结果.xlsx"

# 跳过前2行（如第1行脏数据，第2行为表头）
python scripts/concat_excel_sheets.py \
    -s "D:/源文件夹" \
    -o "D:/输出/合并结果.xlsx" \
    -r 2
```

| 参数 | 说明 | 必填 | 默认值 |
|---|---|---|---|
| `-s` / `--src` | 源文件夹路径 | ✅ | — |
| `-o` / `--out` | 输出 Excel 路径 | ✅ | — |
| `-r` / `--start_row` | 有效数据起始行（从 1 开始） | ❌ | `1` |

---

### ✂️ Excel 多 Sheet 按列拆分

```bash
# 按"城市"列拆分，每个城市生成一个独立 Excel 文件
python scripts/excel_sheet_splitter.py split \
    --input "D:/数据.xlsx" \
    --col "城市" \
    --output_dir "D:/拆分结果/"
```

| 参数 | 说明 | 必填 |
|---|---|---|
| `--input` | 源 Excel 路径 | ✅ |
| `--col` | 拆分依据列名 | ✅ |
| `--output_dir` | 输出目录 | ✅ |

---

### 🔗 Excel 按文件名重组

```bash
python scripts/excel_sheet_splitter.py merge \
    --input_dir "D:/拆分结果/" \
    --output_dir "D:/重组结果/"
```

| 参数 | 说明 | 必填 |
|---|---|---|
| `--input_dir` | 来源目录 | ✅ |
| `--output_dir` | 输出目录 | ✅ |

---

### 🔗 Excel 关联计算（与乘积结果）

```bash
python scripts/excel_join_multiply.py \
    --file_a "D:/数据A.xlsx" \
    --file_b "D:/数据B.xlsx" \
    --out "D:/输出/结果.xlsx" \
    --ka "供应商编码" \
    --kb "编码" \
    --ca "数量" \
    --cb "单价" \
    --res "金额"

# 自定义输出列顺序（逗号分隔）
python scripts/excel_join_multiply.py \
    --file_a "D:/数据A.xlsx" --file_b "D:/数据B.xlsx" \
    --out "D:/输出/结果.xlsx" \
    --ka "编码" --kb "编码" --ca "数量" --cb "单价" --res "金额" \
    --cols "编码,名称,数量,单价,金额"
```

| 参数 | 说明 | 必填 | 默认值 |
|---|---|---|---|
| `--file_a` | Excel A 文件路径 | ✅ | — |
| `--file_b` | Excel B 文件路径 | ✅ | — |
| `--out` | 输出 Excel 路径 | ✅ | — |
| `--ka` | A 文件关联键列名 | ❌ | `a` |
| `--kb` | B 文件关联键列名 | ❌ | `b` |
| `--ca` | A 文件乘数列名 | ❌ | `c` |
| `--cb` | B 文件乘数列名 | ❌ | `c` |
| `--res` | 结果列名 | ❌ | `乘积结果` |
| `--cols` | 输出列名（逗号分隔） | ❌ | A 所有列+结果列 |

---

### 🔎 Excel XLOOKUP 跨表查找

```bash
python scripts/excel_xlookup.py \
    --main "D:/主表.xlsx" \
    --side "D:/副表.xlsx" \
    --out "D:/输出/结果.xlsx" \
    --lookup_col "编码" \
    --match_col "编码" \
    --result_cols "名称,单价"
```

| 参数 | 说明 | 必填 |
|---|---|---|
| `--main` | 主文件路径 | ✅ |
| `--side` | 副文件路径 | ✅ |
| `--out` | 输出 Excel 路径 | ✅ |
| `--lookup_col` | 主文件查找列 | ✅ |
| `--match_col` | 副文件匹配列 | ✅ |
| `--result_cols` | 副文件查询列（逗号分隔） | ✅ |
| `--main_sheet` | 主文件 Sheet 名（可选） | ❌ |
| `--side_sheet` | 副文件 Sheet 名（可选） | ❌ |

---

### ⚖️ Excel 按比例分配（A→B 值拆分）

```bash
python scripts/excel_proportional_alloc.py \
    --file_a "D:/门店数据.xlsx" \
    --file_b "D:/总金额.xlsx" \
    --out "D:/输出/分配结果.xlsx" \
    --ka "城市" \
    --wa "门店面积" \
    --kb "城市" \
    --vb "待分配金额" \
    --res "分配金额"

# 指定小数位数和自定义输出列
python scripts/excel_proportional_alloc.py \
    --file_a "D:/门店数据.xlsx" --file_b "D:/总金额.xlsx" \
    --out "D:/输出/分配结果.xlsx" \
    --ka "城市" --wa "面积" --kb "城市" --vb "金额" --res "分配额" \
    --decimals 4 \
    --cols "门店名称,城市,面积,分配额"
```

| 参数 | 说明 | 必填 | 默认值 |
|---|---|---|---|
| `--file_a` | Excel A 文件路径（含权重） | ✅ | — |
| `--file_b` | Excel B 文件路径（含待分配值） | ✅ | — |
| `--out` | 输出 Excel 路径 | ✅ | — |
| `--ka` | A 文件关联键列名 | ✅ | — |
| `--wa` | A 文件权重列名 | ✅ | — |
| `--kb` | B 文件关联键列名 | ✅ | — |
| `--vb` | B 文件待分配值列名 | ✅ | — |
| `--res` | 结果列名 | ❌ | `分配结果` |
| `--cols` | 输出列名（逗号分隔） | ❌ | A 所有列+结果列 |
| `--decimals` | 结果小数位数 | ❌ | `2` |

---

### 🔢 Excel 数值列条件分组

```bash
python scripts/excel_group_assign.py \
    -s "D:/数据.xlsx" \
    -o "D:/输出/分组结果.xlsx" \
    -c "金额" \
    -r '[{"label":"A组","min":0,"max":1000},{"label":"B组","min":1000,"max":5000}]'
```

| 参数 | 说明 | 必填 | 默认值 |
|---|---|---|---|
| `-s` / `--src` | 源 Excel 路径 | ✅ | — |
| `-o` / `--out` | 输出 Excel 路径 | ✅ | — |
| `-c` / `--col` | 目标数值列名 | ✅ | — |
| `-r` / `--rules` | 分组规则 JSON | ✅ | — |
| `--group_col_name` | 组编码列名 | ❌ | `组编码` |
| `--concat_col` | 汇总拼接列（可选） | ❌ | — |

---

### 📐 Excel 列顺序重组

```bash
python scripts/excel_column_reorder.py \
    -s "D:/数据.xlsx" \
    -o "D:/输出/重组结果.xlsx" \
    -f '["姓名","部门","金额","日期"]'
```

| 参数 | 说明 | 必填 |
|---|---|---|
| `-s` / `--src` | 源 Excel 路径 | ✅ |
| `-o` / `--out` | 输出 Excel 路径 | ✅ |
| `-f` / `--fields` | 有序字段名 JSON 数组 | ✅ |
| `--sheet` | 指定 Sheet 名称（可选） | ❌ |

---

### 🧩 跨 Sheet 多列自由组合

```bash
python scripts/cross_sheet_merge.py \
    -s "D:/数据.xlsx" \
    -o "D:/输出/合并结果.xlsx" \
    -c '[{"sheet":"Sheet1","col":"姓名"},{"sheet":"Sheet2","col":"金额"}]' \
    --out_sheet "汇总"
```

| 参数 | 说明 | 必填 | 默认值 |
|---|---|---|---|
| `-s` / `--src` | 源 Excel 路径 | ✅ | — |
| `-o` / `--out` | 输出 Excel 路径 | ✅ | — |
| `-c` / `--config` | 列配置 JSON | ✅ | — |
| `--out_sheet` | 输出 Sheet 名 | ❌ | `合并结果` |
| `--keep_header` | 保留原始表头 | ❌ | 关闭 |

---

### 🔍 多条件筛选 Excel 记录

条件 JSON 格式：`[{"col": "列名", "op": "操作符", "val": "目标值"}, ...]`

支持的操作符：

| 操作符 | 含义 |
|---|---|
| `contains` | 包含（支持逗号分隔多个值） |
| `not_contains` | 不包含 |
| `equals` | 等于（支持逗号分隔多个值） |
| `not_equals` | 不等于 |
| `gt` | 大于 |
| `lt` | 小于 |
| `gte` | 大于等于 |
| `lte` | 小于等于 |
| `is_empty` | 为空 |
| `is_not_empty` | 不为空 |

```bash
# 筛选金额 > 1000 的记录
python scripts/filter_excel_by_condition.py \
    -s "D:/数据.xlsx" \
    -o "D:/筛选结果.xlsx" \
    -c '[{"col":"金额","op":"gt","val":"1000"}]'

# 多条件（AND）：金额 > 1000 且 状态等于"已付款"
python scripts/filter_excel_by_condition.py \
    -s "D:/数据.xlsx" \
    -o "D:/筛选结果.xlsx" \
    -c '[{"col":"金额","op":"gt","val":"1000"},{"col":"状态","op":"equals","val":"已付款"}]'

# 从条件 Excel 文档导入筛选条件
python scripts/filter_excel_by_condition.py \
    -s "D:/数据.xlsx" \
    -o "D:/筛选结果.xlsx" \
    --cond-file "D:/筛选条件.xlsx"
```

| 参数 | 说明 | 必填 |
|---|---|---|
| `-s` / `--src` | 源 Excel 文件路径 | ✅ |
| `-o` / `--out` | 输出 Excel 路径 | ✅ |
| `-c` / `--conditions` | 条件 JSON 字符串 | ❌* |
| `--cond-file` | 条件 Excel 文件路径 | ❌* |

> *`-c` 与 `--cond-file` 二选一。

---

### ✏️ 批量重命名 Sheet

```bash
# 删除 Sheet 名中的"（草稿）"字样
python scripts/rename_excel_sheets.py \
    -f "D:/数据.xlsx" \
    -o "D:/输出/" \
    -r "（草稿）"

# 在 Sheet 名后面追加"_2024"
python scripts/rename_excel_sheets.py \
    -f "D:/数据.xlsx" \
    -o "D:/输出/" \
    -a "_2024" \
    -p back
```

| 参数 | 说明 | 必填 | 默认值 |
|---|---|---|---|
| `-f` / `--file` | 源 Excel 文件路径 | ✅ | — |
| `-o` / `--out` | 输出路径（文件或目录） | ✅ | — |
| `-r` / `--remove` | 从 Sheet 名中删除的字符串 | ❌* | — |
| `-a` / `--add` | 向 Sheet 名追加的字符串 | ❌* | — |
| `-p` / `--pos` | 追加位置 `front` 或 `back` | ❌ | `back` |

> *`-r` 与 `-a` 至少需要指定一个。

---

### 🏷️ 批量重命名 Excel 表头

```bash
python scripts/rename_excel_headers.py \
    -s "D:/数据.xlsx" \
    -m "D:/字段映射表.xlsx" \
    -o "D:/输出/数据_新表头.xlsx"
```

| 参数 | 说明 | 必填 | 默认值 |
|---|---|---|---|
| `-s` / `--src` | 源 Excel 文件或目录 | ✅ | — |
| `-m` / `--mapping` | 映射关系 Excel 文件路径 | ✅ | — |
| `-o` / `--out` | 输出路径（文件夹或文件） | ✅ | — |
| `--old_idx` | 旧名称列索引（0-based） | ❌ | `0` |
| `--new_idx` | 新名称列索引（0-based） | ❌ | `1` |

---

### 📍 提取特定列到新 Sheet

```bash
python scripts/collect_excel_column.py \
    -f "D:/数据.xlsx" \
    -o "D:/输出/" \
    -c "P"
```

| 参数 | 说明 | 必填 | 默认值 |
|---|---|---|---|
| `-f` / `--file` | 源 Excel 文件路径 | ✅ | — |
| `-o` / `--out` | 输出路径（文件或目录） | ✅ | — |
| `-c` / `--column` | 目标列字母（如 `A`, `P`, `BD`） | ✅ | — |
| `-s` / `--sheet` | 汇总结果所在的新 Sheet 名称 | ❌ | `Collected_Data` |

---

### 🧹 文件名降噪清洗

```bash
# 保留中文，删除所有英文/数字/符号
python scripts/clean_filenames.py \
    -d "D:/待处理文件夹"

# 额外删除指定字符串
python scripts/clean_filenames.py \
    -d "D:/待处理文件夹" \
    --remove "草稿" "副本" "（1）"

# 清洗后在文件名末尾追加字符
python scripts/clean_filenames.py \
    -d "D:/待处理文件夹" \
    -a "_已整理" \
    -p suffix
```

| 参数 | 说明 | 必填 | 默认值 |
|---|---|---|---|
| `-d` / `--dir` | 目标目录路径 | ✅ | — |
| `--remove` | 要删除的特定字符串（可多个） | ❌ | — |
| `-a` / `--add` | 要追加的字符串 | ❌ | — |
| `-p` / `--pos` | 追加位置 `prefix` / `suffix` | ❌ | `suffix` |

---

### 🏷️ 文件重命名为目录名

```bash
python scripts/rename_files_to_dirname.py \
    -d "D:/待处理文件夹"
```

| 参数 | 说明 | 必填 |
|---|---|---|
| `-d` / `--dir` | 目标目录路径（递归处理所有子目录） | ✅ |

---

### 批量 CSV 转 Excel（强制文本格式）

```bash
python scripts/csv_to_excel_text.py \
    -s "D:/CSV文件夹" \
    -o "D:/输出/汇总.xlsx"
```

| 参数 | 说明 | 必填 |
|---|---|---|
| `-s` / `--src_dir` | CSV 源文件夹路径 | ✅ |
| `-o` / `--out_file` | 输出 Excel 路径 | ✅ |

---

### 🖼️ OFD 转图片

```bash
# 默认输出 JPG
python scripts/ofd_to_image.py \
    -s "D:/OFD文件夹" \
    -o "D:/图片输出/"

# 输出 PNG
python scripts/ofd_to_image.py \
    -s "D:/OFD文件夹" \
    -o "D:/图片输出/" \
    -f png
```

| 参数 | 说明 | 必填 | 默认值 |
|---|---|---|---|
| `-s` / `--src` | OFD 源文件夹路径 | ✅ | — |
| `-o` / `--out` | 图片输出目录 | ✅ | — |
| `-f` / `--format` | 输出格式 `jpg` 或 `png` | ❌ | `jpg` |

---

### 📰 OFD 发票全自动提取

```bash
python scripts/extract_ofd_invoice.py \
    -s "D:/OFD发票/" \
    -o "D:/发票汇总.xlsx"
```

| 参数 | 说明 | 必填 |
|---|---|---|
| `-s` / `--src` | OFD 票据源文件夹路径 | ✅ |
| `-o` / `--out` | 生成的汇总 Excel 路径 | ✅ |

---

### 💰 佣金电子签日报表生成

```bash
python scripts/accounts_daily_report.py \
    -b "D:/数据/bzlist202603.xlsx" \
    -s "D:/数据/" \
    -d "佣金支付日报表.xlsx" \
    -t "GX-SF006佣金外包费支付日报表202604.xlsx"
```

| 参数 | 说明 | 必填 | 默认值 |
|---|---|---|---|
| `-b` / `--bzlist` | 报账单台账 Excel 文件路径 | ✅ | — |
| `-s` / `--src_dir` | 日报表和模板文件所在目录 | ✅ | — |
| `-d` / `--daily` | 佣金支付日报表文件名 | ✅ | — |
| `-t` / `--template` | 电子签模板文件名 | ✅ | — |
| `-g` / `--group_col` | 台账中的分组列名 | ❌ | `是否已报账` |
| `-i` / `--id_col` | 报账单号列名 | ❌ | `财辅报账单号` |
| `--amount_col` | 累加金额列名 | ❌ | `实付金额(元)` |
| `--startrow` | 写入模板的起始行（0-indexed） | ❌ | `3` |
| `--skiprows` | 读取日报表时跳过的行数 | ❌ | `1` |

---

## 📋 发票大盘（内置页面）

访问 `http://127.0.0.1:5000` → 切换到发票管理页面。

- **发票数据库管理**：通过 SQLite 持久化存储，支持按卖家名称、发票号、文件名多字段关键字搜索
- **实时统计**：实时聚合总金额、总税额、发票数量
- **一键导出**：将当前筛选结果导出为标准化 Excel 文件

---

## 🔄 工作流系统

访问 `http://127.0.0.1:5000/workflows` 进入工作流管理页面。

工作流编辑器（`/workflow/editor`）支持可视化搭建任务链，工作流引擎支持以下执行模式：

| 模式 | 说明 |
|---|---|
| **顺序（Sequential）** | 任务按定义顺序依次执行 |
| **并行（Parallel）** | 多个子任务同时执行 |
| **条件（Conditional）** | 根据前置任务结果选择执行分支 |
| **循环（Loop）** | 重复执行指定任务 N 次 |

工作流配置与执行记录通过 SQLite 持久化，服务重启后不丢失。

---

## ⏰ 定时调度

通过 Web 界面或 API 创建定时任务，支持三种触发方式：

| 类型 | 示例 | 说明 |
|---|---|---|
| **Cron** | `0 9 * * *` | 每天 09:00 执行 |
| **Interval** | `minutes=30` | 每隔 30 分钟执行 |
| **Date** | 指定日期时间 | 一次性定时执行 |

调度配置通过 SQLite 持久化（`data/schedules.db`），服务重启后自动恢复。

---

## 📥 任务队列

支持将任务加入 SQLite 持久化队列（`data/task_queue.db`），由后台 Worker 线程按优先级 FIFO 顺序处理。

- 支持单条入队与批量入队
- 支持优先级排序（`priority` 越大越优先）
- 任务超时限制：1 小时
- 状态追踪：`pending → running → completed / failed`

---

## 🤖 自然语言路由 Agent

`scripts/agent.py` 提供基于关键词的自然语言任务路由能力：

- 输入中文描述（如"把这个文件夹里的 Excel 按城市拆分"）
- 自动匹配最相关的任务 ID 和参数
- 支持从文本中提取路径、列名、数值等参数

```python
from scripts.agent import match_task, parse_params_from_text

# 匹配任务
results = match_task("按比例分配门店金额")
# [{'id': 'excel_proportional_alloc', 'name': '⚖️ Excel 按比例分配...', 'score': 1.0, ...}]

# 提取参数
params = parse_params_from_text(
    "将 D:/门店数据.xlsx 按城市权重分配 D:/总金额.xlsx 的金额",
    task_id="excel_proportional_alloc"
)
```

---

## 🧩 核心模块说明

### 路由蓝图（routes/）

| 蓝图 | URL 前缀 | 主要功能 |
|---|---|---|
| `tasks_bp` | `/api/tasks` | 任务元数据、SSE 执行日志 |
| `excel_bp` | `/api/excel` | Excel 专项操作 |
| `pdf_bp` | `/api/pdf` | PDF 上传、转换、管理 |
| `invoice_bp` | `/api/invoices` | 发票 CRUD & 导出 |
| `workflow_bp` | `/api/workflows` | 工作流 CRUD & 执行 |
| `schedules_bp` | `/api/schedules` | 定时任务管理 |
| `queue_bp` | `/api/queue` | 任务队列操作 |
| `logs_bp` | `/api/logs` | 执行日志查询 |
| `preview_bp` | `/api/preview` | 文件预览 |

### 工具层（utils/）

| 模块 | 功能 |
|---|---|
| `error_handler.py` | 统一错误处理，标准化 JSON 错误响应 |
| `rate_limiter.py` | 请求限流（默认 100次/60秒，可通过环境变量调整） |
| `security.py` | 路径安全校验，防止目录遍历攻击 |
| `excel_utils.py` | Excel 通用工具（读取、写入、列处理） |
| `file_utils.py` | 文件操作工具（扫描、复制、路径处理） |

---

## 🧪 测试

项目使用 **pytest** 作为测试框架：

```bash
# 运行 scripts/ 下所有测试
pytest scripts/ -v

# 运行 tests/ 下 API 测试
pytest tests/ -v

# 运行指定测试文件
pytest scripts/test_accounts_daily_report.py -v

# 查看测试概览
pytest --tb=short -q
```

已覆盖测试的模块：

| 测试文件 | 覆盖模块 |
|---|---|
| `test_accounts_daily_report.py` | 分组逻辑、金额累加、容错处理 |
| `test_merge_excel_sheets.py` | Sheet 合并、重复 Sheet 名处理 |
| `test_concat_excel_sheets.py` | 多文件内容追加合并 |
| `test_collect_excel_column.py` | 列提取 |
| `test_filter_excel_by_condition.py` | 条件筛选（含运算符组合） |（位于 `scripts/`，引用路径无 `test_` 前缀文件）
| `test_excel_proportional_alloc.py` | 比例分配计算精度 |
| `test_excel_join_multiply.py` | 关联键匹配与乘积计算 |
| `test_excel_xlookup.py` | XLOOKUP 跨表查找 |
| `test_excel_column_reorder.py` | 列顺序重组 |
| `test_cross_sheet_merge.py` | 跨 Sheet 多列组合 |
| `test_clean_filenames.py` | 文件名清洗规则 |
| `test_rename_excel_sheets.py` | Sheet 重命名逻辑 |
| `test_rename_files_to_dirname.py` | 文件重命名 |
| `test_copy_excel_files.py` | 关键字筛选拷贝 |
| `test_csv_to_excel_text.py` | CSV 转 Excel 文本格式 |
| `tests/test_api.py` | Flask 路由 API 集成测试 |

---

## ⚙️ 配置

所有配置均支持通过**环境变量**覆盖：

| 变量 | 说明 | 默认值 |
|---|---|---|
| `HOST` | 服务监听地址 | `127.0.0.1` |
| `PORT` | 服务监听端口 | `5000` |
| `DEBUG` | 调试模式 | `true` |
| `MAX_CONTENT_LENGTH` | 最大文件上传大小（字节） | `52428800`（50MB） |
| `DB_PATH` | 发票数据库路径 | `invoices_archive.db` |
| `PDF_UPLOAD_DIR` | PDF 上传存储目录 | `uploads/pdfs` |
| `LOG_LEVEL` | 日志级别 | `DEBUG`（调试模式）/ `INFO` |
| `LOG_FILE` | 日志文件路径 | `app.log` |
| `RATE_LIMIT_ENABLED` | 是否启用请求限流 | `true` |
| `RATE_LIMIT_DEFAULT` | 限流阈值（次/分钟） | `100` |
| `RATE_LIMIT_WINDOW` | 限流时间窗口（秒） | `60` |

> ⚠️ **生产环境**请务必将 `DEBUG` 设为 `false`，并将 `DB_PATH` 设为绝对路径。

---

## 📦 依赖清单

| 包 | 用途 |
|---|---|
| `Flask` | Web 服务框架 |
| `openpyxl` | Excel 读写（.xlsx） |
| `pandas` | 数据处理核心 |
| `xlrd` | Excel 旧格式读取（.xls） |
| `easyofd` | OFD 文件解析 |
| `Pillow` | 图片处理（OFD 转图片） |
| `APScheduler` | 定时任务调度 |
| `click` | 命令行参数解析 |
| `pypdf` | PDF 工具库 |
| `pytest` | 单元测试框架 |
| `pdf2docx` *(可选)* | PDF 转 Word |
| `PyMuPDF` *(可选)* | PDF 转图片 |

---

## 📝 License

本项目为内部工具，仅供业务团队使用。
