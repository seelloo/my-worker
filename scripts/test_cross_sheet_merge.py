"""
test_cross_sheet_merge.py
--------------------------
对 cross_sheet_merge 核心函数的单元测试。
"""

import sys
import os
from pathlib import Path

import pytest
import openpyxl

# ── 保证可以 import scripts 包 ─────────────────────────────────
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from scripts.cross_sheet_merge import (
    get_sheet_names,
    get_sheet_columns,
    cross_sheet_merge,
)


# ─── Fixtures ────────────────────────────────────────────────────────────────

@pytest.fixture
def sample_excel(tmp_path):
    """
    创建一个包含 3 个 Sheet 的测试 Excel 文件。

    Sheet1: A=name, B=age
    Sheet2: A=city, B=score
    Sheet3: A=code, B=value, C=flag
    """
    wb = openpyxl.Workbook()

    ws1 = wb.active
    ws1.title = "Sheet1"
    ws1.append(["name", "age"])
    ws1.append(["Alice", 30])
    ws1.append(["Bob", 25])

    ws2 = wb.create_sheet("Sheet2")
    ws2.append(["city", "score"])
    ws2.append(["Beijing", 90])
    ws2.append(["Shanghai", 85])

    ws3 = wb.create_sheet("Sheet3")
    ws3.append(["code", "value", "flag"])
    ws3.append(["X001", 100, True])
    ws3.append(["X002", 200, False])
    ws3.append(["X003", 300, True])

    out = tmp_path / "sample.xlsx"
    wb.save(out)
    return str(out)


# ─── get_sheet_names ─────────────────────────────────────────────────────────

def test_get_sheet_names(sample_excel):
    names = get_sheet_names(sample_excel)
    assert names == ["Sheet1", "Sheet2", "Sheet3"]


# ─── get_sheet_columns ───────────────────────────────────────────────────────

def test_get_sheet_columns_sheet1(sample_excel):
    cols = get_sheet_columns(sample_excel, "Sheet1")
    assert len(cols) == 2
    assert cols[0] == {"letter": "A", "header": "name"}
    assert cols[1] == {"letter": "B", "header": "age"}


def test_get_sheet_columns_sheet3(sample_excel):
    cols = get_sheet_columns(sample_excel, "Sheet3")
    assert len(cols) == 3
    assert cols[2]["letter"] == "C"


def test_get_sheet_columns_nonexistent_sheet(sample_excel):
    cols = get_sheet_columns(sample_excel, "不存在的Sheet")
    assert cols == []


# ─── cross_sheet_merge ───────────────────────────────────────────────────────

def test_basic_merge_skip_header(sample_excel, tmp_path):
    """基本合并：跳过表头，两个输出列"""
    out_file = str(tmp_path / "output.xlsx")
    config = {
        "src_file": sample_excel,
        "out_file": out_file,
        "out_sheet": "结果",
        "skip_header": True,
        "columns": [
            {
                "header": "姓名合并",
                "segments": [
                    {"sheet": "Sheet1", "col": "A"},   # name: Alice, Bob
                ]
            },
            {
                "header": "城市合并",
                "segments": [
                    {"sheet": "Sheet2", "col": "A"},   # city: Beijing, Shanghai
                ]
            },
        ]
    }
    result = cross_sheet_merge(config)
    assert result is True

    wb = openpyxl.load_workbook(out_file)
    ws = wb["结果"]

    # 第1行为表头
    assert ws.cell(1, 1).value == "姓名合并"
    assert ws.cell(1, 2).value == "城市合并"

    # 数据行
    assert ws.cell(2, 1).value == "Alice"
    assert ws.cell(3, 1).value == "Bob"
    assert ws.cell(2, 2).value == "Beijing"
    assert ws.cell(3, 2).value == "Shanghai"


def test_merge_vertical_stacking(sample_excel, tmp_path):
    """竖向拼接：同一输出列由多个 Sheet 的同列垂直堆叠"""
    out_file = str(tmp_path / "stack.xlsx")
    config = {
        "src_file": sample_excel,
        "out_file": out_file,
        "out_sheet": "堆叠结果",
        "skip_header": True,
        "columns": [
            {
                "header": "所有名称",
                "segments": [
                    {"sheet": "Sheet1", "col": "A"},   # Alice, Bob
                    {"sheet": "Sheet2", "col": "A"},   # Beijing, Shanghai
                    {"sheet": "Sheet3", "col": "A"},   # X001, X002, X003
                ]
            }
        ]
    }
    result = cross_sheet_merge(config)
    assert result is True

    wb = openpyxl.load_workbook(out_file)
    ws = wb["堆叠结果"]

    values = [ws.cell(r, 1).value for r in range(2, ws.max_row + 1)]
    assert "Alice" in values
    assert "Bob" in values
    assert "Beijing" in values
    assert "Shanghai" in values
    assert "X001" in values
    assert len(values) == 7  # 2+2+3


def test_merge_keep_header(sample_excel, tmp_path):
    """保留表头模式：keep_header=False 即 skip_header=False"""
    out_file = str(tmp_path / "with_header.xlsx")
    config = {
        "src_file": sample_excel,
        "out_file": out_file,
        "out_sheet": "保留表头",
        "skip_header": False,
        "columns": [
            {
                "header": "含表头",
                "segments": [{"sheet": "Sheet1", "col": "A"}]
            }
        ]
    }
    result = cross_sheet_merge(config)
    assert result is True

    wb = openpyxl.load_workbook(out_file)
    ws = wb["保留表头"]
    # 第2行应该是 "name"（原始表头未跳过）
    assert ws.cell(2, 1).value == "name"


def test_merge_missing_sheet_skipped(sample_excel, tmp_path, capsys):
    """不存在的 Sheet 片段被跳过，其他片段正常处理"""
    out_file = str(tmp_path / "skip_missing.xlsx")
    config = {
        "src_file": sample_excel,
        "out_file": out_file,
        "out_sheet": "跳过测试",
        "skip_header": True,
        "columns": [
            {
                "header": "测试列",
                "segments": [
                    {"sheet": "Sheet1", "col": "A"},
                    {"sheet": "不存在的Sheet", "col": "A"},  # 应被跳过
                ]
            }
        ]
    }
    result = cross_sheet_merge(config)
    assert result is True

    captured = capsys.readouterr()
    assert "不存在的Sheet" in captured.out
    assert "跳过" in captured.out


def test_merge_invalid_col_skipped(sample_excel, tmp_path, capsys):
    """无效列字母被跳过，不应崩溃"""
    out_file = str(tmp_path / "invalid_col.xlsx")
    config = {
        "src_file": sample_excel,
        "out_file": out_file,
        "out_sheet": "无效列测试",
        "skip_header": True,
        "columns": [
            {
                "header": "正常列",
                "segments": [{"sheet": "Sheet1", "col": "A"}]
            },
            {
                "header": "无效列",
                "segments": [{"sheet": "Sheet1", "col": "ZZZ"}]
            }
        ]
    }
    # 列 ZZZ 超出范围应被容错处理（返回空值），不会 crash
    result = cross_sheet_merge(config)
    # 只要不抛出异常即可
    assert isinstance(result, bool)


def test_merge_nonexistent_src(tmp_path):
    """源文件不存在时返回 False"""
    config = {
        "src_file": str(tmp_path / "not_exist.xlsx"),
        "out_file": str(tmp_path / "out.xlsx"),
        "out_sheet": "结果",
        "skip_header": True,
        "columns": [{"header": "x", "segments": [{"sheet": "S", "col": "A"}]}]
    }
    result = cross_sheet_merge(config)
    assert result is False


def test_merge_empty_columns(sample_excel, tmp_path):
    """columns 为空时返回 False"""
    config = {
        "src_file": sample_excel,
        "out_file": str(tmp_path / "out.xlsx"),
        "out_sheet": "结果",
        "skip_header": True,
        "columns": []
    }
    result = cross_sheet_merge(config)
    assert result is False


def test_merge_unequal_col_lengths(sample_excel, tmp_path):
    """各输出列数据量不等时，不足的列用 None 补齐"""
    out_file = str(tmp_path / "unequal.xlsx")
    config = {
        "src_file": sample_excel,
        "out_file": out_file,
        "out_sheet": "长短测试",
        "skip_header": True,
        "columns": [
            {
                "header": "长列",
                "segments": [
                    {"sheet": "Sheet1", "col": "A"},   # 2 行
                    {"sheet": "Sheet3", "col": "A"},   # 3 行  → 合计 5 行
                ]
            },
            {
                "header": "短列",
                "segments": [
                    {"sheet": "Sheet2", "col": "A"},   # 2 行
                ]
            }
        ]
    }
    result = cross_sheet_merge(config)
    assert result is True

    wb = openpyxl.load_workbook(out_file)
    ws = wb["长短测试"]

    # 最大行数应为 5（长列），短列超出部分为 None
    assert ws.max_row == 6  # 1表头 + 5数据
    assert ws.cell(6, 2).value is None  # 短列第5行为补位 None
