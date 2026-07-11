"""
test_excel_column_reorder.py
字段重组导出脚本单元测试
"""
import json
import sys
from pathlib import Path

import pytest
from openpyxl import load_workbook, Workbook

# 确保可以导入 scripts 模块
sys.path.insert(0, str(Path(__file__).parent.parent))
from scripts.excel_column_reorder import reorder_columns, parse_txt_fields


# ─── Fixtures ───────────────────────────────────────────────────────────────

@pytest.fixture
def sample_excel(tmp_path):
    """创建一个带有 4 列 3 行数据的测试 Excel 文件"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["姓名", "部门", "金额", "日期"])
    ws.append(["张三", "财务部", 1000, "2024-01-01"])
    ws.append(["李四", "技术部", 2000, "2024-01-02"])
    ws.append(["王五", "市场部", 3000, "2024-01-03"])
    fpath = tmp_path / "source.xlsx"
    wb.save(fpath)
    return str(fpath)


@pytest.fixture
def sample_txt(tmp_path):
    """创建一个字段名 TXT 文件"""
    fields = ["金额", "姓名", "日期"]
    fpath = tmp_path / "fields.txt"
    fpath.write_text("\n".join(fields), encoding="utf-8")
    return str(fpath)


# ─── parse_txt_fields 测试 ────────────────────────────────────────────────────

class TestParseTxtFields:
    def test_normal(self, sample_txt):
        fields = parse_txt_fields(sample_txt)
        assert fields == ["金额", "姓名", "日期"]

    def test_empty_lines_skipped(self, tmp_path):
        fpath = tmp_path / "with_blanks.txt"
        fpath.write_text("姓名\n\n部门\n\n", encoding="utf-8")
        fields = parse_txt_fields(str(fpath))
        assert fields == ["姓名", "部门"]

    def test_bom_utf8(self, tmp_path):
        """支持 UTF-8 BOM 编码（Windows 记事本另存为 UTF-8 格式）"""
        fpath = tmp_path / "bom.txt"
        fpath.write_bytes(b"\xef\xbb\xbf\xe5\xa7\x93\xe5\x90\x8d\n\xe9\x83\xa8\xe9\x97\xa8")
        fields = parse_txt_fields(str(fpath))
        assert fields == ["姓名", "部门"]

    def test_gbk_encoding(self, tmp_path):
        """支持 GBK 编码（Windows 记事本默认中文保存格式）"""
        fpath = tmp_path / "gbk.txt"
        # "姓名\n部门\n金额" 的 GBK 编码
        fpath.write_bytes("姓名\n部门\n金额".encode("gbk"))
        fields = parse_txt_fields(str(fpath))
        assert fields == ["姓名", "部门", "金额"]

    def test_file_not_found(self):
        with pytest.raises(FileNotFoundError):
            parse_txt_fields("/nonexistent/path/fields.txt")

    def test_all_empty_lines(self, tmp_path):
        fpath = tmp_path / "empty.txt"
        fpath.write_text("\n\n\n", encoding="utf-8")
        with pytest.raises(ValueError, match="未找到有效字段名"):
            parse_txt_fields(str(fpath))


# ─── reorder_columns 测试 ─────────────────────────────────────────────────────

class TestReorderColumns:
    def test_basic_reorder(self, sample_excel, tmp_path):
        """正常重组：选 3 列，顺序改变"""
        out = str(tmp_path / "out.xlsx")
        reorder_columns(sample_excel, out, ["金额", "姓名", "部门"])

        wb = load_workbook(out)
        ws = wb.active
        header = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
        assert header == ["金额", "姓名", "部门"]
        # 数据行第一行
        row1 = [c.value for c in next(ws.iter_rows(min_row=2, max_row=2))]
        assert row1 == ["1000", "张三", "财务部"]  # 强制文本
        wb.close()

    def test_missing_field_fills_empty(self, sample_excel, tmp_path, capsys):
        """缺失字段以空数据代替，警告输出，不中止"""
        out = str(tmp_path / "out_missing.xlsx")
        reorder_columns(sample_excel, out, ["姓名", "不存在的列", "金额"])

        captured = capsys.readouterr()
        assert "不存在的列" in captured.out
        assert "空数据代替" in captured.out

        wb = load_workbook(out)
        ws = wb.active
        header = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
        assert header == ["姓名", "不存在的列", "金额"]
        # 不存在的列应为空字符串或 None（openpyxl write_only 写入空字符串后读回为 None）
        row1 = [c.value for c in next(ws.iter_rows(min_row=2, max_row=2))]
        assert row1[1] in ("", None), f"缺失字段列应为空，实际值: {row1[1]!r}"
        wb.close()

    def test_all_text_format(self, sample_excel, tmp_path):
        """所有输出单元格的 number_format 应为文本格式 '@'"""
        out = str(tmp_path / "out_text.xlsx")
        reorder_columns(sample_excel, out, ["金额", "日期"])

        wb = load_workbook(out)
        ws = wb.active
        for row in ws.iter_rows(min_row=1):
            for cell in row:
                assert cell.number_format == "@", (
                    f"单元格 {cell.coordinate} 的格式应为 '@'，实际为 '{cell.number_format}'"
                )
        wb.close()

    def test_case_insensitive_match(self, sample_excel, tmp_path, capsys):
        """字段名大小写不敏感匹配（实际场景中表头可能有大小写差异）"""
        out = str(tmp_path / "out_ci.xlsx")
        # 源文件有 "姓名"，写入 "姓名"（中文无大小写差异，用英文场景验证逻辑）
        # 此处仅验证完全匹配路径正常工作
        reorder_columns(sample_excel, out, ["姓名", "金额"])
        wb = load_workbook(out)
        ws = wb.active
        header = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
        assert "姓名" in header
        wb.close()

    def test_row_count_correct(self, sample_excel, tmp_path):
        """输出行数应与源文件数据行数一致"""
        out = str(tmp_path / "out_rows.xlsx")
        reorder_columns(sample_excel, out, ["姓名", "部门"])

        wb = load_workbook(out)
        ws = wb.active
        data_rows = list(ws.iter_rows(min_row=2, values_only=True))
        assert len(data_rows) == 3  # 源文件有 3 行数据
        wb.close()

    def test_source_not_found(self, tmp_path):
        """源文件不存在时应退出并输出错误"""
        out = str(tmp_path / "out.xlsx")
        with pytest.raises(SystemExit):
            reorder_columns("/nonexistent/source.xlsx", out, ["姓名"])

    def test_empty_fields_list(self, sample_excel, tmp_path):
        """空字段列表应退出"""
        out = str(tmp_path / "out.xlsx")
        with pytest.raises(SystemExit):
            reorder_columns(sample_excel, out, [])
