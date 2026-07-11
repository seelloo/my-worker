import pytest
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from pathlib import Path
from scripts.copy_excel_files import get_unique_path, copy_excel_files

def test_get_unique_path(tmp_path):
    # 目标文件不存在，应直接返回原路径
    target = tmp_path / "test.xlsx"
    assert get_unique_path(target) == target
    
    # 文件已存在，应返回加 "_1" 后缀的路径
    target.touch()
    assert get_unique_path(target) == tmp_path / "test_1.xlsx"
    
    # _1 后缀文件也存在，应该返回加 "_2" 后缀的路径
    (tmp_path / "test_1.xlsx").touch()
    assert get_unique_path(target) == tmp_path / "test_2.xlsx"

def test_copy_excel_files(tmp_path):
    # 建立源和目标根目录
    src_dir = tmp_path / "source"
    dest_dir = tmp_path / "destination"
    src_dir.mkdir()
    
    # 建立多级子目录模拟递归搜索
    sub1 = src_dir / "folder_a" / "folder_b"
    sub2 = src_dir / "folder_c"
    sub1.mkdir(parents=True)
    sub2.mkdir(parents=True)
    
    # 生成测试用的 Excel 文件
    # 1. 正常包含关键词和正确后缀 (不一定要以XXX开头)
    (src_dir / "XXX_file1.xlsx").touch()
    (sub1 / "prefix_XXX_file2.xls").touch()
    
    # 2. 会发生重名冲突的文件（在不同目录下有相同的文件名）
    (sub2 / "XXX_file1.xlsx").touch()
    
    # 3. 错误的关键词或后缀，不应被复制
    (sub1 / "YYY_file1.xlsx").touch()               # 关键词不对
    (sub2 / "XXX_file3.txt").touch()                # 后缀不对
    (src_dir / "XXX_file4.csv").touch()             # 后缀不对
    
    # 4. 包含我们需要排除的过滤关键字
    (sub2 / "XXX_exclude_me.xlsx").touch()          # 含有 exclude 关键字
    (src_dir / "bad_XXX_file.xls").touch()          # 含有 bad 关键字
    
    # 运行复制脚本的主体逻辑 (搜索包含 XXX，但排除包含 bad 或者 me 的（测试单一排除情况）)
    copy_excel_files(str(src_dir), str(dest_dir), keyword="XXX", exclude_keyword="bad")
    
    # 验证目标文件夹内容
    assert dest_dir.exists()
    dest_files = list(dest_dir.iterdir())
    
    # 有4个文件会被复制：
    # XXX_file1.xlsx, prefix_XXX_file2.xls, 重名冲突的 XXX_file1.xlsx, 还有 XXX_exclude_me.xlsx (因为并没有被排除，仅仅排除了包含 'bad' 的)
    assert len(dest_files) == 4
    
    expected_names = {"XXX_file1.xlsx", "prefix_XXX_file2.xls", "XXX_file1_1.xlsx", "XXX_exclude_me.xlsx"}
    actual_names = {f.name for f in dest_files}
    assert actual_names == expected_names
