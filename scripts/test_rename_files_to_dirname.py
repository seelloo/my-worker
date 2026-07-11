import pytest
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from pathlib import Path
from scripts.rename_files_to_dirname import rename_files_to_dirname, get_unique_filename

def test_get_unique_filename(tmp_path):
    target_dir = tmp_path / "test_dir"
    target_dir.mkdir()
    
    # 第一次获取，应该直接是 base_name + extension
    new_path1 = get_unique_filename(target_dir, "newname", ".txt")
    assert new_path1.name == "newname.txt"
    new_path1.touch()
    
    # 第二次获取，存在冲突，应该加 _1
    new_path2 = get_unique_filename(target_dir, "newname", ".txt")
    assert new_path2.name == "newname_1.txt"
    new_path2.touch()
    
    # 第三次获取，存在冲突，应该加 _2
    new_path3 = get_unique_filename(target_dir, "newname", ".txt")
    assert new_path3.name == "newname_2.txt"

def test_rename_files_to_dirname_basic(tmp_path):
    # 创建目录结构： tmp_path / folderA / file1.txt
    folder_a = tmp_path / "folderA"
    folder_a.mkdir()
    file1 = folder_a / "file1.txt"
    file1.touch()
    
    rename_files_to_dirname(str(tmp_path))
    
    # 期望 file1.txt 变成 folderA.txt
    assert not file1.exists()
    assert (folder_a / "folderA.txt").exists()

def test_rename_files_to_dirname_multiple_same_ext(tmp_path):
    # 创建目录结构：
    # tmp_path / folderB / file1.png
    # tmp_path / folderB / file2.png
    # tmp_path / folderB / file3.png
    folder_b = tmp_path / "folderB"
    folder_b.mkdir()
    (folder_b / "file1.png").touch()
    (folder_b / "file2.png").touch()
    (folder_b / "file3.png").touch()
    
    rename_files_to_dirname(str(tmp_path))
    
    # 期望结果：
    # folderB.png
    # folderB_1.png
    # folderB_2.png
    assert (folder_b / "folderB.png").exists()
    assert (folder_b / "folderB_1.png").exists()
    assert (folder_b / "folderB_2.png").exists()

def test_rename_files_to_dirname_skip_already_renamed(tmp_path):
    # 创建目录结构，其中一个文件已经符合命名规则
    folder_c = tmp_path / "folderC"
    folder_c.mkdir()
    file_renamed = folder_c / "folderC.txt"
    file_renamed.touch()
    
    file_other = folder_c / "other.txt"
    file_other.touch()
    
    # 由于先遍历的区别，如果不严格测试，可能把 other.txt 重命名为 folderC.txt 或 folderC_1.txt
    rename_files_to_dirname(str(tmp_path))
    
    # 确保存续
    assert file_renamed.exists()
    # file_other 应被重命名为 folderC_1.txt
    assert (folder_c / "folderC_1.txt").exists()
    assert not file_other.exists()

def test_rename_files_to_dirname_nested_folders(tmp_path):
    # tmp_path / parent / child / data.csv
    parent = tmp_path / "parent_dir"
    parent.mkdir()
    child = parent / "child_dir"
    child.mkdir()
    file1 = child / "data.csv"
    file1.touch()
    
    rename_files_to_dirname(str(tmp_path))
    
    # data.csv 的父目录是 child_dir
    assert not file1.exists()
    assert (child / "child_dir.csv").exists()
