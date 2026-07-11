import pytest
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from scripts.clean_filenames import clean_filenames
def test_clean_filenames_basic(tmp_path):
    # 包含中英文、数字、符号
    file1 = tmp_path / "报告2023_Final版(终极).txt"
    file1.touch()
    
    clean_filenames(str(tmp_path))
    
    # 期望："报告2023_Final版(终极)"
    # 删除数字：报告_Final版(终极)
    # 删除字母：报告_版(终极)
    # 删除下划线和符号：报告版终极
    assert not file1.exists()
    assert (tmp_path / "报告版终极.txt").exists()

def test_clean_filenames_custom_strings(tmp_path):
    # 用户指定的特殊词语
    file2 = tmp_path / "极密核心资料【绝密】不可外泄ABC.docx"
    file2.touch()
    
    clean_filenames(str(tmp_path), remove_strs=["极密", "不可外泄"])
    
    # "极密核心资料【绝密】不可外泄ABC"
    # 删除特定词："核心资料【绝密】ABC"
    # 删除字母："核心资料【绝密】"
    # 删除符号："核心资料绝密"
    assert not file2.exists()
    assert (tmp_path / "核心资料绝密.docx").exists()

def test_clean_filenames_all_erased(tmp_path):
    # 如果文件全是英文字母或符号，清洗后为空
    file3 = tmp_path / "123_abc_!!!.png"
    file3.touch()
    
    clean_filenames(str(tmp_path))
    
    # 全部被删光了，应该使用保底名 "已清洗文件"
    assert not file3.exists()
    assert (tmp_path / "已清洗文件.png").exists()

def test_clean_filenames_spaces_handling(tmp_path):
    # 中英文夹杂，清理后确保空格被合理保留
    file4 = tmp_path / "这是 我们的 测试 文件 v2.0 !.pdf"
    file4.touch()
    
    clean_filenames(str(tmp_path))
    
    # 删除 "v" "2" "0" "!" "."
    # 剩下 "这是 我们的 测试 文件   "
    # strip 处理成 "这是 我们的 测试 文件"
    assert not file4.exists()
    assert (tmp_path / "这是 我们的 测试 文件.pdf").exists()

def test_clean_filenames_conflict(tmp_path):
    # 多个文件洗完后撞车
    (tmp_path / "图片2021.jpg").touch()
    (tmp_path / "图片2022.jpg").touch()
    
    clean_filenames(str(tmp_path))
    
    # 两者都会变成 "图片"
    # 第一者： 图片.jpg, 第二者：图片_1.jpg
    assert (tmp_path / "图片.jpg").exists()
    assert (tmp_path / "图片_1.jpg").exists()

def test_clean_filenames_add_string(tmp_path):
    # 洗净后追加前缀或后缀
    file_front = tmp_path / "123旧文件A.txt"
    file_front.touch()
    
    file_back = tmp_path / "456旧文件B.txt"
    file_back.touch()
    
    clean_filenames(str(tmp_path), add_str="新版_", add_position="front")
    # 期望：123旧文件A -> 旧文件 -> 新版_旧文件
    #      456旧文件B -> 旧文件 -> 新版_旧文件_1
    
    assert (tmp_path / "新版_旧文件.txt").exists()
    assert (tmp_path / "新版_旧文件_1.txt").exists()
    
    # 再测一次后置
    file_back2 = tmp_path / "xyz测试C.txt"
    file_back2.touch()
    
    clean_filenames(str(tmp_path), add_str="_OK", add_position="back")
    assert (tmp_path / "测试_OK.txt").exists()
