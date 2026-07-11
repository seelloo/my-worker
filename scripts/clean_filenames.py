import os
import sys
import re
import argparse
from pathlib import Path

# 将项目根目录加入模块搜索路径，保证直接运行本脚本时能正确执行绝对导入
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from scripts.rename_files_to_dirname import get_unique_filename
def clean_filenames(target_dir: str, remove_strs: list = None, add_str: str = "", add_position: str = "back"):
    """
    遍历指定文件夹下的所有文件，清洗文件名：
    1. 优先删除 `remove_strs` 里指定的最多4个特定字符串。
    2. 删除所有数字和英文字母。
    3. 强力删除所有标点符号（中英文符号），仅保留中文字符等有效的 Unicode 词汇和空格。
    4. 追加 `add_str` 到文件名的前端或后端。
    如清洗后重名，自动进行后缀编号区分。
    """
    target_path = Path(target_dir).resolve()
    
    if not target_path.exists() or not target_path.is_dir():
        print(f"❌ 找不到指定的目录: {target_dir}")
        return

    print(f"🧹 开始扫描并清洗目录: {target_path}")
    
    if remove_strs is None:
        remove_strs = []
        
    # 过滤掉空字符串和 None
    valid_remove_strs = [s for s in remove_strs if s]

    renamed_count = 0
    error_count = 0
    
    all_files = [f for f in target_path.rglob('*') if f.is_file()]
    
    # 设定的清洗正则：
    # [a-zA-Z0-9_] 负责干掉所有英文、数字和下划线
    # [^\w\s] 负责干掉所有非词元（即任何标点符号、 emoji 等），因为 \w 包含中文，所以中文能幸存
    clean_pattern = re.compile(r'[a-zA-Z0-9_]|[^\w\s]')
    
    for file_path in all_files:
        try:
            parent_dir = file_path.parent
            original_stem = file_path.stem
            extension = file_path.suffix
            
            new_stem = original_stem
            
            # 1. 优先抹去用户指定的自定义字符串
            for user_str in valid_remove_strs:
                new_stem = new_stem.replace(user_str, '')
                
            # 2. 正则表达式强力抹去字母、数字和所有符号
            new_stem = clean_pattern.sub('', new_stem)
            
            # 3. 清理可能因为删词导致的多余首尾空格，或连续多个空格变成一个
            new_stem = re.sub(r'\s+', ' ', new_stem).strip()
            
            # 4. 如果全被删空了，给个默认保底防止文件无名
            if not new_stem:
                new_stem = "已清洗文件"
                
            # 5. 追加指定的字符串
            if add_str:
                if add_position == 'front':
                    new_stem = add_str + new_stem
                else:
                    new_stem = new_stem + add_str
                
            # 如果名字完全没变（比如原本就是纯中文没符号且没新增），没必要动它
            if new_stem == original_stem:
                continue
                
            new_file_path = get_unique_filename(parent_dir, new_stem, extension)
            
            file_path.rename(new_file_path)
            renamed_count += 1
            print(f"✨ 净化成功: '{file_path.name}' -> '{new_file_path.name}'")
            
        except Exception as e:
            print(f"❌ 清洗文件 '{file_path}' 时出错: {e}")
            error_count += 1

    print(f"\n🎉 批量文件名降噪/清洗完毕！")
    print(f"👉 成功净化文件数: {renamed_count}")
    print(f"👉 失败/错误数: {error_count}")

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="批量清洗文件名，删除英文、数字、符号及特定字符串。")
    parser.add_argument("-d", "--dir", required=True, help="目标目录绝对路径")
    
    # 允许用户通过 --remove 传入多个要删除的字符串，或者保留旧的 -s1..-s4 风格
    parser.add_argument("--remove", nargs='*', default=[], help="需要删除的特定字符串列表")
    parser.add_argument("-s1", "--str1", default="", help="需要删除的特定字符串 1")
    parser.add_argument("-s2", "--str2", default="", help="需要删除的特定字符串 2")
    parser.add_argument("-s3", "--str3", default="", help="需要删除的特定字符串 3")
    parser.add_argument("-s4", "--str4", default="", help="需要删除的特定字符串 4")
    
    # 解决 --add 歧义问题，改用更明确且不冲突的短名
    parser.add_argument("-a", "--add", dest="add_str", default="", help="需要添加的额外字符串")
    parser.add_argument("--add_str", help=argparse.SUPPRESS) # 隐藏旧参数防止歧义
    
    # 增加对 prefix/suffix 的兼容性
    parser.add_argument("-p", "--pos", dest="add_position", default="back", 
                        choices=["front", "back", "prefix", "suffix"], 
                        help="添加字符串的位置 (front/back/prefix/suffix)")
    parser.add_argument("--add_position", help=argparse.SUPPRESS) # 隐藏旧参数防止歧义
    
    args = parser.parse_args()
    
    # 合并所有可能的删除字符串
    remove_list = list(args.remove)
    for s in [args.str1, args.str2, args.str3, args.str4]:
        if s:
            remove_list.append(s)
            
    # 标准化位置映射
    pos_map = {"prefix": "front", "suffix": "back", "front": "front", "back": "back"}
    final_pos = pos_map.get(args.add_position, "back")
    
    clean_filenames(args.dir, remove_list, args.add_str, final_pos)
