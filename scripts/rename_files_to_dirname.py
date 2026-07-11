import os
import argparse
from pathlib import Path

def get_unique_filename(directory: Path, base_name: str, extension: str) -> Path:
    """
    检查提供的新文件路径是否已存在。如果存在，自动在文件名后追加 _1, _2 等后缀。
    """
    counter = 1
    new_path = directory / f"{base_name}{extension}"
    while new_path.exists():
        new_path = directory / f"{base_name}_{counter}{extension}"
        counter += 1
    return new_path

def rename_files_to_dirname(target_dir: str):
    """
    递归遍历目标目录，将所有文件重命名为其直接父目录的名称。
    如遇同后缀重名，自动追加数字后缀。
    """
    target_path = Path(target_dir).resolve()
    
    if not target_path.exists() or not target_path.is_dir():
        print(f"❌ 找不到指定的目录: {target_dir}")
        return

    print(f"📂 开始遍历目录: {target_path}")
    
    renamed_count = 0
    error_count = 0
    
    # 递归遍历所有文件
    # 使用 list() 防止重命名过程中对遍历器产生影响
    all_files = [f for f in target_path.rglob('*') if f.is_file()]
    
    for file_path in all_files:
        try:
            parent_dir = file_path.parent
            parent_name = parent_dir.name
            
            # 如果文件就在指定的根目录下，其父目录名就是根目录名
            if parent_name == "":
                # 理论上 resolve() 之后很少出现为空，但为了安全起见
                continue
                
            original_stem = file_path.stem
            extension = file_path.suffix
            
            # 如果文件名恰好和父目录名完全一致，说明已经改过了或者本来就是这样，跳过
            # 注意：如果不加数字后缀的判断，这里只判断最简单的一致情况
            if original_stem == parent_name:
                continue
                
            new_file_path = get_unique_filename(parent_dir, parent_name, extension)
            
            # 执行重命名
            file_path.rename(new_file_path)
            renamed_count += 1
            print(f"🔄 已重命名: '{file_path.name}' -> '{new_file_path.name}' (位于 '{parent_dir}')")
            
        except Exception as e:
            print(f"❌ 重命名文件 '{file_path}' 时出错: {e}")
            error_count += 1

    print(f"\n🎉 批量重命名任务完成！")
    print(f"👉 成功重命名文件数: {renamed_count}")
    print(f"👉 失败/错误数: {error_count}")

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="遍历指定目录，将所有文件重命名为包含它们的文件夹名称。")
    parser.add_argument("-d", "--dir", required=True, help="目标目录绝对路径")
    
    args = parser.parse_args()
    rename_files_to_dirname(args.dir)
