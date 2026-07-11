import argparse
import shutil
import concurrent.futures
from pathlib import Path

import threading

reserved_paths = set()
reserved_lock = threading.Lock()

def get_unique_path(target_path: Path) -> Path:
    directory = target_path.parent
    stem = target_path.stem
    suffix = target_path.suffix
    
    with reserved_lock:
        counter = 0
        while True:
            new_path = target_path if counter == 0 else directory / f"{stem}_{counter}{suffix}"
            if not new_path.exists() and new_path not in reserved_paths:
                reserved_paths.add(new_path)
                return new_path
            counter += 1

def copy_single_file(file_path: Path, dest_path: Path, exclude_keyword: str = None):
    """
    单文件复制逻辑，供线程池调用。
    """
    if exclude_keyword and exclude_keyword in file_path.name:
        return 0
    
    base_target_path = dest_path / file_path.name
    unique_target_path = get_unique_path(base_target_path)
    
    try:
        shutil.copy2(file_path, unique_target_path)
        print(f"✅ 已复制: {file_path.name}")
        return 1
    except Exception as e:
        print(f"❌ 复制 '{file_path.name}' 失败: {e}")
        return 0

def copy_excel_files(src_dir: str, dest_dir: str, keyword: str, exclude_keyword: str = None):
    """
    性能优化版：引入 ThreadPoolExecutor 实现多线程并发复制。
    针对大量小文件场景，能显著榨干 IO 性能。
    """
    src_path = Path(src_dir).resolve()
    dest_path = Path(dest_dir).resolve()
    
    if not src_path.exists() or not src_path.is_dir():
        print(f"错误: 源目录 '{src_dir}' 不存在。")
        return
        
    dest_path.mkdir(parents=True, exist_ok=True)
    
    # 扫描待复制文件列表
    print(f"🔍 正在扫荡源目录并匹配关键词: '{keyword}'...")
    candidates = [f for f in src_path.rglob(f"*{keyword}*") if f.is_file() and f.suffix.lower() in ('.xls', '.xlsx')]
    
    if not candidates:
        print("未发现匹配条件的 Excel 文件。")
        return

    print(f"🚀 启动线程池并发复制引擎 (共 {len(candidates)} 个候选文件)...")
    
    processed_count = 0
    # 使用线程池加速 IO 操作 (Windows 下线程对 IO 比较友好)
    with concurrent.futures.ThreadPoolExecutor(max_workers=8) as executor:
        # 提交任务
        future_to_file = {executor.submit(copy_single_file, f, dest_path, exclude_keyword): f for f in candidates}
        for future in concurrent.futures.as_completed(future_to_file):
            processed_count += future.result()
                
    print(f"\n🎉 复制大功告成！全量并发处理了 {processed_count} 个目标文件。")

def main():
    parser = argparse.ArgumentParser(description="并发性能版：Excel 文件递归搜索与平铺复制。")
    parser.add_argument("-s", "--src", required=True)
    parser.add_argument("-d", "--dest", required=True)
    parser.add_argument("-k", "--keyword", required=True)
    parser.add_argument("-e", "--exclude", required=False)
    
    args = parser.parse_args()
    copy_excel_files(args.src, args.dest, args.keyword, args.exclude)

if __name__ == "__main__":
    main()
