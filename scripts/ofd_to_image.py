import os
import argparse
from pathlib import Path
from easyofd.ofd import OFD

def convert_ofd_to_image(src_dir: str, out_dir: str, img_format: str = 'jpg'):
    """
    遍历源文件夹，将所有 .ofd 文件转换为图片。
    img_format: 'jpg' 或 'png'
    """
    src_path = Path(src_dir).resolve()
    out_path = Path(out_dir).resolve()
    
    if not src_path.exists() or not src_path.is_dir():
        print(f"❌ 源目录不存在: {src_dir}")
        return
    
    if not out_path.exists():
        out_path.mkdir(parents=True)
        print(f"📁 已创建输出目录: {out_path}")

    ofd_files = list(src_path.glob("*.ofd"))
    if not ofd_files:
        print(f"ℹ️ 在 {src_dir} 中未找到 .ofd 文件")
        return

    print(f"🚀 开始转换 {len(ofd_files)} 个 OFD 文件...")
    
    converted_count = 0
    total_pages = 0
    
    for ofd_file in ofd_files:
        try:
            print(f"📖 正在解析: {ofd_file.name}")
            
            # 初始化 OFD 对象并读取文件
            ofd_obj = OFD()
            ofd_obj.read(str(ofd_file), fmt="path")
            
            # to_jpg() 返回 PIL Image 对象列表
            # 虽然方法名叫 to_jpg，但返回的是 PIL 图片，可以保存为任意格式
            images = ofd_obj.to_jpg()
            
            base_name = ofd_file.stem
            for i, img in enumerate(images):
                page_num = i + 1
                # 构造输出文件名：原文件名_页码.格式
                target_filename = f"{base_name}_{page_num}.{img_format.lower()}"
                target_file_path = out_path / target_filename
                
                # 保存图片
                if img_format.lower() == 'png':
                    img.save(target_file_path, 'PNG')
                else:
                    img.save(target_file_path, 'JPEG', quality=95)
                    
                total_pages += 1
            
            converted_count += 1
            print(f"✅ 完成: {ofd_file.name} (共 {len(images)} 页)")
            
        except Exception as e:
            print(f"❌ 转换 {ofd_file.name} 时出错: {str(e)}")

    print(f"\n🎉 批量转换任务结束！")
    print(f"👉 成功转换文件数: {converted_count}")
    print(f"👉 生成图片总页数: {total_pages}")
    print(f"📂 输出目录: {out_path}")

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="批量将 OFD 文件转换为图片 (JPG/PNG)")
    parser.add_argument("-s", "--src", required=True, help="源文件夹路径 (包含 .ofd)")
    parser.add_argument("-o", "--out", required=True, help="输出文件夹路径")
    parser.add_argument("-f", "--format", default="jpg", choices=["jpg", "png"], help="输出图片格式 (jpg/png)")
    
    args = parser.parse_args()
    convert_ofd_to_image(args.src, args.out, args.format)
