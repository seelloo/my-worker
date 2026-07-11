import os
from test_file import zhcFileOpera

def test():
    fo = zhcFileOpera()
    
    # 请将下面的 image_path 替换为您实际想要测试的图片路径
    image_path = 'test_table_image.png' 
    output_excel = 'test_table_output.xlsx'
    
    if not os.path.exists(image_path):
        print(f"提示: {image_path} 不存在。请准备一张带有表格的图片，将其命名为 {image_path}，然后再次运行此脚本。")
        return
        
    print("--------------------------------------------------")
    print(f"正在测试图片: {image_path}")
    print(f"目标保存位置: {output_excel}")
    print("--------------------------------------------------")
    
    success = fo.extract_table_from_image(image_path, output_excel)
    
    if success:
        print("🎉 测试成功！已生成 Excel 文件。")
    else:
        print("💥 测试失败，请查看上方错误信息。")

if __name__ == '__main__':
    test()
