import os
import shutil
import time
from pathlib import Path
from typing import Union
import pandas as pd



class zhcFileOpera:
    """
    文件与目录操作工具类。
    涵盖：路径探测、文件遍历、批量拷贝、批量创建目录、
          TXT 读写、路径转 URL、以及批量 PDF → Word 转换。
    """

    # ------------------------------------------------------------------ #
    #  路径探测
    # ------------------------------------------------------------------ #

    @staticmethod
    def path_type(path: str) -> str:
        """
        判断路径指向的是文件、目录还是其他（不存在）。
        返回值: 'file' | 'dir' | 'other'
        """
        if os.path.isfile(path):
            return 'file'
        if os.path.isdir(path):
            return 'dir'
        return 'other'

    # ------------------------------------------------------------------ #
    #  目录遍历（统一入口，原来三个方法合并）
    # ------------------------------------------------------------------ #

    @staticmethod
    def scan(path: str,
             include_files: bool = True,
             include_dirs: bool = False,
             name_only: bool = False,
             ext_filter: tuple = ()) -> list:
        """
        递归扫描目录，返回结果列表。

        参数:
            path        -- 要扫描的根目录
            include_files -- 是否返回文件（默认 True）
            include_dirs  -- 是否返回子目录（默认 False）
            name_only     -- True 时只返回文件名，False 时返回完整路径
            ext_filter    -- 扩展名白名单元组，如 ('.pdf', '.xlsx')；空表示不过滤

        示例:
            # 获取目录下所有 PDF 完整路径
            scan('/data', ext_filter=('.pdf',))
            # 获取所有子目录
            scan('/data', include_files=False, include_dirs=True)
        """
        result = []
        # 使用 os.walk 递归遍历指定路径下的所有层级目录
        for root, dirs, files in os.walk(path):
            # 1. 收集文件
            if include_files:
                for f in files:
                    # 若配置了扩展名白名单，且当前文件后缀不在名单内，则跳过
                    if ext_filter and Path(f).suffix.lower() not in ext_filter:
                        continue
                    # 根据参数决定是仅保留文件名，还是保留完整绝对/相对路径
                    result.append(f if name_only else os.path.join(root, f))
            # 2. 收集目录
            if include_dirs:
                for d in dirs:
                    result.append(d if name_only else os.path.join(root, d))

        # 兼容原有的目录查询逻辑：如果只查目录且未查出任何子目录，则将自身路径作为结果返回
        if include_dirs and not include_files and not result:
            result.append(path)

        return result

    def show(self, path: str) -> None:
        """打印目录下所有文件和子目录的完整路径。"""
        for item in self.scan(path, include_files=True, include_dirs=True):
            print(item)

    # ------------------------------------------------------------------ #
    #  TXT 读写
    # ------------------------------------------------------------------ #

    @staticmethod
    def read_txt(filepath: str, encoding: str = 'gbk') -> list:
        """
        读取 TXT 文件，按换行符拆分，自动过滤空行。
        返回非空字符串列表。
        """
        with open(filepath, 'r', encoding=encoding) as f:
            lines = f.read().splitlines()
        return [line for line in lines if line.strip()]

    @staticmethod
    def write_txt(filepath: str, data: list, encoding: str = 'utf-8') -> None:
        """
        将列表写入 TXT 文件，每项占一行。
        自动清理列表项中的方括号、单引号、逗号。
        """
        cleaned = [
            str(item).replace('[', '').replace(']', '')
                     .replace("'", '').replace(',', '')
            for item in data
        ]
        with open(filepath, 'w', encoding=encoding) as f:
            f.writelines(line + '\n' for line in cleaned)
        print(f"保存成功 → {filepath}")

    # ------------------------------------------------------------------ #
    #  批量拷贝文件
    # ------------------------------------------------------------------ #

    def bat_copy(self,
                 file_list: Union[str, list],
                 source_path: str,
                 target_path: str) -> None:
        """
        批量拷贝文件。

        参数:
            file_list   -- 要拷贝的文件名列表，或包含文件名的 TXT 文件路径
            source_path -- 在此目录（含子目录）中搜索文件
            target_path -- 拷贝目标目录
        """
        # 统一来源：如果是合法的本地 TXT 文件，则读取文件中的多行文本作为列表
        if isinstance(file_list, str) and os.path.isfile(file_list):
            names = self.read_txt(file_list)
        else:
            names = list(file_list)

        # 获取源路径下的所有子目录
        sub_dirs = self.scan(source_path, include_files=False, include_dirs=True)
        # 确保目标目录已存在，不存在则递归创建
        os.makedirs(target_path, exist_ok=True)

        # 遍历每一个子目录
        for directory in sub_dirs:
            # 遍历用户想要找的文件名称
            for name in names:
                src = os.path.join(directory, name)
                dst = os.path.join(target_path, name)
                try:
                    # 尝试复制该文件（保留原文件的属性，如修改时间）
                    shutil.copy2(src, dst)
                    print(f"  ✓ {src}")
                except (FileNotFoundError, NotADirectoryError, OSError):
                    # 如果子目录中不存在该文件或发生普通 IO 错误，直接静默跳过
                    pass
            print(f"[done] {directory}")

    # ------------------------------------------------------------------ #
    #  批量创建目录
    # ------------------------------------------------------------------ #

    def make_dirs(self,
                  parent_path: str,
                  name_source: Union[str, list, dict]) -> list:
        """
        在 parent_path 下批量创建子目录。

        参数:
            parent_path  -- 父目录路径
            name_source  -- 子目录名来源，支持三种类型：
                            list  -- 直接使用列表
                            dict  -- 使用字典的 key
                            str   -- 视为 TXT 文件路径，读取后使用
        返回已创建的目录路径列表。
        """
        if isinstance(name_source, list):
            names = name_source
        elif isinstance(name_source, dict):
            names = list(name_source.keys())
        elif isinstance(name_source, str) and os.path.isfile(name_source):
            names = self.read_txt(name_source)
        else:
            raise ValueError(f"不支持的 name_source 类型或路径不存在: {name_source}")

        created = []
        for name in names:
            folder = os.path.join(parent_path, name)
            os.makedirs(folder, exist_ok=True)
            print(f"  ✓ {folder}")
            created.append(folder)
        return created

    # ------------------------------------------------------------------ #
    #  路径转 HTML 链接
    # ------------------------------------------------------------------ #

    @staticmethod
    def paths_to_html_links(file_paths: list, base_url: str, strip_levels: int = 0) -> list:
        """
        将本地路径列表转换为 HTML <a> 标签列表。

        参数:
            file_paths   -- 本地文件路径列表
            base_url     -- URL 前缀，如 'http://example.com/files'
            strip_levels -- 从路径左侧去掉的层级数（用于拼接相对 URL）
        """
        links = []
        for path in file_paths:
            parts = Path(path).parts[strip_levels:]
            url_path = '/'.join(parts)
            full_url = f"{base_url.rstrip('/')}/{url_path}"
            links.append(f'<p><a href="{full_url}">{path}</a></p>')
        return links

    # ------------------------------------------------------------------ #
    #  批量 PDF → Word (docx)
    # ------------------------------------------------------------------ #

    @staticmethod
    def _convert_one_pdf(input_pdf: str, output_docx: str) -> None:
        """
        将单个 PDF 转换为 docx。
        依赖: pdf2docx (pip install pdf2docx)
        """
        try:
            from pdf2docx import Converter
        except ImportError:
            raise ImportError("请先安装依赖: pip install pdf2docx")

        cv = Converter(input_pdf)
        try:
            cv.convert(output_docx)
        finally:
            cv.close()

    def bat_pdf_to_word(self,
                        pdf_source: str,
                        output_dir: str,
                        overwrite: bool = False) -> dict:
        """
        批量将目录下所有 PDF 转换为 Word（.docx）。

        参数:
            pdf_source -- PDF 文件所在目录（或单个 PDF 文件路径）
            output_dir -- 输出 docx 文件的目录
            overwrite  -- 是否覆盖已存在的 docx 文件，默认 False

        返回 dict: {'success': [...], 'skipped': [...], 'failed': {...}}
        """
        # 确定待处理文件列表
        if os.path.isfile(pdf_source):
            pdf_files = [pdf_source]
        else:
            pdf_files = self.scan(pdf_source, ext_filter=('.pdf',))

        os.makedirs(output_dir, exist_ok=True)

        result = {'success': [], 'skipped': [], 'failed': {}}
        total = len(pdf_files)

        print(f"共检测到 {total} 个 PDF 文件，开始转换…\n")
        start = time.time()

        for idx, pdf_path in enumerate(pdf_files, 1):
            stem = Path(pdf_path).stem
            docx_path = os.path.join(output_dir, stem + '.docx')

            print(f"[{idx}/{total}] {Path(pdf_path).name}", end=' → ')

            if os.path.exists(docx_path) and not overwrite:
                print("跳过（已存在）")
                result['skipped'].append(pdf_path)
                continue

            try:
                self._convert_one_pdf(pdf_path, docx_path)
                print(f"✓ 完成")
                result['success'].append(docx_path)
            except Exception as e:
                print(f"✗ 失败: {e}")
                result['failed'][pdf_path] = str(e)

        elapsed = time.time() - start
        print(f"\n转换完毕 | 成功: {len(result['success'])}  "
              f"跳过: {len(result['skipped'])}  "
              f"失败: {len(result['failed'])}  "
              f"耗时: {elapsed:.1f}s")
        return result

    # ------------------------------------------------------------------ #
    #  基于 Excel 的批量移动与重命名
    # ------------------------------------------------------------------ #

    def bat_move_and_rename_by_excel(
        self,
        agent_excel_path: str,
        src_base_dir: str,
        target_base_dir: str,
        mode: str = 'mapping',
        month_str: str = '',
        category_mapping: dict = None,
        src_col_idx: int = 2,
        target_folder_col_idx: int = 4,
        file_ext: str = '.rar',
        is_move: bool = False
    ) -> None:
        """
        基于代理商信息 Excel，批量拷贝（或移动）并重命名文件。支持两种模式：
        
        模式 1 ('mapping'): 基于分类映射字典的多类别分发
            - 遍历 category_mapping，为每个分类拷贝文件。
            - 假定：Excel 中索引 1 为源文件基础名，索引 2 为目标子目录标识。
            - 需要提供 month_str 参数。
            
        模式 2 ('direct'): 基于 Excel 直接映射的单一文件分发
            - 直接遍历 Excel 行数据。
            - 假定：src_col_idx 指定源文件名，target_folder_col_idx 指定目标子文件夹。
            - 需要提供 file_ext 参数（如 '.rar'）。

        参数:
            agent_excel_path: 包含代理商/部门信息的 Excel 路径。
            src_base_dir:     源文件基础目录。
            target_base_dir:  目标文件基础目录。
            mode:             运行模式，'mapping' 或 'direct'。
            month_str:        年月字符串（'mapping' 模式用）。
            category_mapping: 分类映射字典（'mapping' 模式用）。
            src_col_idx:      源文件名所在的列索引（'direct' 模式用）。
            target_folder_col_idx: 目标子文件夹所在的列索引（'direct' 模式用）。
            file_ext:         文件后缀名（'direct' 模式用）。
            is_move:          True 表示剪切移动，False 表示复制（默认）。
        """
        
        try:
            # 利用 Pandas 读取 Excel，并立即转化为二维 list 以提高遍历性能
            bzlist_info = pd.read_excel(agent_excel_path)
            bzbm_list = bzlist_info.values.tolist()
        except Exception as e:
            print(f"❌ 读取 Excel 数据失败: {e}")
            return

        action_name = "移动" if is_move else "复制"
        success_count = 0
        print(f"🚀 开始批量{action_name}操作 (模式: {mode})，共有 {len(bzbm_list)} 条记录...")

        # ---------------- 模式 1：依据字典映射进行多分类循环 ----------------
        if mode == 'mapping':
            # 如果没有传入字典，则使用项目内默认的类目映射规则
            if category_mapping is None:
                category_mapping = {
                    'KPI': 'KPI', '补回': '补回', '佣金清单': '酬金',
                    '代理商': '代理商', '店中商': '店中商', '工差': '工差',
                    '积分清单': '积分', '门店运补': '运补'
                }
            # 外层循环：遍历分类映射
            for name, directory in category_mapping.items():
                # 内层循环：遍历 Excel 中的每一行数据
                for item in bzbm_list:
                    # 确保该行有效，能够提取到源标识（索引1）和目标标识（索引2）
                    if len(item) < 3:
                        continue

                    # 提取并清理字符串前后的空格
                    src_id = str(item[1]).strip()
                    target_id = str(item[2]).strip()

                    # 组装完整的源文件名与目标文件名
                    src_filename = f"{src_id}.xlsx"
                    target_filename = f"{month_str}-{target_id}-{name}.xlsx"

                    # 拼接最终的源绝对路径
                    src_path = os.path.join(src_base_dir, directory, src_filename)
                    # 拼接目标文件夹的绝对路径及目标文件路径
                    target_dir_full = os.path.join(target_base_dir, target_id)
                    target_path = os.path.join(target_dir_full, target_filename)

                    # 检查目标文件夹是否存在，如果不存在则自动创建
                    os.makedirs(target_dir_full, exist_ok=True)

                    try:
                        # 根据 is_move 标志执行 移动(剪切) 或 复制
                        if is_move:
                            shutil.move(src_path, target_path)
                        else:
                            shutil.copy2(src_path, target_path)
                        success_count += 1
                    except FileNotFoundError:
                        # 找不到源文件是常态，静默处理以继续下一个
                        pass
                    except (NotADirectoryError, OSError) as e:
                        print(f"⚠️ {action_name}出错 [{src_filename}]: {e}")

        # ---------------- 模式 2：单一列直接直连匹配提取 ----------------
        elif mode == 'direct':
            for item in bzbm_list:
                # 检查该行是否含有配置所对应的目标列
                if len(item) <= max(src_col_idx, target_folder_col_idx):
                    continue

                # 直接按指定列索引提取源文件名与目标子文件夹名
                src_val = str(item[src_col_idx]).strip()
                target_folder_val = str(item[target_folder_col_idx]).strip()

                # 此模式下不改变文件名，仅补充后缀，并且放在指定的子文件夹中
                src_filename = f"{src_val}{file_ext}"
                target_filename = src_filename

                src_path = os.path.join(src_base_dir, src_filename)
                target_dir_full = os.path.join(target_base_dir, target_folder_val)
                target_path = os.path.join(target_dir_full, target_filename)

                os.makedirs(target_dir_full, exist_ok=True)

                try:
                    if is_move:
                        shutil.move(src_path, target_path)
                    else:
                        shutil.copy2(src_path, target_path)
                    success_count += 1
                except FileNotFoundError:
                        pass
                except (NotADirectoryError, OSError) as e:
                    print(f"⚠️ {action_name}出错 [{src_filename}]: {e}")
        else:
            print(f"❌ 不支持的模式: {mode}")
            return

        print(f"✅ 批量{action_name}完成！成功处理了 {success_count} 个文件。")

    # ------------------------------------------------------------------ #
    #  图片表格提取到 Excel
    # ------------------------------------------------------------------ #

    def extract_table_from_image(self, image_path: str, output_excel: str) -> bool:
        """
        从图片中提取表格并另存为 Excel 文件。
        依赖: img2table, paddleocr
        
        参数:
            image_path: 输入图片的绝对或相对路径
            output_excel: 输出 Excel 的路径
            
        返回:
            bool: 转换是否成功
        """
        try:
            import os
            os.environ["FLAGS_enable_pir_api"] = "0"
            os.environ["FLAGS_use_mkldnn"] = "0"
            os.environ["FLAGS_enable_mkldnn"] = "0"
            from img2table.document import Image
            from img2table.ocr import PaddleOCR
        except ImportError:
            print("❌ 缺少必要的依赖库，请先执行: pip install \"img2table[paddle]\" paddlepaddle paddleocr")
            return False

        if not os.path.exists(image_path):
            print(f"❌ 找不到图片文件: {image_path}")
            return False

        try:
            print(f"🚀 开始提取图片表格: {image_path}")
            # 实例化 OCR 引擎
            # lang='ch' 支持中英文混合识别
            # 关闭 mkldnn 避免在部分 CPU/Windows 环境下报 OneDnnContext 错误
            ocr = PaddleOCR(lang="ch", kw={"enable_mkldnn": False})

            # 实例化文档
            img = Image(src=image_path)

            # 提取并保存到 Excel
            # implicit_rows/borderless_tables 等参数可根据实际图片质量调整
            img.to_xlsx(dest=output_excel,
                        ocr=ocr,
                        implicit_rows=False,
                        borderless_tables=False,
                        min_confidence=50)
            print(f"✅ 成功提取表格并保存至: {output_excel}")
            return True
        except Exception as e:
            print(f"❌ 提取表格时发生错误: {e}")
            return False



# --------------------------------------------------------------------------- #
if __name__ == '__main__':
    fo = zhcFileOpera()
    print('zhcFileOpera 已就绪')
    fo.extract_table_from_image("活动2026.png", "2026_output.xlsx")
