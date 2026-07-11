"""
文件处理工具类
提供统一的文件路径处理、文件存在性检查、目录创建等功能
"""

import sys
from pathlib import Path
from typing import Optional, Union


class FileUtils:
    """文件处理工具类"""

    @staticmethod
    def resolve_path(path: Union[str, Path]) -> Path:
        """解析并返回绝对路径"""
        return Path(path).resolve()

    @staticmethod
    def ensure_directory_exists(directory: Union[str, Path]) -> None:
        """确保目录存在，如果不存在则创建"""
        path = FileUtils.resolve_path(directory)
        path.mkdir(parents=True, exist_ok=True)

    @staticmethod
    def get_output_path(input_path: Union[str, Path], output_spec: Union[str, Path]) -> Path:
        """
        确定输出路径
        如果output_spec是目录，则在目录下使用输入文件的名称
        如果是文件路径，则直接使用
        """
        input_path = FileUtils.resolve_path(input_path)
        output_spec = FileUtils.resolve_path(output_spec)

        if output_spec.is_dir() or not output_spec.suffix:
            FileUtils.ensure_directory_exists(output_spec)
            return output_spec / input_path.name
        else:
            FileUtils.ensure_directory_exists(output_spec.parent)
            return output_spec

    @staticmethod
    def handle_existing_file(output_path: Path) -> Path:
        """
        处理已存在的输出文件，避免覆盖
        如果文件已存在，在文件名后添加后缀
        """
        if not output_path.exists():
            return output_path

        stem = output_path.stem
        ext = output_path.suffix
        counter = 1

        while True:
            candidate = output_path.parent / f"{stem}_{counter}{ext}"
            if not candidate.exists():
                return candidate
            counter += 1

    @staticmethod
    def validate_file_exists(file_path: Union[str, Path]) -> bool:
        """验证文件是否存在"""
        return FileUtils.resolve_path(file_path).is_file()

    @staticmethod
    def validate_directory_exists(directory: Union[str, Path]) -> bool:
        """验证目录是否存在"""
        return FileUtils.resolve_path(directory).is_dir()