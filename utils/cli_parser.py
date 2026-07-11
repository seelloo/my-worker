"""
命令行参数处理工具类
提供统一的命令行参数解析功能
"""

import argparse
import sys
from typing import Any, Dict, List, Optional


class CLIParser:
    """命令行参数处理工具类"""

    @staticmethod
    def create_parser(description: str) -> argparse.ArgumentParser:
        """创建基本的命令行参数解析器"""
        return argparse.ArgumentParser(description=description)

    @staticmethod
    def add_common_arguments(parser: argparse.ArgumentParser) -> None:
        """添加常见的命令行参数"""
        parser.add_argument("-f", "--file", required=True, help="源文件路径")
        parser.add_argument("-o", "--out", required=True, help="输出文件保存的目录或指定的文件路径")

    @staticmethod
    def add_sheet_arguments(parser: argparse.ArgumentParser) -> None:
        """添加Sheet相关的参数"""
        parser.add_argument("-t", "--sheet", required=False, default="", help="目标Sheet名称（选填）")

    @staticmethod
    def add_column_arguments(parser: argparse.ArgumentParser) -> None:
        """添加列相关的参数"""
        parser.add_argument("-c", "--column", required=True, help="目标列字母 (例如: P, A, BD)")
        parser.add_argument("-s", "--sheet", default="Collected_Data", help="汇总结果所放置的新 Sheet 名称 (默认: Collected_Data)")

    @staticmethod
    def add_rename_arguments(parser: argparse.ArgumentParser) -> None:
        """添加重命名相关的参数"""
        parser.add_argument("-r", "--remove", help="需要从 Sheet 名称中被完全去除的字符串")
        parser.add_argument("-a", "--add", help="要在 Sheet 名称中增加的字符串")
        parser.add_argument("-p", "--pos", choices=["front", "back"], default="back", help="增加字符串的位置: front(前面) 或 back(后面)，默认为 back")

    @staticmethod
    def parse_arguments(parser: argparse.ArgumentParser) -> argparse.Namespace:
        """解析命令行参数"""
        return parser.parse_args()

    @staticmethod
    def validate_column_argument(column: str) -> None:
        """验证列参数"""
        if not column.isalpha():
            print("错误: 目标列必须是英文字母，例如 'P'。")
            sys.exit(1)

    @staticmethod
    def validate_rename_arguments(remove_str: Optional[str], add_str: Optional[str]) -> None:
        """验证重命名参数"""
        if not remove_str and not add_str:
            print("错误: 必须指定至少一个操作参数: --remove 或 --add")
            sys.exit(1)