"""
安全工具模块
提供路径遍历防护、文件验证等安全功能
"""
from __future__ import annotations

# The rest of your imports...

import os
import re
from flask import jsonify
from werkzeug.utils import secure_filename as werkzeug_secure_filename
from config import ALLOWED_EXTENSIONS


def validate_path(path: str, allow_relative: bool = False) -> tuple[bool, str]:
    """
    验证路径安全性，防止路径遍历攻击

    Returns:
        (is_valid, error_message)
    """
    if not path:
        return False, "路径不能为空"

    # 防止空字节注入
    if '\x00' in path:
        return False, "路径包含非法字符"

    # 标准化路径并检查遍历
    try:
        abs_path = os.path.abspath(path)
        if not allow_relative:
            # 检查是否包含 ../ 或 ..\\
            if '..' in path or path.startswith('/') and not os.path.isabs(path) is False:
                # 更严格的检查
                normalized = os.path.normpath(path)
                if '..' in normalized:
                    return False, "路径不允许包含上级目录"

        # 检查路径长度
        if len(path) > 4096:
            return False, "路径长度超限"

        return True, ""
    except (ValueError, OSError) as e:
        return False, f"路径无效: {str(e)}"


def validate_filename(filename: str) -> tuple[bool, str]:
    """
    验证文件名安全性

    Returns:
        (is_valid, error_message)
    """
    if not filename:
        return False, "文件名不能为空"

    # 使用 werkzeug 的安全文件名
    safe_name = werkzeug_secure_filename(filename)

    if not safe_name:
        return False, "文件名包含非法字符"

    if len(safe_name) > 255:
        return False, "文件名过长"

    # 防止隐藏文件
    if safe_name.startswith('.'):
        return False, "不允许创建隐藏文件"

    return True, ""


def validate_file_extension(filename: str, allowed: set = None) -> tuple[bool, str]:
    """
    验证文件扩展名

    Returns:
        (is_valid, error_message)
    """
    allowed = allowed or ALLOWED_EXTENSIONS

    if '.' not in filename:
        return False, "文件没有扩展名"

    ext = filename.rsplit('.', 1)[1].lower()
    if ext not in allowed:
        return False, f"不支持的文件类型: {ext}，允许的类型: {', '.join(allowed)}"

    return True, ""


def validate_file_path(file_path: str, must_exist: bool = True, allowed_dirs: list = None) -> tuple[bool, str]:
    """
    综合验证文件路径安全性

    Args:
        file_path: 文件路径
        must_exist: 是否必须存在
        allowed_dirs: 允许的目录列表（为空则不限制）

    Returns:
        (is_valid, error_message)
    """
    # 1. 验证路径安全性
    is_valid, msg = validate_path(file_path)
    if not is_valid:
        return False, msg

    # 2. 验证文件名
    filename = os.path.basename(file_path)
    is_valid, msg = validate_filename(filename)
    if not is_valid:
        return False, msg

    # 3. 验证扩展名
    is_valid, msg = validate_file_extension(filename)
    if not is_valid:
        return False, msg

    # 4. 检查文件是否存在
    if must_exist and not os.path.isfile(file_path):
        return False, f"文件不存在: {filename}"

    # 5. 检查是否在允许的目录内
    if allowed_dirs:
        abs_file = os.path.abspath(file_path)
        is_allowed = any(os.path.commonpath([abs_file, os.path.abspath(d)]) == os.path.abspath(d) for d in allowed_dirs)
        if not is_allowed:
            return False, "文件不在允许的目录范围内"

    return True, ""


def validate_directory_path(dir_path: str, must_exist: bool = True) -> tuple[bool, str]:
    """
    验证目录路径安全性

    Returns:
        (is_valid, error_message)
    """
    # 验证路径安全性
    is_valid, msg = validate_path(dir_path)
    if not is_valid:
        return False, msg

    # 检查目录是否存在
    if must_exist and not os.path.isdir(dir_path):
        return False, f"目录不存在: {dir_path}"

    return True, ""


def require_params(data: dict, required_fields: list) -> tuple[bool, str]:
    """
    验证必需参数

    Returns:
        (is_valid, error_message)
    """
    missing = [f for f in required_fields if not data.get(f)]
    if missing:
        return False, f"缺少必需参数: {', '.join(missing)}"
    return True, ""
