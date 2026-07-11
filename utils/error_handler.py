"""
错误处理工具类
提供统一的错误处理和日志记录功能
"""

import sys
import traceback
import logging
from flask import jsonify, current_app
from werkzeug.exceptions import HTTPException


logger = logging.getLogger(__name__)


class APIException(Exception):
    """API 统一异常类"""

    def __init__(self, message: str, status_code: int = 500, error_code: str = None):
        super().__init__(message)
        self.message = message
        self.status_code = status_code
        self.error_code = error_code or f"ERR_{status_code}"

    def to_dict(self):
        return {
            "status": "error",
            "message": self.message,
            "error_code": self.error_code
        }


def register_error_handlers(app):
    """注册 Flask 统一错误处理器"""

    @app.errorhandler(APIException)
    def handle_api_exception(e):
        """处理自定义 API 异常"""
        current_app.logger.warning(f"API异常: {e.message} (code: {e.error_code})")
        return jsonify(e.to_dict()), e.status_code

    @app.errorhandler(400)
    def bad_request(e):
        return jsonify({
            "status": "error",
            "message": "请求参数错误",
            "error_code": "ERR_400"
        }), 400

    @app.errorhandler(404)
    def not_found(e):
        return jsonify({
            "status": "error",
            "message": "请求的资源不存在",
            "error_code": "ERR_404"
        }), 404

    @app.errorhandler(422)
    def unprocessable(e):
        return jsonify({
            "status": "error",
            "message": "请求数据无法处理",
            "error_code": "ERR_422"
        }), 422

    @app.errorhandler(500)
    def internal_error(e):
        current_app.logger.error(f"服务器内部错误: {str(e)}")
        return jsonify({
            "status": "error",
            "message": "服务器内部错误，请稍后重试",
            "error_code": "ERR_500"
        }), 500

    @app.errorhandler(Exception)
    def handle_exception(e):
        """处理所有未捕获的异常"""
        if isinstance(e, HTTPException):
            return e

        current_app.logger.error(f"未捕获的异常: {str(e)}\n{traceback.format_exc()}")
        return jsonify({
            "status": "error",
            "message": "服务器内部错误，请稍后重试",
            "error_code": "ERR_INTERNAL"
        }), 500


class ErrorHandler:
    """错误处理工具类"""

    @staticmethod
    def handle_file_not_found(file_path: str) -> None:
        """处理文件不存在错误"""
        print(f"错误: 源文件 '{file_path}' 不存在或不是一个有效的文件。")
        sys.exit(1)

    @staticmethod
    def handle_directory_not_found(directory_path: str) -> None:
        """处理目录不存在错误"""
        print(f"错误: 源目录 '{directory_path}' 不存在或不是有效目录。")
        sys.exit(1)

    @staticmethod
    def handle_file_read_error(file_path: str, error: Exception) -> None:
        """处理文件读取错误"""
        print(f"❌ 无法读取文件 '{file_path}': {error}")
        sys.exit(1)

    @staticmethod
    def handle_file_save_error(file_path: str, error: Exception) -> None:
        """处理文件保存错误"""
        print(f"❌ 无法保存文件 '{file_path}': {error}")
        sys.exit(1)

    @staticmethod
    def handle_excel_processing_error(file_path: str, error: Exception) -> None:
        """处理Excel处理错误"""
        print(f"❌ 处理 '{file_path}' 发生严重错误: {error}")
        sys.exit(1)

    @staticmethod
    def handle_workflow_error(message: str) -> None:
        """处理工作流错误"""
        print(f"❌ 工作流错误: {message}")

    @staticmethod
    def print_info(message: str) -> None:
        """打印信息消息"""
        print(message)

    @staticmethod
    def print_success(message: str) -> None:
        """打印成功消息"""
        print(f"✅ {message}")

    @staticmethod
    def print_error(message: str) -> None:
        """打印错误消息"""
        print(f"❌ {message}")

    @staticmethod
    def print_skip(message: str) -> None:
        """打印跳过消息"""
        print(f"⚠️ 跳过: {message}")