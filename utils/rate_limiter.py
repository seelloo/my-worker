"""
请求限流工具
提供基于内存的简单限流功能
"""

import time
from functools import wraps
from threading import Lock
from flask import request, jsonify, g


class RateLimiter:
    """简单内存限流器"""

    def __init__(self):
        self._requests = {}
        self._lock = Lock()

    def is_allowed(self, key: str, max_requests: int, window_seconds: int) -> bool:
        """检查请求是否允许"""
        now = time.time()
        window_start = now - window_seconds

        with self._lock:
            if key not in self._requests:
                self._requests[key] = []

            # 清理过期记录
            self._requests[key] = [
                t for t in self._requests[key] if t > window_start
            ]

            if len(self._requests[key]) < max_requests:
                self._requests[key].append(now)
                return True
            return False

    def get_remaining(self, key: str, max_requests: int, window_seconds: int) -> int:
        """获取剩余请求次数"""
        now = time.time()
        window_start = now - window_seconds

        with self._lock:
            if key not in self._requests:
                return max_requests

            valid_requests = [
                t for t in self._requests[key] if t > window_start
            ]
            return max(0, max_requests - len(valid_requests))


# 全局限流器实例
limiter = RateLimiter()


def rate_limit(max_requests: int = 100, window_seconds: int = 60, key_func=None):
    """
    限流装饰器

    Args:
        max_requests: 时间窗口内最大请求数
        window_seconds: 时间窗口大小（秒）
        key_func: 获取限流 key 的函数，默认用 IP 地址
    """
    def decorator(f):
        @wraps(f)
        def wrapped(*args, **kwargs):
            # 获取限流 key
            if key_func:
                key = key_func()
            else:
                key = request.remote_addr or 'unknown'

            if not limiter.is_allowed(key, max_requests, window_seconds):
                remaining = 0
            else:
                remaining = limiter.get_remaining(key, max_requests, window_seconds)

            # 添加限流响应头
            g.rate_limit_remaining = remaining
            g.rate_limit_limit = max_requests

            response = f(*args, **kwargs)

            # 为响应添加限流头
            if hasattr(response, 'headers'):
                response.headers['X-RateLimit-Limit'] = str(max_requests)
                response.headers['X-RateLimit-Remaining'] = str(remaining)
                response.headers['X-RateLimit-Window'] = str(window_seconds)

            return response
        return wrapped
    return decorator


def register_rate_limit_handlers(app):
    """注册限流错误处理器"""

    @app.errorhandler(429)
    def rate_limit_exceeded(e):
        return jsonify({
            "status": "error",
            "message": "请求过于频繁，请稍后再试",
            "error_code": "ERR_RATE_LIMIT"
        }), 429
