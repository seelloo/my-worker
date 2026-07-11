import os
import logging

# 数据库配置
DB_PATH = os.getenv('DB_PATH', 'invoices_archive.db')

# 服务器配置
HOST = os.getenv('HOST', '127.0.0.1')
PORT = int(os.getenv('PORT', 5000))
DEBUG = os.getenv('DEBUG', 'true').lower() == 'true'

# 文件上传配置
MAX_CONTENT_LENGTH = int(os.getenv('MAX_CONTENT_LENGTH', 50 * 1024 * 1024))  # 50MB
ALLOWED_EXTENSIONS = {'pdf', 'xlsx', 'xls', 'csv', 'ofd'}

# PDF 存储目录
PDF_UPLOAD_DIR = os.getenv('PDF_UPLOAD_DIR', 'uploads/pdfs')

# 日志配置
LOG_LEVEL = os.getenv('LOG_LEVEL', 'INFO' if not DEBUG else 'DEBUG')
LOG_FORMAT = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
LOG_FILE = os.getenv('LOG_FILE', 'app.log')

# 限流配置
RATE_LIMIT_ENABLED = os.getenv('RATE_LIMIT_ENABLED', 'true').lower() == 'true'
RATE_LIMIT_DEFAULT = int(os.getenv('RATE_LIMIT_DEFAULT', 100))  # 默认每分钟100次
RATE_LIMIT_WINDOW = int(os.getenv('RATE_LIMIT_WINDOW', 60))  # 时间窗口60秒
