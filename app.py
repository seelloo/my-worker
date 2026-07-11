import os
import sys

# 强制将项目根目录加入模块搜索路径，解决部分环境下的 ModuleNotFoundError
project_root = os.path.dirname(os.path.abspath(__file__))
if project_root not in sys.path:
    sys.path.insert(0, project_root)

import logging
from flask import Flask, jsonify, render_template

# 配置
from config import HOST, PORT, DEBUG, MAX_CONTENT_LENGTH, LOG_LEVEL, LOG_FORMAT, LOG_FILE, RATE_LIMIT_ENABLED
import psutil

# 统一错误处理
from utils.error_handler import register_error_handlers

# 限流
from utils.rate_limiter import register_rate_limit_handlers

# 配置日志
def setup_logging():
    logging.basicConfig(
        level=getattr(logging, LOG_LEVEL),
        format=LOG_FORMAT
    )
    # 文件日志
    file_handler = logging.FileHandler(LOG_FILE, encoding='utf-8')
    file_handler.setLevel(getattr(logging, LOG_LEVEL))
    file_handler.setFormatter(logging.Formatter(LOG_FORMAT))
    logging.getLogger().addHandler(file_handler)

setup_logging()
logger = logging.getLogger(__name__)

# 工作流系统
from workflows.api.workflow_api import workflow_bp
from workflows.models.database import init_database

# Blueprints
from routes.tasks import tasks_bp
from routes.excel import excel_bp
from routes.pdf import pdf_bp
from routes.invoice import invoice_bp
from routes.schedules import schedules_bp
from routes.logs import logs_bp
from routes.queue import queue_bp
from routes.preview import preview_bp
from routes.github import github_bp

app = Flask(__name__)

# 注册统一错误处理器
register_error_handlers(app)

# 注册限流处理器
if RATE_LIMIT_ENABLED:
    register_rate_limit_handlers(app)
    logger.info(f"限流已启用: {RATE_LIMIT_ENABLED} requests per minute")

# 注册蓝图
app.register_blueprint(tasks_bp)
app.register_blueprint(excel_bp)
app.register_blueprint(pdf_bp)
app.register_blueprint(invoice_bp)
app.register_blueprint(workflow_bp)
app.register_blueprint(schedules_bp)
app.register_blueprint(logs_bp)
app.register_blueprint(queue_bp)
app.register_blueprint(preview_bp)
app.register_blueprint(github_bp)

# 初始化工作流数据库
try:
    init_database()
except Exception as e:
    print(f"⚠️ 工作流数据库初始化失败: {e}")


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/workflows')
def workflow_list():
    return render_template('workflow_list.html')


@app.route('/workflow/editor')
def workflow_editor():
    return render_template('workflow_editor.html')


@app.route('/api/system/status')
def system_status():
    try:
        cpu_percent = psutil.cpu_percent(interval=None)
        memory = psutil.virtual_memory()
        mem_percent = memory.percent
        return jsonify({
            'status': 'success',
            'data': {
                'cpu': cpu_percent,
                'memory': mem_percent
            }
        })
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

if __name__ == '__main__':
    app.config['MAX_CONTENT_LENGTH'] = MAX_CONTENT_LENGTH
    print(f"Task Engine Server v3.0 [Blueprint Mode] is running on http://{HOST}:{PORT}")
    app.run(debug=DEBUG, host=HOST, port=PORT)
