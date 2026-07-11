import os
import sys
import time
import subprocess
from flask import Blueprint, Response, jsonify, request

from scripts.task_registry import TASKS, get_task_metadata
from scripts.excel_join_multiply import get_excel_header
from scripts.validator import validate_params
from scripts.job_logger import create_log, finish_log
from utils.flask_utils import stream_task
from utils.rate_limiter import rate_limit
from utils.security import validate_file_path, require_params
from config import RATE_LIMIT_ENABLED, RATE_LIMIT_DEFAULT, RATE_LIMIT_WINDOW

tasks_bp = Blueprint('tasks', __name__, url_prefix='/api')


def api_rate_limit():
    """API 专用限流装饰器"""
    if RATE_LIMIT_ENABLED:
        return rate_limit(max_requests=RATE_LIMIT_DEFAULT, window_seconds=RATE_LIMIT_WINDOW)
    return lambda f: f


@tasks_bp.route('/task_list', methods=['GET'])
def api_task_list():
    """动态获取所有注册任务"""
    response = jsonify({"status": "success", "tasks": get_task_metadata()})
    response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    return response


@tasks_bp.route('/run_task/<task_id>', methods=['POST'])
@api_rate_limit()
def api_run_task(task_id):
    """通过隔离引擎运行指定任务"""
    if task_id not in TASKS:
        return jsonify({"status": "error", "message": f"任务ID {task_id} 不存在"}), 404

    info = TASKS[task_id]
    data = request.json or {}

    is_valid, err_msg = validate_params(task_id, data, TASKS)
    if not is_valid:
        return jsonify({"status": "error", "message": f"输入校验不通过: {err_msg}"}), 422

    if "cmd_builder" in info:
        cmd = info["cmd_builder"](data)

        # Create log entry before execution
        log_id = create_log(task_id, info["name"], params=data, cmd=cmd)

        # Wrap streaming response to capture output and save to log
        return stream_task_with_log(cmd, log_id)

    return jsonify({"status": "error", "message": "该任务未配置有效的子进程指令生成器"}), 500


def stream_task_with_log(cmd, log_id):
    """Stream task output and save to log on completion."""
    if cmd[0] == "python":
        cmd[0] = sys.executable

    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"

    # 强制将工作目录设为项目根目录，确保子进程能找到 utils / scripts 等包
    project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    # 注入 PYTHONPATH，让子进程 Python 解释器也能找到项目根下的所有包
    existing_pythonpath = env.get("PYTHONPATH", "")
    env["PYTHONPATH"] = project_root + (os.pathsep + existing_pythonpath if existing_pythonpath else "")

    process = subprocess.Popen(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        bufsize=1,
        universal_newlines=True,
        encoding='utf-8',
        env=env,
        cwd=project_root
    )

    output_lines = []

    def generate():
        start = time.time()
        yield f"🚀 [Sandbox] 独立任务进程已启动: {' '.join(cmd)}\n"
        for line in iter(process.stdout.readline, ""):
            output_lines.append(line)
            yield line
        process.stdout.close()
        return_code = process.wait()
        duration = time.time() - start
        status = "completed" if return_code == 0 else "failed"

        if return_code == 0:
            yield "\n✅ 物理隔离进程执行圆满成功！\n"
        else:
            yield f"\n❌ 任务进程异常终止 (Status: {return_code})，请检查上方日志。\n"

        # Save to log
        finish_log(
            log_id, status,
            stdout="".join(output_lines),
            stderr="",
            duration=round(duration, 2)
        )

    return Response(generate(), mimetype='text/plain')


@tasks_bp.route('/get_columns', methods=['POST'])
def api_get_columns():
    """获取 Excel 表头（供关联计算使用）"""
    data = request.json or {}
    path = data.get('path', '').strip()

    # 安全验证
    is_valid, err_msg = validate_file_path(path, must_exist=True)
    if not is_valid:
        return jsonify({"status": "error", "message": err_msg}), 400

    try:
        return jsonify({"status": "success", "columns": get_excel_header(path)})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500
