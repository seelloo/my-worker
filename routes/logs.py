"""Job execution logs API routes."""
from flask import Blueprint, jsonify, request

from scripts.job_logger import get_log, list_logs, get_stats, delete_old_logs

logs_bp = Blueprint("logs", __name__, url_prefix="/api/logs")


@logs_bp.route("", methods=["GET"])
def api_list_logs():
    task_id = request.args.get("task_id")
    status = request.args.get("status")
    limit = request.args.get("limit", 50, type=int)
    offset = request.args.get("offset", 0, type=int)

    logs = list_logs(task_id=task_id, status=status, limit=limit, offset=offset)
    return jsonify({"status": "success", "logs": logs, "total": len(logs)})


@logs_bp.route("/<log_id>", methods=["GET"])
def api_get_log(log_id):
    log = get_log(log_id)
    if not log:
        return jsonify({"status": "error", "message": "Log not found"}), 404
    return jsonify({"status": "success", "log": log})


@logs_bp.route("/stats", methods=["GET"])
def api_get_stats():
    return jsonify({"status": "success", "stats": get_stats()})


@logs_bp.route("/cleanup", methods=["POST"])
def api_cleanup():
    days = request.json.get("days", 30) if request.json else 30
    deleted = delete_old_logs(days=days)
    return jsonify({"status": "success", "deleted": deleted})
