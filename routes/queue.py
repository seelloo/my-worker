"""Task Queue API routes."""
from flask import Blueprint, jsonify, request

from scripts.task_queue import (
    enqueue, batch_enqueue, get_all, get_stats, remove, retry_task,
    clear_completed, clear_all, start_processing, stop_processing, is_processing
)

queue_bp = Blueprint("queue", __name__, url_prefix="/api/queue")


@queue_bp.route("", methods=["GET"])
def api_get_queue():
    status = request.args.get("status")
    items = get_all(status=status)
    return jsonify({"status": "success", "queue": items, "total": len(items)})


@queue_bp.route("", methods=["POST"])
def api_enqueue():
    data = request.json or {}
    task_id = data.get("task_id")
    params = data.get("params", {})
    priority = data.get("priority", 0)

    if not task_id:
        return jsonify({"status": "error", "message": "缺少 task_id"}), 400

    result = enqueue(task_id, params, priority)
    status_code = 201 if result.get("status") == "ok" else 400
    return jsonify(result), status_code


@queue_bp.route("/batch", methods=["POST"])
def api_batch_enqueue():
    data = request.json or {}
    tasks = data.get("tasks", [])
    if not tasks:
        return jsonify({"status": "error", "message": "缺少 tasks 数组"}), 400

    result = batch_enqueue(tasks)
    return jsonify(result), 201


@queue_bp.route("/<queue_id>", methods=["DELETE"])
def api_remove(queue_id):
    result = remove(queue_id)
    status_code = 200 if result.get("status") == "ok" else 400
    return jsonify(result), status_code


@queue_bp.route("/<queue_id>/retry", methods=["POST"])
def api_retry(queue_id):
    result = retry_task(queue_id)
    status_code = 200 if result.get("status") == "ok" else 400
    return jsonify(result), status_code


@queue_bp.route("/clear", methods=["POST"])
def api_clear():
    data = request.json or {}
    scope = data.get("scope", "completed")
    if scope == "all":
        deleted = clear_all()
    else:
        deleted = clear_completed()
    return jsonify({"status": "success", "deleted": deleted})


@queue_bp.route("/start", methods=["POST"])
def api_start():
    result = start_processing()
    return jsonify(result)


@queue_bp.route("/stop", methods=["POST"])
def api_stop():
    result = stop_processing()
    return jsonify(result)


@queue_bp.route("/status", methods=["GET"])
def api_status():
    return jsonify({
        "status": "success",
        "stats": get_stats(),
        "processing": is_processing()
    })
