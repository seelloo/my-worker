"""Schedule management API routes."""
from flask import Blueprint, jsonify, request

from scripts.scheduler import (
    create_schedule, list_schedules, toggle_schedule,
    delete_schedule, run_task_now
)

schedules_bp = Blueprint("schedules", __name__, url_prefix="/api/schedules")


@schedules_bp.route("", methods=["GET"])
def api_list_schedules():
    return jsonify({"status": "success", "schedules": list_schedules()})


@schedules_bp.route("", methods=["POST"])
def api_create_schedule():
    data = request.json or {}
    task_id = data.get("task_id")
    params = data.get("params", {})
    trigger_type = data.get("trigger_type", "cron")
    trigger_config = data.get("trigger_config", {})

    if not task_id:
        return jsonify({"status": "error", "message": "缺少 task_id"}), 400

    result = create_schedule(task_id, params, trigger_type, **trigger_config)
    if result["status"] == "ok":
        return jsonify(result), 201
    return jsonify(result), 400


@schedules_bp.route("/<schedule_id>", methods=["PUT"])
def api_update_schedule(schedule_id):
    data = request.json or {}
    schedules = list_schedules()
    current = next((s for s in schedules if s["id"] == schedule_id), None)
    if not current:
        return jsonify({"status": "error", "message": "Schedule not found"}), 404

    if "enabled" in data:
        toggle_schedule(schedule_id, data["enabled"])

    return jsonify({"status": "success", "message": "Schedule updated"})


@schedules_bp.route("/<schedule_id>/toggle", methods=["POST"])
def api_toggle_schedule(schedule_id):
    data = request.json or {}
    enabled = data.get("enabled", True)
    result = toggle_schedule(schedule_id, enabled)
    status_code = 200 if result["status"] == "ok" else 400
    return jsonify(result), status_code


@schedules_bp.route("/<schedule_id>", methods=["DELETE"])
def api_delete_schedule(schedule_id):
    result = delete_schedule(schedule_id)
    return jsonify(result)


@schedules_bp.route("/run_now", methods=["POST"])
def api_run_now():
    """Execute a task immediately without creating a schedule."""
    data = request.json or {}
    task_id = data.get("task_id")
    params = data.get("params", {})

    if not task_id:
        return jsonify({"status": "error", "message": "缺少 task_id"}), 400

    result = run_task_now(task_id, params)
    status_code = 200 if result.get("status") in ("completed", "running") else 400
    return jsonify(result), status_code
