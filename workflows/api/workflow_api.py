"""
工作流API
提供工作流的创建、执行、查询和管理接口
"""

from flask import Blueprint, jsonify, request

from scripts.workflow_engine import WorkflowEngine
from scripts.task_registry import TASKS
from workflows.models.workflow import Workflow
from utils.workflow_utils import WorkflowReportGenerator

workflow_bp = Blueprint('workflow', __name__)
workflow_engine = WorkflowEngine(TASKS)


@workflow_bp.route('/api/workflows', methods=['GET'])
def list_workflows():
    workflows = []
    for wf_id, workflow in workflow_engine.workflow_store.items():
        workflows.append({
            "id": wf_id,
            "name": workflow.name,
            "description": workflow.description,
            "status": workflow.status.value,
            "task_count": len(workflow.tasks),
            "created_at": workflow.created_at.isoformat(),
            "updated_at": workflow.updated_at.isoformat()
        })
    return jsonify({"status": "success", "workflows": workflows})


@workflow_bp.route('/api/workflows', methods=['POST'])
def create_workflow():
    try:
        data = request.json
        if not data:
            return jsonify({"status": "error", "message": "请求数据为空"}), 400
        workflow = workflow_engine.create_workflow(data)
        return jsonify({"status": "success", "workflow_id": workflow.id, "workflow": workflow.to_dict()}), 201
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 400


@workflow_bp.route('/api/workflows/<workflow_id>', methods=['GET'])
def get_workflow(workflow_id):
    workflow = workflow_engine.workflow_store.get(workflow_id)
    if not workflow:
        return jsonify({"status": "error", "message": "工作流不存在"}), 404
    return jsonify({"status": "success", "workflow": workflow.to_dict()})


@workflow_bp.route('/api/workflows/<workflow_id>', methods=['PUT'])
def update_workflow(workflow_id):
    try:
        data = request.json
        if not data:
            return jsonify({"status": "error", "message": "请求数据为空"}), 400
        existing = workflow_engine.workflow_store.get(workflow_id)
        if not existing:
            return jsonify({"status": "error", "message": "工作流不存在"}), 404
        updated = Workflow.from_dict(data)
        workflow_engine.workflow_store[workflow_id] = updated
        return jsonify({"status": "success", "workflow": updated.to_dict()})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 400


@workflow_bp.route('/api/workflows/<workflow_id>', methods=['DELETE'])
def delete_workflow(workflow_id):
    if workflow_id not in workflow_engine.workflow_store:
        return jsonify({"status": "error", "message": "工作流不存在"}), 404
    del workflow_engine.workflow_store[workflow_id]
    return jsonify({"status": "success"})


@workflow_bp.route('/api/workflows/<workflow_id>/execute', methods=['POST'])
def execute_workflow(workflow_id):
    try:
        execution = workflow_engine.execute_workflow(workflow_id)
        if not execution:
            return jsonify({"status": "error", "message": "工作流不存在"}), 404
        return jsonify({"status": "success", "execution_id": execution.id, "execution": execution.to_dict()})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


@workflow_bp.route('/api/workflows/executions/<execution_id>', methods=['GET'])
def get_execution(execution_id):
    execution = workflow_engine.execution_store.get(execution_id)
    if not execution:
        return jsonify({"status": "error", "message": "执行记录不存在"}), 404
    return jsonify({"status": "success", "execution": execution.to_dict()})


@workflow_bp.route('/api/workflows/executions/<execution_id>/report', methods=['GET'])
def get_execution_report(execution_id):
    execution = workflow_engine.execution_store.get(execution_id)
    if not execution:
        return jsonify({"status": "error", "message": "执行记录不存在"}), 404
    report = WorkflowReportGenerator.generate(execution)
    return jsonify({"status": "success", "report": report})


@workflow_bp.route('/api/workflows/executions', methods=['GET'])
def list_executions():
    executions = []
    for exec_id, execution in workflow_engine.execution_store.items():
        executions.append({
            "id": exec_id,
            "workflow_id": execution.workflow_id,
            "workflow_name": execution.workflow_name,
            "status": execution.status,
            "start_time": execution.start_time.isoformat() if execution.start_time else None,
            "end_time": execution.end_time.isoformat() if execution.end_time else None,
            "duration": execution.duration,
            "task_count": len(execution.task_executions)
        })
    return jsonify({"status": "success", "executions": executions})


@workflow_bp.route('/api/workflows/tasks', methods=['GET'])
def list_tasks():
    tasks = []
    for task_id, task_def in TASKS.items():
        tasks.append({
            "id": task_id,
            "name": task_def.get("name", ""),
            "category": task_def.get("category", ""),
            "description": task_def.get("description", ""),
            "params": task_def.get("params", [])
        })
    return jsonify({"status": "success", "tasks": tasks})