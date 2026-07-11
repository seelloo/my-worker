"""
工作流引擎核心类
提供工作流的创建、执行和管理功能
"""

import sys
import json
import subprocess
from typing import Dict, Any, List, Optional
from datetime import datetime

from workflows.models.workflow import (
    Workflow, WorkflowTask, WorkflowExecution, TaskExecution,
    TaskType
)
from scripts.task_registry import TASKS
from utils.error_handler import ErrorHandler


class ExecutionContext:
    """工作流执行上下文，用于在任务间传递变量"""

    def __init__(self):
        self.variables: Dict[str, Any] = {}
        self.current_task: Optional[str] = None

    def set(self, name: str, value: Any) -> None:
        self.variables[name] = value

    def get(self, name: str) -> Optional[Any]:
        return self.variables.get(name)


class ResultHandler:
    """执行结果处理器"""

    def handle(self, execution: WorkflowExecution) -> None:
        self._save_to_file(execution)
        self._save_to_db(execution)
        self._gen_report(execution)

    def _save_to_file(self, execution: WorkflowExecution) -> None:
        pass

    def _save_to_db(self, execution: WorkflowExecution) -> None:
        pass

    def _gen_report(self, execution: WorkflowExecution) -> None:
        pass


class WorkflowEngine:
    """工作流引擎核心类"""

    def __init__(self, task_registry: Dict[str, Any]):
        self.task_registry = task_registry
        self.workflow_store: Dict[str, Workflow] = {}
        self.execution_store: Dict[str, WorkflowExecution] = {}

    def create_workflow(self, workflow_data: Dict[str, Any]) -> Workflow:
        try:
            workflow = Workflow.from_dict(workflow_data)
            workflow.validate()
            self.workflow_store[workflow.id] = workflow
            return workflow
        except Exception as e:
            ErrorHandler.handle_workflow_error(f"创建工作流失败: {e}")
            raise

    def execute_workflow(self, workflow_id: str) -> Optional[WorkflowExecution]:
        if workflow_id not in self.workflow_store:
            ErrorHandler.handle_workflow_error(f"工作流 {workflow_id} 不存在")
            return None

        workflow = self.workflow_store[workflow_id]
        execution = WorkflowExecution(
            workflow_id=workflow.id,
            workflow_name=workflow.name
        )

        try:
            execution.status = "running"
            execution.start_time = datetime.now()
            self._execute_tasks(workflow, execution)
            execution.status = "completed"
            execution.end_time = datetime.now()
            execution.duration = (execution.end_time - execution.start_time).total_seconds()
        except Exception as e:
            execution.status = "failed"
            execution.end_time = datetime.now()
            execution.duration = (execution.end_time - execution.start_time).total_seconds() if execution.start_time else 0
            execution.error = str(e)
            ErrorHandler.handle_workflow_error(f"工作流执行失败: {e}")

        self.execution_store[execution.id] = execution
        return execution

    def _execute_tasks(self, workflow: Workflow, execution: WorkflowExecution) -> None:
        task_map = {task.id: task for task in workflow.tasks}
        for task in workflow.tasks:
            if task.type == TaskType.SEQUENTIAL:
                self._execute_sequential_task(task, task_map, execution)
            elif task.type == TaskType.PARALLEL:
                self._execute_parallel_tasks(task, task_map, execution)
            elif task.type == TaskType.CONDITIONAL:
                self._execute_conditional_task(task, task_map, execution)
            elif task.type == TaskType.LOOP:
                self._execute_loop_task(task, task_map, execution)

    def _execute_sequential_task(self, task: WorkflowTask, task_map: Dict[str, WorkflowTask], execution: WorkflowExecution) -> None:
        task_execution = TaskExecution(
            execution_id=execution.id,
            task_id=task.task_id,
            task_name=self.task_registry[task.task_id]["name"]
        )
        try:
            task_execution.status = "running"
            task_execution.start_time = datetime.now()
            cmd = self._build_command(task)
            result = self._run_command(cmd)
            task_execution.status = "completed"
            task_execution.end_time = datetime.now()
            task_execution.duration = (task_execution.end_time - task_execution.start_time).total_seconds()
            task_execution.result = result
            execution.add_task_execution(task_execution)
            execution.add_log(f"顺序任务 {task.id} 执行完成")
        except Exception as e:
            task_execution.status = "failed"
            task_execution.end_time = datetime.now()
            task_execution.duration = (task_execution.end_time - task_execution.start_time).total_seconds() if task_execution.start_time else 0
            task_execution.error = str(e)
            task_execution.stderr = [str(e)]
            execution.add_task_execution(task_execution)
            execution.add_log(f"顺序任务 {task.id} 执行失败: {e}", level="error")

    def _execute_parallel_tasks(self, task: WorkflowTask, task_map: Dict[str, WorkflowTask], execution: WorkflowExecution) -> None:
        task_executions: List[TaskExecution] = []
        try:
            for sub_task_id in task.config.get("tasks", []):
                sub_task = task_map[sub_task_id]
                sub_execution = TaskExecution(
                    execution_id=execution.id,
                    task_id=sub_task.task_id,
                    task_name=self.task_registry[sub_task.task_id]["name"]
                )
                sub_execution.status = "running"
                sub_execution.start_time = datetime.now()
                try:
                    cmd = self._build_command(sub_task)
                    result = self._run_command(cmd)
                    sub_execution.status = "completed"
                    sub_execution.end_time = datetime.now()
                    sub_execution.duration = (sub_execution.end_time - sub_execution.start_time).total_seconds()
                    sub_execution.result = result
                except Exception as e:
                    sub_execution.status = "failed"
                    sub_execution.end_time = datetime.now()
                    sub_execution.duration = (sub_execution.end_time - sub_execution.start_time).total_seconds() if sub_execution.start_time else 0
                    sub_execution.error = str(e)
                    sub_execution.stderr = [str(e)]
                task_executions.append(sub_execution)
                execution.add_log(f"并行子任务 {sub_task.id} 执行完成")

            for sub_execution in task_executions:
                execution.add_task_execution(sub_execution)
        except Exception as e:
            execution.add_log(f"并行任务执行失败: {e}", level="error")
            raise

    def _execute_conditional_task(self, task: WorkflowTask, task_map: Dict[str, WorkflowTask], execution: WorkflowExecution) -> None:
        task_execution = TaskExecution(
            execution_id=execution.id,
            task_id=task.task_id,
            task_name=self.task_registry[task.task_id]["name"]
        )
        try:
            task_execution.status = "running"
            task_execution.start_time = datetime.now()
            condition_result = self._evaluate_condition(task.conditions, execution)

            if condition_result:
                branch_task = task_map[task.conditions["true_branch"]["task_id"]]
                cmd = self._build_command(branch_task)
                result = self._run_command(cmd)
                task_execution.result = {
                    "condition_met": True,
                    "executed_branch": "true_branch",
                    "result": result
                }
            else:
                branch_task = task_map[task.conditions["false_branch"]["task_id"]]
                cmd = self._build_command(branch_task)
                result = self._run_command(cmd)
                task_execution.result = {
                    "condition_met": False,
                    "executed_branch": "false_branch",
                    "result": result
                }

            task_execution.status = "completed"
            task_execution.end_time = datetime.now()
            task_execution.duration = (task_execution.end_time - task_execution.start_time).total_seconds()
            execution.add_task_execution(task_execution)
            execution.add_log(f"条件任务 {task.id} 执行完成，条件结果: {condition_result}")
        except Exception as e:
            task_execution.status = "failed"
            task_execution.end_time = datetime.now()
            task_execution.duration = (task_execution.end_time - task_execution.start_time).total_seconds() if task_execution.start_time else 0
            task_execution.error = str(e)
            task_execution.stderr = [str(e)]
            execution.add_task_execution(task_execution)
            execution.add_log(f"条件任务 {task.id} 执行失败: {e}", level="error")

    def _execute_loop_task(self, task: WorkflowTask, task_map: Dict[str, WorkflowTask], execution: WorkflowExecution) -> None:
        task_execution = TaskExecution(
            execution_id=execution.id,
            task_id=task.task_id,
            task_name=self.task_registry[task.task_id]["name"]
        )
        try:
            task_execution.status = "running"
            task_execution.start_time = datetime.now()
            loop_count = task.loop_config.get("count", 1)
            results = []

            for i in range(loop_count):
                execution.add_log(f"循环执行 {i + 1}/{loop_count}")
                loop_task = task_map[task.loop_config["task_id"]]
                cmd = self._build_command(loop_task)
                result = self._run_command(cmd)
                results.append(result)

            task_execution.result = {
                "loop_type": task.loop_config.get("type", "count"),
                "loop_count": loop_count,
                "results": results
            }
            task_execution.status = "completed"
            task_execution.end_time = datetime.now()
            task_execution.duration = (task_execution.end_time - task_execution.start_time).total_seconds()
            execution.add_task_execution(task_execution)
            execution.add_log(f"循环任务 {task.id} 执行完成，共执行 {loop_count} 次")
        except Exception as e:
            task_execution.status = "failed"
            task_execution.end_time = datetime.now()
            task_execution.duration = (task_execution.end_time - task_execution.start_time).total_seconds() if task_execution.start_time else 0
            task_execution.error = str(e)
            task_execution.stderr = [str(e)]
            execution.add_task_execution(task_execution)
            execution.add_log(f"循环任务 {task.id} 执行失败: {e}", level="error")

    def _build_command(self, task: WorkflowTask) -> List[str]:
        task_def = self.task_registry[task.task_id]
        cmd_builder = task_def.get("cmd_builder")
        if cmd_builder:
            params = task.config.copy()
            params["task_id"] = task.task_id
            return cmd_builder(params)
        return ["python", task_def["script"], *self._build_args(task.config)]

    def _build_args(self, config: Dict[str, Any]) -> List[str]:
        args: List[str] = []
        for key, value in config.items():
            if value is not None:
                args.extend([f"--{key}", str(value)])
        return args

    def _run_command(self, cmd: List[str]) -> Dict[str, Any]:
        try:
            process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                bufsize=1,
                universal_newlines=True,
                encoding='utf-8'
            )
            stdout, stderr = process.communicate()
            return {
                "return_code": process.returncode,
                "stdout": stdout,
                "stderr": stderr
            }
        except Exception as e:
            raise RuntimeError(f"命令执行失败: {e}")

    def _evaluate_condition(self, conditions: Dict[str, Any], execution: WorkflowExecution) -> bool:
        condition_type = conditions.get("type", "success")
        if condition_type == "success":
            previous_task_id = conditions.get("task_id")
            if not previous_task_id:
                return True
            for task_exec in execution.task_executions:
                if task_exec.task_id == previous_task_id:
                    return task_exec.status == "completed"
            return False
        return False

    def get_workflow_status(self, workflow_id: str) -> Dict[str, Any]:
        if workflow_id not in self.workflow_store:
            return {"status": "not_found"}
        workflow = self.workflow_store[workflow_id]
        return {
            "id": workflow.id,
            "name": workflow.name,
            "status": workflow.status.value,
            "created_at": workflow.created_at.isoformat(),
            "updated_at": workflow.updated_at.isoformat()
        }

    def get_execution_status(self, execution_id: str) -> Dict[str, Any]:
        if execution_id not in self.execution_store:
            return {"status": "not_found"}
        return self.execution_store[execution_id].to_dict()


if __name__ == '__main__':
    import argparse

    parser = argparse.ArgumentParser(description="工作流引擎 CLI")
    parser.add_argument("-w", "--workflow-id", required=True)
    parser.add_argument("-c", "--config", required=True)
    args = parser.parse_args()

    config = json.loads(args.config)
    workflow = Workflow.from_dict(config)
    engine = WorkflowEngine(TASKS)
    engine.workflow_store[workflow.id] = workflow
    execution = engine.execute_workflow(workflow.id)
    if execution:
        result = json.dumps(execution.to_dict(), ensure_ascii=False, indent=2)
        print(result)
    else:
        print("工作流执行失败", file=sys.stderr)
        sys.exit(1)
