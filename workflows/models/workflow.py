"""
工作流数据模型
定义工作流、任务和执行记录的数据结构
"""

import uuid
from datetime import datetime
from typing import List, Dict, Any, Optional
from dataclasses import dataclass, field
from enum import Enum


class WorkflowStatus(str, Enum):
    """工作流状态枚举"""
    DRAFT = "draft"          # 草稿
    ACTIVE = "active"        # 激活
    INACTIVE = "inactive"    # 非激活
    RUNNING = "running"      # 运行中
    COMPLETED = "completed"  # 已完成
    FAILED = "failed"        # 失败
    CANCELLED = "cancelled"  # 已取消


class TaskType(str, Enum):
    """任务类型枚举"""
    SEQUENTIAL = "sequential"    # 顺序执行
    PARALLEL = "parallel"        # 并行执行
    CONDITIONAL = "conditional"  # 条件分支
    LOOP = "loop"               # 循环执行


@dataclass
class WorkflowTask:
    """工作流任务定义"""
    type: TaskType
    task_id: str  # 注册的任务ID
    id: str = field(default_factory=lambda: str(uuid.uuid4()))
    config: Dict[str, Any] = field(default_factory=dict)
    position: Dict[str, int] = field(default_factory=dict)
    next_tasks: List[str] = field(default_factory=list)
    conditions: Dict[str, Any] = field(default_factory=dict)
    loop_config: Dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> Dict[str, Any]:
        """转换为字典格式"""
        return {
            "id": self.id,
            "type": self.type.value,
            "task_id": self.task_id,
            "config": self.config,
            "position": self.position,
            "next_tasks": self.next_tasks,
            "conditions": self.conditions,
            "loop_config": self.loop_config
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'WorkflowTask':
        """从字典创建实例"""
        return cls(
            id=data.get("id", str(uuid.uuid4())),
            type=TaskType(data["type"]),
            task_id=data["task_id"],
            config=data.get("config", {}),
            position=data.get("position", {}),
            next_tasks=data.get("next_tasks", []),
            conditions=data.get("conditions", {}),
            loop_config=data.get("loop_config", {})
        )


@dataclass
class Workflow:
    """工作流定义"""
    name: str
    description: str
    id: str = field(default_factory=lambda: str(uuid.uuid4()))
    version: str = "1.0"
    created_at: datetime = field(default_factory=datetime.now)
    updated_at: datetime = field(default_factory=datetime.now)
    config: Dict[str, Any] = field(default_factory=dict)
    tasks: List[WorkflowTask] = field(default_factory=list)
    status: WorkflowStatus = WorkflowStatus.DRAFT

    def add_task(self, task: WorkflowTask) -> None:
        """添加任务到工作流"""
        self.tasks.append(task)
        self.updated_at = datetime.now()

    def validate(self) -> bool:
        """验证工作流配置的合法性"""
        # 检查任务ID是否唯一
        task_ids = [task.id for task in self.tasks]
        if len(task_ids) != len(set(task_ids)):
            raise ValueError("任务ID必须唯一")

        # 检查任务类型和配置
        for task in self.tasks:
            if not task.task_id:
                raise ValueError(f"任务 {task.id} 缺少task_id")

            if task.type == TaskType.CONDITIONAL and not task.conditions:
                raise ValueError(f"条件任务 {task.id} 必须有条件配置")

            if task.type == TaskType.LOOP and not task.loop_config:
                raise ValueError(f"循环任务 {task.id} 必须有循环配置")

        return True

    def to_dict(self) -> Dict[str, Any]:
        """转换为字典格式"""
        return {
            "id": self.id,
            "name": self.name,
            "description": self.description,
            "version": self.version,
            "created_at": self.created_at.isoformat(),
            "updated_at": self.updated_at.isoformat(),
            "config": self.config,
            "tasks": [task.to_dict() for task in self.tasks],
            "status": self.status.value
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'Workflow':
        """从字典创建实例"""
        tasks = [WorkflowTask.from_dict(task_data) for task_data in data.get("tasks", [])]
        return cls(
            id=data.get("id", str(uuid.uuid4())),
            name=data["name"],
            description=data.get("description", ""),
            version=data.get("version", "1.0"),
            created_at=datetime.fromisoformat(data["created_at"]),
            updated_at=datetime.fromisoformat(data["updated_at"]),
            config=data.get("config", {}),
            tasks=tasks,
            status=WorkflowStatus(data["status"])
        )


@dataclass
class TaskExecution:
    """任务执行记录"""
    execution_id: str
    task_id: str
    task_name: str
    id: str = field(default_factory=lambda: str(uuid.uuid4()))
    status: str = "pending"
    start_time: Optional[datetime] = None
    end_time: Optional[datetime] = None
    duration: float = 0.0
    result: Optional[Dict[str, Any]] = None
    error: Optional[str] = None
    stdout: List[str] = field(default_factory=list)
    stderr: List[str] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        """转换为字典格式"""
        return {
            "id": self.id,
            "execution_id": self.execution_id,
            "task_id": self.task_id,
            "task_name": self.task_name,
            "status": self.status,
            "start_time": self.start_time.isoformat() if self.start_time else None,
            "end_time": self.end_time.isoformat() if self.end_time else None,
            "duration": self.duration,
            "result": self.result,
            "error": self.error,
            "stdout": self.stdout,
            "stderr": self.stderr
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'TaskExecution':
        """从字典创建实例"""
        start_time = datetime.fromisoformat(data["start_time"]) if data["start_time"] else None
        end_time = datetime.fromisoformat(data["end_time"]) if data["end_time"] else None
        return cls(
            id=data["id"],
            execution_id=data["execution_id"],
            task_id=data["task_id"],
            task_name=data["task_name"],
            status=data["status"],
            start_time=start_time,
            end_time=end_time,
            duration=data["duration"],
            result=data.get("result"),
            error=data.get("error"),
            stdout=data.get("stdout", []),
            stderr=data.get("stderr", [])
        )


@dataclass
class WorkflowExecution:
    """工作流执行记录"""
    workflow_id: str
    workflow_name: str
    id: str = field(default_factory=lambda: str(uuid.uuid4()))
    status: str = "pending"
    start_time: Optional[datetime] = None
    end_time: Optional[datetime] = None
    duration: float = 0.0
    result_file: Optional[str] = None
    report_file: Optional[str] = None
    logs: List[Dict[str, Any]] = field(default_factory=list)
    task_executions: List[TaskExecution] = field(default_factory=list)

    def add_task_execution(self, task_execution: TaskExecution) -> None:
        """添加任务执行记录"""
        self.task_executions.append(task_execution)

    def add_log(self, message: str, level: str = "info") -> None:
        """添加日志记录"""
        self.logs.append({
            "timestamp": datetime.now().isoformat(),
            "level": level,
            "message": message
        })

    def to_dict(self) -> Dict[str, Any]:
        """转换为字典格式"""
        return {
            "id": self.id,
            "workflow_id": self.workflow_id,
            "workflow_name": self.workflow_name,
            "status": self.status,
            "start_time": self.start_time.isoformat() if self.start_time else None,
            "end_time": self.end_time.isoformat() if self.end_time else None,
            "duration": self.duration,
            "result_file": self.result_file,
            "report_file": self.report_file,
            "logs": self.logs,
            "task_executions": [task.to_dict() for task in self.task_executions]
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'WorkflowExecution':
        """从字典创建实例"""
        start_time = datetime.fromisoformat(data["start_time"]) if data["start_time"] else None
        end_time = datetime.fromisoformat(data["end_time"]) if data["end_time"] else None
        task_executions = [TaskExecution.from_dict(task_data) for task_data in data.get("task_executions", [])]
        return cls(
            id=data["id"],
            workflow_id=data["workflow_id"],
            workflow_name=data["workflow_name"],
            status=data["status"],
            start_time=start_time,
            end_time=end_time,
            duration=data["duration"],
            result_file=data.get("result_file"),
            report_file=data.get("report_file"),
            logs=data.get("logs", []),
            task_executions=task_executions
        )