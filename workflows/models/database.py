"""
工作流数据库模型
提供工作流和执行记录的持久化存储
"""

import json
import sqlite3
from pathlib import Path
from typing import Dict, Any, List, Optional
from datetime import datetime

from workflows.models.workflow import Workflow, WorkflowExecution, TaskExecution


DB_PATH = Path("data/workflows.db")


def get_connection():
    """获取数据库连接"""
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    return conn


def init_database():
    """初始化数据库表"""
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS workflows (
            id TEXT PRIMARY KEY,
            name TEXT NOT NULL,
            description TEXT,
            version TEXT DEFAULT '1.0',
            config TEXT,
            status TEXT DEFAULT 'draft',
            created_at TEXT,
            updated_at TEXT
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS workflow_executions (
            id TEXT PRIMARY KEY,
            workflow_id TEXT NOT NULL,
            workflow_name TEXT,
            status TEXT DEFAULT 'pending',
            start_time TEXT,
            end_time TEXT,
            duration REAL DEFAULT 0,
            result_file TEXT,
            report_file TEXT,
            logs TEXT,
            task_executions TEXT,
            error TEXT,
            FOREIGN KEY (workflow_id) REFERENCES workflows(id)
        )
    """)

    conn.commit()
    conn.close()


class WorkflowStore:
    """工作流持久化存储"""

    @staticmethod
    def save_workflow(workflow: Workflow) -> None:
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute(
            "INSERT OR REPLACE INTO workflows VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
            (workflow.id, workflow.name, workflow.description, workflow.version,
             json.dumps(workflow.config, ensure_ascii=False), workflow.status.value,
             workflow.created_at.isoformat(), workflow.updated_at.isoformat())
        )
        conn.commit()
        conn.close()

    @staticmethod
    def get_workflow(workflow_id: str) -> Optional[Workflow]:
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM workflows WHERE id = ?", (workflow_id,))
        row = cursor.fetchone()
        conn.close()
        if not row:
            return None
        return Workflow(id=row["id"], name=row["name"], description=row["description"],
                        version=row["version"], config=json.loads(row["config"]) if row["config"] else {},
                        status=row["status"],
                        created_at=datetime.fromisoformat(row["created_at"]),
                        updated_at=datetime.fromisoformat(row["updated_at"]))

    @staticmethod
    def list_workflows() -> List[Workflow]:
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM workflows ORDER BY updated_at DESC")
        rows = cursor.fetchall()
        conn.close()
        return [Workflow(id=r["id"], name=r["name"], description=r["description"],
                         version=r["version"], config=json.loads(r["config"]) if r["config"] else {},
                         status=r["status"], created_at=datetime.fromisoformat(r["created_at"]),
                         updated_at=datetime.fromisoformat(r["updated_at"])) for r in rows]

    @staticmethod
    def delete_workflow(workflow_id: str) -> None:
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute("DELETE FROM workflows WHERE id = ?", (workflow_id,))
        conn.commit()
        conn.close()


class ExecutionStore:
    """执行记录持久化存储"""

    @staticmethod
    def save_execution(execution: WorkflowExecution) -> None:
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute(
            "INSERT OR REPLACE INTO workflow_executions VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
            (execution.id, execution.workflow_id, execution.workflow_name, execution.status,
             execution.start_time.isoformat() if execution.start_time else None,
             execution.end_time.isoformat() if execution.end_time else None,
             execution.duration, execution.result_file, execution.report_file,
             json.dumps(execution.logs, ensure_ascii=False),
             json.dumps([te.to_dict() for te in execution.task_executions], ensure_ascii=False),
             execution.error)
        )
        conn.commit()
        conn.close()

    @staticmethod
    def get_execution(execution_id: str) -> Optional[WorkflowExecution]:
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM workflow_executions WHERE id = ?", (execution_id,))
        row = cursor.fetchone()
        conn.close()
        if not row:
            return None
        execution = WorkflowExecution(id=row["id"], workflow_id=row["workflow_id"],
                                       workflow_name=row["workflow_name"], status=row["status"],
                                       start_time=datetime.fromisoformat(row["start_time"]) if row["start_time"] else None,
                                       end_time=datetime.fromisoformat(row["end_time"]) if row["end_time"] else None,
                                       duration=row["duration"], result_file=row["result_file"],
                                       report_file=row["report_file"])
        if row["logs"]:
            execution.logs = json.loads(row["logs"])
        if row["task_executions"]:
            for td in json.loads(row["task_executions"]):
                execution.add_task_execution(TaskExecution.from_dict(td))
        execution.error = row["error"]
        return execution

    @staticmethod
    def list_executions() -> List[WorkflowExecution]:
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM workflow_executions ORDER BY start_time DESC")
        rows = cursor.fetchall()
        conn.close()
        executions = []
        for row in rows:
            execution = WorkflowExecution(id=row["id"], workflow_id=row["workflow_id"],
                                           workflow_name=row["workflow_name"], status=row["status"],
                                           start_time=datetime.fromisoformat(row["start_time"]) if row["start_time"] else None,
                                           end_time=datetime.fromisoformat(row["end_time"]) if row["end_time"] else None,
                                           duration=row["duration"])
            if row["task_executions"]:
                for td in json.loads(row["task_executions"]):
                    execution.add_task_execution(TaskExecution.from_dict(td))
            executions.append(execution)
        return executions