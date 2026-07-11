"""
Task Queue — SQLite-backed job queue with background worker.

Supports:
  - Enqueue single or batch tasks
  - FIFO sequential processing in background thread
  - Status tracking (pending / running / completed / failed)
  - Priority ordering
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
import threading
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from uuid import uuid4

from scripts.task_registry import TASKS
from scripts.job_logger import create_log, finish_log
from scripts.validator import validate_params

QUEUE_DB = Path(__file__).parent.parent / "data" / "task_queue.db"
_lock = threading.Lock()
_worker_thread: threading.Thread | None = None
_worker_stop = threading.Event()

STATUS_PENDING = "pending"
STATUS_RUNNING = "running"
STATUS_COMPLETED = "completed"
STATUS_FAILED = "failed"


def _get_db():
    if not hasattr(_get_db, "_conn") or _get_db._conn is None:
        QUEUE_DB.parent.mkdir(parents=True, exist_ok=True)
        import sqlite3
        conn = sqlite3.connect(str(QUEUE_DB), check_same_thread=False)
        conn.row_factory = sqlite3.Row
        conn.execute("""
            CREATE TABLE IF NOT EXISTS task_queue (
                id TEXT PRIMARY KEY,
                task_id TEXT NOT NULL,
                task_name TEXT,
                params TEXT NOT NULL,
                status TEXT NOT NULL DEFAULT 'pending',
                priority INTEGER DEFAULT 0,
                log_id TEXT,
                error TEXT,
                created_at TEXT NOT NULL,
                started_at TEXT,
                finished_at TEXT
            )
        """)
        conn.execute("CREATE INDEX IF NOT EXISTS idx_queue_status ON task_queue(status)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_queue_priority ON task_queue(priority DESC)")
        conn.commit()
        _get_db._conn = conn
    return _get_db._conn


_get_db._conn = None


def _now() -> str:
    return datetime.now(timezone.utc).isoformat()


def enqueue(task_id: str, params: dict, priority: int = 0) -> dict:
    if task_id not in TASKS:
        return {"status": "error", "message": f"未知任务: {task_id}"}

    is_valid, err_msg = validate_params(task_id, params, TASKS, skip_exists=True)
    if not is_valid:
        return {"status": "error", "message": f"参数校验失败: {err_msg}"}

    queue_id = str(uuid4())
    task_name = TASKS[task_id]["name"]
    conn = _get_db()
    with _lock:
        conn.execute(
            """INSERT INTO task_queue
               (id, task_id, task_name, params, status, priority, created_at)
               VALUES (?, ?, ?, ?, ?, ?, ?)""",
            (queue_id, task_id, task_name,
             json.dumps(params, ensure_ascii=False),
             STATUS_PENDING, priority, _now())
        )
        conn.commit()
    return {"status": "ok", "queue_id": queue_id}


def batch_enqueue(tasks: list[dict]) -> dict:
    results = [enqueue(t.get("task_id"), t.get("params", {}), t.get("priority", 0)) for t in tasks]
    success_count = sum(1 for r in results if r.get("status") == "ok")
    return {"status": "ok", "total": len(tasks), "enqueued": success_count, "results": results}


def dequeue() -> dict | None:
    conn = _get_db()
    with _lock:
        row = conn.execute(
            """SELECT * FROM task_queue WHERE status = ?
               ORDER BY priority DESC, created_at ASC LIMIT 1""",
            (STATUS_PENDING,)
        ).fetchone()
        if not row:
            return None
        conn.execute(
            "UPDATE task_queue SET status=?, started_at=? WHERE id=?",
            (STATUS_RUNNING, _now(), row["id"])
        )
        conn.commit()
    d = dict(row)
    d["params"] = json.loads(d["params"])
    d["status"] = STATUS_RUNNING
    return d


def get_all(status: str | None = None) -> list[dict]:
    conn = _get_db()
    query = "SELECT * FROM task_queue WHERE 1=1"
    params: list = []
    if status:
        query += " AND status=?"
        params.append(status)
    query += " ORDER BY priority DESC, created_at ASC"
    rows = conn.execute(query, params).fetchall()
    result = []
    for row in rows:
        d = dict(row)
        d["params"] = json.loads(d["params"])
        result.append(d)
    return result


def get_stats() -> dict:
    conn = _get_db()
    row = conn.execute("""
        SELECT
            COUNT(*) as total,
            SUM(CASE WHEN status='pending' THEN 1 ELSE 0 END) as pending,
            SUM(CASE WHEN status='running' THEN 1 ELSE 0 END) as running,
            SUM(CASE WHEN status='completed' THEN 1 ELSE 0 END) as completed,
            SUM(CASE WHEN status='failed' THEN 1 ELSE 0 END) as failed
        FROM task_queue
    """).fetchone()
    return dict(row) if row else {}


def remove(queue_id: str) -> dict:
    conn = _get_db()
    with _lock:
        row = conn.execute("SELECT status FROM task_queue WHERE id=?", (queue_id,)).fetchone()
        if not row:
            return {"status": "error", "message": "队列项不存在"}
        if row["status"] == STATUS_RUNNING:
            return {"status": "error", "message": "无法删除正在执行的任务"}
        conn.execute("DELETE FROM task_queue WHERE id=?", (queue_id,))
        conn.commit()
    return {"status": "ok"}


def retry_task(queue_id: str) -> dict:
    """将失败的任务重置为 pending，使 Worker 可重新捡起执行。"""
    conn = _get_db()
    with _lock:
        row = conn.execute("SELECT status FROM task_queue WHERE id=?", (queue_id,)).fetchone()
        if not row:
            return {"status": "error", "message": "队列项不存在"}
        if row["status"] != STATUS_FAILED:
            return {"status": "error", "message": f"只有失败的任务可以重试（当前状态: {row['status']}）"}
        conn.execute(
            "UPDATE task_queue SET status=?, error=NULL, started_at=NULL, finished_at=NULL WHERE id=?",
            (STATUS_PENDING, queue_id)
        )
        conn.commit()
    return {"status": "ok"}


def clear_completed() -> int:
    conn = _get_db()
    with _lock:
        cursor = conn.execute(
            "DELETE FROM task_queue WHERE status IN (?, ?)",
            (STATUS_COMPLETED, STATUS_FAILED)
        )
        conn.commit()
        return cursor.rowcount


def clear_all() -> int:
    conn = _get_db()
    with _lock:
        cursor = conn.execute(
            "DELETE FROM task_queue WHERE status != ?",
            (STATUS_RUNNING,)
        )
        conn.commit()
        return cursor.rowcount


def _execute_task(item: dict) -> None:
    task_id = item["task_id"]
    info = TASKS[task_id]
    params = item["params"]
    cmd = info.get("cmd_builder", lambda p: [])(params)
    # 将命令头的裸 "python" 替换为当前解释器路径，保证子进程使用 venv
    if cmd and cmd[0] in ("python", "python3"):
        cmd = [sys.executable] + cmd[1:]
    log_id = create_log(task_id, info["name"], params=params, cmd=cmd)
    start = time.time()

    conn = _get_db()
    
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"

    # 强制将工作目录设为项目根目录，确保子进程能找到 utils / scripts 等包
    project_root = str(Path(__file__).parent.parent.resolve())

    # 注入 PYTHONPATH，让子进程 Python 解释器也能找到项目根下的所有包
    existing_pythonpath = env.get("PYTHONPATH", "")
    env["PYTHONPATH"] = project_root + (os.pathsep + existing_pythonpath if existing_pythonpath else "")

    try:
        proc = subprocess.run(
            cmd, capture_output=True, text=True, timeout=3600,
            encoding="utf-8", errors="replace", env=env,
            cwd=project_root
        )
        duration = time.time() - start
        status = STATUS_COMPLETED if proc.returncode == 0 else STATUS_FAILED
        finish_log(log_id, status, stdout=proc.stdout, stderr=proc.stderr, duration=duration)

        with _lock:
            conn.execute(
                "UPDATE task_queue SET status=?, log_id=?, error=?, finished_at=? WHERE id=?",
                (status, log_id,
                 proc.stderr[:500] if proc.returncode != 0 else None,
                 _now(), item["id"])
            )
            conn.commit()
    except subprocess.TimeoutExpired:
        finish_log(log_id, STATUS_FAILED, stderr="任务执行超时 (1h)")
        with _lock:
            conn.execute(
                "UPDATE task_queue SET status=?, error=?, finished_at=? WHERE id=?",
                (STATUS_FAILED, "执行超时", _now(), item["id"])
            )
            conn.commit()
    except Exception as e:
        finish_log(log_id, STATUS_FAILED, stderr=str(e))
        with _lock:
            conn.execute(
                "UPDATE task_queue SET status=?, error=?, finished_at=? WHERE id=?",
                (STATUS_FAILED, str(e)[:500], _now(), item["id"])
            )
            conn.commit()


def _worker_loop() -> None:
    while not _worker_stop.is_set():
        try:
            item = dequeue()
            if item:
                _execute_task(item)
            else:
                time.sleep(2)
        except Exception as e:
            # 捕获所有意外异常，防止线程崩溃导致 Worker 无法重启
            import traceback
            print(f"[Worker] ⚠️ 意外异常，线程继续运行: {e}")
            traceback.print_exc()
            time.sleep(2)


def start_processing() -> dict:
    global _worker_thread, _worker_stop
    if _worker_thread and _worker_thread.is_alive():
        return {"status": "ok", "message": "Worker is already running"}
    # 无论线程是否存活，先重置停止标志，确保新线程能正常进入循环
    _worker_stop.clear()
    _worker_thread = threading.Thread(target=_worker_loop, daemon=True)
    _worker_thread.start()
    return {"status": "ok", "message": "Worker started"}


def stop_processing() -> dict:
    global _worker_stop
    _worker_stop.set()
    return {"status": "ok", "message": "Stop signal sent"}


def is_processing() -> bool:
    return _worker_thread is not None and _worker_thread.is_alive()
