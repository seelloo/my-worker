"""
Scheduler — APScheduler-based job scheduling with SQLite persistence.

Supports:
  - Cron-style schedules (e.g. "0 9 * * *" = 每天 9:00)
  - Interval-style schedules (e.g. every 30 minutes)
  - One-shot delayed execution
"""
from __future__ import annotations

import json
import subprocess
import sys
import threading
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from uuid import uuid4

try:
    from apscheduler.schedulers.background import BackgroundScheduler
    from apscheduler.triggers.cron import CronTrigger
    from apscheduler.triggers.interval import IntervalTrigger
    from apscheduler.triggers.date import DateTrigger
except ImportError:
    print("[WARN] APScheduler not installed. Run: pip install APScheduler")
    sys.exit(1)

from scripts.task_registry import TASKS
from scripts.job_logger import create_log, finish_log
from scripts.validator import validate_params


# ── Persistent Schedule Store ─────────────────────────────────────────────────

SCHEDULE_DB = Path(__file__).parent.parent / "data" / "schedules.db"
_lock = threading.Lock()


def _get_db():
    if not hasattr(_get_db, "_conn") or _get_db._conn is None:
        SCHEDULE_DB.parent.mkdir(parents=True, exist_ok=True)
        import sqlite3
        conn = sqlite3.connect(str(SCHEDULE_DB), check_same_thread=False)
        conn.row_factory = sqlite3.Row
        conn.execute("""
            CREATE TABLE IF NOT EXISTS schedules (
                id TEXT PRIMARY KEY,
                task_id TEXT NOT NULL,
                task_name TEXT,
                trigger_type TEXT NOT NULL,  -- 'cron' | 'interval' | 'date'
                trigger_config TEXT NOT NULL,
                params TEXT NOT NULL,
                enabled INTEGER DEFAULT 1,
                last_run TEXT,
                next_run TEXT,
                created_at TEXT NOT NULL
            )
        """)
        conn.commit()
        _get_db._conn = conn
    return _get_db._conn


_get_db._conn = None  # type: ignore[attr-defined]


def _now() -> str:
    return datetime.now(timezone.utc).isoformat()


# ── Scheduler Wrapper ─────────────────────────────────────────────────────────

_scheduler: BackgroundScheduler | None = None
_started = False


def get_scheduler() -> BackgroundScheduler:
    global _scheduler, _started
    if not _started:
        _scheduler = BackgroundScheduler()
        _scheduler.start()
        _started = True
        _restore_schedules()
    return _scheduler


def _restore_schedules() -> None:
    """Restore persisted schedules on startup."""
    conn = _get_db()
    rows = conn.execute("SELECT * FROM schedules WHERE enabled=1").fetchall()
    for row in rows:
        try:
            cfg = json.loads(row["trigger_config"])
            params = json.loads(row["params"])
            _add_job(
                schedule_id=row["id"],
                task_id=row["task_id"],
                params=params,
                trigger_type=row["trigger_type"],
                **cfg
            )
        except Exception as e:
            print(f"⚠️ Failed to restore schedule {row['id']}: {e}")


def _run_task(task_id: str, params: dict, schedule_id: str | None = None) -> dict:
    """Execute a task via subprocess and log the result."""
    if task_id not in TASKS:
        return {"status": "error", "message": f"未知任务: {task_id}"}

    info = TASKS[task_id]
    is_valid, err_msg = validate_params(task_id, params, TASKS)
    if not is_valid:
        return {"status": "error", "message": f"参数校验失败: {err_msg}"}

    cmd = info.get("cmd_builder", lambda p: [])(params)
    if not cmd:
        return {"status": "error", "message": "任务未配置 cmd_builder"}

    log_id = create_log(task_id, info["name"], params=params, cmd=cmd)
    start = time.time()

    try:
        proc = subprocess.run(
            cmd, capture_output=True, text=True, timeout=3600,
            encoding="utf-8", errors="replace"
        )
        duration = time.time() - start
        status = "completed" if proc.returncode == 0 else "failed"
        finish_log(
            log_id, status,
            stdout=proc.stdout, stderr=proc.stderr,
            duration=duration
        )

        # Update last_run in schedule store
        if schedule_id:
            conn = _get_db()
            with _lock:
                conn.execute(
                    "UPDATE schedules SET last_run=? WHERE id=?",
                    (_now(), schedule_id)
                )
                conn.commit()

        return {
            "status": status,
            "log_id": log_id,
            "return_code": proc.returncode,
            "duration": round(duration, 2),
        }
    except subprocess.TimeoutExpired:
        finish_log(log_id, "failed", stderr="任务执行超时 (1h)")
        return {"status": "timeout", "log_id": log_id}
    except Exception as e:
        finish_log(log_id, "failed", stderr=str(e))
        return {"status": "error", "message": str(e)}


def _add_job(schedule_id: str, task_id: str, params: dict,
             trigger_type: str, **trigger_config) -> None:
    """Internal: register a job with APScheduler."""
    scheduler = get_scheduler()
    task_name = TASKS[task_id]["name"] if task_id in TASKS else task_id

    def wrapper():
        _run_task(task_id, params, schedule_id=schedule_id)

    if trigger_type == "cron":
        trigger = CronTrigger(**trigger_config)
    elif trigger_type == "interval":
        trigger = IntervalTrigger(**trigger_config)
    elif trigger_type == "date":
        trigger = DateTrigger(run_date=trigger_config.get("run_date"))
    else:
        raise ValueError(f"不支持的触发器类型: {trigger_type}")

    scheduler.add_job(wrapper, trigger, id=schedule_id, replace_existing=True)


# ── Public API ────────────────────────────────────────────────────────────────


def create_schedule(task_id: str, params: dict, trigger_type: str,
                    **trigger_config) -> dict:
    """Create and persist a new schedule."""
    if task_id not in TASKS:
        return {"status": "error", "message": f"未知任务: {task_id}"}

    schedule_id = str(uuid4())
    task_name = TASKS[task_id]["name"]

    # Persist to DB
    conn = _get_db()
    with _lock:
        conn.execute(
            """INSERT INTO schedules
               (id, task_id, task_name, trigger_type, trigger_config, params,
                enabled, created_at)
               VALUES (?, ?, ?, ?, ?, ?, 1, ?)""",
            (schedule_id, task_id, task_name, trigger_type,
             json.dumps(trigger_config, ensure_ascii=False),
             json.dumps(params, ensure_ascii=False),
             _now())
        )
        conn.commit()

    # Register with scheduler
    try:
        _add_job(schedule_id, task_id, params, trigger_type, **trigger_config)
        return {"status": "ok", "schedule_id": schedule_id}
    except Exception as e:
        return {"status": "error", "message": str(e)}


def list_schedules() -> list[dict]:
    """List all persisted schedules."""
    conn = _get_db()
    rows = conn.execute("SELECT * FROM schedules ORDER BY created_at DESC").fetchall()
    result = []
    for row in rows:
        d = dict(row)
        d["trigger_config"] = json.loads(d["trigger_config"])
        d["params"] = json.loads(d["params"])
        result.append(d)
    return result


def toggle_schedule(schedule_id: str, enabled: bool) -> dict:
    """Enable or disable a schedule."""
    conn = _get_db()
    with _lock:
        conn.execute(
            "UPDATE schedules SET enabled=? WHERE id=?",
            (1 if enabled else 0, schedule_id)
        )
        conn.commit()

    scheduler = get_scheduler()
    if enabled:
        # Re-add to scheduler
        row = conn.execute(
            "SELECT * FROM schedules WHERE id=?", (schedule_id,)
        ).fetchone()
        if row:
            cfg = json.loads(row["trigger_config"])
            params = json.loads(row["params"])
            try:
                _add_job(schedule_id, row["task_id"], params,
                         row["trigger_type"], **cfg)
                return {"status": "ok", "action": "enabled"}
            except Exception as e:
                return {"status": "error", "message": str(e)}
        return {"status": "error", "message": "Schedule not found"}
    else:
        try:
            scheduler.remove_job(schedule_id)
            return {"status": "ok", "action": "disabled"}
        except Exception:
            return {"status": "error", "message": "Failed to remove job"}


def delete_schedule(schedule_id: str) -> dict:
    """Delete a schedule."""
    conn = _get_db()
    with _lock:
        conn.execute("DELETE FROM schedules WHERE id=?", (schedule_id,))
        conn.commit()

    scheduler = get_scheduler()
    try:
        scheduler.remove_job(schedule_id)
    except Exception:
        pass
    return {"status": "ok", "action": "deleted"}


def run_task_now(task_id: str, params: dict) -> dict:
    """Execute a task immediately (one-shot, not persisted as schedule)."""
    return _run_task(task_id, params)
