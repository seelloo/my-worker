"""
Execution Logger — structured job execution logging with SQLite persistence.

Records: start/end time, status, params, stdout/stderr, output files.
"""
from __future__ import annotations

import json
import sqlite3
import threading
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from uuid import uuid4

DB_PATH = Path(__file__).parent.parent / "data" / "job_logs.db"
_lock = threading.Lock()


def _get_db() -> sqlite3.Connection:
    """Thread-safe singleton SQLite connection."""
    if not hasattr(_get_db, "_conn") or _get_db._conn is None:
        DB_PATH.parent.mkdir(parents=True, exist_ok=True)
        conn = sqlite3.connect(str(DB_PATH), check_same_thread=False)
        conn.row_factory = sqlite3.Row
        _get_db._conn = conn
        _init_schema(conn)
    return _get_db._conn


_get_db._conn = None  # type: ignore[attr-defined]


def _init_schema(conn: sqlite3.Connection) -> None:
    conn.execute("""
        CREATE TABLE IF NOT EXISTS job_logs (
            id TEXT PRIMARY KEY,
            task_id TEXT NOT NULL,
            task_name TEXT,
            status TEXT NOT NULL DEFAULT 'running',
            params TEXT,
            cmd TEXT,
            stdout TEXT,
            stderr TEXT,
            output_files TEXT,
            duration REAL,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
    """)
    conn.execute("""
        CREATE INDEX IF NOT EXISTS idx_job_logs_task_id ON job_logs(task_id)
    """)
    conn.execute("""
        CREATE INDEX IF NOT EXISTS idx_job_logs_status ON job_logs(status)
    """)
    conn.execute("""
        CREATE INDEX IF NOT EXISTS idx_job_logs_created ON job_logs(created_at)
    """)
    conn.commit()


# ── Public API ────────────────────────────────────────────────────────────────


def create_log(task_id: str, task_name: str, params: dict | None = None,
               cmd: list[str] | None = None) -> str:
    """Create a new job log entry. Returns the log ID."""
    now = _now()
    log_id = str(uuid4())
    conn = _get_db()
    with _lock:
        conn.execute(
            """INSERT INTO job_logs
               (id, task_id, task_name, status, params, cmd, created_at, updated_at)
               VALUES (?, ?, ?, 'running', ?, ?, ?, ?)""",
            (log_id, task_id, task_name,
             json.dumps(params, ensure_ascii=False) if params else None,
             json.dumps(cmd, ensure_ascii=False) if cmd else None,
             now, now)
        )
        conn.commit()
    return log_id


def update_log(log_id: str, status: str, stdout: str | None = None,
               stderr: str | None = None, output_files: list[str] | None = None,
               duration: float | None = None) -> None:
    """Update an existing job log entry."""
    now = _now()
    conn = _get_db()
    with _lock:
        conn.execute(
            """UPDATE job_logs
               SET status=?, stdout=?, stderr=?, output_files=?,
                   duration=?, updated_at=?
               WHERE id=?""",
            (status, stdout, stderr,
             json.dumps(output_files) if output_files else None,
             duration, now, log_id)
        )
        conn.commit()


def finish_log(log_id: str, status: str, stdout: str = "",
               stderr: str = "", output_files: list[str] | None = None,
               duration: float = 0.0) -> None:
    """Mark a job as completed or failed."""
    update_log(log_id, status, stdout=stdout, stderr=stderr,
               output_files=output_files, duration=duration)


def get_log(log_id: str) -> dict | None:
    """Retrieve a single job log by ID."""
    conn = _get_db()
    row = conn.execute("SELECT * FROM job_logs WHERE id=?", (log_id,)).fetchone()
    return dict(row) if row else None


def list_logs(task_id: str | None = None, status: str | None = None,
              limit: int = 50, offset: int = 0) -> list[dict]:
    """List job logs with optional filters."""
    conn = _get_db()
    query = "SELECT * FROM job_logs WHERE 1=1"
    params: list[Any] = []

    if task_id:
        query += " AND task_id=?"
        params.append(task_id)
    if status:
        query += " AND status=?"
        params.append(status)

    query += " ORDER BY created_at DESC LIMIT ? OFFSET ?"
    params.extend([limit, offset])

    rows = conn.execute(query, params).fetchall()
    return [dict(r) for r in rows]


def get_stats() -> dict:
    """Get execution statistics."""
    conn = _get_db()
    row = conn.execute("""
        SELECT
            COUNT(*) as total,
            SUM(CASE WHEN status='completed' THEN 1 ELSE 0 END) as completed,
            SUM(CASE WHEN status='failed' THEN 1 ELSE 0 END) as failed,
            SUM(CASE WHEN status='running' THEN 1 ELSE 0 END) as running,
            AVG(CASE WHEN duration IS NOT NULL THEN duration END) as avg_duration
        FROM job_logs
    """).fetchone()
    return dict(row) if row else {}


def delete_old_logs(days: int = 30) -> int:
    """Delete logs older than N days. Returns deleted count."""
    cutoff = datetime.now(timezone.utc).timestamp() - days * 86400
    cutoff_str = datetime.fromtimestamp(cutoff, tz=timezone.utc).isoformat()
    conn = _get_db()
    with _lock:
        cursor = conn.execute("DELETE FROM job_logs WHERE created_at < ?", (cutoff_str,))
        conn.commit()
        return cursor.rowcount


# ── Helpers ───────────────────────────────────────────────────────────────────


def _now() -> str:
    return datetime.now(timezone.utc).isoformat()


def close_db() -> None:
    """Close the database connection (for cleanup)."""
    conn = getattr(_get_db, "_conn", None)
    if conn:
        conn.close()
        _get_db._conn = None  # type: ignore[attr-defined]
