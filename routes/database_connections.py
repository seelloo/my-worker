"""
数据库连接配置管理 API
提供连接的增删改查和测试功能
"""
import uuid
from datetime import datetime
from flask import Blueprint, jsonify, request

from workflows.models.database import ConnectionStore, ConnectionConfig
from utils.rate_limiter import rate_limit
from config import RATE_LIMIT_ENABLED, RATE_LIMIT_DEFAULT, RATE_LIMIT_WINDOW

db_conn_bp = Blueprint("database_connections", __name__, url_prefix="/api")


def _rate_limit():
    if RATE_LIMIT_ENABLED:
        return rate_limit(max_requests=RATE_LIMIT_DEFAULT, window_seconds=RATE_LIMIT_WINDOW)
    return lambda f: f


@db_conn_bp.route("/connections", methods=["GET"])
def api_list_connections():
    """列出所有已保存的数据库连接（密码脱敏）"""
    try:
        conns = ConnectionStore.list_connections(redact_password=True)
        items = [c.to_dict(redact_password=True) for c in conns]
        return jsonify({"status": "success", "connections": items})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


@db_conn_bp.route("/connections", methods=["POST"])
@_rate_limit()
def api_save_connection():
    """新建或更新一个数据库连接"""
    data = request.json or {}

    name = (data.get("name") or "").strip()
    host = (data.get("host") or "").strip()
    user = (data.get("user") or "").strip()
    password = (data.get("password") or "").strip()
    database_name = (data.get("database_name") or data.get("database") or "").strip()

    if not all([name, host, user, database_name]):
        missing = [f for f, v in [("name", name), ("host", host), ("user", user), ("database_name", database_name)] if not v]
        return jsonify({"status": "error", "message": f"缺少必需字段: {', '.join(missing)}"}), 400

    # 编辑模式：检查同名冲突（排除自身）
    existing_id = data.get("id", "")
    existing = ConnectionStore.get_connection_by_name(name)
    if existing and existing.id != existing_id:
        return jsonify({"status": "error", "message": f"连接名「{name}」已存在"}), 409

    now = datetime.now()

    # 编辑时如果密码为空，保留原密码
    if existing_id:
        existing_conn = ConnectionStore.get_connection(existing_id)
        if not password and existing_conn:
            password = existing_conn.password

    conn = ConnectionConfig(
        id=existing_id or str(uuid.uuid4()),
        name=name,
        db_type=data.get("db_type", "mysql"),
        host=host,
        port=data.get("port", 3306),
        user=user,
        password=password,
        database_name=database_name,
        charset=data.get("charset", "utf8mb4"),
        timeout=data.get("timeout", 30),
        created_at=existing_conn.created_at if existing_id and existing_conn else now,
        updated_at=now,
    )

    ConnectionStore.save_connection(conn)
    return jsonify({"status": "success", "message": "已保存", "id": conn.id}), 201


@db_conn_bp.route("/connections/<conn_id>", methods=["DELETE"])
@_rate_limit()
def api_delete_connection(conn_id):
    """删除一个数据库连接"""
    store = ConnectionStore.get_connection(conn_id)
    if not store:
        return jsonify({"status": "error", "message": "连接不存在"}), 404

    ConnectionStore.delete_connection(conn_id)
    return jsonify({"status": "success", "message": "已删除"})


@db_conn_bp.route("/connections/<conn_id>/test", methods=["POST"])
def api_test_connection(conn_id):
    """测试连接是否可用"""
    config = ConnectionStore.get_connection(conn_id)
    if not config:
        return jsonify({"status": "error", "message": "连接不存在"}), 404

    if config.db_type != "mysql":
        return jsonify({"status": "error", "message": f"暂未支持 db_type={config.db_type} 的测试"}), 501

    try:
        import pymysql
        conn = pymysql.connect(
            host=config.host,
            port=config.port,
            user=config.user,
            password=config.password,
            database=config.database_name,
            charset=config.charset,
            connect_timeout=config.timeout,
        )
        with conn.cursor() as cursor:
            cursor.execute("SELECT 1 AS ok")
            cursor.fetchone()
        conn.close()
        return jsonify({"status": "success", "message": f"连接成功 — {config.host}:{config.port}/{config.database_name}"})
    except Exception as e:
        return jsonify({"status": "error", "message": f"连接失败: {e}"}), 500
