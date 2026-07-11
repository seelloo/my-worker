"""
Unified CLI for the Data Worker system.

Usage:
    python cli.py task list
    python cli.py task run <task_id> --param key=value
    python cli.py workflow run <workflow_id> --config '{"..."}'
    python cli.py schedule create <task_id> --cron "0 9 * * *"
    python cli.py schedule list
    python cli.py schedule toggle <schedule_id> --enable
    python cli.py schedule delete <schedule_id>
    python cli.py log list [--task <task_id>] [--status completed]
    python cli.py log show <log_id>
    python cli.py log stats
    python cli.py agent "帮我合并文件夹里的所有Excel"
"""
from __future__ import annotations

import json
import os
import re
import sys
import time
from pathlib import Path

import click

# Force UTF-8 output on Windows
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

# Ensure project root is on path
ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from scripts.task_registry import TASKS, get_task_metadata
from scripts.scheduler import create_schedule, list_schedules, toggle_schedule, delete_schedule, run_task_now
from scripts.job_logger import get_log, list_logs, get_stats, delete_old_logs
from scripts.agent import match_task, parse_params_from_text


@click.group()
def cli():
    """数据员工 CLI — 统一管理所有数据处理任务"""
    pass


# ── Task Commands ─────────────────────────────────────────────────────────────

@cli.group()
def task():
    """任务管理"""
    pass


@task.command("list")
@click.option("--category", "-c", help="按分类过滤")
def task_list(category):
    """列出所有可用任务"""
    tasks = get_task_metadata()
    if category:
        tasks = [t for t in tasks if t.get("category") == category]

    click.echo("")
    for t in tasks:
        name = re.sub(r'[^\w\s一-鿿\-\(\)\.\,]', '', t['name'])
        click.echo(f"  {t['id']:<30} {name}")
        click.echo(f"  {'':30} {t['description']}")
        click.echo()
    click.echo(f"共 {len(tasks)} 个任务")


@task.command("show")
@click.argument("task_id")
def task_show(task_id):
    """查看任务详情"""
    if task_id not in TASKS:
        click.echo(f"❌ 未知任务: {task_id}")
        sys.exit(1)

    info = TASKS[task_id]
    click.echo(f"\n📋 {info['name']}")
    click.echo(f"   ID: {task_id}")
    click.echo(f"   分类: {info['category']}")
    click.echo(f"   描述: {info['description']}")
    click.echo(f"   脚本: {info['script']}")
    click.echo(f"\n   参数:")
    for p in info.get("params", []):
        req = "必填" if p.get("required", True) else "选填"
        default = f" (默认: {p.get('default')})" if "default" in p else ""
        click.echo(f"     - {p['label']} ({p['id']}) [{p['type']}] {req}{default}")
    click.echo()


@task.command("run")
@click.argument("task_id")
@click.option("--param", "-p", multiple=True, help="参数 (格式: key=value)")
@click.option("--json-file", "-j", type=click.Path(exists=True), help="从 JSON 文件读取参数")
def task_run(task_id, param, json_file):
    """执行单个任务"""
    if task_id not in TASKS:
        click.echo(f"❌ 未知任务: {task_id}")
        sys.exit(1)

    # Build params dict
    params = {}
    if json_file:
        with open(json_file, "r", encoding="utf-8") as f:
            params = json.load(f)
    for p in param:
        key, _, value = p.partition("=")
        params[key.strip()] = value.strip()

    info = TASKS[task_id]
    click.echo(f"🚀 执行任务: {info['name']}")
    click.echo(f"   参数: {json.dumps(params, ensure_ascii=False)}")

    result = run_task_now(task_id, params)
    click.echo(f"\n   状态: {result.get('status', 'unknown')}")
    if result.get("duration"):
        click.echo(f"   耗时: {result['duration']}s")
    if result.get("log_id"):
        click.echo(f"   日志ID: {result['log_id']}")
    if result.get("status") != "completed":
        click.echo(f"   详情: {result.get('message', '未知错误')}")
        sys.exit(1)


# ── Schedule Commands ─────────────────────────────────────────────────────────

@cli.group()
def schedule():
    """定时任务管理"""
    pass


@schedule.command("create")
@click.argument("task_id")
@click.option("--cron", help="Cron 表达式 (e.g. '0 9 * * *')")
@click.option("--interval", help="间隔 (e.g. '30m', '1h', '1d')")
@click.option("--param", "-p", multiple=True, help="参数 (格式: key=value)")
@click.option("--json-file", "-j", type=click.Path(exists=True), help="从 JSON 文件读取参数")
def schedule_create(task_id, cron, interval, param, json_file):
    """创建定时任务"""
    if task_id not in TASKS:
        click.echo(f"❌ 未知任务: {task_id}")
        sys.exit(1)

    # Build params
    params = {}
    if json_file:
        with open(json_file, "r", encoding="utf-8") as f:
            params = json.load(f)
    for p in param:
        key, _, value = p.partition("=")
        params[key.strip()] = value.strip()

    if cron:
        parts = cron.split()
        if len(parts) != 5:
            click.echo("❌ Cron 表达式格式不正确，应为 5 个字段: 分 时 日 月 周")
            sys.exit(1)
        trigger_type = "cron"
        trigger_config = {
            "minute": parts[0], "hour": parts[1],
            "day": parts[2], "month": parts[3], "day_of_week": parts[4]
        }
    elif interval:
        trigger_type = "interval"
        trigger_config = _parse_interval(interval)
    else:
        click.echo("❌ 必须提供 --cron 或 --interval 参数")
        sys.exit(1)

    result = create_schedule(task_id, params, trigger_type, **trigger_config)
    if result["status"] == "ok":
        click.echo(f"✅ 定时任务已创建: {result['schedule_id']}")
    else:
        click.echo(f"❌ 创建失败: {result.get('message')}")
        sys.exit(1)


@schedule.command("list")
def schedule_list():
    """列出所有定时任务"""
    schedules = list_schedules()
    if not schedules:
        click.echo("暂无定时任务")
        return

    click.echo(f"\n{'ID':<38} {'任务':<25} {'类型':<10} {'启用':<4}")
    click.echo(f"{'='*77}")
    for s in schedules:
        icon = "✅" if s["enabled"] else "⏸️"
        click.echo(f"{s['id']:<38} {s['task_name']:<25} {s['trigger_type']:<10} {icon}")
        click.echo(f"{'':38} 配置: {json.dumps(s['trigger_config'], ensure_ascii=False)}")
        click.echo()


@schedule.command("toggle")
@click.argument("schedule_id")
@click.option("--enable", is_flag=True, help="启用")
@click.option("--disable", is_flag=True, help="禁用")
def schedule_toggle(schedule_id, enable, disable):
    """启用/禁用定时任务"""
    if disable:
        enabled = False
    else:
        enabled = True
    result = toggle_schedule(schedule_id, enabled)
    action = "启用" if enabled else "禁用"
    if result["status"] == "ok":
        click.echo(f"✅ 已{action}: {schedule_id}")
    else:
        click.echo(f"❌ 操作失败: {result.get('message')}")
        sys.exit(1)


@schedule.command("delete")
@click.argument("schedule_id")
def schedule_delete(schedule_id):
    """删除定时任务"""
    result = delete_schedule(schedule_id)
    if result["status"] == "ok":
        click.echo(f"✅ 已删除: {schedule_id}")
    else:
        click.echo(f"❌ 删除失败: {result.get('message')}")
        sys.exit(1)


# ── Log Commands ──────────────────────────────────────────────────────────────

@cli.group()
def log():
    """执行日志"""
    pass


@log.command("list")
@click.option("--task", "-t", help="按任务 ID 过滤")
@click.option("--status", "-s", help="按状态过滤 (completed/failed/running)")
@click.option("--limit", "-n", default=20, help="显示条数")
def log_list(task, status, limit):
    """列出执行日志"""
    logs = list_logs(task_id=task, status=status, limit=limit)
    if not logs:
        click.echo("暂无日志记录")
        return

    click.echo(f"\n{'时间':<26} {'任务':<25} {'状态':<12} {'耗时(s)':<8}")
    click.echo(f"{'='*71}")
    for entry in logs:
        t = entry["created_at"][:19]
        status_icon = {"completed": "✅", "failed": "❌", "running": "🔄"}.get(
            entry["status"], entry["status"]
        )
        dur = f"{entry['duration']:.1f}" if entry["duration"] else "-"
        click.echo(f"{t:<26} {entry['task_name']:<25} {status_icon:<12} {dur}")


@log.command("show")
@click.argument("log_id")
def log_show(log_id):
    """查看日志详情"""
    entry = get_log(log_id)
    if not entry:
        click.echo("❌ 日志不存在")
        sys.exit(1)

    click.echo(f"\n📋 {entry['task_name']}")
    click.echo(f"   状态: {entry['status']}")
    click.echo(f"   创建: {entry['created_at']}")
    click.echo(f"   耗时: {entry['duration'] or '-'}s")
    if entry.get("params"):
        click.echo(f"   参数: {entry['params']}")
    if entry.get("stdout"):
        click.echo(f"\n--- 标准输出 ---")
        click.echo(entry["stdout"][:2000])
    if entry.get("stderr"):
        click.echo(f"\n--- 标准错误 ---")
        click.echo(entry["stderr"][:2000])


@log.command("stats")
def log_stats():
    """查看执行统计"""
    stats = get_stats()
    click.echo(f"\n📊 执行统计")
    click.echo(f"   总计: {stats.get('total', 0)}")
    click.echo(f"   成功: {stats.get('completed', 0)}")
    click.echo(f"   失败: {stats.get('failed', 0)}")
    click.echo(f"   运行中: {stats.get('running', 0)}")
    avg = stats.get("avg_duration")
    click.echo(f"   平均耗时: {avg:.1f}s" if avg else "   平均耗时: -")


@log.command("cleanup")
@click.option("--days", "-d", default=30, help="保留天数")
def log_cleanup(days):
    """清理过期日志"""
    deleted = delete_old_logs(days=days)
    click.echo(f"✅ 已清理 {deleted} 条 {days} 天前的日志")


# ── Agent Command ──────────────────────────────────────────────────────────────

@cli.command()
@click.argument("text")
@click.option("--dry-run", is_flag=True, help="仅匹配任务，不执行")
def agent(text, dry_run):
    """自然语言执行任务"""
    click.echo(f"🤖 解析: {text}")

    matches = match_task(text, top_k=3)
    if not matches:
        click.echo("❌ 未找到匹配的任务。可用任务列表: python cli.py task list")
        sys.exit(1)

    best = matches[0]
    click.echo(f"   匹配: {best['name']} ({best['score']:.0%} 置信度)")

    parsed_params = parse_params_from_text(text, best["task_id"])
    if parsed_params:
        click.echo(f"   参数: {json.dumps(parsed_params, ensure_ascii=False)}")
    else:
        click.echo("   ⚠️  未提取到参数，请补充后重试")
        sys.exit(1)

    if dry_run:
        click.echo("   (dry-run 模式，未实际执行)")
        return

    result = run_task_now(best["id"], parsed_params)
    click.echo(f"\n   状态: {result.get('status', 'unknown')}")
    if result.get("duration"):
        click.echo(f"   耗时: {result['duration']}s")


# ── Helpers ────────────────────────────────────────────────────────────────────

def _parse_interval(spec: str) -> dict:
    """Parse interval string like '30m', '1h', '1d' into APScheduler kwargs."""
    spec = spec.strip().lower()
    if spec.endswith("m"):
        return {"minutes": int(spec[:-1])}
    elif spec.endswith("h"):
        return {"hours": int(spec[:-1])}
    elif spec.endswith("d"):
        return {"days": int(spec[:-1])}
    else:
        click.echo(f"❌ 间隔格式不正确: {spec} (支持: 30m, 1h, 1d)")
        sys.exit(1)


if __name__ == "__main__":
    cli()
