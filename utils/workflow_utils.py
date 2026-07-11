"""
工作流工具类
提供工作流相关的通用工具函数
"""

from typing import Dict, Any
from datetime import datetime

from workflows.models.workflow import WorkflowExecution


class WorkflowReportGenerator:
    """工作流执行报告生成器"""

    @staticmethod
    def generate(execution: WorkflowExecution) -> str:
        """生成执行报告"""
        lines = []
        lines.append("=" * 50)
        lines.append("工作流执行报告")
        lines.append("=" * 50)
        lines.append(f"工作流名称: {execution.workflow_name}")
        lines.append(f"工作流ID: {execution.workflow_id}")
        lines.append(f"执行ID: {execution.id}")
        lines.append(f"状态: {execution.status}")
        lines.append(f"开始时间: {execution.start_time}")
        lines.append(f"结束时间: {execution.end_time}")
        lines.append(f"持续时间: {execution.duration:.2f}秒")
        lines.append("")
        lines.append("-" * 50)
        lines.append("任务执行详情:")
        lines.append("-" * 50)
        for te in execution.task_executions:
            lines.append(f"  [{te.status}] {te.task_name} ({te.duration:.2f}秒)")
            if te.error:
                lines.append(f"    错误: {te.error}")
        lines.append("")
        lines.append("=" * 50)
        lines.append("日志记录:")
        lines.append("=" * 50)
        for log in execution.logs:
            lines.append(f"  [{log['level']}] {log['message']}")
        return "\n".join(lines)

