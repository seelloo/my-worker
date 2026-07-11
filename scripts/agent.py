"""
Agent — natural language routing to tasks.

Parses user intent, matches to the best task, extracts parameters,
and delegates to the task engine.
"""
from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

from scripts.task_registry import TASKS


# ── Keyword → Task ID mapping ─────────────────────────────────────────────────

_KEYWORD_MAP: dict[str, list[str]] = {
    # Excel operations
    "复制": ["copy_excel"],
    "筛选 excel": ["copy_excel", "filter_excel"],
    "sheet合并": ["merge_sheets"],
    "合并sheet": ["merge_sheets"],
    "excel合并": ["merge_sheets", "concat_excel"],
    "合并文件夹": ["concat_excel"],
    "文件夹合并": ["concat_excel"],
    "全合并": ["concat_excel"],
    "行追加": ["concat_excel"],
    "关联": ["excel_join_multiply", "excel_xlookup"],
    "乘积": ["excel_join_multiply"],
    "xlookup": ["excel_xlookup"],
    "查找": ["excel_xlookup"],
    "按比例": ["excel_proportional_alloc"],
    "比例分配": ["excel_proportional_alloc"],
    "权重": ["excel_proportional_alloc"],
    "分组": ["excel_group_assign"],
    "组编号": ["excel_group_assign"],
    "列顺序": ["excel_column_reorder"],
    "列重组": ["excel_column_reorder"],
    "拆分": ["excel_sheet_split"],
    "sheet拆分": ["excel_sheet_split"],
    "重组": ["excel_sheet_merge", "cross_sheet_merge"],
    "跨sheet": ["cross_sheet_merge"],
    "列拼接": ["cross_sheet_merge"],
    "表头": ["rename_excel_headers"],
    "sheet重命名": ["rename_excel_sheets"],
    "提取列": ["collect_column"],
    "条件筛选": ["filter_excel"],
    "过滤": ["filter_excel"],

    # File operations
    "文件名": ["clean_filenames", "rename_to_dir"],
    "清洗文件名": ["clean_filenames"],
    "重命名为目录": ["rename_to_dir"],

    # Format conversion
    "csv转excel": ["csv_to_excel"],
    "csv转xlsx": ["csv_to_excel"],
    "ofd转图片": ["ofd_to_image"],
    "ofd转jpg": ["ofd_to_image"],
    "发票提取": ["ofd_extract"],
    "电子发票": ["ofd_extract"],

    # PDF
    "pdf转word": ["pdf_convert_word"],
    "pdf转docx": ["pdf_convert_word"],
    "pdf加密": ["pdf_add_password"],
    "pdf密码": ["pdf_add_password"],

    # Business
    "日报表": ["accounts_daily_report"],
    "佣金": ["accounts_daily_report"],
}


def match_task(text: str, top_k: int = 1) -> list[dict[str, Any]]:
    """
    Match user text to the most relevant task(s).

    Returns list of dicts with: id, name, score, matched_keywords.
    """
    text_lower = text.lower()
    scores: dict[str, tuple[float, list[str]]] = {}

    for keyword, task_ids in _KEYWORD_MAP.items():
        if keyword in text_lower:
            for tid in task_ids:
                prev_score, prev_kw = scores.get(tid, (0.0, []))
                # Longer keywords get higher weight
                weight = len(keyword) / 10.0
                scores[tid] = (prev_score + weight, prev_kw + [keyword])

    if not scores:
        # Fallback: fuzzy match against task descriptions
        for tid, info in TASKS.items():
            desc = info.get("description", "").lower()
            name = info.get("name", "").lower()
            for word in text_lower.split():
                if word in desc or word in name:
                    prev_score, prev_kw = scores.get(tid, (0.0, []))
                    scores[tid] = (prev_score + 0.3, prev_kw + [word])

    # Normalize scores to 0-1 range
    max_score = max(s for s, _ in scores.values())
    ranked = []
    for tid, (score, keywords) in sorted(scores.items(), key=lambda x: -x[1][0]):
        info = TASKS.get(tid, {})
        normalized = min(score / max_score, 1.0)
        ranked.append({
            "id": tid,
            "name": info.get("name", tid),
            "score": normalized,
            "matched_keywords": list(set(keywords)),
        })

    return ranked[:top_k]


def parse_params_from_text(text: str, task_id: str) -> dict[str, Any]:
    """
    Extract parameters from natural language text based on task definition.

    Looks for:
      - File paths (Windows/Unix style)
      - Directory paths
      - Column names (quoted strings)
      - Numbers
      - Keywords
    """
    if task_id not in TASKS:
        return {}

    params = TASKS[task_id].get("params", [])
    result: dict[str, Any] = {}

    # Extract all quoted strings
    quoted = re.findall(r'["“”]([^"“”]+)["“”]', text)

    # Extract paths (Windows and Unix style)
    win_paths = re.findall(r'([A-Za-z]:[/\\][^\s:"\']+(?:\.[a-zA-Z0-9]+)?)', text)
    unix_paths = re.findall(r'(/[^\s:"\']+(?:\.[a-zA-Z0-9]+)?)', text)
    all_paths = win_paths + unix_paths

    # Extract numbers
    numbers = re.findall(r'(\d+(?:\.\d+)?)', text)

    for p in params:
        pid = p["id"]
        ptype = p.get("type", "text")
        label = p.get("label", "").lower()

        # Try to match parameter from label keywords in text
        label_words = [w for w in re.split(r'[\s/()]+', label) if len(w) > 1]
        for word in label_words:
            if word.lower() in text.lower():
                # Found label hint in text — try to extract value
                if "路径" in label or "目录" in label or "file" in pid:
                    if all_paths:
                        result[pid] = all_paths.pop(0) if all_paths else ""
                        break
                elif ptype == "number":
                    if numbers:
                        result[pid] = float(numbers.pop(0))
                        break

    # For "keyword" type params, extract remaining unquoted words
    for p in params:
        pid = p["id"]
        if pid not in result and "关键字" in p.get("label", ""):
            # Extract Chinese keywords (2-6 chars) from text
            keywords = re.findall(r'[一-鿿]{2,6}', text)
            if keywords:
                # Filter out common words already in task description
                task_desc = TASKS[task_id].get("description", "")
                filtered = [k for k in keywords if k not in task_desc]
                if filtered:
                    result[pid] = filtered[0]

    return result
