"""GitHub Trending — 使用 GitHub Search API 模拟热门项目排行"""

import time
from datetime import datetime, timedelta, timezone
from flask import Blueprint, jsonify, request

try:
    import requests as _requests
    HAS_REQUESTS = True
except ImportError:
    HAS_REQUESTS = False

github_bp = Blueprint("github", __name__, url_prefix="/api/github")

# ── 内存缓存 ──────────────────────────────────────────────────────────────────
_cache: dict = {}          # key → (timestamp, data)
_CACHE_TTL = 30 * 60      # 30 分钟

def _cache_get(key: str):
    entry = _cache.get(key)
    if entry and time.time() - entry[0] < _CACHE_TTL:
        return entry[1]
    return None

def _cache_set(key: str, data):
    _cache[key] = (time.time(), data)


# ── 时间范围映射 ───────────────────────────────────────────────────────────────
def _date_from(since: str) -> str:
    """返回 YYYY-MM-DD 格式的起始日期"""
    now = datetime.now(timezone.utc)
    delta_map = {"daily": 1, "weekly": 7, "monthly": 30}
    days = delta_map.get(since, 1)
    return (now - timedelta(days=days)).strftime("%Y-%m-%d")


# ── 语言映射（GitHub API 语言参数区分大小写）──────────────────────────────────
_LANG_MAP = {
    "all": "",
    "python": "Python",
    "javascript": "JavaScript",
    "typescript": "TypeScript",
    "go": "Go",
    "rust": "Rust",
    "java": "Java",
    "cpp": "C++",
    "csharp": "C#",
    "shell": "Shell",
    "vue": "Vue",
}


@github_bp.route("/trending", methods=["GET"])
def api_trending():
    if not HAS_REQUESTS:
        return jsonify({
            "status": "error",
            "message": "缺少依赖 requests，请运行: pip install requests"
        }), 500

    lang_key = request.args.get("lang", "all").lower()
    since = request.args.get("since", "daily")
    token = request.args.get("token", "").strip()   # 可选：传入 PAT 提升限额

    cache_key = f"{lang_key}:{since}"
    cached = _cache_get(cache_key)
    if cached:
        return jsonify({"status": "success", "cached": True, **cached})

    # 构建查询
    date_from = _date_from(since)
    lang_val = _LANG_MAP.get(lang_key, "")
    q_parts = [f"created:>{date_from}", "stars:>10"]
    if lang_val:
        q_parts.append(f"language:{lang_val}")
    query = " ".join(q_parts)

    headers = {
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }
    if token:
        headers["Authorization"] = f"Bearer {token}"

    try:
        resp = _requests.get(
            "https://api.github.com/search/repositories",
            params={
                "q": query,
                "sort": "stars",
                "order": "desc",
                "per_page": 25,
            },
            headers=headers,
            timeout=15,
        )

        if resp.status_code == 403:
            rate = resp.headers.get("X-RateLimit-Reset", "")
            reset_time = ""
            if rate:
                reset_time = datetime.fromtimestamp(int(rate)).strftime("%H:%M")
            return jsonify({
                "status": "error",
                "message": f"GitHub API 限流，请稍后重试（限额将于 {reset_time} 重置）。可提供 Personal Access Token 提升额度。",
                "rate_limited": True,
            }), 429

        if resp.status_code != 200:
            return jsonify({
                "status": "error",
                "message": f"GitHub API 返回 {resp.status_code}"
            }), resp.status_code

        raw = resp.json()
        items = []
        for repo in raw.get("items", []):
            items.append({
                "id": repo["id"],
                "name": repo["name"],
                "full_name": repo["full_name"],
                "url": repo["html_url"],
                "description": repo.get("description") or "",
                "language": repo.get("language") or "Unknown",
                "stars": repo.get("stargazers_count", 0),
                "forks": repo.get("forks_count", 0),
                "watchers": repo.get("watchers_count", 0),
                "open_issues": repo.get("open_issues_count", 0),
                "avatar": repo["owner"]["avatar_url"],
                "owner": repo["owner"]["login"],
                "created_at": repo.get("created_at", ""),
                "updated_at": repo.get("updated_at", ""),
                "topics": repo.get("topics", [])[:5],
            })

        # 剩余配额
        remaining = int(resp.headers.get("X-RateLimit-Remaining", -1))
        limit = int(resp.headers.get("X-RateLimit-Limit", -1))

        payload = {
            "repos": items,
            "total_count": raw.get("total_count", 0),
            "lang": lang_key,
            "since": since,
            "date_from": date_from,
            "rate_remaining": remaining,
            "rate_limit": limit,
        }
        _cache_set(cache_key, payload)
        return jsonify({"status": "success", "cached": False, **payload})

    except _requests.exceptions.Timeout:
        return jsonify({"status": "error", "message": "请求 GitHub API 超时"}), 504
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500
