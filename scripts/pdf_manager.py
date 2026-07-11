"""
PDF 文件管理核心模块
支持：列出文件、删除页面、提取页面、重命名、删除文件
"""
from __future__ import annotations

# The rest of your imports...

import os
import math
from pathlib import Path

try:
    from pypdf import PdfReader, PdfWriter
except ImportError:
    raise ImportError("请安装 pypdf: pip install pypdf")

try:
    from pdf2docx import Converter
except ImportError:
    pass  # Allow graceful degradation if pdf2docx is not installed yet

# 文件统一存储目录
PDF_UPLOAD_DIR = Path(__file__).parent.parent / "uploads" / "pdf"
WORD_UPLOAD_DIR = Path(__file__).parent.parent / "uploads" / "word"


def ensure_upload_dir():
    """确保上传目录存在"""
    PDF_UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    WORD_UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    return PDF_UPLOAD_DIR



def list_pdfs() -> list[dict]:
    """列出所有已上传的 PDF 文件及其元数据"""
    ensure_upload_dir()
    result = []
    for f in sorted(PDF_UPLOAD_DIR.glob("*.pdf"), key=lambda x: x.stat().st_mtime, reverse=True):
        info = {
            "filename": f.name,
            "size_bytes": f.stat().st_size,
            "size_str": _human_size(f.stat().st_size),
            "mtime": f.stat().st_mtime,
            "pages": 0,
        }
        try:
            reader = PdfReader(str(f))
            info["pages"] = len(reader.pages)
        except Exception:
            info["pages"] = -1  # 损坏或加密文件
        result.append(info)
    return result


def save_uploaded_pdf(file_storage, filename: str) -> dict:
    """
    保存上传的 PDF 文件（接收 Flask FileStorage 对象）
    返回文件信息
    """
    ensure_upload_dir()
    # 安全化文件名，防路径穿越
    safe_name = _safe_filename(filename)
    dest = PDF_UPLOAD_DIR / safe_name

    # 若同名文件已存在，自动加序号
    if dest.exists():
        stem = dest.stem
        suffix = dest.suffix
        counter = 1
        while dest.exists():
            dest = PDF_UPLOAD_DIR / f"{stem}_{counter}{suffix}"
            counter += 1

    file_storage.save(str(dest))

    reader = PdfReader(str(dest))
    return {
        "filename": dest.name,
        "size_str": _human_size(dest.stat().st_size),
        "pages": len(reader.pages),
    }


def delete_pdf_file(filename: str) -> bool:
    """永久删除指定 PDF 文件"""
    path = _resolve_safe_path(filename)
    if not path:
        return False
    path.unlink()
    return True


def rename_pdf(old_name: str, new_name: str) -> dict:
    """重命名 PDF 文件"""
    old_path = _resolve_safe_path(old_name)
    if not old_path:
        raise FileNotFoundError(f"文件不存在: {old_name}")

    # 确保新名称有 .pdf 后缀
    new_safe = _safe_filename(new_name)
    if not new_safe.lower().endswith(".pdf"):
        new_safe += ".pdf"

    new_path = PDF_UPLOAD_DIR / new_safe
    if new_path.exists():
        raise FileExistsError(f"目标文件名已存在: {new_safe}")

    old_path.rename(new_path)
    return {"old_name": old_name, "new_name": new_safe}


def delete_pages(filename: str, pages_to_remove: list[int]) -> dict:
    """
    从 PDF 中删除指定页码（1-indexed），保存覆盖原文件
    pages_to_remove: 1-indexed 页码列表
    """
    path = _resolve_safe_path(filename)
    if not path:
        raise FileNotFoundError(f"文件不存在: {filename}")

    reader = PdfReader(str(path))
    total = len(reader.pages)

    # 转为 0-indexed 集合并过滤非法值
    remove_set = {p - 1 for p in pages_to_remove if 1 <= p <= total}
    if not remove_set:
        raise ValueError("没有有效的页码可删除")
    if len(remove_set) >= total:
        raise ValueError("不能删除全部页面，至少保留 1 页")

    writer = PdfWriter()
    for i, page in enumerate(reader.pages):
        if i not in remove_set:
            writer.add_page(page)

    with open(str(path), "wb") as f:
        writer.write(f)

    return {
        "filename": filename,
        "removed_pages": sorted(p + 1 for p in remove_set),
        "remaining_pages": total - len(remove_set),
    }


def extract_pages(filename: str, page_start: int, page_end: int, new_filename: str) -> dict:
    """
    提取指定页码范围，另存为新文件（1-indexed，含首尾）
    """
    path = _resolve_safe_path(filename)
    if not path:
        raise FileNotFoundError(f"文件不存在: {filename}")

    reader = PdfReader(str(path))
    total = len(reader.pages)

    if not (1 <= page_start <= page_end <= total):
        raise ValueError(f"页码范围无效: {page_start}-{page_end}，文件共 {total} 页")

    # 安全化新文件名
    safe_new = _safe_filename(new_filename)
    if not safe_new.lower().endswith(".pdf"):
        safe_new += ".pdf"

    new_path = PDF_UPLOAD_DIR / safe_new
    if new_path.exists():
        raise FileExistsError(f"目标文件名已存在: {safe_new}")

    writer = PdfWriter()
    for i in range(page_start - 1, page_end):
        writer.add_page(reader.pages[i])

    with open(str(new_path), "wb") as f:
        writer.write(f)

    return {
        "new_filename": safe_new,
        "extracted_pages": page_end - page_start + 1,
        "size_str": _human_size(new_path.stat().st_size),
    }


def get_pdf_path(filename: str) -> Path | None:
    """获取 PDF 文件的安全路径（用于预览/下载）"""
    return _resolve_safe_path(filename)


# ─── 内部工具函数 ────────────────────────────────────────────────────────────

def _resolve_safe_path(filename: str) -> Path | None:
    """安全解析文件路径，防路径穿越攻击"""
    ensure_upload_dir()
    safe = _safe_filename(filename)
    path = PDF_UPLOAD_DIR / safe
    # 确保解析后的路径仍在 upload 目录内
    try:
        path.resolve().relative_to(PDF_UPLOAD_DIR.resolve())
    except ValueError:
        return None
    return path if path.exists() else None


def _safe_filename(name: str) -> str:
    """移除路径分隔符等危险字符"""
    return Path(name).name


def _human_size(size_bytes: int) -> str:
    """将字节数转为人类可读格式"""
    if size_bytes == 0:
        return "0 B"
    units = ["B", "KB", "MB", "GB"]
    i = int(math.floor(math.log(size_bytes, 1024)))
    i = min(i, len(units) - 1)
    p = math.pow(1024, i)
    return f"{size_bytes / p:.1f} {units[i]}"


def add_password(filename: str, user_password: str, owner_password: str = "") -> dict:
    """
    为 PDF 文件添加密码保护，生成新的加密文件（原文件保留）。
    - user_password:  打开文件时所需的密码
    - owner_password: 权限密码（可选；若留空则与 user_password 相同）
    生成的文件命名规则：{stem}_locked.pdf
    """
    path = _resolve_safe_path(filename)
    if not path:
        raise FileNotFoundError(f"文件不存在: {filename}")

    if not user_password:
        raise ValueError("用户密码不能为空")

    # 确保权限密码不空（pypdf 要求）
    owner_pwd = owner_password.strip() if owner_password.strip() else user_password

    # 生成带 _locked 后缀的新文件名
    stem = path.stem
    locked_name = f"{stem}_locked.pdf"
    locked_path = PDF_UPLOAD_DIR / locked_name

    # 若同名已存在则加序号
    counter = 1
    while locked_path.exists():
        locked_path = PDF_UPLOAD_DIR / f"{stem}_locked_{counter}.pdf"
        counter += 1
    locked_name = locked_path.name

    reader = PdfReader(str(path))
    writer = PdfWriter()

    # 复制所有页面
    for page in reader.pages:
        writer.add_page(page)

    # 复制元数据（如有）
    if reader.metadata:
        writer.add_metadata(reader.metadata)

    # 使用 AES-256 加密（pypdf >= 3.x 默认）
    writer.encrypt(
        user_password=user_password,
        owner_password=owner_pwd,
        use_128bit=False,   # False → 使用 256-bit AES（更安全）
    )

    with open(str(locked_path), "wb") as f:
        writer.write(f)

    return {
        "original_filename": filename,
        "locked_filename": locked_name,
        "size_str": _human_size(locked_path.stat().st_size),
    }



def get_word_path(filename: str) -> Path | None:
    """获取生成的 Word 文件的安全路径"""
    ensure_upload_dir()
    safe = _safe_filename(filename)
    path = WORD_UPLOAD_DIR / safe
    try:
        path.resolve().relative_to(WORD_UPLOAD_DIR.resolve())
    except ValueError:
        return None
    return path if path.exists() else None


def convert_to_word(filename: str) -> dict:
    """
    将指定的 PDF 转换为 Word (.docx)
    转换后的文件存放在 WORD_UPLOAD_DIR 下
    """
    pdf_path = _resolve_safe_path(filename)
    if not pdf_path:
        raise FileNotFoundError(f"未找到指定的 PDF 文件: {filename}")

    # 确定生成的 docx 文件名
    pdf_stem = pdf_path.stem
    docx_filename = f"{pdf_stem}.docx"
    docx_path = WORD_UPLOAD_DIR / docx_filename

    # 使用 pdf2docx 执行转换
    cv = Converter(str(pdf_path))
    # start=0, end=None 表示转换所有页面
    cv.convert(str(docx_path), start=0, end=None)
    cv.close()

    return {
        "pdf_filename": filename,
        "docx_filename": docx_filename,
        "size_str": _human_size(docx_path.stat().st_size)
    }
