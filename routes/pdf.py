import os
from flask import Blueprint, jsonify, request, send_file

from scripts.pdf_manager import (
    save_uploaded_pdf, list_pdfs, get_pdf_path, delete_pdf_file,
    rename_pdf, delete_pages, extract_pages, convert_to_word,
    get_word_path, add_password
)

pdf_bp = Blueprint('pdf', __name__, url_prefix='/api')


@pdf_bp.route('/pdf/upload', methods=['POST'])
def api_pdf_upload():
    """上传 PDF 文件（支持多文件）"""
    if 'files' not in request.files:
        return jsonify({"status": "error", "message": "没有收到文件"}), 400
    files = request.files.getlist('files')
    results = []
    errors = []
    for f in files:
        if not f.filename or not f.filename.lower().endswith('.pdf'):
            errors.append(f"{f.filename}: 不是有效的 PDF 文件")
            continue
        try:
            info = save_uploaded_pdf(f, f.filename)
            results.append(info)
        except Exception as e:
            errors.append(f"{f.filename}: {str(e)}")
    return jsonify({"status": "success", "uploaded": results, "errors": errors})


@pdf_bp.route('/pdf/list', methods=['GET'])
def api_pdf_list():
    """获取已上传的所有 PDF 文件列表"""
    try:
        return jsonify({"status": "success", "files": list_pdfs()})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


@pdf_bp.route('/pdf/preview/<path:filename>', methods=['GET'])
def api_pdf_preview(filename):
    """返回 PDF 文件内容供浏览器内嵌预览"""
    path = get_pdf_path(filename)
    if not path:
        return jsonify({"status": "error", "message": "文件不存在"}), 404
    return send_file(str(path), mimetype='application/pdf')


@pdf_bp.route('/pdf/download/<path:filename>', methods=['GET'])
def api_pdf_download(filename):
    """下载 PDF 文件"""
    print(f"[DEBUG] Requested PDF download: {filename}")
    path = get_pdf_path(filename)
    print(f"[DEBUG] Resolved Path: {path}")
    if not path:
        return jsonify({"status": "error", "message": "文件不存在"}), 404
    try:
        return send_file(str(path), as_attachment=True, download_name=filename)
    except TypeError:
        return send_file(str(path), as_attachment=True, attachment_filename=filename)


@pdf_bp.route('/pdf/delete/<path:filename>', methods=['DELETE'])
def api_pdf_delete(filename):
    """删除 PDF 文件"""
    try:
        if delete_pdf_file(filename):
            return jsonify({"status": "success", "message": f"已删除: {filename}"})
        return jsonify({"status": "error", "message": "文件不存在"}), 404
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


@pdf_bp.route('/pdf/rename', methods=['POST'])
def api_pdf_rename():
    """重命名 PDF 文件"""
    data = request.json or {}
    old_name = data.get('old_name', '').strip()
    new_name = data.get('new_name', '').strip()
    if not old_name or not new_name:
        return jsonify({"status": "error", "message": "参数缺失"}), 400
    try:
        result = rename_pdf(old_name, new_name)
        return jsonify({"status": "success", **result})
    except (FileNotFoundError, FileExistsError) as e:
        return jsonify({"status": "error", "message": str(e)}), 400
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


@pdf_bp.route('/pdf/delete_pages', methods=['POST'])
def api_pdf_delete_pages():
    """删除 PDF 中指定页码"""
    data = request.json or {}
    filename = data.get('filename', '').strip()
    pages = data.get('pages', [])
    if not filename or not pages:
        return jsonify({"status": "error", "message": "参数缺失"}), 400
    try:
        result = delete_pages(filename, [int(p) for p in pages])
        return jsonify({"status": "success", **result})
    except (FileNotFoundError, ValueError) as e:
        return jsonify({"status": "error", "message": str(e)}), 400
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


@pdf_bp.route('/pdf/extract_pages', methods=['POST'])
def api_pdf_extract_pages():
    """提取 PDF 中指定页码范围为新文件"""
    data = request.json or {}
    filename = data.get('filename', '').strip()
    page_start = data.get('page_start')
    page_end = data.get('page_end')
    new_filename = data.get('new_filename', '').strip()
    if not all([filename, page_start, page_end, new_filename]):
        return jsonify({"status": "error", "message": "参数缺失"}), 400
    try:
        result = extract_pages(filename, int(page_start), int(page_end), new_filename)
        return jsonify({"status": "success", **result})
    except (FileNotFoundError, FileExistsError, ValueError) as e:
        return jsonify({"status": "error", "message": str(e)}), 400
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


@pdf_bp.route('/pdf/to_word', methods=['POST'])
def api_pdf_to_word():
    """将 PDF 文件转换为 Word"""
    data = request.json or {}
    filename = data.get('filename', '').strip()
    if not filename:
        return jsonify({"status": "error", "message": "缺少 filename 参数"}), 400
    try:
        result = convert_to_word(filename)
        return jsonify({"status": "success", **result})
    except FileNotFoundError as e:
        return jsonify({"status": "error", "message": str(e)}), 404
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


@pdf_bp.route('/pdf/download_word/<path:filename>', methods=['GET'])
def api_pdf_download_word(filename):
    """下载生成的 Word 文件"""
    print(f"[DEBUG] Requested Word download: {filename}")
    path = get_word_path(filename)
    print(f"[DEBUG] Resolved Word Path: {path}")
    if not path:
        return jsonify({"status": "error", "message": "Word 文件不存在"}), 404
    try:
        return send_file(str(path), as_attachment=True, download_name=filename)
    except TypeError:
        return send_file(str(path), as_attachment=True, attachment_filename=filename)


@pdf_bp.route('/pdf/add_password', methods=['POST'])
def api_pdf_add_password():
    """为 PDF 添加密码保护，生成加密副本"""
    data = request.json or {}
    filename = data.get('filename', '').strip()
    user_password = data.get('user_password', '').strip()
    owner_password = data.get('owner_password', '').strip()
    if not filename or not user_password:
        return jsonify({"status": "error", "message": "文件名和用户密码不能为空"}), 400
    try:
        result = add_password(filename, user_password, owner_password)
        return jsonify({"status": "success", **result})
    except (FileNotFoundError, ValueError) as e:
        return jsonify({"status": "error", "message": str(e)}), 400
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500
