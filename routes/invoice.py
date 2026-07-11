import os
import sqlite3
from flask import Blueprint, jsonify, request, send_file

from config import DB_PATH
from scripts.extract_ofd_invoice import save_xlsx

invoice_bp = Blueprint('invoice', __name__, url_prefix='/api')


@invoice_bp.route('/get_invoices', methods=['GET'])
def api_get_invoices():
    """获取发票大盘数据（支持多维过滤与实时统计）"""
    db_path = DB_PATH
    keyword = request.args.get('keyword', '')
    try:
        if not os.path.exists(db_path):
            return jsonify({"status": "success", "data": [], "stats": {"total_amount": 0, "total_tax": 0, "count": 0}})

        conn = sqlite3.connect(db_path)
        try:
            conn.row_factory = sqlite3.Row

            query = 'SELECT * FROM invoices'
            params = []
            if keyword:
                query += ' WHERE seller_name LIKE ? OR invoice_number LIKE ? OR file_name LIKE ?'
                params = [f'%{keyword}%', f'%{keyword}%', f'%{keyword}%']

            query += ' ORDER BY created_at DESC LIMIT 500'
            rows = conn.execute(query, params).fetchall()

            stats_query = 'SELECT SUM(total_amount), SUM(tax_amount), COUNT(*) FROM invoices'
            if keyword:
                stats_query += ' WHERE seller_name LIKE ? OR invoice_number LIKE ? OR file_name LIKE ?'
                stats = conn.execute(stats_query, params).fetchone()
            else:
                stats = conn.execute(stats_query).fetchone()
        finally:
            conn.close()

        return jsonify({
            "status": "success",
            "data": [dict(r) for r in rows],
            "stats": {
                "total_amount": round(stats[0] or 0, 2),
                "total_tax": round(stats[1] or 0, 2),
                "count": stats[2] or 0
            }
        })
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


@invoice_bp.route('/export_invoices', methods=['GET'])
def api_export_invoices():
    """将当前筛选的发票记录导出为标准化 Excel"""
    db_path = DB_PATH
    keyword = request.args.get('keyword', '')

    try:
        if not os.path.exists(db_path):
            return jsonify({"status": "error", "message": "数据库不存在"}), 404

        conn = sqlite3.connect(db_path)
        try:
            conn.row_factory = sqlite3.Row

            query = 'SELECT * FROM invoices'
            params = []
            if keyword:
                query += ' WHERE seller_name LIKE ? OR invoice_number LIKE ? OR file_name LIKE ?'
                params = [f'%{keyword}%', f'%{keyword}%', f'%{keyword}%']

            query += ' ORDER BY created_at DESC'
            rows = [dict(r) for r in conn.execute(query, params).fetchall()]
        finally:
            conn.close()

        if not rows:
            return jsonify({"status": "error", "message": "没有可导出的数据"}), 400

        import tempfile
        fd, temp_path = tempfile.mkstemp(suffix='.xlsx')
        os.close(fd)

        save_xlsx(rows, temp_path)

        return send_file(
            temp_path,
            as_attachment=True,
            download_name=f"Invoice_Export_{keyword or 'All'}.xlsx",
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500
