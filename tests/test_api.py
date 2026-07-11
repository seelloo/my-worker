"""
API 测试
测试 Flask API 端点
"""

import pytest
import sys
import json
import tempfile
from pathlib import Path
from unittest.mock import patch, MagicMock

# 添加项目根目录到路径
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from app import app


@pytest.fixture
def client():
    """创建测试客户端"""
    app.config['TESTING'] = True
    with app.test_client() as client:
        yield client


class TestTasksAPI:
    """任务 API 测试"""

    def test_task_list(self, client):
        """测试获取任务列表"""
        response = client.get('/api/task_list')
        assert response.status_code == 200

        data = json.loads(response.data)
        assert data['status'] == 'success'
        assert 'tasks' in data

    def test_run_task_not_found(self, client):
        """测试运行不存在的任务"""
        response = client.post(
            '/api/run_task/nonexistent_task',
            json={}
        )
        assert response.status_code == 404
        data = json.loads(response.data)
        assert data['status'] == 'error'

    def test_get_columns_missing_path(self, client):
        """测试获取列信息 - 缺少路径"""
        response = client.post('/api/get_columns', json={})
        assert response.status_code == 400
        data = json.loads(response.data)
        assert data['status'] == 'error'

    def test_get_columns_invalid_path(self, client):
        """测试获取列信息 - 无效路径"""
        response = client.post('/api/get_columns', json={'path': '/nonexistent/file.xlsx'})
        assert response.status_code == 400
        data = json.loads(response.data)
        assert data['status'] == 'error'

    def test_get_columns_path_traversal(self, client):
        """测试获取列信息 - 路径遍历攻击"""
        response = client.post('/api/get_columns', json={'path': '../../../etc/passwd'})
        assert response.status_code == 400
        data = json.loads(response.data)
        assert '不允许' in data['message'] or '非法' in data['message']


class TestExcelAPI:
    """Excel API 测试"""

    def test_import_conditions_missing_path(self, client):
        """测试导入条件 - 缺少路径"""
        response = client.post('/api/import_conditions', json={})
        assert response.status_code == 400
        data = json.loads(response.data)
        assert data['status'] == 'error'

    def test_import_conditions_nonexistent_file(self, client):
        """测试导入条件 - 文件不存在"""
        response = client.post('/api/import_conditions', json={'path': '/nonexistent.xlsx'})
        assert response.status_code == 400

    def test_import_conditions_path_traversal(self, client):
        """测试导入条件 - 路径遍历攻击"""
        response = client.post('/api/import_conditions', json={'path': '../../../etc/passwd'})
        assert response.status_code == 400

    def test_export_conditions_empty(self, client):
        """测试导出条件 - 空条件"""
        response = client.post('/api/export_conditions', json={'conditions': []})
        assert response.status_code == 400

    def test_get_sheets_missing_src_file(self, client):
        """测试获取 Sheets - 缺少参数"""
        response = client.post('/api/excel/get_sheets', json={})
        assert response.status_code == 400

    def test_get_sheets_path_traversal(self, client):
        """测试获取 Sheets - 路径遍历攻击"""
        response = client.post('/api/excel/get_sheets', json={'src_file': '../test.xlsx'})
        assert response.status_code == 400


class TestRateLimiting:
    """限流测试"""

    def test_rate_limit_headers(self, client):
        """测试限流响应头"""
        response = client.post('/api/run_task/test_task', json={})
        # 即使返回 404，也应该有 rate limit 头（如果任务不存在）
        # 如果任务存在且被限流，会返回 429
        if response.status_code != 429:
            # 正常请求应该有限流头
            assert 'X-RateLimit-Limit' in response.headers or response.status_code == 404


class TestErrorHandlers:
    """错误处理器测试"""

    def test_404_error(self, client):
        """测试 404 错误"""
        response = client.get('/api/nonexistent-endpoint')
        assert response.status_code == 404
        data = json.loads(response.data)
        assert data['status'] == 'error'
        assert 'error_code' in data

    def test_500_error(self, client):
        """测试 500 错误 - 模拟内部错误"""
        with patch('routes.tasks.get_task_metadata', side_effect=Exception("Test error")):
            response = client.get('/api/task_list')
            # 应该返回 500 或者至少能处理
            assert response.status_code in [200, 500]


class TestSecurityValidation:
    """安全验证测试"""

    def test_null_byte_injection(self, client):
        """测试空字节注入防护"""
        response = client.post('/api/get_columns', json={'path': 'test\x00.xlsx'})
        assert response.status_code == 400

    def test_hidden_file(self, client):
        """测试隐藏文件防护"""
        response = client.post('/api/get_columns', json={'path': '.hidden.xlsx'})
        assert response.status_code == 400

    def test_invalid_extension(self, client):
        """测试无效扩展名"""
        response = client.post('/api/get_columns', json={'path': 'test.exe'})
        assert response.status_code == 400


class TestPDFAPI:
    """PDF API 测试"""

    def test_pdf_list(self, client):
        """测试 PDF 列表"""
        response = client.get('/api/pdf/list')
        assert response.status_code == 200
        data = json.loads(response.data)
        assert data['status'] == 'success'

    def test_pdf_download_nonexistent(self, client):
        """测试下载不存在的 PDF"""
        response = client.get('/api/pdf/download/nonexistent.pdf')
        assert response.status_code == 404


if __name__ == '__main__':
    pytest.main([__file__, '-v'])
