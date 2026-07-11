"""
scripts/conftest.py
-------------------
将项目根目录注入 sys.path，使 scripts/ 下所有测试文件可以直接
    from scripts.xxx import yyy
而无需在每个测试文件顶部重复写 sys.path.insert(...)。

pytest 会在收集测试前自动加载同目录的 conftest.py。
"""
import sys
from pathlib import Path

# 项目根目录（scripts/ 的上一级）
_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))
