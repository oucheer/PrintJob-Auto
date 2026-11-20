import sys
from pathlib import Path

# 让测试可以直接 `pytest tests/` 运行，无需手动设置 PYTHONPATH
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
