from __future__ import annotations

import importlib
import sys
from pathlib import Path

ORIGINAL_SRC = Path(__file__).resolve().parents[3]


def import_original(module_name: str):
    """从项目根目录加载原程序，确保维护根目录源码即可同步 CLI/API 行为。"""
    source = str(ORIGINAL_SRC)
    if source not in sys.path:
        sys.path.insert(0, source)
    return importlib.import_module(module_name)
