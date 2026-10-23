from __future__ import annotations

import importlib
import sys
from pathlib import Path

ORIGINAL_SRC = Path(__file__).resolve().parents[3] / "legacy" / "original_src" / "检验报告模板转换"


def import_original(module_name: str):
    """从复制的原实现加载模块，转换规则本身不在此层修改。"""
    source = str(ORIGINAL_SRC)
    if source not in sys.path:
        sys.path.insert(0, source)
    return importlib.import_module(module_name)
