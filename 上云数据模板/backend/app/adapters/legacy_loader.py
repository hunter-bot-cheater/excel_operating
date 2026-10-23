from __future__ import annotations

import importlib
import sys
from pathlib import Path

ORIGINAL_SRC = Path(__file__).resolve().parents[3] / "legacy" / "original_src"


def import_original(module_name: str):
    source = str(ORIGINAL_SRC)
    if source not in sys.path:
        sys.path.insert(0, source)
    return importlib.import_module(module_name)
