"""旧核心模块兼容入口；真实实现保留在 legacy/original_src/converter.py。"""
import importlib.util
from pathlib import Path

_path = Path(__file__).resolve().parent / "legacy" / "original_src" / "converter.py"
_spec = importlib.util.spec_from_file_location("_legacy_converter", _path)
_module = importlib.util.module_from_spec(_spec)
assert _spec and _spec.loader
_spec.loader.exec_module(_module)
globals().update({name: value for name, value in vars(_module).items() if not name.startswith("__")})
