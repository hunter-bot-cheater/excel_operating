"""旧核心模块兼容入口；实际实现保持在 legacy/original_src/audit_fail.py。"""
import importlib.util
from pathlib import Path

_path = Path(__file__).resolve().parent / "legacy" / "original_src" / "audit_fail.py"
_spec = importlib.util.spec_from_file_location("_legacy_audit_fail", _path)
_module = importlib.util.module_from_spec(_spec)
assert _spec and _spec.loader
_spec.loader.exec_module(_module)
globals().update({name: value for name, value in vars(_module).items() if not name.startswith("__")})
