"""旧 FastAPI 入口兼容文件；原 API 实现保持在 legacy/original_src/api.py。"""
import importlib.util
import runpy
import sys
from pathlib import Path

_source = Path(__file__).resolve().parent / "legacy" / "original_src"

if __name__ == "__main__":
    sys.path.insert(0, str(_source))
    runpy.run_path(str(_source / "api.py"), run_name="__main__")
else:
    _spec = importlib.util.spec_from_file_location("_legacy_audit_api", _source / "api.py")
    _module = importlib.util.module_from_spec(_spec)
    assert _spec and _spec.loader
    sys.path.insert(0, str(_source))
    _spec.loader.exec_module(_module)
    globals().update({name: value for name, value in vars(_module).items() if not name.startswith("__")})
