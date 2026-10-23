"""旧 `extract_core` 导入路径兼容入口，真实实现位于 legacy/original_src。"""
from pathlib import Path

_legacy_package = Path(__file__).resolve().parents[1] / "legacy" / "original_src" / "extract_core"
__path__ = [str(_legacy_package)]
__file__ = str(_legacy_package / "__init__.py")
exec(compile(__file__ and Path(__file__).read_text(encoding="utf-8"), __file__, "exec"), globals(), globals())
