from __future__ import annotations

from pathlib import Path
from typing import Callable

from ..adapters.legacy_loader import import_original
from ..domain.contracts import ScreeningRequest


def screen(request: ScreeningRequest, on_progress: Callable[[int, str], None] | None = None) -> dict:
    """调用原 `audit_paths`；判定、命名、格式处理均由复制的原代码负责。"""
    audit_fail = import_original("audit_fail")
    result = audit_fail.audit_paths([str(path) for path in request.input_paths], str(request.output_dir), on_progress=on_progress)
    outputs = sorted(str(path) for path in request.output_dir.glob("*") if path.is_file())
    return {"legacy_result": result, "output_dir": str(request.output_dir), "outputs": outputs}
