from __future__ import annotations

from pathlib import Path
from typing import Callable

from ..adapters.legacy_loader import import_original
from ..domain.contracts import ReportGenerationRequest


def generate(request: ReportGenerationRequest, on_progress: Callable[[int, str], None] | None = None) -> dict:
    """直接委派给旧 `extract_core.api.run_batch`，不改变匹配、抽取和回写规则。"""
    legacy_api = import_original("extract_core.api")
    result = legacy_api.run_batch(str(request.bom_path), [str(path) for path in request.report_paths], str(request.output_dir), request.round_no, on_progress=on_progress)
    outputs = sorted(str(path) for path in request.output_dir.rglob("*") if path.is_file())
    return {"legacy_result": result, "output_dir": str(request.output_dir), "outputs": outputs}
