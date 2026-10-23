from __future__ import annotations

from pathlib import Path
from typing import Callable

from ..adapters.legacy_loader import import_original
from ..domain.contracts import SpecialFormatConversionRequest


def convert(request: SpecialFormatConversionRequest, on_progress: Callable[[int, str], None] | None = None) -> dict:
    """原样委派给旧 `convert_file`，保留模板、判定、命名和 COM 兜底规则。"""
    converter = import_original("converter")
    result = converter.convert_file(str(request.input_path), str(request.output_dir), on_progress=on_progress, force_openpyxl=request.force_openpyxl)
    return {"legacy_result": result, "output_dir": str(request.output_dir), "outputs": [str(Path(result["output"]))]}
