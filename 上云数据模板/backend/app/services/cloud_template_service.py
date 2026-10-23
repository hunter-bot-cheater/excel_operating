from __future__ import annotations

from pathlib import Path
from typing import Callable

from ..adapters.legacy_loader import import_original
from ..domain.contracts import CloudTemplateRequest


def convert(request: CloudTemplateRequest, on_progress: Callable[[int, str], None] | None = None) -> dict:
    """调用原 `converter.convert_file`；模板字段、文件名和格式处理不在包装层改写。"""
    converter = import_original("converter")
    result = converter.convert_file(str(request.input_path), str(request.output_dir), on_progress=on_progress, default_year=request.default_year)
    return {"legacy_result": result, "output_dir": str(request.output_dir), "outputs": [str(Path(result["output"]))]}
