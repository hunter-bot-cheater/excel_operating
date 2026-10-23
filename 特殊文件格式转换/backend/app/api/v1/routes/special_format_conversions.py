from __future__ import annotations

import shutil
import uuid
from pathlib import Path

from fastapi import APIRouter, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse

from ....core.config import settings
from ....domain.contracts import SpecialFormatConversionRequest
from ....schemas.responses import envelope
from ....services.special_format_service import convert
from ....services.task_registry import TaskRegistry

router = APIRouter(prefix="/special-format-conversions", tags=["special-format-conversions"])
tasks = TaskRegistry(settings.max_workers)
SUPPORTED_EXTENSIONS = {".xls", ".xlsx", ".xlsm"}


@router.post("/convert")
async def create_conversion_task(file: UploadFile = File(...), force_openpyxl: bool = Form(False)) -> dict:
    if Path(file.filename or "").suffix.lower() not in SUPPORTED_EXTENSIONS:
        raise HTTPException(422, detail="only .xls, .xlsx and .xlsm are supported")
    staging = settings.runtime_dir / "uploads" / uuid.uuid4().hex
    output_dir = staging / "outputs"
    output_dir.mkdir(parents=True)
    input_path = staging / Path(file.filename or "input.xlsx").name
    with input_path.open("wb") as handle:
        shutil.copyfileobj(file.file, handle)
    submitted = tasks.submit(lambda progress: convert(SpecialFormatConversionRequest(input_path, output_dir, force_openpyxl), progress))
    return envelope(submitted, submitted["trace_id"], "accepted")


@router.get("/tasks/{task_id}")
def get_task(task_id: str) -> dict:
    task = tasks.get(task_id)
    if task is None:
        raise HTTPException(404, detail="task not found")
    return envelope(task, task["trace_id"])


@router.get("/tasks/{task_id}/downloads/{filename}")
def download_output(task_id: str, filename: str):
    task = tasks.get(task_id)
    if task is None:
        raise HTTPException(404, detail="task not found")
    for output in (task.get("result") or {}).get("outputs", []):
        path = Path(output)
        if path.name == Path(filename).name and path.is_file():
            return FileResponse(path, filename=path.name)
    raise HTTPException(404, detail="output not found")
