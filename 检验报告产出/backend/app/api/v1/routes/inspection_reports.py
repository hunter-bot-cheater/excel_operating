from __future__ import annotations

import shutil
import uuid
from pathlib import Path

from fastapi import APIRouter, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse

from ....core.config import settings
from ....domain.contracts import ReportGenerationRequest
from ....schemas.responses import envelope
from ....services.report_generation_service import generate
from ....services.task_registry import TaskRegistry

router = APIRouter(prefix="/inspection-reports", tags=["inspection-reports"])
tasks = TaskRegistry(settings.max_workers)
SUPPORTED_EXTENSIONS = {".xls", ".xlsx", ".xlsm"}


@router.post("/generate")
async def create_generation_task(bom: UploadFile = File(...), reports: list[UploadFile] = File(...), round_no: int | None = Form(None)) -> dict:
    all_files = [bom, *reports]
    invalid = [item.filename for item in all_files if Path(item.filename or "").suffix.lower() not in SUPPORTED_EXTENSIONS]
    if invalid:
        raise HTTPException(422, detail=f"unsupported files: {invalid}")
    staging = settings.runtime_dir / "uploads" / uuid.uuid4().hex
    reports_dir, output_dir = staging / "reports", staging / "outputs"
    reports_dir.mkdir(parents=True)
    output_dir.mkdir()
    bom_path = staging / Path(bom.filename or "bom.xlsx").name
    with bom_path.open("wb") as handle:
        shutil.copyfileobj(bom.file, handle)
    report_paths: list[Path] = []
    for item in reports:
        destination = reports_dir / Path(item.filename or "report.xlsx").name
        with destination.open("wb") as handle:
            shutil.copyfileobj(item.file, handle)
        report_paths.append(destination)
    submitted = tasks.submit(lambda progress: generate(ReportGenerationRequest(bom_path, report_paths, output_dir, round_no), progress))
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
