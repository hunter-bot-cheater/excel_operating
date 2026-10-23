from __future__ import annotations

import shutil
from pathlib import Path

from fastapi import APIRouter, File, HTTPException, UploadFile
from fastapi.responses import FileResponse

from ....core.config import settings
from ....domain.contracts import ScreeningRequest
from ....schemas.responses import envelope
from ....services.screening_service import screen
from ....services.task_registry import TaskRegistry


router = APIRouter(prefix="/nonconforming", tags=["nonconforming"])
tasks = TaskRegistry(settings.max_workers)
SUPPORTED_EXTENSIONS = {".xls", ".xlsx", ".xlsm"}


@router.post("/screen")
async def create_screening_task(files: list[UploadFile] = File(...)) -> dict:
    invalid = [item.filename for item in files if Path(item.filename or "").suffix.lower() not in SUPPORTED_EXTENSIONS]
    if invalid:
        raise HTTPException(422, detail=f"unsupported files: {invalid}")
    task_dir = settings.runtime_dir / "uploads"  # 任务 ID 在实际提交后确定，先以 uuid 目录替代由 registry 返回。
    # 每个请求使用独立 staging 目录，上传文件从不覆盖原始文件。
    import uuid
    staging = task_dir / uuid.uuid4().hex
    staging.mkdir(parents=True, exist_ok=False)
    input_paths: list[Path] = []
    for item in files:
        safe_name = Path(item.filename or "upload.xlsx").name
        destination = staging / safe_name
        with destination.open("wb") as handle:
            shutil.copyfileobj(item.file, handle)
        input_paths.append(destination)
    output_dir = staging / "outputs"
    output_dir.mkdir()
    submitted = tasks.submit(lambda progress: screen(ScreeningRequest(input_paths, output_dir), progress))
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
