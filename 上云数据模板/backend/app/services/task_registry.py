from __future__ import annotations

import threading
import traceback
import uuid
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timezone
from typing import Any, Callable


class TaskRegistry:
    def __init__(self, max_workers: int = 1):
        self._tasks: dict[str, dict[str, Any]] = {}
        self._lock = threading.Lock()
        self._pool = ThreadPoolExecutor(max_workers=max_workers, thread_name_prefix="cloud-template")

    def submit(self, work: Callable[[Callable[[int, str], None]], dict[str, Any]]) -> dict[str, str]:
        task_id, trace_id = uuid.uuid4().hex, uuid.uuid4().hex
        with self._lock:
            self._tasks[task_id] = {"task_id": task_id, "trace_id": trace_id, "status": "pending", "progress": 0, "message": "queued", "created_at": datetime.now(timezone.utc).isoformat(), "result": None}
        def run() -> None:
            def progress(value: int, message: str) -> None:
                with self._lock:
                    self._tasks[task_id].update(progress=max(0, min(100, int(value))), message=message)
            with self._lock:
                self._tasks[task_id].update(status="running", message="running")
            try:
                result = work(progress)
                with self._lock:
                    self._tasks[task_id].update(status="succeeded", progress=100, message="completed", result=result)
            except Exception as exc:
                with self._lock:
                    self._tasks[task_id].update(status="failed", message=str(exc), error_type=type(exc).__name__, traceback=traceback.format_exc())
        self._pool.submit(run)
        return {"task_id": task_id, "trace_id": trace_id}

    def get(self, task_id: str) -> dict[str, Any] | None:
        with self._lock:
            task = self._tasks.get(task_id)
            return dict(task) if task else None
