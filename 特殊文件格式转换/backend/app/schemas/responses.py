from typing import Any


def envelope(data: Any, trace_id: str, message: str = "ok", code: int = 0) -> dict[str, Any]:
    return {"code": code, "message": message, "data": data, "trace_id": trace_id}
