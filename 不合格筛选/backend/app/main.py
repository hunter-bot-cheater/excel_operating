from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from .api.v1.routes.nonconforming import router
from .core.config import settings
from .schemas.responses import envelope

app = FastAPI(title="不合格筛选 API", version="0.1.0")
app.add_middleware(CORSMiddleware, allow_origins=list(settings.cors_origins), allow_credentials=False, allow_methods=["*"], allow_headers=["*"])
app.include_router(router, prefix="/api/v1")


@app.get("/api/v1/health")
def health() -> dict:
    return envelope({"status": "ok", "auth_enabled": settings.auth_enabled}, "health")
