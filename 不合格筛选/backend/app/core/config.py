from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[3]


@dataclass(frozen=True)
class Settings:
    project_root: Path = PROJECT_ROOT
    output_dir: Path = Path(os.getenv("APP_OUTPUT_DIR", str(PROJECT_ROOT / "output")))
    log_dir: Path = Path(os.getenv("APP_LOG_DIR", str(PROJECT_ROOT / "logs")))
    runtime_dir: Path = Path(os.getenv("APP_RUNTIME_DIR", str(PROJECT_ROOT / "backend" / ".runtime")))
    cors_origins: tuple[str, ...] = tuple(filter(None, os.getenv("APP_CORS_ORIGINS", "").split(",")))
    auth_enabled: bool = os.getenv("APP_AUTH_ENABLED", "false").lower() == "true"
    max_workers: int = 1  # Excel COM 自动化必须串行，保持原行为。


settings = Settings()
for _path in (settings.output_dir, settings.log_dir, settings.runtime_dir):
    _path.mkdir(parents=True, exist_ok=True)
