import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "backend"))
import uvicorn

uvicorn.run("app.main:app", host="127.0.0.1", port=8000, reload=False, workers=1)
