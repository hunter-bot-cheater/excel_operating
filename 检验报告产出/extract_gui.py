"""旧 GUI 文件名兼容入口。"""
import runpy
import sys
from pathlib import Path

source = Path(__file__).resolve().parent / "legacy" / "original_src"
sys.path.insert(0, str(source))
runpy.run_path(str(source / "extract_gui.py"), run_name="__main__")
