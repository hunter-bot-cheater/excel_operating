# -*- coding: utf-8 -*-
"""提取文件数据工具 —— 核心逻辑层（前后端分离，供 GUI / 网页端共用）。

对外统一入口见 api.py：
    from extract_core.api import run_batch, process_report
"""
from .api import run_batch, process_report

__all__ = ["run_batch", "process_report"]
