# -*- coding: utf-8 -*-
"""统一的错误类型：每个错误带 error_code（前端可映射）与中文 reason（一眼可知原因）。"""
from typing import Optional


class ExtractError(Exception):
    """提取流程中的可控错误。

    code  : 机器可读错误码（FILE_PARSE_ERROR / BOM_NO_MATCH / K_TOO_LARGE ...）
    reason: 给人看的中文原因，应一眼可知问题所在。
    """

    def __init__(self, code: str, reason: str, *, file: Optional[str] = None):
        self.code = code
        self.reason = reason
        self.file = file
        super().__init__(reason)

    def __str__(self):
        if self.file:
            return f"[{self.code}] {self.reason} ｜ 文件：{self.file}"
        return f"[{self.code}] {self.reason}"
