# -*- coding: utf-8 -*-
"""报告文件名解析。

格式：图号_组件描述_总件数件_日期-已用(微信后缀).xlsx
  - 「图号」可能本身含 - 或 _（甚至多个 -，如 JH10-93021NDB-DM、93021WKD_prt、
    93010698-（1）），也可能没有分隔符（纯数字如 93126082）。
  - 因此**不再拆分图号与描述**：front = `_N件_` 之前的全部（图号_描述整体），
    后续匹配以 front 整体与 BOM 的「图号_组件描述」拼接键比对。
例：93019282-DM_翻转横撑_21件_26.4.09-12(1).xlsx
  -> bh(front)=93019282-DM_翻转横撑, total=21, date=26.4.09, used=12
"""
import os
import re

from .errors import ExtractError

_WECHAT = re.compile(r"\s*\(\d+\)\s*$")  # 微信 "(1)" 后缀
# front_总件数件_日期(-已用)?
# 已用序号后缀四种写法都认：-3 / -出3 / -出3件 / -已出3件（均表示已用到第3件）
_NAME_RE = re.compile(
    r"^(?P<front>.+)_(?P<total>\d+)件_(?P<date>[\d.]+)"
    r"(?:-(?:已出|出)?(?P<used>\d+)件?)?$"
)


def parse_report_filename(path: str) -> "ParseResult":
    if not path or not isinstance(path, str):
        raise ExtractError("FILE_PARSE_ERROR", "文件路径为空")
    base = os.path.basename(path)
    name, ext = os.path.splitext(base)
    if ext.lower() not in (".xlsx", ".xls", ".xlsm"):
        raise ExtractError(
            "FILE_PARSE_ERROR",
            f"不支持的扩展名：{ext or '（无）'}（仅支持 Excel 的 .xls/.xlsx/.xlsm；"
            f".doc/.docx 等 Word 文件或非表格文件不在处理范围）",
            file=base,
        )
    # 去掉微信 "(1)" 后缀
    name = _WECHAT.sub("", name).strip()
    if not name:
        raise ExtractError("FILE_PARSE_ERROR", f"文件名解析后为空：{base}", file=base)

    m = _NAME_RE.match(name)
    if not m:
        raise ExtractError(
            "FILE_PARSE_ERROR",
            f"文件名格式无法识别：{base}（应为 图号_描述_N件_日期-已用；"
            f"若为「出货报告清单」等不含 _N件_ 命名的文件，本工具暂不支持，请改名或换标准测量报告）",
            file=base,
        )
    front = m.group("front").strip()  # 图号_组件描述 整体
    if not front:
        raise ExtractError("FILE_PARSE_ERROR", f"图号解析为空：{base}", file=base)
    try:
        total = int(m.group("total"))
    except ValueError:
        raise ExtractError("FILE_PARSE_ERROR", f"总件数不是数字：{base}", file=base)
    date = m.group("date").strip()
    if m.group("used") is not None:
        used = int(m.group("used"))
        if used == 0:
            # 后缀 "-0" 表示已用完：已用量 = 总件数（剩余 0）
            used = total
    else:
        used = 0  # 无后缀 = 全新未用

    from .models import ParseResult
    return ParseResult(bh=front, desc="", total=total, date=date, used=used, raw_name=base)
