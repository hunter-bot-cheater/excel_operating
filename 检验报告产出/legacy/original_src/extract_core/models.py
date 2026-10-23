# -*- coding: utf-8 -*-
"""数据模型。"""
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple


@dataclass
class ParseResult:
    """报告文件名解析结果。"""
    bh: str            # 图号（原样）
    desc: str          # 组件描述
    total: int         # 总件数
    date: str          # 日期（如 26.4.09）
    used: int          # 已用件数（缺省 0）
    raw_name: str      # 原始文件名


@dataclass
class BomRow:
    """BOM 中一行（匹配所需：图号、组件描述、数量）。

    图号列可能缺失（WYS 型 BOM，图号嵌在组件描述里），此时 bh 为空串。
    row/sheet 用于回写定位（Excel 1-based 行号 + 工作表名）。
    """
    bh: str = ""
    desc: str = ""
    qty: Optional[int] = None
    row: int = 0            # Excel 1-based 行号（回写用；0 表示未知）
    sheet: str = ""         # 所在工作表名（回写用）
    # 同一 (图号,描述,数量) 在其它的出现位置 [(sheet, row), ...]（不含上面的首个）。
    # 用于 BOM 含多张相同数据表时，回写「第N套」到所有副本，保持同步。
    locs: List[Tuple[str, int]] = field(default_factory=list)


@dataclass
class BomMeta:
    """BOM 表头结构（回写定位「第N套」列用；坐标均 1-based）。"""
    sheet: str = ""                          # 命中的工作表名
    header_row: int = 0                      # 表头行（1-based）
    bh_col: Optional[int] = None             # 图号列
    desc_col: Optional[int] = None           # 组件描述列
    qty_col: Optional[int] = None            # 组件数量列
    headers: Dict[int, str] = field(default_factory=dict)  # 列号 -> 表头文本


@dataclass
class Candidate:
    """候选来源报告文件（已解析文件名）。run_batch 期间实时更新 used/remaining。"""
    path: str
    bh: str
    total: int
    used: int
    remaining: int
    exhausted: bool
    parse: "ParseResult"


@dataclass
class SheetLayout:
    """单张工作表的布局探测结果（坐标均为 1-based）。"""
    name: str
    dwg: str                                   # 该表 DWG No（图号）
    seq_row: int                              # 测量序号所在行
    sample_cols: List[int]                    # 测量数据列（局部列号，按序）
    seq_numbers: List[int]                    # 与 sample_cols 一一对应的全局件号
    result_col: Optional[int]                 # 结果列
    max_col: Optional[int]                    # 最大值列
    min_col: Optional[int]                    # 最小值列
    total_cell: Optional[Tuple[int, int]]     # (行,列) 总数
    sample_cell: Optional[Tuple[int, int]]    # (行,列) 检查数
    data_first_row: int
    data_last_row: int
    part: str = ""                             # 该表零件名称（Part Description）
    # 测量列原始读数缓存：列号 -> [data_first_row..data_last_row 的值]
    col_values: Dict[int, list] = field(default_factory=dict)
    # 各测量列在源报告序号行上的原始文本（照抄用，可溯源）
    seq_labels: List[str] = field(default_factory=list)


@dataclass
class ReportInfo:
    path: str
    parse: ParseResult
    sheets: List[SheetLayout]
    # 全局件号 -> (工作表名, 局部列号)
    sample_map: Dict[int, Tuple[str, int]] = field(default_factory=dict)
    # 全局件号 -> 测量读数列表（与该表 data_first_row 起逐行对齐）——探测时一次性读出，
    # 供生成新报告时直接使用，避免逐格 COM 读取
    values: Dict[int, list] = field(default_factory=dict)
