# -*- coding: utf-8 -*-
"""BOM 读取与「报告 → BOM 行」匹配。

兼容两种 BOM 形态（表头均在第一行）：
  A. PM-DM 型：有「图号」列 + 「组件描述」列 + 「组件数量」列。
  B. WYS  型：无图号列，「组件描述」列中嵌入图号（如「远端轴套_93126082_Pla,PTFE-V0」），
              数量列表头为「组件数量」。

列定位规则：
  - 图号列：表头含「图号」（可缺失）。
  - 组件描述列：表头为「组件描述」或「物料长描述」（可缺失；前者优先）。
  - 数量列：表头**必须精确为「组件数量」**（不模糊匹配「数量」，避免
    「组件计量单位数量」「基本数量」等误命中）；找不到则报 BOM_READ_ERROR，绝不猜测。

匹配（match_report，多信号交叉验证，前级命中即止）：
  报告携带三条身份信号（按可信度排序）：
    dwg   报告内部 DWG No（图号）——图纸真值，最强
    part  报告内部零件名称（Part Description）——图纸真值，强
    front 文件名「图号_描述」整体——用户命名，可能缩写/带前后缀/加版本号
  匹配原则：**联合信号优先于单一信号**（图号+描述同时对上，胜于只有图号对上）。
    S1 内部双键全等：canon(图号)==canon(dwg) 且 canon(描述)==canon(part)
    S2 文件名双键联合：①拼接全等 → 图号段×描述段联合（文本/数字匹配 + 描述命中）
    S3 内部单键：内部图号全等 + 零件名收敛；零件名对不上任何行时不采信图号
    S4 文件名单键：图号前缀 + 描述校验 → 分段集合（WYS 型）→ 数字键（≥5 位）
  任何一级命中多个**不等价**行（图号/描述/数量任一不同）时明确报 BOM_MULTI_MATCH，
  绝不瞎选。match_front 为无内部信号的兼容入口。

canon 归一化：去全部空白、全角→半角（括号/逗号/横线等）、转小写。
"""
import re
from typing import List, Optional, Tuple

import openpyxl

from .errors import ExtractError
from .models import BomMeta, BomRow

_FULL2HALF = {
    ord("（"): "(", ord("）"): ")", ord("，"): ",", ord("．"): ".",
    ord("－"): "-", ord("＿"): "_", ord("　"): "", ord("\u00a0"): "",
    ord("!"): "!", ord("？"): "?", ord("："): ":",
}

_QTY_HEADER = "组件数量"


def canon(s) -> str:
    """匹配归一化：去空白、全角→半角、转小写。"""
    if s is None:
        return ""
    s = str(s).translate(_FULL2HALF)
    s = re.sub(r"\s+", "", s)
    return s.lower()


# 数字键最少位数：太短（如标准件的 GB/T 119.1 → 1191）误撞概率高，不参与
_DIGITS_MIN = 5


def digits_key(s) -> str:
    """把字符串里的全部数字拼成一个键，用于「报告只写了图号的一部分」的匹配。

    例：`JH-YCL-93010410（按图加工）` -> `93010410`；`93010410` -> `93010410`；
        `9301X01XM4` -> `9301014`（两侧同样处理，故仍可相等比对）。
    """
    return "".join(re.findall(r"\d+", canon(s)))


def _norm(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def _is_int(v):
    if isinstance(v, bool):
        return False
    if isinstance(v, int):
        return True
    if isinstance(v, float):
        return v.is_integer()
    if isinstance(v, str):
        s = v.strip()
        return s.lstrip("-").isdigit()
    return False


def _to_int(v):
    if isinstance(v, float):
        return int(v)
    return int(str(v).strip())


# --------------------------------------------------------------------------- #
# 读取
# --------------------------------------------------------------------------- #
_MAX_R = 2000   # 最多扫描行数
_MAX_C = 60     # 最多扫描列数（WYS 有 51 列）


def _cell(grid, r, c):
    """0-based 取值，越界返回 None。"""
    if 0 <= r < len(grid) and 0 <= c < len(grid[r]):
        return grid[r][c]
    return None


def _locate_cols(grid):
    """返回 (hdr_row, bh_col, desc_col, qty_col)，均 0-based，可为 None。"""
    hdr_row = bh_col = desc_col = qty_col = None
    # 1) 表头行：前 12 行内出现 图号/组件描述/组件数量 任一表头
    for r in range(min(len(grid), 12)):
        for c in range(len(grid[r])):
            h = canon(grid[r][c])
            if h in ("图号", "组件描述", "组件数量"):
                hdr_row = r
                break
        if hdr_row is not None:
            break
    if hdr_row is None:
        return None, None, None, None

    # 2) 图号列 / 组件描述列（物料长描述同样可作描述列；组件描述优先）
    for c in range(len(grid[hdr_row])):
        h = canon(_cell(grid, hdr_row, c))
        if bh_col is None and "图号" in h:
            bh_col = c
        if desc_col is None and h == "组件描述":
            desc_col = c
    if desc_col is None:
        for c in range(len(grid[hdr_row])):
            if canon(_cell(grid, hdr_row, c)) == "物料长描述":
                desc_col = c
                break

    # 3) 数量列：**只认表头精确为「组件数量」的列**；否则视为未找到（由调用方报错，绝不降级适配）
    for c in range(len(grid[hdr_row])):
        if canon(_cell(grid, hdr_row, c)) == _QTY_HEADER:
            qty_col = c
            break
    return hdr_row, bh_col, desc_col, qty_col


def read_bom_ex(path: str) -> Tuple[List[BomRow], BomMeta]:
    """读取 BOM，返回 (rows, meta)。

    rows：图号/描述/数量 + Excel 1-based 行号(row) + 工作表名(sheet)；
    meta：回写定位「第N套」列所需的表头结构（取首个有效工作表）。
    """
    try:
        # 不用 read_only：普通模式才能读 row_dimensions.hidden，
        # 从而识别「被 Excel 自动筛选隐藏」的行（用户只按可见行匹配/回写）。
        wb = openpyxl.load_workbook(path, data_only=True)
    except Exception as e:  # noqa: BLE001
        raise ExtractError("BOM_READ_ERROR", f"BOM 文件读取失败：{e}", file=path)

    rows: List[BomRow] = []
    meta = BomMeta()
    missing_qty_reason = None
    try:
        for ws in wb.worksheets:
            # 一次性流式读入（read_only 下 ws.cell() 随机访问极慢，禁用）
            grid: List[tuple] = []
            for row in ws.iter_rows(min_row=1, max_row=_MAX_R, min_col=1,
                                    max_col=_MAX_C, values_only=True):
                grid.append(row)
            if not grid:
                continue

            hdr_row, bh_col, desc_col, qty_col = _locate_cols(grid)
            if hdr_row is None or (bh_col is None and desc_col is None):
                continue
            if qty_col is None:
                # 找到了图号/组件描述，但表头没有「组件数量」列 -> 明确报错（不猜、不模糊匹配）
                hdrs = "、".join(_norm(_cell(grid, hdr_row, c))
                                for c in range(len(grid[hdr_row]))
                                if canon(_cell(grid, hdr_row, c)))
                missing_qty_reason = (
                    "BOM 未找到数量列：表头第一行没有【组件数量】列。"
                    "本工具只认精确的「组件数量」四个字，不做模糊匹配、也不做内容推断降级——"
                    "请检查 BOM 表头（数量列必须命名为「组件数量」）。"
                    f"（本表表头为：{hdrs}）"
                )
                continue

            # 首个有效表的表头结构作为回写定位依据
            if not meta.headers:
                meta.sheet = ws.title
                meta.header_row = hdr_row + 1
                meta.bh_col = (bh_col + 1) if bh_col is not None else None
                meta.desc_col = (desc_col + 1) if desc_col is not None else None
                meta.qty_col = qty_col + 1
                for c in range(len(grid[hdr_row])):
                    t = _norm(_cell(grid, hdr_row, c))
                    if t:
                        meta.headers[c + 1] = t

            # 收集本表「被隐藏」的行号（1-based）——Excel 自动筛选 / 手动行隐藏都会置 hidden。
            # 用户按筛选后可见行匹配，被隐藏的行不参与匹配/建任务/回写。
            hidden_rows = set()
            try:
                for rd in ws.row_dimensions.values():
                    if getattr(rd, "hidden", False):
                        hidden_rows.add(int(rd.index))
            except Exception:  # noqa: BLE001
                hidden_rows = set()

            # **只认第一张有效数据表**：表头有效且数量列存在、且贡献了至少一行数据的
            # 第一张表，就是唯一的匹配/建任务/回写真相源；后续所有表一律忽略。
            n_before = len(rows)
            for r in range(hdr_row + 1, len(grid)):
                if (r + 1) in hidden_rows:
                    continue  # 被 Excel 筛选隐藏的行不参与
                bh = _norm(_cell(grid, r, bh_col)) if bh_col is not None else ""
                desc = _norm(_cell(grid, r, desc_col)) if desc_col is not None else ""
                if not bh and not desc:
                    continue
                qty_raw = _cell(grid, r, qty_col)
                if qty_raw is None or str(qty_raw).strip() == "":
                    continue  # 数量为空 -> 不参与
                if not _is_int(qty_raw):
                    continue
                rows.append(BomRow(bh=bh, desc=desc, qty=_to_int(qty_raw),
                                   row=r + 1, sheet=ws.title))
            if len(rows) > n_before:
                break  # 第一张有效数据表已处理完毕，忽略其余表单
    finally:
        wb.close()

    if not rows:
        if missing_qty_reason:
            raise ExtractError("BOM_READ_ERROR", missing_qty_reason, file=path)
        raise ExtractError(
            "BOM_READ_ERROR",
            "BOM 中未解析到任何「图号/组件描述 + 组件数量」行"
            "（需表头含 图号 或 组件描述，且表头第一行有精确的「组件数量」列名）",
            file=path,
        )
    return rows, meta


def read_bom(path: str) -> List[BomRow]:
    """读取 BOM，返回含 (图号, 组件描述, 数量) 的行列表（键与数量均非空者）。"""
    return read_bom_ex(path)[0]


# --------------------------------------------------------------------------- #
# 回写定位：「第N套」列 / 插入位置
# --------------------------------------------------------------------------- #
_ROUND_RE = re.compile(r"^第\s*(\d+|[一二两三四五六七八九十]+)\s*套$")
# 「第X套」列右侧的无关列（图纸/材质报告/尺寸报告等）——新列须插在它们左边
_IRRELEVANT_KWS = ("图纸", "材质报告", "尺寸报告")

_CN_DIGIT = {"一": 1, "二": 2, "两": 2, "三": 3, "四": 4,
             "五": 5, "六": 6, "七": 7, "八": 8, "九": 9}


def _cn_to_int(s: str) -> Optional[int]:
    """把「三/十一/二十/二十三」等中文数字转为阿拉伯整数（1~99）。"""
    if not s:
        return None
    if "十" in s:
        left, _, right = s.partition("十")
        tens = _CN_DIGIT.get(left, 1) if left else 1   # 「十X」= 一十几
        ones = _CN_DIGIT.get(right, 0) if right else 0
        return tens * 10 + ones
    if len(s) == 2:   # 「二十」「三十一」之外的两位（如 两三 不规范，按十位+个位）
        t = _CN_DIGIT.get(s[0])
        o = _CN_DIGIT.get(s[1])
        if t is not None and o is not None:
            return t * 10 + o
        return None
    return _CN_DIGIT.get(s)


def parse_round_header(header) -> Optional[int]:
    """表头若为「第N套」返回 N（阿拉伯/中文数字均可，如「第3套」「第三套」），否则 None。"""
    m = _ROUND_RE.match(canon(header))
    if not m:
        return None
    g = m.group(1)
    if g.isdigit():
        try:
            return int(g)
        except ValueError:
            return None
    return _cn_to_int(g)


def find_round_col(meta: BomMeta, n: int) -> Optional[int]:
    """返回 BOM 中「第n套」列的 1-based 列号；不存在返回 None。"""
    for c in sorted(meta.headers):
        if parse_round_header(meta.headers[c]) == n:
            return c
    return None


def find_insert_col(meta: BomMeta, n: int) -> int:
    """计算插入「第n套」列的 1-based 位置。

    规则：插在所有「第X套」列的最右边；若无「第X套」列，则插在
    「图纸/材质报告/尺寸报告」等无关列的最左边。
    """
    round_cols = [c for c, h in meta.headers.items() if parse_round_header(h) is not None]
    other_cols = [c for c, h in meta.headers.items()
                  if any(k in canon(h) for k in _IRRELEVANT_KWS)]
    if round_cols:
        return max(round_cols) + 1
    if other_cols:
        return min(other_cols)
    if meta.qty_col:
        return meta.qty_col + 1
    return (max(meta.headers) + 1) if meta.headers else 1



# --------------------------------------------------------------------------- #
# 匹配（报告 -> BOM 行，多信号交叉验证）
# --------------------------------------------------------------------------- #
def _substr(a: str, b: str, min_len: int) -> bool:
    """a 与 b 互为连续子串（任一包含另一），且较短侧长度 >= min_len，防短码误撞。"""
    if not a or not b:
        return False
    if min(len(a), len(b)) < min_len:
        return False
    return a in b or b in a


def _multi_error(front: str, rows: List[BomRow], level: str) -> ExtractError:
    detail = "；".join(f"「{r.bh}／{r.desc}」×{r.qty}（第{r.row}行）" for r in rows)
    return ExtractError(
        "BOM_MULTI_MATCH",
        f"报告「{front}」按{level}匹配到多个不同 BOM 行，存在歧义，无法确定是哪一行：{detail}"
        f"（请在报告文件名/内部图号里写全，或在 BOM 中把这些行区分开）",
    )


def match_report(front: str, rows: List[BomRow], dwg: str = "", part: str = "") -> List[BomRow]:
    """报告 -> BOM 行匹配（多信号交叉验证；前级命中即止）。

    信号（按可信度排序）：
      dwg   报告内部 DWG No（图号）
      part  报告内部零件名称（Part Description）
      front 文件名「图号_描述」整体（用户命名，可能缩写/带前后缀/加版本号）

    匹配原则：**联合信号优先于单一信号**（图号+描述同时对上，胜于只有图号对上）。
      S1 内部双键全等：canon(图号)==canon(dwg) 且 canon(描述)==canon(part)
      S2 文件名双键联合：
         ① 拼接全等（原第 1 级）→ L2 图号段×描述段联合（文本/数字匹配 + 描述命中）
      S3 内部单键：内部图号全等 + 零件名收敛（互为子串）；
         零件名对不上任何行、或文件名联合已指向别的行时不采信图号
      S4 文件名单键：L3 图号前缀 + 描述校验 → L4 分段集合（WYS 型，原③）→
         L5 数字键（≥5 位，原④）

    **唯一行规则（用户口径 v6.2）**：BOM 只取第一张有效数据表，且该表内**图号唯一**。
    任何一级只要命中**多于一行**（无论描述/数量是否相同，含完全相同的行）-> 直接报
    `BOM_MULTI_MATCH` 并列出候选行，**绝不拆多任务、绝不静默选一行**；
    仅当「描述精确收敛到唯一一行」时才放行。全部未命中返回 []。
    """
    cf = canon(front)
    if not cf:
        return []
    cdwg, cpart = canon(dwg), canon(part)

    # ---- S1: 内部双键全等（最强信号；两份内部真值互相印证，直接采信）----
    if cdwg and cpart:
        hits = [r for r in rows
                if r.bh and r.desc and canon(r.bh) == cdwg and canon(r.desc) == cpart]
        if hits:
            if len(hits) == 1:
                return hits
            raise _multi_error(front, hits, "内部图号+零件名精确匹配")

    # ---- S1b: 内部双信号「双向互为子串」（P0 放宽级）----
    # 报告内部 dwg 常带前后缀/注解（如 JH-YCL-93019282-DM(按图加工)），BOM 图号只是其中一段
    # 93019282-DM。要求两边都成立才命中：
    #   图号：canon(bh) 与 canon(dwg) 互为连续子串；且零件名 canon(desc) 与 canon(part) 互为连续子串。
    # 两侧都卡最短长度，防短码/单字误撞；命中多行仍按唯一行规则报错，绝不挑行。
    if cdwg and cpart:
        hits = [r for r in rows
                if r.bh and r.desc
                and _substr(canon(r.bh), cdwg, min_len=5)
                and _substr(canon(r.desc), cpart, min_len=2)]
        if hits:
            if len(hits) == 1:
                return hits
            raise _multi_error(front, hits, "内部图号+零件名双向包含")

    # ---- S2a: ① 拼接全等（原第 1 级）----
    hits = [r for r in rows
            if r.bh and canon(f"{r.bh}_{r.desc}" if r.desc else r.bh) == cf]
    if hits:
        if len(hits) == 1:
            return hits
        raise _multi_error(front, hits, "①拼接全等")

    # ---- S2b: 文件名双键联合（图号段 × 描述段）----
    segs = [t for t in cf.split("_") if t]
    if len(segs) >= 2:
        bh_seg, desc_seg = segs[0], "_".join(segs[1:])
        fkey = digits_key(bh_seg)
        cands = []
        for r in rows:
            if not r.bh or not r.desc:
                continue
            cbh, cdesc = canon(r.bh), canon(r.desc)
            bh_ok = (cbh == bh_seg or bh_seg.startswith(cbh)
                     or (len(fkey) >= _DIGITS_MIN and digits_key(r.bh) == fkey))
            if not bh_ok:
                continue
            # 描述命中分两档：**精确全等 > 子串包含**（防「X」被「X1」行干扰）
            desc_exact = bool(desc_seg) and (cdesc == desc_seg or desc_seg == cdesc)
            desc_contain = (cdesc in cf) or (desc_seg and desc_seg in cdesc)
            if desc_exact:
                cands.append((r, True))
            elif desc_contain:
                cands.append((r, False))
        if cands:
            exact = [r for r, ex in cands if ex]
            pool = exact or [r for r, _ex in cands]
            if len(pool) == 1:
                return pool
            raise _multi_error(front, pool, "「图号×描述」联合匹配")

    # ---- S3: 内部单键（图号全等 + 零件名收敛）----
    # 只有文件名联合信号无结果时才采信内部图号；零件名对不上任何行（信号冲突）
    # 或多行同图号时继续向下（最终由 S4 唯一行规则裁定：多行 -> 报错）。
    if cdwg:
        hits = [r for r in rows if r.bh and canon(r.bh) == cdwg]
        if hits:
            if cpart:
                narrowed = [r for r in hits
                            if r.desc and (cpart in canon(r.desc) or canon(r.desc) in cpart)]
                if narrowed:
                    if len(narrowed) == 1:
                        return narrowed
                    raise _multi_error(front, narrowed, "内部图号匹配+零件名收敛")
                # 零件名对不上任何行：内部图号与零件名互相矛盾 -> 不采信图号，继续向下
                # （防 93019282-DM 图号 + 翻转横撑 误锁「弹簧压板」行）
            elif len(hits) == 1:
                # 无零件名、图号唯一命中：此时文件名联合已无结果 -> 采信内部图号
                return hits
            # 多行同图号、无零件名可收敛 -> 由 S4 唯一行规则裁定

    # ---- S4a: 图号前缀 + 描述校验（原②；多行同图号 -> 直接报错）----
    # 用户口径：BOM 第一张有效数据表内图号唯一。同图号多行视为数据异常，
    # 唯一放行口是「描述精确收敛到唯一一行」；否则一律 BOM_MULTI_MATCH 报错。
    pref = []
    for r in rows:
        if r.bh:
            cbh = canon(r.bh)
            if cf == cbh or cf.startswith(cbh + "_"):
                pref.append((cbh, r))
    if pref:
        longest = max(len(cbh) for cbh, _r in pref)
        selected = [r for cbh, r in pref if len(cbh) == longest]
        if len(selected) == 1:
            return selected  # 图号唯一确定 -> 直接命中
        longest_cbh = max((cbh for cbh, _r in pref if len(cbh) == longest), key=len)
        desc_part = cf[len(longest_cbh):].strip("_")
        # 描述收敛：**精确全等优先于子串包含**（防「保护盖」被「保护盖1」行干扰）
        narrowed_exact = [r for r in selected if r.desc and canon(r.desc) == desc_part]
        narrowed = narrowed_exact or [r for r in selected if r.desc and canon(r.desc) in cf]
        if len(narrowed) == 1:
            return narrowed
        raise _multi_error(front, selected, "图号前缀匹配（BOM 第一张有效数据表中该图号存在多行）")

    # ---- S4b: 分段集合（原第 3 级，WYS 型）----
    ftoks = [t for t in cf.split("_") if t]
    if ftoks:
        hits = []
        for r in rows:
            if not r.desc:
                continue
            dtoks = {t for t in canon(r.desc).split("_") if t}
            if dtoks and all(t in dtoks for t in ftoks):
                hits.append(r)
        if hits:
            if len(hits) == 1:
                return hits
            raise _multi_error(front, hits, "分段集合匹配")

    # ---- S4c: 数字键（原第 4 级；仅当前面全部未命中才启用）----
    # 图号可能藏在文件名任意段里（如 `91010070_E.1_20260814_Pd93010410_保护盖1`，
    # 93010410 在第四段），故对**每一段**取数字键，与 BOM 图号比对。
    segs_all = [t for t in cf.split("_") if t]
    fkeys = {digits_key(t) for t in segs_all}
    fkeys = {k for k in fkeys if len(k) >= _DIGITS_MIN}
    if fkeys:
        # WYS 型 BOM：图号列可能为空，真正的图号写在「描述」里（如 描述=Pd主腔体_91010070）。
        # 故行图号取「图号列数字键 或 描述列数字键」，二者任一落在文件名数字键里即为候选。
        def _row_dig(r):
            return digits_key(r.bh) or digits_key(r.desc)
        cands = [r for r in rows
                 if _row_dig(r) and _row_dig(r) in fkeys]
        if cands:
            # 用组件描述收敛：**精确全等优先**，其次「描述出现在文件名 front 里」
            desc_exact = [r for r in cands if r.desc and canon(r.desc) ==
                          "_".join(t for t in cf.split("_")[1:])]
            narrowed = desc_exact or [r for r in cands if r.desc and canon(r.desc) in cf]
            pool = narrowed or cands
            if len(pool) == 1:
                return pool
            # 再用「图号文本本身出现在 front 里」收敛
            exact = [r for r in pool if canon(r.bh) and canon(r.bh) in cf]
            if len(exact) == 1:
                return exact
            if exact:
                pool = exact
            raise _multi_error(
                front, pool,
                f"图号数字 {sorted(fkeys)} 匹配（BOM 第一张有效数据表中该图号存在多行）")

    return []


def match_front(front: str, rows: List[BomRow]) -> List[BomRow]:
    """兼容入口：仅用文件名 front 匹配（不携带报告内部 DWG No / 零件名称）。"""
    return match_report(front, rows)


def match_qty(bh: str, bom_rows: List[BomRow]) -> int:
    """兼容旧接口：等价于 match_front 的单数量版本（0/多命中报错）。"""
    hits = match_front(bh, bom_rows)
    if not hits:
        raise ExtractError("BOM_NO_MATCH", f"「{bh}」在 BOM 中未找到匹配行")
    if len(hits) > 1:
        raise ExtractError(
            "BOM_MULTI_MATCH",
            f"「{bh}」在 BOM 中匹配到 {len(hits)} 行，存在歧义",
        )
    return hits[0].qty
