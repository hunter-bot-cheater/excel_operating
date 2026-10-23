#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
审核不合格数据（按数据反推 / v3 行驱动版）
================================

设计目标
--------
针对检验报告（.xls / .xlsx / .xlsm）：
  * 按报告内“自身公式 / 结果”判定的合格状态，命中不合格则：
      1) 在【原文件名最前面】加 `不合格_` 前缀（命名体现）。
      2) 在报告内部对**同一行的【结果列单元格】** 标红 + 标 F：
         - 实测值超出 [下限, 上限] 时：把该行结果列单元格标红（亮红底色，字体不变），
           若该单元格当前是 P / 合格 / 空（且不是公式），自动改为 F。
         - 结果列本身已是不合格（不合格 / NG / Fail / F 等）时：标红，保持原值/公式。
         - **测量列单元格保持原样**，不追加 F、不标红，便于复检核对原始读数。
  * **扫描所有工作表**：从 v3.1 起，程序会遍历文件里的全部工作表，不再只读 Sheet1。
    多张工作表合并判定：任一表有越界即判不合格，标红也只标在对应表的对应单元格。
  * 其余格式（字体、列宽、行高、合并单元格、图片等）一律原样保留。
  * 合格文件：逐字节复制（shutil.copy2），零格式化风险。

判定口径（v3：只认上下限标签列，结构定位）
----------------------------------
1. 用 Excel COM 打开（**临时副本 + 强制重算**），保证 .Value 拿到的是
   公式最新计算结果。消除“同一份文件在不同机器/Excel 状态被判出不同结果”的问题。
2. **只靠标签定位角色列**（标签在表内任意行匹配，列号随报告自动找，不写死）：
   - 上限列：某列含 上限值/上限/max/ucl/upper 等标签，且该列有数值；
   - 下限列：含 下限值/下限/min/lcl/lower 等标签，且该列有数值；
   - 结果列：含 result/结果/判定/结论 等标签（或整列多为 P/F 判定词）。
3. **测量列 = 公差列右侧、结果列左侧之间的列**（结构区间）。
   **不做“数值占比≥50%”的投票**：数据区若大部分是 OK/目视文字，占比法会把
   测量列剔除；这里无论列内数值多少都参与检查，非数值单元格自动忽略。
4. **逐行校验**：该行上限列 / 下限列的值 = 该行公差 [mn, mx]，写反自动交换；
   测量列数值落在区间外 → 不合格，结果列标红（P/合格/空且非公式时改 F）。
5. 有公差行只看越界，结果列文本（含公式算出的 P/F）不参与判不合格，
   避免不同机器 Excel 公式重算状态差异导致误判。
6. 找不到上下限标签列时无法做数值校验，仅保留结果列文本兜底。

命名规则
--------
  不合格： out = <目录>/不合格_<原名>.<ext>
  合格  ： out = <目录>/<原名>.<ext>
"""

import os
import re
import shutil
import time

try:
    import win32com.client
except Exception:
    win32com = None


# ---- 颜色（OLE 颜色：低位字节为 R）----
# 亮红底色（255,0,0）-> BGR 0x0000FF
RED_FILL = 0x0000FF

# SaveAs 文件格式
_EXT_FMT = {".xlsx": 51, ".xlsm": 52, ".xls": 56}

# ---- Excel 错误值编码 ----
# COM 读取含错误（#VALUE!/#NAME?/#DIV/0! 等）的单元格时，返回的是这些负整数而非字符串。
# 不拦截会被当成超大负数参与越界比较 → 误判不合格。
# 常见触发：表格用了 LET/TEXTJOIN/XLOOKUP 等新函数，而打开端 Excel 版本较旧（返回 #NAME?）。
_XL_ERROR_CODES = frozenset({
    -2146826281,  # #DIV/0!
    -2146826246,  # #N/A
    -2146826259,  # #NAME?
    -2146826288,  # #NULL!
    -2146826252,  # #NUM!
    -2146826265,  # #REF!
    -2146826273,  # #VALUE!
})

# ---- 合格 / 不合格 词表 ----
_FAIL_WORDS = ("不合格", "ng", "n.g", "fail", "reject", "拒收", "超差", "超上限", "超下限", "f")
_PASS_WORDS = ("合格", "ok", "pass", "p", "good", "accept", "接受", "是", "y", "yes")

# 结果列定位关键字（不依赖表头行，可在表内任意行匹配）
_RESULT_HEADERS = ("result", "结果", "判定", "结论")


def _norm(v):
    """转小写字符串；None/空返回 ''。"""
    if v is None:
        return ""
    return str(v).strip().lower()


def _strip_symbols(t):
    """去复选框/标点，仅留字母与中日韩文字，便于精确匹配。"""
    return re.sub(r"[^a-z\u4e00-\u9fff]", "", t)


def _tokens(v):
    """把单元格文本切成词元。"""
    t = _norm(v)
    return re.findall(r"[a-z]+|\d+(?:\.\d+)?|[\u4e00-\u9fff]+", t)


def _is_fail_text(v, strict=False):
    """单元格文本是否表达“不合格”。"""
    t = _norm(v)
    if not t:
        return False
    t2 = _strip_symbols(t)
    if t2 in ("不合格", "ng", "fail", "reject", "nogo"):
        return True
    if not strict and t2 == "f":
        return True
    if "不合格" in t and "合格" not in t:
        return True
    if any(k in t for k in ("超差", "超上限", "超下限", "拒收")):
        return True
    bad = {"ng", "fail", "failed", "reject", "rejected", "nogo"}
    return any(tk in bad for tk in _tokens(v))


def _is_pass_text(v):
    t = _norm(v)
    if not t:
        return False
    t2 = _strip_symbols(t)
    if t2 in ("合格", "ok", "pass", "p", "good", "accept", "是", "yes", "y"):
        return True
    if "ok" in t or "合格" in t:
        return True
    return False


def _fnum(x):
    """解析数值，去掉单位/°/空格；失败返回 None。

    重要：Excel 错误值（#VALUE!/#NAME?/#DIV/0!...）经 COM 读取时返回的是特定负整数，
    若不加拦截会被当成"超大负数"参与越界比较 → 把错误单元格误判成不合格。
    典型触发场景：表格用了 LET/TEXTJOIN/XLOOKUP 等新版函数，而打开端的 Excel
    版本较旧不支持（返回 #NAME?）→ 同一份文件在不同电脑被判出不同结果。
    """
    if x is None:
        return None
    # Excel 错误值（xlErrValue / xlErrName / xlErrDiv0 等）→ 一律视为无效，不参与判定
    if isinstance(x, (int, float)) and not isinstance(x, bool):
        try:
            xi = int(x)
        except (ValueError, OverflowError):
            xi = None
        if xi is not None and xi in _XL_ERROR_CODES:
            return None
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).strip()
    if not s:
        return None
    # 单元格显示形态的错误标记（如 '#NAME?'、'#VALUE!'）也直接判无效
    if s.startswith("#") and s.endswith(("?", "!")):
        return None
    s = re.sub(r"[°\s]", "", s)
    m = re.search(r"-?\d+(?:\.\d+)?", s)
    if not m:
        return None
    try:
        return float(m.group(0))
    except Exception:
        return None


# ============================================================
# COM 读写
# ============================================================

def _sheet_grid(ws):
    """从单个工作表对象读取全部 UsedRange，规整成 (grid, nrows, ncols)。"""
    vals = ws.UsedRange.Value
    if vals is None:
        return (None, 0, 0)
    if not isinstance(vals, (tuple, list)):
        vals = [[vals]]
    rows = []
    for r in vals:
        if not isinstance(r, (tuple, list)):
            r = [r]
        rows.append(list(r))
    nrows = len(rows)
    ncols = max((len(r) for r in rows), default=0)
    for r in rows:
        if len(r) < ncols:
            r.extend([None] * (ncols - len(r)))
    return (rows, nrows, ncols)


def _read_grid(path, sheet_idx=1):
    """COM 打开（临时副本 + 强制重算），返回指定工作表 (grid, nrows, ncols)。

    关键改进（v3 修复环境差异）：
    - 用临时副本 + ReadOnly=False 打开 → 让 Excel 自动重算公式，避免
      .Value 拿到未计算的旧缓存值（不同机器 Excel 缓存状态不同 → 同一份
      文件被判出不同结果）。
    - 关闭时不保存（False）→ 源文件不动。
    """
    if win32com is None:
        raise RuntimeError("缺少 win32com 模块，无法读取 Excel")
    xl = None
    wb = None
    tmp_path = None
    try:
        try:
            xl = win32com.client.DispatchEx("Excel.Application")
        except Exception:
            xl = win32com.client.Dispatch("Excel.Application")
        for _attr, _val in (("Visible", False), ("DisplayAlerts", False),
                             ("ScreenUpdating", False), ("EnableEvents", False)):
            try:
                setattr(xl, _attr, _val)
            except Exception:
                pass

        # 用临时副本打开：保证公式被 Excel 自动重算 + 不锁源文件
        tmp_dir = os.environ.get("TEMP") or os.path.expanduser("~")
        tmp_path = os.path.join(tmp_dir, f"_audit_read_{os.getpid()}_{int(time.time()*1000) % 100000}.xlsx")
        shutil.copy2(os.path.abspath(path), tmp_path)

        wb = xl.Workbooks.Open(os.path.abspath(tmp_path), ReadOnly=False)
        # 强制全表重算（保证公式最新值）
        try:
            xl.Calculate
        except Exception:
            pass
        try:
            xl.CalculateUntilAsyncDone
        except Exception:
            pass

        ws = wb.Sheets(sheet_idx)
        return _sheet_grid(ws)
    finally:
        try:
            if wb is not None:
                wb.Close(False)  # 不保存任何修改
        except Exception:
            pass
        try:
            if xl is not None:
                xl.Quit()
        except Exception:
            pass
        if tmp_path and os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except Exception:
                pass


# ============================================================
# 列角色定位（与 D:\文件转换\converter.py 完全一致的判定逻辑）
# ============================================================
# 索引约定：grid 一律用 0-based（行、列都从 0 开始），与 converter.py 的 sv(r,c) 一致。
# 输出给 COM 标红时再 +1 转成 1-based。

def _sv(grid, r, c):
    """取单元格文本（strip），越界返回 ''。0-based。"""
    try:
        v = grid[r][c]
        return "" if v is None else str(v).strip()
    except Exception:
        return ""


def _detect_layout(grid, nrows, ncols):
    """探测表头/子表头/数据起始行与各关键列（对齐 converter.py 的 _detect_layout）。

    返回 dict，或 None（未识别）。所有列、行索引均为 0-based。
    """
    def sv(r, c):
        return _sv(grid, r, c)

    # 1) 表头行：含 "item no"
    header_row = None
    for r in range(min(nrows, 25)):
        for c in range(ncols):
            if "item no" in sv(r, c).lower():
                header_row = r
                break
        if header_row is not None:
            break
    if header_row is None:
        return None

    def find_col(row, *keywords):
        for c in range(ncols):
            vl = sv(row, c).lower()
            if vl and any(k in vl for k in keywords):
                return c
        return None

    item_col = find_col(header_row, "item no")
    instr_col = find_col(header_row, "instrument")

    # Total Qty 数值
    total_qty = 1
    tq_col = find_col(header_row, "total qty", "总数")
    if tq_col is not None:
        for off in (1, 2, 3):
            try:
                total_qty = int(float(sv(header_row, tq_col + off)))
                break
            except (ValueError, TypeError, IndexError):
                pass

    # 2) 子表头行：含 公称/上限/下限/dimension/max/min
    sub_row = None
    for r in range(header_row + 1, header_row + 4):
        if r >= nrows:
            break
        if (find_col(r, "公称", "dimension") is not None
                or find_col(r, "上限", "max") is not None
                or find_col(r, "下限", "min") is not None):
            sub_row = r
            break
    if sub_row is None:
        sub_row = header_row + 1

    spec_col = find_col(sub_row, "公称", "dimension")
    max_col = find_col(sub_row, "上限", "max")
    min_col = find_col(sub_row, "下限", "min")
    result_col = find_col(sub_row, "result", "p合格", "(p")

    # 3) 实测数据列候选：instr+1 .. result-1，排除 最大/最小
    reading_cols = []
    if instr_col is not None and result_col is not None:
        for c in range(instr_col + 1, result_col):
            lab = sv(sub_row, c) or ""
            if lab and "最大" not in lab and "最小" not in lab:
                reading_cols.append(c)
    elif instr_col is not None:
        for c in range(instr_col + 1, instr_col + 1 + max(total_qty, 1)):
            reading_cols.append(c)

    # 实测列必须在子表头行有非空标签（对应 converter.py 的 eff_reading）
    if reading_cols:
        reading_cols = [c for c in reading_cols if sv(sub_row, c)]

    data_start = sub_row + 1
    return {
        "header_row": header_row,
        "sub_row": sub_row,
        "item_col": item_col,
        "spec_col": spec_col,
        "max_col": max_col,
        "min_col": min_col,
        "instr_col": instr_col,
        "result_col": result_col,
        "reading_cols": reading_cols,
        "data_start": data_start,
    }


# ---- 兜底：找不到标准 Item No 表头时，用标签定位（0-based 版） ----

_MIN_LABEL_GROUPS = (("下限值",), ("下限",), ("min", "lcl", "lower"))
_MAX_LABEL_GROUPS = (("上限值",), ("上限",), ("max", "ucl", "upper"))


def _find_label_col(grid, nrows, ncols, groups):
    """按标签组优先级定位某角色列（0-based）；要求该列至少 1 行是数值。"""
    for group in groups:
        for c in range(ncols):
            parts = []
            for r in range(nrows):
                v = _sv(grid, r, c)
                if v:
                    parts.append(v.lower())
            hv = " ".join(parts)
            if not hv:
                continue
            if any(k in hv for k in group):
                n_val = sum(1 for r in range(nrows) if _fnum(_sv(grid, r, c)) is not None)
                if n_val >= 1:
                    return c
    return None


def _find_result_col_fallback(grid, nrows, ncols):
    """兜底结果列定位（0-based），优先级：result/结果 > 判定/结论 > P/F 列。"""
    for c in range(ncols):
        for r in range(nrows):
            t = _sv(grid, r, c).lower()
            if t and ("result" in t or "结果" in t):
                return c
    for c in range(ncols):
        for r in range(nrows):
            t = _sv(grid, r, c).lower()
            if t and ("判定" in t or "结论" in t):
                return c
    _JUDGE = {"p", "f", "pass", "ng", "fail", "ok", "合格", "不合格",
              "accept", "reject", "nogo", "good", "y", "yes"}
    for c in range(ncols):
        hit = tot = num = 0
        for r in range(nrows):
            v = _sv(grid, r, c)
            if not v:
                continue
            tot += 1
            if _strip_symbols(v.lower()) in _JUDGE:
                hit += 1
            if _fnum(v) is not None:
                num += 1
        if tot >= 3 and hit / tot >= 0.7 and num / tot < 0.3:
            return c
    return None


# ============================================================
# 扫描核心
# ============================================================

def _scan_core(grid, nrows, ncols, min_col, max_col, result_col, reading_cols,
               data_start, result_is_layout):
    """核心越界判定（对齐 converter.py 的 _has_failed）。0-based 索引。

    规则与 converter.py 完全一致：
      - 该行 min 与 max 都为 None → 跳过（无公差，不参与判定）；
      - min/max 只有一个有值 → 仍按该单边判定（converter.py 的单边逻辑）；
      - max < min → 自动交换（源文件偶尔填反）；
      - 读数 OK/合格/PASS/P/空 → 跳过；
      - 读数无法转数值（带单位文本等）→ 跳过；
      - v < min 或 v > max → 越界。
    """
    if not reading_cols:
        return [], "未识别到实测列"
    if min_col is None and max_col is None:
        return [], "未识别到上限/下限列"

    fails = []
    notes = []
    for r in range(data_start, nrows):
        minv = _fnum(_sv(grid, r, min_col)) if min_col is not None else None
        maxv = _fnum(_sv(grid, r, max_col)) if max_col is not None else None
        # 智能识别：源文件偶尔把 max/min 填反（max<min），自动交换保证区间合法
        if minv is not None and maxv is not None and maxv < minv:
            minv, maxv = maxv, minv
        if minv is None and maxv is None:
            continue  # 无公差，不参与判定

        oob_col = None
        oob_val = None
        for c in reading_cols:
            rs = _sv(grid, r, c)
            if not rs:
                continue
            if rs.upper() in ("OK", "合格", "PASS", "P", ""):
                continue
            v = _fnum(rs)
            if v is None:
                continue
            if (minv is not None and v < minv) or (maxv is not None and v > maxv):
                oob_col = c
                oob_val = v
                break

        if oob_col is not None:
            lo_str = "-" if minv is None else str(minv)
            hi_str = "-" if maxv is None else str(maxv)
            if result_col is not None:
                fails.append({"r": r + 1, "c": result_col + 1,
                              "value": _sv(grid, r, result_col),
                              "kind": "reading", "is_result": True})
            else:
                fails.append({"r": r + 1, "c": oob_col + 1, "value": oob_val,
                              "kind": "reading", "is_result": False})
            notes.append(f"第{r + 1}行 测量值 {oob_val} 超出公差 [{lo_str}, {hi_str}]")

    if notes:
        return fails, "；".join(notes)
    return fails, "未检出不合格（逐行按上下限校验，无越界）"


def _scan_one_sheet(grid, nrows, ncols):
    """对单张表扫描越界，返回 (fails, reason)。"""
    layout = _detect_layout(grid, nrows, ncols)
    if layout is not None:
        max_col = layout["max_col"]
        min_col = layout["min_col"]
        result_col = layout["result_col"]
        reading_cols = layout["reading_cols"]
        data_start = layout["data_start"]

        if max_col is None and min_col is None:
            return [], "未识别到标准表头（找不到 Item No 表头），无法做数值校验"
        if not reading_cols:
            return [], "未识别到实测数据列"

        fails, reason = _scan_core(grid, nrows, ncols, min_col, max_col,
                                   result_col, reading_cols, data_start,
                                   result_is_layout=True)
        return fails, reason

    # ---- 兜底：没有标准 Item No 表头，改用标签定位 ----
    max_col = _find_label_col(grid, nrows, ncols, _MAX_LABEL_GROUPS)
    min_col = _find_label_col(grid, nrows, ncols, _MIN_LABEL_GROUPS)
    result_col = _find_result_col_fallback(grid, nrows, ncols)

    reading_cols = []
    if result_col is not None and result_col > max(max_col or 0, min_col or 0):
        for c in range(max(max_col or 0, min_col or 0) + 1, result_col):
            reading_cols.append(c)
    elif result_col is not None:
        for c in range(max(max_col or 0, min_col or 0) + 1, ncols):
            if c != result_col:
                reading_cols.append(c)
    else:
        return [], "未识别到标准表头，且未找到结果列，无法做数值校验"

    fails, reason = _scan_core(grid, nrows, ncols, min_col, max_col,
                               result_col, reading_cols, 0,
                               result_is_layout=False)
    return fails, reason


def scan_fails(path):
    """扫描所有工作表中的不合格单元格，返回 (fails, reason)。

    v3：
      1) 全表定位上限/下限/结果三列的角色（标签匹配，与行号无关）；
      2) 测量列 = 公差列右侧、结果列左侧的列（结构区间，不做占比投票）；
      3) 逐行：上限/下限列的值就是该行公差（写反自动交换），
         测量列数值若落在 [mn, mx] 之外 → 不合格，结果列标红 + 改 F；
      4) 有公差行只看越界，结果列文本不参与判不合格（修复环境差异）；
      5) 遍历全部工作表，任一表有越界即判不合格。
    """
    if win32com is None:
        raise RuntimeError("缺少 win32com 模块，无法读取 Excel")

    xl = None
    wb = None
    tmp_path = None
    try:
        try:
            xl = win32com.client.DispatchEx("Excel.Application")
        except Exception:
            xl = win32com.client.Dispatch("Excel.Application")
        for _attr, _val in (("Visible", False), ("DisplayAlerts", False),
                             ("ScreenUpdating", False), ("EnableEvents", False)):
            try:
                setattr(xl, _attr, _val)
            except Exception:
                pass

        tmp_dir = os.environ.get("TEMP") or os.path.expanduser("~")
        tmp_path = os.path.join(tmp_dir, f"_audit_read_{os.getpid()}_{int(time.time()*1000) % 100000}.xlsx")
        shutil.copy2(os.path.abspath(path), tmp_path)

        wb = xl.Workbooks.Open(os.path.abspath(tmp_path), ReadOnly=False)
        try:
            xl.Calculate
        except Exception:
            pass
        try:
            xl.CalculateUntilAsyncDone
        except Exception:
            pass

        all_fails = []
        all_reasons = []
        sheet_count = wb.Sheets.Count
        for sheet_idx in range(1, sheet_count + 1):
            ws = wb.Sheets(sheet_idx)
            grid, nrows, ncols = _sheet_grid(ws)
            if not grid or nrows == 0:
                continue
            fails, reason = _scan_one_sheet(grid, nrows, ncols)
            if fails:
                # 为每个 fail 记录它所在的工作表
                for f in fails:
                    f["sheet"] = sheet_idx
                all_fails.extend(fails)
                all_reasons.append(f"工作表 {sheet_idx}（{ws.Name}）: {reason}")

        if all_fails:
            reason = "；".join(all_reasons)
        else:
            reason = "未检出不合格（已扫描全部工作表，无越界）"
        return all_fails, reason
    finally:
        try:
            if wb is not None:
                wb.Close(False)
        except Exception:
            pass
        try:
            if xl is not None:
                xl.Quit()
        except Exception:
            pass
        if tmp_path and os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except Exception:
                pass


def detect_fail(path):
    """判定文件是否含不合格数据。返回 (is_fail, reason)。（兼容旧调用）"""
    fails, reason = scan_fails(path)
    return (len(fails) > 0, reason)


def _mark_fails(ws, fails):
    """对 fails 中的【目标单元格】应用红底色 + 标 F（字体保持不变）。"""
    _PASS_TOKENS = ("p", "pass", "ok", "合格", "good", "accept", "y", "yes")
    for f in fails:
        r, c = f["r"], f["c"]
        try:
            cell = ws.Cells(r, c)
            cell.Interior.Pattern = 1  # xlSolid
            cell.Interior.Color = RED_FILL
            # 用户要求：只把单元格底色标红，字体保持不变
            if f.get("kind") == "reading" and f.get("is_result"):
                if not bool(cell.HasFormula):
                    raw = cell.Value
                    s = "" if raw is None else str(raw).strip()
                    t2 = _strip_symbols(s.lower())
                    if t2 in _PASS_TOKENS or s == "":
                        cell.Value = "F"
        except Exception:
            pass


def _mark_and_save(input_path, fails, out_path):
    """打开源文件（读写），按 fail 里的 sheet 标红标 F，SaveAs 到 out_path。"""
    if win32com is None:
        raise RuntimeError("缺少 win32com 模块，无法写入 Excel")
    xl = None
    wb = None
    tmp_in = None
    try:
        try:
            xl = win32com.client.DispatchEx("Excel.Application")
        except Exception:
            xl = win32com.client.Dispatch("Excel.Application")
        for _attr, _val in (("Visible", False), ("DisplayAlerts", False),
                             ("ScreenUpdating", False), ("EnableEvents", False),
                             ("AskToUpdateLinks", False)):
            try:
                setattr(xl, _attr, _val)
            except Exception:
                pass
        # 用临时副本打开，避免源文件被 Excel 锁
        tmp_dir = os.environ.get("TEMP") or os.path.expanduser("~")
        tmp_in = os.path.join(tmp_dir, f"_audit_mark_{os.getpid()}_{int(time.time()*1000) % 100000}.xlsx")
        shutil.copy2(os.path.abspath(input_path), tmp_in)
        wb = xl.Workbooks.Open(os.path.abspath(tmp_in), ReadOnly=False)
        try:
            xl.Calculate
        except Exception:
            pass

        # 按工作表分组标红
        by_sheet = {}
        for f in fails:
            sheet_idx = f.get("sheet", 1)
            by_sheet.setdefault(sheet_idx, []).append(f)
        for sheet_idx, sheet_fails in sorted(by_sheet.items()):
            if 1 <= sheet_idx <= wb.Sheets.Count:
                ws = wb.Sheets(sheet_idx)
                _mark_fails(ws, sheet_fails)

        ext = os.path.splitext(out_path)[1].lower()
        fmt = _EXT_FMT.get(ext, 51)
        wb.SaveAs(os.path.abspath(out_path), FileFormat=fmt)
    finally:
        try:
            if wb is not None:
                wb.Close(False)
        except Exception:
            pass
        try:
            if xl is not None:
                xl.Quit()
        except Exception:
            pass
        if tmp_in and os.path.exists(tmp_in):
            try:
                os.remove(tmp_in)
            except Exception:
                pass


def _with_fail_prefix(out_path):
    """在文件名最前加 不合格_（已有则不加）。"""
    d, f = os.path.split(out_path)
    if f.startswith("不合格_"):
        return out_path
    return os.path.join(d, "不合格_" + f)


def _alt_path(out_path, n):
    """目标被占用时的备选名：name(2).ext / name(3).ext ...（Windows 惯例）。"""
    d, f = os.path.split(out_path)
    base, ext = os.path.splitext(f)
    return os.path.join(d, f"{base}({n}){ext}")


def _write_output(write_fn, first_target, max_try=20):
    """尝试写输出；目标被占用（已打开/只读）时自动换 name(2)、name(3)... 重试。

    write_fn(target) 成功返回 target；全部失败抛最后一次异常。
    """
    last_err = None
    target = first_target
    for attempt in range(max_try):
        if attempt > 0:
            target = _alt_path(first_target, attempt + 1)
        try:
            write_fn(target)
            return target
        except Exception as e:  # noqa: BLE001
            last_err = e
    raise last_err


def audit_file(input_path, out_dir=None, on_progress=None):
    """审核单个文件并输出。"""
    input_path = os.path.abspath(input_path)
    if not os.path.isfile(input_path):
        raise FileNotFoundError(input_path)
    base, ext = os.path.splitext(os.path.basename(input_path))
    od = os.path.abspath(out_dir) if out_dir else os.path.dirname(input_path)
    os.makedirs(od, exist_ok=True)

    fails, reason = scan_fails(input_path)
    is_fail = len(fails) > 0

    if is_fail:
        out_path = _with_fail_prefix(os.path.join(od, base + ext))
    else:
        out_path = os.path.join(od, base + ext)

    copied = False
    marked = False
    if os.path.abspath(out_path) == input_path:
        output = input_path
    else:
        if is_fail:
            output = _write_output(
                lambda t: _mark_and_save(input_path, fails, t), out_path)
            marked = True
        else:
            output = _write_output(
                lambda t: shutil.copy2(input_path, t), out_path)
            copied = True

    if on_progress:
        on_progress(100, f"{'不合格' if is_fail else '合格'} -> {os.path.basename(output)}"
                        + ("（已标红标F）" if marked else ""))
    return {"input": input_path, "output": output, "is_fail": is_fail,
            "reason": reason, "copied": copied, "marked": marked}


def audit_paths(paths, out_dir=None, on_progress=None):
    """批量审核（文件或目录混合）。"""
    import glob

    files = []
    for p in paths:
        p = os.path.abspath(p)
        if os.path.isdir(p):
            for ext in ("*.xls", "*.xlsx", "*.xlsm"):
                files.extend(glob.glob(os.path.join(p, "**", ext), recursive=True))
        elif os.path.isfile(p):
            files.append(p)
    files = sorted(set(files))
    total = len(files)
    fail = 0
    for i, f in enumerate(files, 1):
        try:
            res = audit_file(f, out_dir=out_dir, on_progress=on_progress)
            if res["is_fail"]:
                fail += 1
        except Exception as e:
            if on_progress:
                on_progress(100, f"失败 {os.path.basename(f)}: {e}")
        if on_progress:
            on_progress(int(i / total * 100) if total else 100, f"已处理 {i}/{total}")
    return {"total": total, "fail": fail}


if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("用法: python audit_fail.py <文件或目录> [输出目录]")
        sys.exit(1)

    src = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) > 2 else None
    res = audit_file(src, out_dir=out)
    print(res)
