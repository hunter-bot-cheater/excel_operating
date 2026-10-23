#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
审核不合格数据（按数据反推 / v11 双引擎 + 规范化 + 异常检测修复版）
================================

设计目标（沿用 v3-v9，未变）
--------
针对检验报告（.xls / .xlsx / .xlsm）：
  * 按报告内“自身公式 / 结果”判定的合格状态，命中不合格则：
      1) 在【原文件名最前面】加 `不合格_` 前缀（命名体现）。
      2) 在报告内部对**同一行的【结果列单元格】** 标红 + 标 F：
         - 实测值超出 [下限, 上限] 时：把该行结果列单元格标红（亮红底色，字体不变），
           若该单元格当前是 P / 合格 / 空（且不是公式），自动改为 F。
         - 结果列本身已是不合格（不合格 / NG / Fail / F 等）时：标红，保持原值/公式。
         - **测量列单元格保持原样**，不追加 F、不标红，便于复检核对原始读数。
  * **扫描所有工作表**：任一表有越界即判不合格，标红也只标在对应表的对应单元格。
  * 其余格式（字体、列宽、行高、合并单元格、图片等）一律原样保留。

v7 新增：报告规范化（与合格/不合格判定并列，独立展示）
--------------------------------------------------
  * 客户代码 → 统一改为 0998（以文本格式写入，前导 0 不丢失；
    兼容「Customer Code: 客户代号 998」夹标签模板）。
  * 最终判定三选项（Passed 合格 / Failed 不合格 / UAI 让步）任一为空
    → final_missing 报错提示（不改变合格/不合格判定，并列展示）。
  * 上下限列（上限值/下限值）数据区空白单元格 → 自动填入 "/"，并统计
    填写个数 slash_filled（合并单元格/表头页脚不填不统计）。
  * 合格文件不再逐字节复制：所有输出文件（合格/不合格）都执行规范化后写出。

v8 新增（用户要求）：
  * 绝不覆盖原文件：输出到指定目录，目标已存在（含源文件同名）时自动
    在文件名后加 (1)、(2)... 区分。
  * 预加载：文件放入列表即后台预打开（openpyxl 快路径缓存），审核时直接
    命中缓存，省去加载时间。

v9 修复（用户提供真实模板后定位）：
  * 上下限空白只填到「最后一个有效数据行」（结果列有判定 P/F 或任一测量列
    有读数）为止；尾部空白/斜杠行（如检具代码说明行、纯空行）不填 /，
    避免“填了1000个/”。

v6 优化（速度 / 健壮性）
-----------------------
  * 混合引擎：
      - 无公式的 .xlsx/.xlsm（且不含图表/图片/透视/切片器等复杂元素）
        → openpyxl 快路径：读值→判定→标红→保存，全部内存内完成，实测
        单文件 ~0.09s，免启动 Excel，格式保真。
      - 有公式 / .xls / 复杂元素 → Excel COM 慢路径：强制重算，保证跨机器
        判定一致（公式结果以打开端 Excel 重算为准）。
  * ExcelSession：批量审核共享一个 Excel COM 进程。原来每文件启动一次
    Excel（实测 4.6s/次），现在整批只启动 1 次。
  * COM 路径单会话化：不合格文件在同一会话内完成「打开→重算→扫描→标红→另存」。
  * 健壮性：
      - 图表工作表读取失败不再让整个文件报错（跳过并记录）；
      - 临时文件用 uuid 命名，避免多进程碰撞；
      - 目录扫描排除输出目录子树与已生成的 `不合格_` 前缀文件，防止二次扫描。
  * 兼容性：scan_fails / detect_fail / audit_file / audit_paths / _mark_and_save
    签名与行为保持不变；audit_file 新增可选 session 参数（不传则自动管理）。
"""

import os
import re
import shutil
import sys
import json
import queue
import subprocess
import tempfile
import threading
import time
import uuid

try:
    import win32com.client
except Exception:
    win32com = None

try:
    import openpyxl
    from openpyxl.styles import PatternFill
except Exception:
    openpyxl = None


# ---- 颜色（OLE 颜色：低位字节为 R）----
# 亮红底色（255,0,0）-> BGR 0x0000FF；openpyxl 侧用 RGB "FF0000"
RED_FILL = 0x0000FF
RED_FILL_RGB = "FF0000"

# SaveAs 文件格式
_EXT_FMT = {".xlsx": 51, ".xlsm": 52, ".xls": 56}

# ---- Excel 错误值编码 ----
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
    若不拦截会被当成"超大负数"参与越界比较 → 把错误单元格误判成不合格。
    """
    if x is None:
        return None
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
# COM 会话管理（批量只启动一次 Excel）
# ============================================================

def _make_xl():
    """新建一个隐藏的 Excel COM 实例。"""
    if win32com is None:
        raise RuntimeError("缺少 win32com 模块，无法读取 Excel")
    try:  # COM 必须在线程内初始化（GUI worker / API worker 均需）
        import pythoncom
        pythoncom.CoInitialize()
    except Exception:
        pass
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
    # 宏安全：禁用宏自动运行，避免 WPS 弹"启用宏"对话框（不影响数据读取）
    try:
        xl.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
    except Exception:
        pass
    return xl


def _recalc(xl):
    """强制全表重算（保证公式最新值）。"""
    try:
        xl.Calculate()
    except Exception:
        pass
    try:
        xl.CalculateUntilAsyncDone()
    except Exception:
        pass


def _com_speed_on(xl):
    """关闭屏幕刷新 / 事件 / 自动重算，返回原设置（规范化、标红期间减少连锁开销）。"""
    prev = {}
    try:
        prev["screen"] = xl.ScreenUpdating; xl.ScreenUpdating = False
    except Exception:
        pass
    try:
        prev["events"] = xl.EnableEvents; xl.EnableEvents = False
    except Exception:
        pass
    try:
        prev["calc"] = xl.Calculation; xl.Calculation = -4135   # xlCalculationManual
    except Exception:
        pass
    return prev


def _com_speed_off(xl, prev):
    """恢复 _com_speed_on 关闭的设置。"""
    try:
        xl.Calculation = prev.get("calc", -4105)                # xlCalculationAutomatic
    except Exception:
        pass
    try:
        xl.EnableEvents = prev.get("events", True)
    except Exception:
        pass
    try:
        xl.ScreenUpdating = prev.get("screen", True)
    except Exception:
        pass


def _make_tmp(prefix, ext=".xlsx"):
    """生成带 uuid 的临时文件路径（避免多进程碰撞）。"""
    tmp_dir = os.environ.get("TEMP") or os.path.expanduser("~")
    return os.path.join(tmp_dir, f"{prefix}{os.getpid()}_{uuid.uuid4().hex[:8]}{ext}")


class ExcelSession:
    """复用的 Excel COM 进程。批量审核共享一个实例，启动开销从每文件一次降为整批一次。

    用法：
        sess = ExcelSession()
        try:
            audit_file(f, out_dir, session=sess)
        finally:
            sess.quit()
    """

    def __init__(self):
        self._xl = None
        self.xl_pid = None

    def get(self):
        if self._xl is None:
            self._xl = _make_xl()
            try:
                h = self._xl.Hwnd
                if h:
                    import win32process
                    _, self.xl_pid = win32process.GetWindowThreadProcessId(h)
            except Exception:
                self.xl_pid = None
        return self._xl

    def quit(self):
        if self._xl is not None:
            try:
                self._xl.Quit()
            except Exception:
                pass
            self._xl = None


# ============================================================
# 数据网格读取
# ============================================================

def _sheet_grid(ws):
    """从单个 COM 工作表对象读取全部 UsedRange，规整成 (grid, nrows, ncols)。"""
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


def _com_visible_sheet(wb, n):
    """取 COM 工作簿中第 n 个「可见」工作表（1-based）；找不到返回 None。

    隐藏/深度隐藏工作表不参与审核，索引统一按可见表计数，
    与 grids 的 sheet_idx（可见表序号）保持一致。
    """
    seen = 0
    for i in range(1, int(wb.Sheets.Count) + 1):
        try:
            ws = wb.Sheets(i)
        except Exception:
            continue
        try:
            if int(ws.Visible) != -1:
                continue  # 0=隐藏, 2=深度隐藏
        except Exception:
            continue
        seen += 1
        if seen == n:
            return ws
    return None


def _hidden_rows_com(ws, nrows):
    """收集被筛选/手动隐藏的行（返回 0-based 行号集合）。

    Excel 的 Rows(r).Hidden 已综合 AutoFilter 筛选结果与手动隐藏，
    与“用户看得到”的行完全一致。
    """
    hidden = set()
    for r in range(1, nrows + 1):
        try:
            if ws.Rows(r).Hidden:
                hidden.add(r - 1)
        except Exception:
            pass
    return hidden


def _openpyxl_grid(ws):
    """从 openpyxl 工作表读取全部值，规整成 (grid, nrows, ncols, has_formula)。

    与 COM 版同构；顺带检测公式单元格（含公式必须降级 COM 强制重算，
    否则快路径可能读到旧缓存值或把公式文本当数值误判）。
    """
    nrows = ws.max_row
    ncols = ws.max_column
    if not nrows or not ncols:
        return (None, 0, 0, False)
    rows = []
    has_formula = False
    for row in ws.iter_rows(min_row=1, max_row=nrows, min_col=1, max_col=ncols,
                            values_only=True):
        rl = list(row)
        if not has_formula:
            for v in rl:
                if isinstance(v, str) and v.startswith("="):
                    has_formula = True
                    break
        rows.append(rl)
    return (rows, nrows, ncols, has_formula)


# ============================================================
# 列角色定位（与 converter.py 完全一致的判定逻辑）
# ============================================================
# 索引约定：grid 一律用 0-based（行、列都从 0 开始），与 converter.py 的 sv(r,c) 一致。
# 输出给标红时再 +1 转成 1-based。

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

    # 2) 子表头行：含 公称/上限/下限/dimension/max/min。
    #    优先在表头行下方找（旧模板紧邻）；找不到再全表前 30 行找
    #    （新模板子表头在 Item No 表头行的上方，如 93010LBF 模板）。
    sub_row = None
    for r in range(header_row + 1, min(header_row + 4, nrows)):
        if (find_col(r, "公称", "dimension") is not None
                or find_col(r, "上限", "max") is not None
                or find_col(r, "下限", "min") is not None):
            sub_row = r
            break
    if sub_row is None:
        for r in range(min(nrows, 30)):
            if r == header_row:
                continue
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
    if result_col is None:
        # 结果列标签可能不在子表头行（表头区另一行）→ 全表前 30 行兜底
        for r in range(min(nrows, 30)):
            result_col = find_col(r, "result", "p合格", "(p")
            if result_col is not None:
                break

    # Instrument 列：优先子表头行，其次表头行
    if instr_col is None and sub_row is not None and sub_row != header_row:
        instr_col = find_col(sub_row, "instrument")

    # 3) 实测数据列候选：instr+1 .. result-1，排除 最大/最小
    #    注意：不能要求 sub_row 有标签——新模板读数位标签放在 Item No
    #    表头行、子表头行无标签（如 Shutter Storage 模板 G 列测量、第10行空），
    #    否则 reading_cols 会误判为空 → 整表不扫描 → 漏判不合格。
    reading_cols = []
    if instr_col is not None and result_col is not None:
        for c in range(instr_col + 1, result_col):
            lab = sv(sub_row, c) or ""
            if "最大" not in lab and "最小" not in lab:
                reading_cols.append(c)
    elif instr_col is not None:
        for c in range(instr_col + 1, instr_col + 1 + max(total_qty, 1)):
            reading_cols.append(c)

    # 实测列须在 子表头 或 表头 行有非空标签（对应 converter.py 的 eff_reading；
    # 新模板读数位标签（1,2,3…）放在 Item No 表头行，子表头行无标签）
    if reading_cols:
        reading_cols = [c for c in reading_cols
                        if sv(sub_row, c) or (header_row is not None and sv(header_row, c))]

    data_start = sub_row + 1
    # 子表头与数据区之间可能隔了表头/空白区（新模板）→ 前移到第一个有效数据行，
    # 避免把表头区空行当数据行（填 /、补结果列）
    def _is_seq_label_row(rr):
        # 读数位标签行：reading_cols 全是 1..9 小整数序号、item_col 为空
        vals = [sv(rr, c) for c in reading_cols]
        nums = []
        for v in vals:
            if v is None or v == "":
                continue
            try:
                fv = float(v)
            except (ValueError, TypeError):
                return False
            if 1 <= fv <= 9 and fv == int(fv):
                nums.append(fv)
            else:
                return False
        return bool(nums) and not sv(rr, item_col)

    for r in range(data_start, min(data_start + 80, nrows)):
        if _is_seq_label_row(r):
            continue  # 读数位标签行（1,2,3...）不是数据行
        if ((max_col is not None and sv(r, max_col))
                or (min_col is not None and sv(r, min_col))
                or any(sv(r, c) for c in reading_cols)):
            data_start = r
            break
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

_FAIL_TOKENS = ("f", "fail", "ng", "不合格", "不通过", "x")


def _is_fail_token(v):
    """结果列单元格是否为「不合格」判定（原文件检验员已标 F，必须尊重）。"""
    if v is None:
        return False
    s = str(v).strip().lower()
    if not s:
        return False
    return s in _FAIL_TOKENS or s.startswith("f")


def _scan_core(grid, nrows, ncols, min_col, max_col, result_col, reading_cols,
               data_start, result_is_layout, hidden=None):
    """核心越界判定（对齐 converter.py 的 _has_failed）。0-based 索引。

    规则与 converter.py 完全一致：
      - 该行 min 与 max 都为 None → 跳过（无公差，不参与判定）；
      - min/max 只有一个有值 → 仍按该单边判定（converter.py 的单边逻辑）；
      - max < min → 自动交换（源文件偶尔填反）；
      - 读数 OK/合格/PASS/P/空 → 跳过；
      - 读数无法转数值（带单位文本等）→ 跳过；
      - v < min 或 v > max → 越界。
      - 原文件结果列已标 F/不合格 → 直接判不合格（尊重检验员判定）。
    """
    fails = []
    notes = []
    # 原文件结果列已有 F/不合格 → 检验员已判定，直接判不合格。
    # 必须前置：即使 reading_cols / 上下限列未识别（其他模板变体），
    # 也不能漏掉结果列已有的 F（Shutter Storage 漏判教训）。
    if result_col is not None:
        for r in range(data_start, nrows):
            if hidden and r in hidden:
                continue  # 被筛选/隐藏的行：用户看不到，不判断
            rv = _sv(grid, r, result_col)
            if _is_fail_token(rv):
                fails.append({"r": r + 1, "c": result_col + 1,
                              "value": rv, "kind": "reading", "is_result": True})
                notes.append(f"第{r + 1}行 原文件结果列已标F")
    if fails:
        return fails, "；".join(notes)
    if not reading_cols:
        return [], "未识别到实测列"
    if min_col is None and max_col is None:
        return [], "未识别到上限/下限列"

    for r in range(data_start, nrows):
        if hidden and r in hidden:
            continue  # 被筛选/隐藏的行：用户看不到，不判断
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


def _scan_one_sheet(grid, nrows, ncols, hidden=None):
    """对单张表扫描越界，返回 (fails, reason)。hidden: 被筛选/隐藏的 0-based 行集合。"""
    layout = _detect_layout(grid, nrows, ncols)
    if layout is not None:
        max_col = layout["max_col"]
        min_col = layout["min_col"]
        result_col = layout["result_col"]
        reading_cols = layout["reading_cols"]
        data_start = layout["data_start"]

        fails, reason = _scan_core(grid, nrows, ncols, min_col, max_col,
                                   result_col, reading_cols, data_start,
                                   result_is_layout=True, hidden=hidden)
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
                               result_is_layout=False, hidden=hidden)
    return fails, reason


def _scan_workbook(xl, wb):
    """扫描已打开的 COM 工作簿的「可见」工作表，返回 (fails, reason, norm)。

    隐藏/深度隐藏工作表跳过；被筛选或手动隐藏的行不判断。
    """
    all_fails = []
    all_reasons = []
    grids = []
    sheet_count = wb.Sheets.Count
    for i in range(1, sheet_count + 1):
        try:
            ws = wb.Sheets(i)
            if int(ws.Visible) != -1:
                continue  # 隐藏工作表：用户看不到，不判断
        except Exception:
            continue
        sheet_idx = len(grids) + 1  # 可见表序号
        try:
            grid, nrows, ncols = _sheet_grid(ws)
            if not grid or nrows == 0:
                continue
            hidden = _hidden_rows_com(ws, nrows)
            grids.append((grid, nrows, ncols, hidden))
            fails, reason = _scan_one_sheet(grid, nrows, ncols, hidden)
            if fails:
                for f in fails:
                    f["sheet"] = sheet_idx
                all_fails.extend(fails)
                all_reasons.append(f"工作表 {sheet_idx}（{ws.Name}）: {reason}")
        except Exception:
            # 图表工作表等无 UsedRange 的对象：跳过并记录，不让整文件报错
            all_reasons.append(f"工作表 {sheet_idx}: 读取失败（可能为图表工作表），已跳过")

    norm = _collect_norm(grids)
    # 剔除合并单元格（表头/页脚结构性区域，不填 /）
    _drop_merged_slash_com(norm, wb)
    if all_fails:
        reason = "；".join(all_reasons)
    else:
        reason = "未检出不合格（已扫描全部可见工作表，无越界）"
    return all_fails, reason, norm


# ============================================================
# 报告规范化（v7 新增：独立于合格判定的并列功能）
# ------------------------------------------------------------
# 1) 客户代码 → 统一改为 0998（标签右侧第一个非空单元格）
# 2) 最终判定三选项（合格/不合格/UAI）任一为空 → final_missing 报错提示
# 3) 上下限列数据区空白单元格 → 填入 "/"，并统计数量 slash_filled
# 索引约定：探测返回 0-based；应用时 +1 转 1-based。
# ============================================================

CUSTOMER_CODE = "0998"
_CUST_LABEL_ZH = ("客户代码", "客户代号", "客户编号", "客户编码")
_CUST_LABEL_EN = ("customer code", "cust code", "customer no")


def _is_cust_label(text):
    """判断单元格文本是否为客户代码类标签。"""
    t = (text or "").lower()
    return any(k in t for k in _CUST_LABEL_ZH) or any(k in t for k in _CUST_LABEL_EN)


def _find_customer_code(grid, nrows, ncols):
    """定位客户代码值单元格（0-based (r, c)）。

    兼容多种模板布局：
      - 英文标签 "Customer Code:" / 中文标签 "客户代码/客户代号/客户编号"
      - 标签与值之间可能夹着另一个客户代码类标签（如 "Customer Code: 客户代号 998"）
      - 优先用中文标签（通常紧邻值）；标签右侧跳过嵌套标签取首个非标签值
    """
    labels = []
    for r in range(min(nrows, 25)):
        for c in range(ncols):
            t = _sv(grid, r, c)
            if not t:
                continue
            is_zh = any(k in t for k in _CUST_LABEL_ZH)
            is_en = any(k in t for k in _CUST_LABEL_EN)
            if is_zh or is_en:
                labels.append((r, c, is_zh))
    # 中文标签优先（通常紧邻值）；同优先级按行、列序
    labels.sort(key=lambda x: (not x[2], x[0], x[1]))
    for (r, c, _) in labels:
        for cc in range(c + 1, ncols):
            v = _sv(grid, r, cc)
            if not v:
                continue
            if _is_cust_label(v):
                continue  # 中间夹着另一个客户代码标签，继续右移
            return (r, cc)
    return None


def _calc_result_val(grid, r, max_col, min_col, reading_cols):
    """结果列缺失时，按读数与公差推算应补的判定 P/F（v11）。

    - 上下限缺失或非数值 → P（用户口径：空白补 /，结果列直接标 P）
    - 读数为 OK/合格/文字 → P
    - 读数数值越界 → F；否则 P
    """
    minv = _fnum(_sv(grid, r, min_col)) if min_col is not None else None
    maxv = _fnum(_sv(grid, r, max_col)) if max_col is not None else None
    if minv is not None and maxv is not None and maxv < minv:
        minv, maxv = maxv, minv
    for c in reading_cols:
        rs = _sv(grid, r, c)
        if not rs:
            continue
        if rs.upper() in ("OK", "合格", "PASS", "P", ""):
            return "P"
        v = _fnum(rs)
        if v is None:
            return "P"  # 文字读数（合格/目测等）→ P
        if (minv is not None and v < minv) or (maxv is not None and v > maxv):
            return "F"
    return "P"


def _find_total_sample(grid, nrows, ncols):
    """定位 总数(Total Qty) / 检样数(Sample Qty) 右侧数值单元格（表头区前 30 行）。

    返回 {"total": (r, c, old) | None, "sample": (r, c, old) | None}（0-based）。
    """
    out = {"total": None, "sample": None}
    for r in range(min(nrows, 30)):
        for c in range(ncols):
            t = _sv(grid, r, c).lower()
            if not t:
                continue
            if out["total"] is None and ("total qty" in t or "总数" in t):
                for cc in range(c + 1, min(c + 6, ncols)):
                    v = _sv(grid, r, cc)
                    if v:
                        out["total"] = (r, cc, v)
                        break
            elif out["sample"] is None and ("sample qty" in t or "检样数" in t):
                for cc in range(c + 1, min(c + 6, ncols)):
                    v = _sv(grid, r, cc)
                    if v:
                        out["sample"] = (r, cc, v)
                        break
    return out


def _find_drawing_no(grid, nrows, ncols):
    """定位图号（WG No / 图号）值单元格。返回 (r, c, value) | None（0-based）。"""
    for r in range(min(nrows, 30)):
        for c in range(ncols):
            t = _sv(grid, r, c).lower()
            if not t:
                continue
            if ("wg no" in t or "图号" in t) and "零件名称" not in t:
                for cc in range(c + 1, min(c + 6, ncols)):
                    v = _sv(grid, r, cc)
                    if v and not _is_cust_label(v):
                        # 图号取第一行（换行后的备注/说明不混入）
                        first = str(v).strip().split("\n")[0].strip()
                        # 数字图号去 .0（93121255.0 → 93121255）
                        try:
                            fv = float(first)
                            if fv == int(fv):
                                first = str(int(fv))
                        except (ValueError, TypeError):
                            pass
                        return (r, cc, first)
    return None


def _probe_norm_sheet(grid, nrows, ncols, sheet_idx, hidden=None):
    """单表规范化探测。返回：
    {
      "customer_code": {"sheet","r","c","old"} | None,
      "final_cells": [(sheet, r, c), ...],
      "final_missing": bool,
      "slash_cells": [(sheet, r, c), ...],
    }
    """
    out = {"customer_code": None, "final_cells": [], "final_missing": False,
           "slash_cells": [], "result_fill": [], "total_qty": [],
           "sample_qty": [], "drawing_no": None, "gap_move": [],
           "label_erase": [], "n_meas": 0}

    # 1) 客户代码：标签右侧第一个非空值单元格（跳过嵌套标签）
    cc = _find_customer_code(grid, nrows, ncols)
    if cc is not None:
        out["customer_code"] = {"sheet": sheet_idx, "r": cc[0], "c": cc[1],
                                "old": _sv(grid, cc[0], cc[1])}

    # 2) 最终判定三选项：标签正下方连续 3 格
    fd = None
    for r in range(min(nrows, 30)):
        for c in range(ncols):
            t = _sv(grid, r, c).lower()
            if t and ("final disposition" in t or "最终判定" in t):
                fd = (r, c)
                break
        if fd is not None:
            break
    if fd is not None:
        r0, c0 = fd
        cells = [(r0 + 1, c0), (r0 + 2, c0), (r0 + 3, c0)]
        cells = [(r, c) for (r, c) in cells if r < nrows]
        out["final_cells"] = [(sheet_idx, r, c) for (r, c) in cells]
        out["final_missing"] = any(not _sv(grid, r, c) for (r, c) in cells)

    # 3) 上下限列数据区空白单元格（只统计到「最后一个有效数据行」，
    #    尾部空白/斜杠行不填 /，避免几百上千个空行被误填）
    #    4) 结果列缺失自动补 P/F（v11） 5) 总数/检样数更正（v11）
    #    6) 空列整理（v11）
    layout = _detect_layout(grid, nrows, ncols)
    if layout is not None:
        max_col, min_col = layout["max_col"], layout["min_col"]
        data_start = layout["data_start"]
        result_col = layout["result_col"]
        reading_cols = layout["reading_cols"]
        instr_col = layout.get("instr_col")
        item_col = layout.get("item_col")
        measure_cols = list(range(instr_col + 1, result_col)) if (
            instr_col is not None and result_col is not None) else reading_cols
        if result_col is not None or reading_cols:
            # 有效数据行 = item 编号有值 且 (测量列有值 或 结果列有值)
            # 底部合并大单元格（item 有值但测量列/结果列空）不算数据行
            data_end = data_start
            for r in range(data_start, nrows):
                if hidden and r in hidden:
                    continue  # 筛选隐藏行不计入有效数据行
                row_item = item_col is not None and bool(_sv(grid, r, item_col))
                row_read = any(_sv(grid, r, c) for c in measure_cols)
                row_res = (result_col is not None and bool(_sv(grid, r, result_col)))
                if row_item and (row_read or row_res):
                    data_end = r + 1
        else:
            data_end = nrows  # 无结果/测量列信息时保持原行为

        # 4) 结果列缺失 → 逐行按读数与公差推算 P/F（v11）
        if result_col is not None and reading_cols:
            for r in range(data_start, data_end):
                if hidden and r in hidden:
                    continue
                has_reading = any(_sv(grid, r, c) for c in reading_cols)
                has_lim = ((max_col is not None and _sv(grid, r, max_col)) or
                           (min_col is not None and _sv(grid, r, min_col)))
                if not has_reading and not has_lim:
                    continue  # 空行不补
                if _sv(grid, r, result_col):
                    continue  # 已有判定不补
                out["result_fill"].append({
                    "sheet": sheet_idx, "r": r, "c": result_col,
                    "val": _calc_result_val(grid, r, max_col, min_col, reading_cols)})

        # 5) 有数据的测量列数 N（数据区该列任意非空值即算，v11.2）
        has_data_cols = []
        for c in measure_cols:
            if any(_sv(grid, r, c) for r in range(data_start, data_end)
                   if not (hidden and r in hidden)):
                has_data_cols.append(c)
        n_meas = len(has_data_cols)
        out["n_meas"] = n_meas
        c_last = max(has_data_cols) if has_data_cols else None

        # 5c) 删多余序号标签：N < 测量列数时，删掉第 N+1..末尾的序号单元格
        #     （空白数据列保留、结果列不动，v11.2）
        if measure_cols and n_meas > 0 and n_meas < len(measure_cols):
            label_row = None
            for r in range(layout["header_row"], data_start):
                vals = [_sv(grid, r, c) for c in measure_cols]
                ints = []
                ok = True
                for v in vals:
                    if v is None or v == "":
                        continue
                    try:
                        fv = float(v)
                    except (ValueError, TypeError):
                        ok = False
                        break
                    if 1 <= fv <= 9 and fv == int(fv):
                        ints.append(fv)
                    else:
                        ok = False
                        break
                if ok and ints:
                    label_row = r
                    break
            if label_row is not None:
                for c in measure_cols[n_meas:]:
                    out["label_erase"].append({"sheet": sheet_idx,
                                               "r": label_row, "c": c})

        # 6) 空列整理：仅当有数据列 N > 模板读数位（如 >5）时，
        #    删第 N 列与结果列之间的空列、结果列移到第 N 列右边（v11.2）
        if result_col is not None and c_last is not None and measure_cols:
            template_n = len(measure_cols)
            if n_meas > template_n:
                target = c_last + 1
                if result_col > target:
                    out["gap_move"].append({"sheet": sheet_idx,
                                            "c_old": result_col, "c_new": target,
                                            "gap": result_col - target})
    else:
        max_col = _find_label_col(grid, nrows, ncols, _MAX_LABEL_GROUPS)
        min_col = _find_label_col(grid, nrows, ncols, _MIN_LABEL_GROUPS)
        data_start = 0
        data_end = nrows
    for r in range(data_start, data_end):
        if hidden and r in hidden:
            continue  # 筛选隐藏行不填 /
        for c in (max_col, min_col):
            if c is not None and not _sv(grid, r, c):
                out["slash_cells"].append((sheet_idx, r, c))

    # 7) 图号（WG No / 图号，v11）
    dn = _find_drawing_no(grid, nrows, ncols)
    if dn is not None:
        out["drawing_no"] = {"sheet": sheet_idx, "r": dn[0], "c": dn[1],
                             "value": dn[2]}
    return out


def _collect_norm(grids):
    """汇总各表规范化探测（grids: [(grid, nrows, ncols, hidden), ...]，仅可见表）。"""
    norm = {"customer_code": None, "final_cells": [], "final_missing": False,
            "slash_cells": [], "result_fill": [], "total_qty": [],
            "sample_qty": [], "drawing_no": None, "gap_move": [],
            "label_erase": [], "n_meas_total": 0, "split": []}
    sis = []
    total_n = 0
    for sheet_idx, (grid, nrows, ncols, hidden) in enumerate(grids, 1):
        si = _probe_norm_sheet(grid, nrows, ncols, sheet_idx, hidden)
        sis.append((sheet_idx, grid, nrows, ncols, si))
        total_n += si.get("n_meas", 0)
    norm["n_meas_total"] = total_n
    for sheet_idx, grid, nrows, ncols, si in sis:
        if norm["customer_code"] is None:
            norm["customer_code"] = si["customer_code"]
        if not norm["final_cells"]:
            norm["final_cells"] = si["final_cells"]
            norm["final_missing"] = si["final_missing"]
        norm["slash_cells"].extend(si["slash_cells"])
        norm["result_fill"].extend(si["result_fill"])
        norm["gap_move"].extend(si["gap_move"])
        norm["label_erase"].extend(si["label_erase"])
        if norm["drawing_no"] is None:
            norm["drawing_no"] = si["drawing_no"]
        # v11.3：每张表的 总数/检样数 = 全部可见表有数据测量列数之和
        if total_n > 0:
            ts = _find_total_sample(grid, nrows, ncols)
            for key, lst in (("total", norm["total_qty"]),
                             ("sample", norm["sample_qty"])):
                hit = ts.get(key)
                if hit is None:
                    continue
                oldv = str(hit[2]).strip()
                same = False
                try:
                    same = abs(float(oldv) - total_n) < 1e-9
                except (ValueError, TypeError):
                    same = (oldv == str(total_n))
                if not same:
                    lst.append({"sheet": sheet_idx, "r": hit[0], "c": hit[1],
                                "old": hit[2], "new": total_n})
    return norm


def _norm_summary(norm):
    """把规范化结果转成对外返回的摘要。"""
    cc = norm.get("customer_code")
    rf = norm.get("result_fill", [])
    return {
        "customer_code": {"found": cc is not None, "changed": cc is not None,
                          "value": CUSTOMER_CODE if cc is not None else None},
        "final_missing": bool(norm.get("final_missing")),
        "slash_filled": len(norm.get("slash_cells", [])),
        "result_filled_p": sum(1 for x in rf if x["val"] == "P"),
        "result_filled_f": sum(1 for x in rf if x["val"] == "F"),
        "result_filled_total": len(rf),
        "total_qty": [{"r": t["r"] + 1, "c": t["c"] + 1,
                       "old": t["old"], "new": t["new"]}
                      for t in norm.get("total_qty", [])],
        "sample_qty": [{"r": t["r"] + 1, "c": t["c"] + 1,
                        "old": t["old"], "new": t["new"]}
                       for t in norm.get("sample_qty", [])],
        "drawing_no": norm.get("drawing_no"),
        "gap_move": [{"c_old": g["c_old"] + 1, "c_new": g["c_new"] + 1,
                      "gap": g["gap"]}
                     for g in norm.get("gap_move", [])],
        "label_erase": [{"r": l["r"] + 1, "c": l["c"] + 1}
                        for l in norm.get("label_erase", [])],
    }


YELLOW_FILL_RGB = "FFFF00"   # 结果列缺失自动补 P 的黄底标记


def _move_result_col_openpyxl(ws, c_old, c_new):
    """openpyxl：把结果列（c_old）整列内容与主要格式移到 c_new，清空原列与中间列。

    列号均为 1-based。只清空值（保留列结构，避免破坏右侧其他列）。
    """
    if c_old <= c_new:
        return
    import copy
    for r in range(1, ws.max_row + 1):
        src = ws.cell(r, c_old)
        val = src.value
        if val is None:
            continue
        dst = ws.cell(r, c_new)
        dst.value = val
        try:
            dst.number_format = src.number_format
            dst.font = copy.copy(src.font)
            dst.border = copy.copy(src.border)
            dst.fill = copy.copy(src.fill)
            dst.alignment = copy.copy(src.alignment)
        except Exception:
            pass
        src.value = None
    for c in range(c_new + 1, c_old):
        for r in range(1, ws.max_row + 1):
            ws.cell(r, c).value = None


def _apply_gap_move_openpyxl(wb, norm, fails):
    """执行空列整理：结果列左移到正确位置，并同步 fails / result_fill 的列号。"""
    if not norm or not norm.get("gap_move"):
        return
    vis = [w for w in wb.worksheets if w.sheet_state == "visible"]
    for gm in norm["gap_move"]:
        try:
            ws = vis[gm["sheet"] - 1]
            _move_result_col_openpyxl(ws, gm["c_old"] + 1, gm["c_new"] + 1)
        except Exception:
            continue
        old_c1 = gm["c_old"] + 1
        for f in fails:
            if f.get("sheet") == gm["sheet"] and f.get("c") == old_c1:
                f["c"] = gm["c_new"] + 1
        for rf in norm.get("result_fill", []):
            if rf["sheet"] == gm["sheet"] and rf["c"] == gm["c_old"]:
                rf["c"] = gm["c_new"]


def _move_result_col_com(ws, c_old, c_new):
    """COM：结果列内容（值+主要格式）移到 c_new，清空原列与中间列。1-based。"""
    if c_old <= c_new:
        return
    try:
        nrows = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    except Exception:
        nrows = 300
    for r in range(1, nrows + 1):
        try:
            src = ws.Cells(r, c_old)
            val = src.Value
            if val is None:
                continue
            dst = ws.Cells(r, c_new)
            dst.Value = val
            try:
                dst.NumberFormat = src.NumberFormat
                dst.Font.Name = src.Font.Name
                dst.Font.Size = src.Font.Size
                dst.Font.Bold = src.Font.Bold
                dst.Interior.Color = src.Interior.Color
                dst.Interior.Pattern = src.Interior.Pattern
            except Exception:
                pass
            src.Value = None
            src.ClearFormats()
        except Exception:
            pass
    for c in range(c_new + 1, c_old):
        try:
            ws.Columns(c).ClearContents()
        except Exception:
            pass


def _apply_gap_move_com(wb, norm, fails):
    """COM：执行空列整理，同步 fails / result_fill 列号。"""
    if not norm or not norm.get("gap_move"):
        return
    for gm in norm["gap_move"]:
        try:
            ws = _com_visible_sheet(wb, gm["sheet"])
            if ws is None:
                continue
            _move_result_col_com(ws, gm["c_old"] + 1, gm["c_new"] + 1)
        except Exception:
            continue
        old_c1 = gm["c_old"] + 1
        for f in fails:
            if f.get("sheet") == gm["sheet"] and f.get("c") == old_c1:
                f["c"] = gm["c_new"] + 1
        for rf in norm.get("result_fill", []):
            if rf["sheet"] == gm["sheet"] and rf["c"] == gm["c_old"]:
                rf["c"] = gm["c_new"]


def _clear_fills_openpyxl(wb):
    """清除所有可见工作表的单元格填充（统一为无填充，视觉即白色）。

    返回被清除底色的单元格数；合并区域按其覆盖格数计。
    """
    total = 0
    for ws in wb.worksheets:
        if ws.sheet_state != "visible":
            continue
        tl_map, non_tl = {}, set()
        for rng in ws.merged_cells.ranges:
            cells = (rng.max_row - rng.min_row + 1) * (rng.max_col - rng.min_col + 1)
            tl_map[(rng.min_row, rng.min_col)] = cells
            for rr in range(rng.min_row, rng.max_row + 1):
                for cc in range(rng.min_col, rng.max_col + 1):
                    if (rr, cc) != (rng.min_row, rng.min_col):
                        non_tl.add((rr, cc))
        for row in ws.iter_rows():
            for cell in row:
                p = (cell.row, cell.column)
                if p in non_tl:
                    continue
                ft = cell.fill.fill_type if cell.fill is not None else None
                if ft not in (None, "none"):
                    total += tl_map.get(p, 1)
                    cell.fill = PatternFill(fill_type=None)
    return total


def _count_fills_openpyxl(wb):
    """只统计各可见工作表带填充底色的单元格数（本地、不清除），返回 {sheet_name: 数量}。

    合并区域按覆盖格数计；供 COM 批量清除时复用数量，避免逐格跨进程统计。
    """
    out = {}
    for ws in wb.worksheets:
        if ws.sheet_state != "visible":
            continue
        tl_map, non_tl = {}, set()
        for rng in ws.merged_cells.ranges:
            cells = (rng.max_row - rng.min_row + 1) * (rng.max_col - rng.min_col + 1)
            tl_map[(rng.min_row, rng.min_col)] = cells
            for rr in range(rng.min_row, rng.max_row + 1):
                for cc in range(rng.min_col, rng.max_col + 1):
                    if (rr, cc) != (rng.min_row, rng.min_col):
                        non_tl.add((rr, cc))
        n = 0
        for row in ws.iter_rows():
            for cell in row:
                p = (cell.row, cell.column)
                if p in non_tl:
                    continue
                ft = cell.fill.fill_type if cell.fill is not None else None
                if ft not in (None, "none"):
                    n += tl_map.get(p, 1)
        out[ws.title] = n
    return out


def _clear_fills_com(wb, counts=None):
    """清除所有可见工作表的单元格填充（直接批量，不统计数量，最快）。"""
    XL_NONE = -4142
    for i in range(1, int(wb.Sheets.Count) + 1):
        try:
            ws = wb.Sheets(i)
        except Exception:
            continue
        try:
            if int(ws.Visible) != -1:
                continue
        except Exception:
            continue
        try:
            ws.UsedRange.Interior.Pattern = XL_NONE
        except Exception:
            pass
    return 0


def _apply_norm_openpyxl(wb, norm):
    """把规范化应用到 openpyxl 工作簿（全表）。

    客户代码以文本格式写入（number_format="@"），保证前导 0 不被 Excel 省略；
    若目标单元格是合并区域的非左上角（MergedCell 只读），改写入合并区域左上角。
    """
    if not norm:
        return
    vis = [w for w in wb.worksheets if w.sheet_state == "visible"]
    n_clear = _clear_fills_openpyxl(wb)
    if n_clear:
        norm["fill_cleared"] = n_clear
    cc = norm.get("customer_code")
    if cc:
        try:
            ws = vis[cc["sheet"] - 1]
            r, c = cc["r"] + 1, cc["c"] + 1
            try:
                cell = ws.cell(r, c)
                cell.number_format = "@"   # 先设文本格式，再赋值（保前导0）
                cell.value = CUSTOMER_CODE
            except Exception:
                # 合并单元格非左上角：写入合并区域左上角
                for rng in ws.merged_cells.ranges:
                    if rng.min_row <= r <= rng.max_row and rng.min_col <= c <= rng.max_col:
                        try:
                            tl = ws.cell(rng.min_row, rng.min_col)
                            tl.number_format = "@"
                            tl.value = CUSTOMER_CODE
                        except Exception:
                            pass
                        break
        except Exception:
            pass
    for sheet_idx, r, c in norm.get("slash_cells", []):
        try:
            vis[sheet_idx - 1].cell(r + 1, c + 1).value = "/"
        except Exception:
            pass
    # v11：结果列缺失自动补 P/F（黄底标记「原文件未标结果列，已自动补全」）
    for rf in norm.get("result_fill", []):
        try:
            cell = vis[rf["sheet"] - 1].cell(rf["r"] + 1, rf["c"] + 1)
            cell.value = rf["val"]
            cell.fill = PatternFill(patternType="solid", fgColor=YELLOW_FILL_RGB)
        except Exception:
            pass
    # v11：总数/检样数更正为实际读数数据列数
    for t in norm.get("total_qty", []):
        try:
            vis[t["sheet"] - 1].cell(t["r"] + 1, t["c"] + 1).value = t["new"]
        except Exception:
            pass
    for t in norm.get("sample_qty", []):
        try:
            vis[t["sheet"] - 1].cell(t["r"] + 1, t["c"] + 1).value = t["new"]
        except Exception:
            pass
    # v11.2：删多余序号标签（清空单元格，数据区空列保留）
    for le in norm.get("label_erase", []):
        try:
            vis[le["sheet"] - 1].cell(le["r"] + 1, le["c"] + 1).value = None
        except Exception:
            pass


def _drop_merged_slash_openpyxl(norm, wb):
    """从 slash_cells 中剔除合并单元格（表头/页脚等结构性区域，不可写、不应填 /）。"""
    if not norm.get("slash_cells"):
        return norm
    merged = set()
    for sidx, ws in enumerate((w for w in wb.worksheets if w.sheet_state == "visible"), 1):
        for rng in ws.merged_cells.ranges:
            for r in range(rng.min_row - 1, rng.max_row):
                for c in range(rng.min_col - 1, rng.max_col):
                    merged.add((sidx, r, c))
    norm["slash_cells"] = [sc for sc in norm["slash_cells"] if sc not in merged]
    return norm


def _drop_merged_slash_com(norm, wb):
    """COM 版：剔除合并单元格（逐格查 MergeCells）。"""
    if not norm.get("slash_cells"):
        return norm
    by_sheet = {}
    for sc in norm["slash_cells"]:
        by_sheet.setdefault(sc[0], []).append(sc)
    keep = []
    for sidx, cells in by_sheet.items():
        try:
            ws = _com_visible_sheet(wb, sidx)
        except Exception:
            ws = None
        if ws is None:
            keep.extend(cells)
            continue
        for sc in cells:
            try:
                if not ws.Cells(sc[1] + 1, sc[2] + 1).MergeCells:
                    keep.append(sc)
            except Exception:
                keep.append(sc)
    norm["slash_cells"] = keep
    return norm


def _apply_norm_com(wb, norm):
    """把规范化应用到 COM 工作簿（全表）。

    客户代码以文本格式写入（NumberFormat="@"），保证前导 0 不被省略；
    若目标单元格在合并区域内，写合并区域左上角。
    """
    if not norm:
        return
    n_clear = _clear_fills_com(wb, norm.get("fill_counts"))
    if n_clear:
        norm["fill_cleared"] = n_clear
    cc = norm.get("customer_code")
    if cc:
        try:
            sh = _com_visible_sheet(wb, cc["sheet"])
            if sh is None:
                raise Exception("no visible sheet")
            cell = sh.Cells(cc["r"] + 1, cc["c"] + 1)
            try:
                cell.NumberFormat = "@"          # 先设文本格式，再赋值
                cell.Value = CUSTOMER_CODE
            except Exception:
                ma = cell.MergeArea
                ma.Cells(1, 1).NumberFormat = "@"
                ma.Cells(1, 1).Value = CUSTOMER_CODE
        except Exception:
            pass
    for sheet_idx, r, c in norm.get("slash_cells", []):
        try:
            sh = _com_visible_sheet(wb, sheet_idx)
            if sh is not None:
                sh.Cells(r + 1, c + 1).Value = "/"
        except Exception:
            pass
    # v11：结果列缺失自动补 P/F（黄底）
    for rf in norm.get("result_fill", []):
        try:
            sh = _com_visible_sheet(wb, rf["sheet"])
            if sh is not None:
                cell = sh.Cells(rf["r"] + 1, rf["c"] + 1)
                cell.Value = rf["val"]
                cell.Interior.Pattern = 1
                cell.Interior.Color = 0x00FFFF  # BGR 黄色
        except Exception:
            pass
    # v11：总数/检样数更正
    for t in norm.get("total_qty", []):
        try:
            sh = _com_visible_sheet(wb, t["sheet"])
            if sh is not None:
                sh.Cells(t["r"] + 1, t["c"] + 1).Value = t["new"]
        except Exception:
            pass
    for t in norm.get("sample_qty", []):
        try:
            sh = _com_visible_sheet(wb, t["sheet"])
            if sh is not None:
                sh.Cells(t["r"] + 1, t["c"] + 1).Value = t["new"]
        except Exception:
            pass
    # v11.2：删多余序号标签（清空单元格）
    for le in norm.get("label_erase", []):
        try:
            sh = _com_visible_sheet(wb, le["sheet"])
            if sh is not None:
                sh.Cells(le["r"] + 1, le["c"] + 1).Value = ""
        except Exception:
            pass


# ============================================================
# openpyxl 快路径（无公式 + 无复杂元素）
# ============================================================

def _openpyxl_risky(wb):
    """openpyxl 往返会丢失/无法保真的复杂元素检测。

    图表/图片/透视表/切片器；以及 AutoFilter（筛选结果 openpyxl 无法计算，
    必须降级 COM 用 Rows.Hidden 判断被筛掉的行）。
    """
    for ws in wb.worksheets:
        if (getattr(ws, "_charts", None) or getattr(ws, "_images", None)
                or getattr(ws, "_pivots", None) or getattr(ws, "_slicers", None)):
            return True
        try:
            if ws.auto_filter and ws.auto_filter.ref:
                return True
        except Exception:
            pass
    return False


# ---- 预加载缓存（v8：文件放入列表即预打开，审核时省去加载时间） ----
_preload_lock = threading.Lock()
_preloaded = {}          # abs_path -> {"wb": wb, "mtime": float}
_preload_pending = {}    # abs_path -> True（正在预加载）
_preload_consumed = set()  # 已被审核流程取走的 abs_path


def _load_openpyxl(path, use_cache=True):
    """加载工作簿；失败返回 None。.xlsm 保留宏。

    use_cache=True 时优先取预加载缓存（取走即删，审核完成后关闭）。
    """
    if openpyxl is None:
        return None
    p = os.path.abspath(path)
    if use_cache:
        with _preload_lock:
            hit = _preloaded.pop(p, None)
        if hit is not None:
            try:
                if abs(os.path.getmtime(p) - hit["mtime"]) < 1e-6:
                    return hit["wb"]
            except Exception:
                pass
            try:
                hit["wb"].close()
            except Exception:
                pass
        # 若该文件正在被预加载，标记为已消费，让预加载线程丢弃自己的 wb（避免泄漏）
        with _preload_lock:
            if p in _preload_pending:
                _preload_consumed.add(p)
    ext = os.path.splitext(p)[1].lower()
    try:
        return openpyxl.load_workbook(p, data_only=False,
                                      keep_vba=(ext == ".xlsm"))
    except Exception:
        return None


def preload_openpyxl(path):
    """预加载一个文件的 openpyxl 工作簿到缓存（仅快路径适用文件）。

    返回 True 表示已缓存或即将由其他线程缓存；False 表示不适用/失败。
    审核时 audit_file 会自动命中缓存，省去加载时间。
    """
    try:
        p = os.path.abspath(path)
        ext = os.path.splitext(p)[1].lower()
        if ext not in (".xlsx", ".xlsm"):
            return False
        with _preload_lock:
            if p in _preloaded or p in _preload_pending:
                return True
            _preload_pending[p] = True
        wb = _load_openpyxl(p, use_cache=False)
        if wb is None:
            with _preload_lock:
                _preload_pending.pop(p, None)
            return False
        if _openpyxl_risky(wb):
            try:
                wb.close()
            except Exception:
                pass
            with _preload_lock:
                _preload_pending.pop(p, None)
            return False
        with _preload_lock:
            _preload_pending.pop(p, None)
            if p in _preload_consumed:
                _preload_consumed.discard(p)
                try:
                    wb.close()
                except Exception:
                    pass
            else:
                _preloaded[p] = {"wb": wb, "mtime": os.path.getmtime(p)}
        return True
    except Exception:
        with _preload_lock:
            _preload_pending.pop(p, None)
        return False


def clear_preloaded(pred):
    """按谓词清理预加载缓存并关闭 wb（pred(abs_path) -> bool）。"""
    with _preload_lock:
        for p in list(_preloaded):
            if pred(p):
                hit = _preloaded.pop(p, None)
                if hit:
                    try:
                        hit["wb"].close()
                    except Exception:
                        pass


def _audit_openpyxl(input_path, wb=None):
    """openpyxl 快路径：扫描全部工作表并探测规范化信息。

    返回 (fails, reason, norm)；若文件含公式或 openpyxl 无法保真的复杂元素，返回 None
    （调用方应降级到 COM 路径）。
    wb 传入时使用外部已加载的工作簿（不负责关闭）；否则内部加载并在结束时关闭。
    """
    own = wb is None
    if own:
        wb = _load_openpyxl(input_path)
    if wb is None:
        return None
    try:
        if _openpyxl_risky(wb):
            return None
        all_fails = []
        all_reasons = []
        grids = []
        for ws in wb.worksheets:
            if ws.sheet_state != "visible":
                continue  # 隐藏工作表：用户看不到，不判断
            grid, nrows, ncols, has_formula = _openpyxl_grid(ws)
            if has_formula:
                return None  # 含公式 → 交给 COM 强制重算，保证判定一致
            if not grid or nrows == 0:
                continue
            # 手动隐藏的行（openpyxl 只能读 row_dimensions；筛选隐藏须靠
            # AutoFilter 降级 COM 由 Rows.Hidden 判断）
            hidden = set()
            try:
                for rr, dim in ws.row_dimensions.items():
                    if dim and dim.hidden:
                        hidden.add(rr - 1)
            except Exception:
                pass
            grids.append((grid, nrows, ncols, hidden))
            fails, reason = _scan_one_sheet(grid, nrows, ncols, hidden)
            if fails:
                sidx = len(grids)
                for f in fails:
                    f["sheet"] = sidx
                all_fails.extend(fails)
                all_reasons.append(f"工作表 {sidx}（{ws.title}）: {reason}")
        norm = _collect_norm(grids)
        # 剔除合并单元格（表头/页脚结构性区域，不填 /）
        _drop_merged_slash_openpyxl(norm, wb)
        if all_fails:
            reason = "；".join(all_reasons)
        else:
            reason = "未检出不合格（openpyxl 快路径，已扫描全部工作表）"
        return all_fails, reason, norm
    finally:
        if own:
            try:
                wb.close()
            except Exception:
                pass


def _mark_openpyxl(ws, fails):
    """对 openpyxl 工作表中的目标单元格应用红底色 + 标 F（字体保持不变）。"""
    _PASS_TOKENS = ("p", "pass", "ok", "合格", "good", "accept", "y", "yes")
    for f in fails:
        r, c = f["r"], f["c"]
        try:
            cell = ws.cell(r, c)
            cell.fill = PatternFill(patternType="solid", fgColor=RED_FILL_RGB)
            # 用户要求：只把单元格底色标红，字体保持不变
            if f.get("kind") == "reading" and f.get("is_result"):
                v = cell.value
                if isinstance(v, str) and v.startswith("="):
                    continue  # 公式单元格保持原样
                s = "" if v is None else str(v).strip()
                t2 = _strip_symbols(s.lower())
                if t2 in _PASS_TOKENS or s == "":
                    cell.value = "F"
        except Exception:
            pass


def _save_openpyxl(input_path, fails, out_path, norm=None, wb=None):
    """用 openpyxl 打开源文件（或使用传入 wb），按 fails 标红标 F，并按 norm 规范化，保存到 out_path。"""
    own = wb is None
    if own:
        wb = _load_openpyxl(input_path)
    if wb is None:
        raise RuntimeError(f"openpyxl 无法打开文件: {input_path}")
    try:
        # v11：空列整理（结果列左移）→ 标红/补值使用移动后的新列号
        if norm:
            _apply_gap_move_openpyxl(wb, norm, fails)
        # 先规范化（结果列补值+黄底/客户代码/填/ /总数更正），再标红（红覆盖黄）
        _apply_norm_openpyxl(wb, norm)
        by_sheet = {}
        for f in fails:
            by_sheet.setdefault(f.get("sheet", 1), []).append(f)
        vis = [w for w in wb.worksheets if w.sheet_state == "visible"]
        for sidx, sheet_fails in by_sheet.items():
            if 1 <= sidx <= len(vis):
                _mark_openpyxl(vis[sidx - 1], sheet_fails)
        wb.save(out_path)
    finally:
        if own:
            try:
                wb.close()
            except Exception:
                pass


# ============================================================
# COM 路径（有公式 / .xls / 复杂元素）
# ============================================================

def _scan_com(xl, path):
    """COM 打开（临时副本 + 强制重算），扫描全部工作表，返回 (fails, reason)。"""
    tmp_path = None
    wb = None
    try:
        ext = os.path.splitext(path)[1].lower() or ".xlsx"
        tmp_path = _make_tmp("_audit_read_", ext)
        shutil.copy2(os.path.abspath(path), tmp_path)
        wb = xl.Workbooks.Open(os.path.abspath(tmp_path), ReadOnly=False)
        _recalc(xl)
        fails, reason, _norm = _scan_workbook(xl, wb)
        return fails, reason
    finally:
        try:
            if wb is not None:
                wb.Close(False)  # 不保存任何修改
        except Exception:
            pass
        if tmp_path and os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except Exception:
                pass


def _mark_save_com(xl, input_path, fails, out_path, norm=None):
    """COM 打开源文件（读写临时副本），按 fails 标红标 F、按 norm 规范化，SaveAs 到 out_path。"""
    tmp_in = None
    wb = None
    try:
        ext = os.path.splitext(out_path)[1].lower() or ".xlsx"
        tmp_in = _make_tmp("_audit_mark_", ext)
        shutil.copy2(os.path.abspath(input_path), tmp_in)
        wb = xl.Workbooks.Open(os.path.abspath(tmp_in), ReadOnly=False)
        _recalc(xl)

        # v11：空列整理（结果列左移）→ 标红/补值使用移动后的新列号
        if norm:
            _apply_gap_move_com(wb, norm, fails)
        # 先规范化（补值+黄底），再标红（红覆盖黄）
        _apply_norm_com(wb, norm)
        # 按工作表分组标红
        by_sheet = {}
        for f in fails:
            sheet_idx = f.get("sheet", 1)
            by_sheet.setdefault(sheet_idx, []).append(f)
        for sheet_idx, sheet_fails in sorted(by_sheet.items()):
            ws = _com_visible_sheet(wb, sheet_idx)
            if ws is not None:
                _mark_fails(ws, sheet_fails)

        fmt = _EXT_FMT.get(ext, 51)
        wb.SaveAs(os.path.abspath(out_path), FileFormat=fmt)
    finally:
        try:
            if wb is not None:
                wb.Close(False)
        except Exception:
            pass
        if tmp_in and os.path.exists(tmp_in):
            try:
                os.remove(tmp_in)
            except Exception:
                pass


# ============================================================
# v11.3 宽表拆分：测量数据列超过每组位数（模板通常5列）时，
# 每 k 列拆成一张标准表，整表复制（保真），结果列随组移动、公式重算
# ------------------------------------------------------------
def _find_group_k(ws, grid, nrows, ncols, layout):
    """动态识别每组读数位 k（不写死5）。
    1) Actual Readings/测量数据 合并标题的列跨度（>=2 采信）；
    2) 表头区测量列内连续序号 1,2,3... 的个数；
    3) 兜底 5。
    """
    first = (layout.get("instr_col") or 0) + 1
    result = layout.get("result_col")
    rng_end = result if result is not None else ncols
    # 1) 合并标题宽度
    for r in range(min(nrows, 30)):
        for c in range(ncols):
            t = _sv(grid, r, c).lower()
            if "actual reading" in t or "测量数据" in t:
                for rng in ws.merged_cells.ranges:
                    if (rng.min_row <= r + 1 <= rng.max_row
                            and rng.min_col <= c + 1 <= rng.max_col):
                        w = rng.max_col - rng.min_col + 1
                        if w >= 2:
                            return w
    # 2) 连续序号
    hr = layout.get("header_row") or 0
    ds = layout.get("data_start") or nrows
    for r in range(hr, ds):
        seq = []
        for c in range(first, rng_end):
            v = _sv(grid, r, c)
            try:
                fv = float(v)
            except (ValueError, TypeError):
                continue
            if fv == int(fv) and 1 <= fv <= 99:
                seq.append(int(fv))
        if len(seq) >= 2 and seq == list(range(1, len(seq) + 1)):
            return len(seq)
    return 5


def _plan_splits(wb):
    """openpyxl 预探测：返回需要拆分的可见表计划列表。
    每项 {name, first_col(1b), k, result_col(1b), n_groups, label_row(1b),
          last_row(1b), M}；无需拆分返回 []。
    """
    plans = []
    shifts = []
    for ws in wb.worksheets:
        if ws.sheet_state != "visible":
            continue
        try:
            grid, nrows, ncols, _f = _openpyxl_grid(ws)
        except Exception:
            continue
        if not grid:
            continue
        lay = _detect_layout(grid, nrows, ncols)
        if not lay:
            continue
        ic, rc, ds = lay.get("instr_col"), lay.get("result_col"), lay.get("data_start")
        if ic is None or rc is None or ds is None:
            continue
        first = ic + 1                      # 0-based 首测量列
        if rc <= first:
            continue
        k = _find_group_k(ws, grid, nrows, ncols, lay)
        c_last = None
        for c in range(first, rc):
            if any(_sv(grid, r, c) for r in range(ds, nrows)):
                c_last = c
        if c_last is None:
            continue
        M = c_last - first + 1              # 实际用到的物理测量列数
        if M <= k:
            # 测量列未超组位：结果列若被推到标准位置（第 k 位右边）以右，记录归位
            std = first + k                 # 0-based 标准结果列
            if rc > std:
                areas = [(a.min_row, a.min_col, a.max_row, a.max_col)
                         for a in ws.merged_cells.ranges]
                shifts.append({"name": ws.title, "d1": std + 1, "d2": rc,
                               "areas": areas})
            continue
        last = ds
        for r in range(ds, nrows):
            # 末行取整个测量区任意列有值的最大行（底部合并说明行测量列为空，不计入）
            if any(_sv(grid, r, c) for c in range(first, rc)):
                last = r
        plans.append({
            "name": ws.title, "first_col": first + 1, "k": k,
            "result_col": rc + 1, "n_groups": (M + k - 1) // k,
            "label_row": ds, "last_row": last + 1, "M": M,
        })
    return plans, shifts


def _shift_cols_com(ws, d1, d2, pre_areas=None):
    """物理删除整列 [d1, d2]（1-based），安全处理横跨合并区。

    pre_areas：可选合并区坐标 [(r1,c1,r2,c2)]（openpyxl 本地预读），传入免逐格遍历；
    否则回退逐格收集。解除与删除区相交的合并、把将随删除丢失的左上角值移到右侧存活列，
    删列后按位移重建合并。用于把结果列从异常靠右位置移回标准位置。
    返回删除的列数。
    """
    if pre_areas is not None:
        areas = [tuple(a) for a in pre_areas]
    else:
        seen, areas = set(), []
        for cell in ws.UsedRange.Cells:
            try:
                if cell.MergeCells:
                    ma = cell.MergeArea
                    k = (ma.Row, ma.Column, ma.Row + ma.Rows.Count - 1,
                         ma.Column + ma.Columns.Count - 1)
                    if k not in seen:
                        seen.add(k)
                        areas.append(k)
            except Exception:
                pass
    delta = d2 - d1 + 1
    hit = []
    for (r1, c1, r2, c2) in areas:
        if c2 < d1 or c1 > d2:
            continue
        hit.append((r1, c1, r2, c2))
        v = ws.Cells(r1, c1).Value
        ws.Range(ws.Cells(r1, c1), ws.Cells(r2, c2)).UnMerge()
        if d1 <= c1 <= d2 and c2 > d2 and v not in (None, ""):
            ws.Cells(r1, d2 + 1).Value = v
    ws.Range(ws.Cells(1, d1), ws.Cells(1, d2)).EntireColumn.Delete()
    for (r1, c1, r2, c2) in hit:
        if c1 < d1 and c2 > d2:
            n1, n2 = c1, c2 - delta
        elif c1 < d1 and c2 <= d2:
            n1, n2 = c1, d1 - 1
        elif d1 <= c1 <= d2 and c2 > d2:
            n1, n2 = d1, c2 - delta
        else:
            n1 = n2 = None
        if n1 and n2 and n2 >= n1 and not (n1 == n2):
            ws.Range(ws.Cells(r1, n1), ws.Cells(r2, n2)).Merge()
    return delta


def _split_wb_com(xl, input_path, out_path, plans, shifts=None, ext=".xlsx"):
    """COM 打开源文件，按 plans 把宽表拆成多张标准表，重算后 SaveAs 到 out_path。

    返回 [(原表名, 组数), ...]。
    """
    wb = None
    made = []
    shift_made = []
    shifts = shifts or []
    try:
        wb = xl.Workbooks.Open(os.path.abspath(input_path), ReadOnly=True, UpdateLinks=0)
        _recalc(xl)
        # 先做结果列归位（与拆分作用于不同的表）
        for sh in shifts:
            wsh = wb.Sheets(sh["name"])
            n = _shift_cols_com(wsh, sh["d1"], sh["d2"], sh.get("areas"))
            shift_made.append((sh["name"], n))
        for plan in plans:
            ws = wb.Sheets(plan["name"])
            orig = ws.Name
            first, k = plan["first_col"], plan["k"]
            res = plan["result_col"]
            lab, lr = plan["label_row"], plan["last_row"]
            ng = plan["n_groups"]

            # T1：复制原表 → 删结果列右侧（辅助列）→ 删 first+k..res-1 多余测量列
            ws.Copy(None, ws)
            t1 = wb.ActiveSheet
            used_last = t1.UsedRange.Column + t1.UsedRange.Columns.Count - 1
            if used_last >= res + 1:
                t1.Range(t1.Cells(1, res + 1), t1.Cells(1, used_last)) \
                    .EntireColumn.Delete()
            if res - 1 >= first + k:
                t1.Range(t1.Cells(1, first + k), t1.Cells(1, res - 1)) \
                    .EntireColumn.Delete()
            t1.Name = "%s-1" % orig
            group_sheets = [t1]

            # T2..Tn：复制 T1（标准布局），替换测量区为该组数据
            for g in range(1, ng):
                t1.Copy(None, group_sheets[-1])
                tg = wb.ActiveSheet
                tg.Name = "%s-%d" % (orig, g + 1)
                s0 = first + g * k
                tg.Range(tg.Cells(lab, first),
                         tg.Cells(lr, first + k - 1)).ClearContents()
                src_end = min(s0 + k - 1, res - 1)
                ws.Range(ws.Cells(lab, s0), ws.Cells(lr, src_end)).Copy(
                    Destination=tg.Cells(lab, first))
                group_sheets.append(tg)
            ws.Delete()                       # 删除原宽表
            made.append((orig, ng))
        xl.CalculateFull()
        fmt = _EXT_FMT.get((ext or ".xlsx").lower(), 51)
        wb.SaveAs(os.path.abspath(out_path), FileFormat=fmt)
    finally:
        if wb is not None:
            try:
                wb.Close(False)
            except Exception:
                pass
    return made, shift_made


def _audit_com_file(xl, input_path, out_dir, base, ext, fill_counts=None,
                    pending_shifts=None):
    """单次打开完成「重算 → 扫描 → 归位 → 规范化 → 标红 → SaveAs」（v14 合并，省去二次打开）。

    返回 (fails, reason, norm, output, marked, copied)。
    扫描为纯读；规范化/标红在同一工作簿完成，最后 SaveAs 到输出目录（不覆盖、同名递增）。
    """
    tmp_path = None
    wb = None
    prev = None
    try:
        ext_ok = ext or ".xlsx"
        tmp_path = _make_tmp("_audit_", ext_ok)
        shutil.copy2(os.path.abspath(input_path), tmp_path)
        prev = _com_speed_on(xl)
        wb = xl.Workbooks.Open(os.path.abspath(tmp_path), ReadOnly=False, UpdateLinks=0)
        # 纯归位：先删空列（复用预读合并区），再重算、扫描
        internal_shift = []
        for sh in (pending_shifts or []):
            wsh = wb.Sheets(sh["name"])
            n = _shift_cols_com(wsh, sh["d1"], sh["d2"], sh.get("areas"))
            internal_shift.append((sh["name"], n))
        _recalc(xl)
        all_fails, reason, norm = _scan_workbook(xl, wb)
        if internal_shift:
            norm["shift_made"] = internal_shift
        if fill_counts:
            norm["fill_counts"] = fill_counts

        is_fail = len(all_fails) > 0
        out_path = os.path.join(out_dir, base + ext)
        if is_fail:
            out_path = _with_fail_prefix(out_path)

        # 同一工作簿内：空列归位 → 规范化（清底色/补值/黄底）→ 标红（红覆盖黄）
        _apply_gap_move_com(wb, norm, all_fails)
        _apply_norm_com(wb, norm)
        by_sheet = {}
        for f in all_fails:
            by_sheet.setdefault(f.get("sheet", 1), []).append(f)
        for sheet_idx, sheet_fails in sorted(by_sheet.items()):
            sh = _com_visible_sheet(wb, sheet_idx)
            if sh is not None:
                _mark_fails(sh, sheet_fails)

        fmt = _EXT_FMT.get(ext_ok.lower(), 51)
        output = _write_output(
            lambda t: wb.SaveAs(os.path.abspath(t), FileFormat=fmt), out_path)
        return all_fails, reason, norm, output, is_fail, (not is_fail)
    finally:
        try:
            if wb is not None:
                wb.Close(False)
        except Exception:
            pass
        if prev is not None:
            _com_speed_off(xl, prev)
        if tmp_path and os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except Exception:
                pass


# ============================================================
# 公共扫描入口
# ============================================================

def scan_fails(path):
    """扫描所有工作表中的不合格单元格，返回 (fails, reason)。

    自动路由：无公式的 .xlsx/.xlsm → openpyxl 快路径；其余 → COM 慢路径。
    """
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        r = _audit_openpyxl(path)
        if r is not None:
            return r[0], r[1]
    sess = ExcelSession()
    try:
        return _scan_com(sess.get(), path)
    finally:
        sess.quit()


def detect_fail(path):
    """判定文件是否含不合格数据。返回 (is_fail, reason)。（兼容旧调用）"""
    fails, reason = scan_fails(path)
    return (len(fails) > 0, reason)


# ============================================================
# 标红（COM 单元格级，供 _mark_save_com 使用）
# ============================================================

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
    """打开源文件（读写），按 fail 里的 sheet 标红标 F，SaveAs 到 out_path。（兼容旧调用）"""
    sess = ExcelSession()
    try:
        _mark_save_com(sess.get(), input_path, fails, out_path)
    finally:
        sess.quit()


# ============================================================
# 输出命名与重试
# ============================================================

def _with_dwg_prefix(out_path):
    """在文件名最前加 图号错误_（已有则不加）。"""
    d, f = os.path.split(out_path)
    if f.startswith("图号错误_"):
        return out_path
    return os.path.join(d, "图号错误_" + f)


def _dwg_bad(norm, base):
    """图号不在文件名子串中 → True（需加 图号错误_ 前缀）。"""
    dn = (norm or {}).get("drawing_no")
    if not dn:
        return False
    dwg = str(dn.get("value", "")).strip()
    return bool(dwg) and dwg not in base


def _v11_apply_result_f(fails, reason, norm):
    """结果列缺失自动补出的 F → 并入 fails（判不合格并标红）。"""
    rf = (norm or {}).get("result_fill", [])
    add = [x for x in rf if x["val"] == "F"]
    if not add:
        return fails, reason, norm
    new_fails = list(fails)
    for x in add:
        new_fails.append({
            "r": x["r"] + 1, "c": x["c"] + 1, "value": "F",
            "kind": "reading", "is_result": True, "sheet": x["sheet"],
        })
    note = "结果列缺失：按读数与公差自动判定，有 %d 行越界（自动标 F）" % len(add)
    if not reason or reason.startswith("未检出"):
        reason = note
    else:
        reason = note + "；" + reason
    return new_fails, reason, norm


def _build_log(fails, reason, norm, base):
    """组装文件级操作日志（v11：点击结果行查看详情）。"""
    log = []
    norm = norm or {}
    cc = norm.get("customer_code")
    if cc:
        log.append({"level": "info", "op": "客户代码",
                    "msg": "已改为 %s（原: %s）" % (CUSTOMER_CODE, cc.get("old"))})
    n_slash = len(norm.get("slash_cells", []))
    if n_slash:
        log.append({"level": "info", "op": "上下限空白",
                    "msg": "已填 / × %d 个" % n_slash})
    if norm.get("final_missing"):
        log.append({"level": "err", "op": "最终判定",
                    "msg": "Passed/Failed/UAI 三选项中有空白格未填写"})
    rf = norm.get("result_fill", [])
    if rf:
        p = sum(1 for x in rf if x["val"] == "P")
        f = sum(1 for x in rf if x["val"] == "F")
        if f:
            log.append({"level": "err", "op": "结果列缺失",
                        "msg": "原文件未标结果列，已自动补全（P×%d、F×%d），其中 %d 行越界" % (p, f, f)})
        else:
            log.append({"level": "warn", "op": "结果列缺失",
                        "msg": "原文件未标结果列，已自动补全（全部 P，黄色标记）"})
    tq = norm.get("total_qty", [])
    sq = norm.get("sample_qty", [])
    if tq or sq:
        def _uniq(lst, label):
            seen = {}
            for t in lst:
                k = (str(t.get("old")), t.get("new"))
                seen[k] = seen.get(k, 0) + 1
            parts = []
            for (o, n), cnt in sorted(seen.items()):
                s = "%s %s → %s" % (label, o, n)
                if cnt > 1:
                    s += "（%d 张表）" % cnt
                parts.append(s)
            return parts
        parts = _uniq(tq, "总数") + _uniq(sq, "检样数")
        log.append({"level": "warn", "op": "总数/检样数",
                    "msg": "、".join(parts) + "（更正为实际测量数据列数）"})
    dn = norm.get("drawing_no")
    if dn:
        dwg = str(dn.get("value", "")).strip()
        if dwg and dwg in base:
            log.append({"level": "info", "op": "图号",
                        "msg": "图号 %s 在文件名中，一致" % dwg})
        else:
            log.append({"level": "err", "op": "图号",
                        "msg": "图号 %s 不在文件名「%s」中，已加「图号错误_」前缀" % (dwg or "(空)", base)})
    if norm.get("label_erase"):
        log.append({"level": "info", "op": "序号标签",
                    "msg": "删除多余序号标签 %d 个（实际有数据列较少，空白数据列保留）" % len(norm["label_erase"])})
    for g in norm.get("gap_move", []):
        log.append({"level": "warn", "op": "空列整理",
                    "msg": "结果列与数据列之间删除 %d 个空列，结果列移至第 %d 列" % (g.get("gap"), g.get("c_new") + 1)})
    for orig, ng in (norm.get("split_made") or []):
        log.append({"level": "warn", "op": "宽表拆分",
                    "msg": "测量列超过每组位数，已将工作表「%s」拆分为 %d 张标准表（%s-1 ～ %s-%d），结果列随组判定"
                           % (orig, ng, orig, orig, ng)})
    for nm, n in (norm.get("shift_made") or []):
        log.append({"level": "warn", "op": "结果列归位",
                    "msg": "工作表「%s」结果列位置异常（右侧存在空白列），已删除 %d 个空白列并将结果列移回标准位置"
                           % (nm, n)})
    fc = norm.get("fill_cleared") or 0
    if fc:
        log.append({"level": "warn", "op": "清除底色",
                    "msg": "源文件有 %d 个单元格带填充底色（人工标记 / 高亮等），已统一清除为白色（无填充）"
                           % fc})
    if fails:
        log.append({"level": "err", "op": "越界判定", "msg": reason})
    if not log:
        log.append({"level": "info", "op": "结论", "msg": "未发现异常"})
    return log


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
    """写输出：目标已存在（含源文件本身）时不覆盖，自动换 name(1)、name(2)... 后写入。

    用户要求：不管合不合格都不覆盖原文件，只输出到指定目录；同名文件在名字后加 (1) 区分。
    write_fn(target) 成功返回 target；全部尝试失败抛最后一次异常。
    """
    last_err = None
    target = first_target
    for attempt in range(max_try):
        if attempt > 0:
            target = _alt_path(first_target, attempt)  # attempt=1 -> name(1).ext
        if os.path.exists(target):
            continue  # 已存在（或与源文件同名）→ 换名，绝不覆盖
        try:
            write_fn(target)
            return target
        except Exception as e:  # noqa: BLE001
            last_err = e
    raise last_err


# ============================================================
# 主流程
# ============================================================

def audit_file(input_path, out_dir=None, on_progress=None, session=None, pre=None):
    """审核单个文件并输出。

    session: 可选的 ExcelSession（批量审核时传入以复用 Excel 进程）；
             不传时自动创建并在结束时退出。
    pre: 可选预检结果（主进程 precheck 产出），提供 plans/shifts/fill_counts
         时跳过重复的 openpyxl 布局探测，供多进程 worker 复用。
    v11.3：测量列超过每组位数（模板通常5列）的宽表，先拆分为多张标准表再审核。
    """
    input_path = os.path.abspath(input_path)
    if not os.path.isfile(input_path):
        raise FileNotFoundError(input_path)
    base, ext = os.path.splitext(os.path.basename(input_path))
    od = os.path.abspath(out_dir) if out_dir else os.path.dirname(input_path)
    os.makedirs(od, exist_ok=True)

    # ---- v15.1 格式自动修复：老 .xls(OLE2) 伪装成 .xlsx 时，先转成真 xlsx 再处理 ----
    # 这类文件 Excel 直接报"格式/扩展名无效"拒绝打开；这里探测文件头自动转换，不报错。
    sess = session
    own_sess = False
    conv_tmp = None
    conv_msgs = []
    try:
        with open(input_path, "rb") as _f:
            _magic = _f.read(8)
    except Exception:
        _magic = b""
    if _magic[:4] == b"\xD0\xCF\x11\xE0":
        # 老 .xls(OLE2)：用 worker 已连好的 WPS/Excel 实例转换为 .xlsx，
        # 转完走 openpyxl 快路径（不再依赖 COM）；转换用 ready 实例，不额外冷启动。
        if sess is None:
            sess = ExcelSession()
            own_sess = True
        try:
            xl = sess.get()
            conv_tmp = _make_tmp("_conv_", ".xls")
            shutil.copyfile(input_path, conv_tmp)
            cwb = xl.Workbooks.Open(os.path.abspath(conv_tmp), ReadOnly=True, UpdateLinks=0)
            conv_xlsx = _make_tmp("_conv_", ".xlsx")
            cwb.SaveAs(os.path.abspath(conv_xlsx), FileFormat=51)
            cwb.Close(False)
            work_path = conv_xlsx
            ext = ".xlsx"
            conv_msgs.append("原文件为老 .xls 格式，已自动转换为 .xlsx 后继续处理")
        except Exception:
            work_path = input_path
    else:
        work_path = input_path

    # ---- v11.3 宽表拆分预处理 ----
    split_tmp = None
    split_made = []
    shift_made = []
    fill_counts = None
    reuse_wb = None
    pending_shifts = None
    if conv_msgs and on_progress:
        try:
            on_progress(0, conv_msgs[-1])
        except Exception:
            pass
    if ext.lower() in (".xlsx", ".xlsm"):
        plans, shifts = [], []
        wbp = None
        pre = pre or {}
        if pre.get("plans") is not None:
            # 主进程已预检：直接复用布局计划（快路径候选仍需 wbp）
            plans = pre.get("plans") or []
            shifts = pre.get("shifts") or []
            fill_counts = pre.get("fill_counts")
            if not plans and not shifts:
                try:
                    wbp = _load_openpyxl(input_path, use_cache=False)
                except Exception:
                    wbp = None
        else:
            try:
                wbp = _load_openpyxl(input_path, use_cache=False)
                if wbp is not None:
                    plans, shifts = _plan_splits(wbp)
            except Exception:
                plans, shifts = [], []
            # 本地预统计底色（不拆分时；归位仅删空列、不改变底色数量，表名一致）
            if wbp is not None and not plans:
                try:
                    fill_counts = _count_fills_openpyxl(wbp)
                except Exception:
                    fill_counts = None
        if plans:
            # 宽表拆分（可能同时含其他表归位）：COM 中转 split_tmp
            if wbp is not None:
                try:
                    wbp.close()
                except Exception:
                    pass
            if sess is None:
                sess = ExcelSession()
                own_sess = True
            try:
                xl = sess.get()
                split_tmp = _make_tmp("_split_", ext)
                split_made, shift_made = _split_wb_com(
                    xl, input_path, split_tmp, plans, shifts, ext)
                work_path = split_tmp
            except Exception:
                work_path = input_path
                split_made, shift_made = [], []
        elif shifts:
            # 纯归位：不中转，归位并入后续单次 COM 打开
            pending_shifts = shifts
            if wbp is not None:
                try:
                    wbp.close()
                except Exception:
                    pass
        else:
            # 不拆分/不归位：复用预处理 wbp 给快路径判定，省去再次加载
            reuse_wb = wbp

    try:
        # ---- openpyxl 快路径（无公式 + 无复杂元素；有删空列归位时跳过，走 COM） ----
        if ext.lower() in (".xlsx", ".xlsm") and not pending_shifts:
            wb = reuse_wb if reuse_wb is not None else _load_openpyxl(work_path)
            if wb is not None:
                r = _audit_openpyxl(work_path, wb=wb)
                if r is not None:
                    fails, reason, norm = r
                    fails, reason, norm = _v11_apply_result_f(fails, reason, norm)
                    norm["split_made"] = split_made
                    norm["shift_made"] = shift_made
                    is_fail = len(fails) > 0
                    out_path = os.path.join(od, base + ext)
                    if is_fail:
                        out_path = _with_fail_prefix(out_path)
                    if _dwg_bad(norm, base):
                        out_path = _with_dwg_prefix(out_path)
                    marked = copied = False
                    try:
                        if is_fail:
                            output = _write_output(
                                lambda t: _save_openpyxl(work_path, fails, t, norm, wb=wb), out_path)
                            marked = True
                        else:
                            # 合格文件也做规范化（客户代码→0998 / 上下限空白填 /）
                            output = _write_output(
                                lambda t: _save_openpyxl(work_path, [], t, norm, wb=wb), out_path)
                            copied = True
                    finally:
                        try:
                            wb.close()
                        except Exception:
                            pass
                    if on_progress:
                        on_progress(100, f"{'不合格' if is_fail else '合格'} -> {os.path.basename(output)}"
                                        + ("（已标红标F）" if marked else ""))
                    return {"input": input_path, "output": output, "is_fail": is_fail,
                            "reason": reason, "copied": copied, "marked": marked,
                            "norm": _norm_summary(norm),
                            "log": _build_log(fails, reason, norm, base)
                                  + ([{"level": "info", "op": "格式修复", "msg": m} for m in conv_msgs] if conv_msgs else [])}
                # 快路径不适用（公式/复杂元素）→ 降级 COM，关闭已加载的 wb
                try:
                    wb.close()
                except Exception:
                    pass

        # ---- COM 慢路径（有公式 / .xls / 复杂元素 / 已拆分） ----
        own = sess is None
        if sess is None:
            sess = ExcelSession()
            own_sess = True
        try:
            xl = sess.get()
            fails, reason, norm, output, marked, copied = _audit_com_file(
                xl, work_path, od, base, ext, fill_counts, pending_shifts)
            fails, reason, norm = _v11_apply_result_f(fails, reason, norm)
            norm["split_made"] = split_made
            norm["shift_made"] = norm.get("shift_made") or shift_made
            is_fail = len(fails) > 0
            # 图号错误 → 输出文件加 图号错误_ 前缀（不覆盖已写文件，改名）
            if _dwg_bad(norm, base):
                pref = _with_dwg_prefix(output)
                n = 0
                newout = pref
                while os.path.exists(newout) and os.path.abspath(newout) != os.path.abspath(output):
                    n += 1
                    newout = _alt_path(pref, n)
                if os.path.abspath(newout) != os.path.abspath(output):
                    shutil.move(output, newout)
                    output = newout
        finally:
            if own:
                sess.quit()

        if on_progress:
            on_progress(100, f"{'不合格' if is_fail else '合格'} -> {os.path.basename(output)}"
                            + ("（已标红标F）" if marked else ""))
        return {"input": input_path, "output": output, "is_fail": is_fail,
                "reason": reason, "copied": copied, "marked": marked,
                "norm": _norm_summary(norm),
                "log": _build_log(fails, reason, norm, base)
                      + ([{"level": "info", "op": "格式修复", "msg": m} for m in conv_msgs] if conv_msgs else [])}
    finally:
        if split_tmp and os.path.exists(split_tmp):
            try:
                os.remove(split_tmp)
            except Exception:
                pass
        if own_sess and sess is not None and session is None:
            try:
                sess.quit()
            except Exception:
                pass


def audit_paths(paths, out_dir=None, on_progress=None):
    """批量审核（文件或目录混合）。整批共享一个 Excel COM 会话。"""
    import glob

    files = []
    od = os.path.abspath(out_dir) if out_dir else None
    for p in paths:
        p = os.path.abspath(p)
        if os.path.isdir(p):
            for ext in ("*.xls", "*.xlsx", "*.xlsm"):
                for f in glob.glob(os.path.join(p, "**", ext), recursive=True):
                    f = os.path.abspath(f)
                    # 排除已生成的审核结果（避免二次扫描）
                    if os.path.basename(f).startswith("不合格_"):
                        continue
                    # 排除输出目录子树（输出目录恰好是被扫目录时由前缀规则兜底）
                    if od and od != p and f.startswith(od + os.sep):
                        continue
                    files.append(f)
        elif os.path.isfile(p):
            files.append(p)
    files = sorted(set(files))
    total = len(files)
    fail = 0
    sess = ExcelSession()
    try:
        for i, f in enumerate(files, 1):
            try:
                res = audit_file(f, out_dir=out_dir, on_progress=on_progress,
                                 session=sess)
                if res["is_fail"]:
                    fail += 1
            except Exception as e:
                if on_progress:
                    on_progress(100, f"失败 {os.path.basename(f)}: {e}")
            if on_progress:
                on_progress(int(i / total * 100) if total else 100, f"已处理 {i}/{total}")
    finally:
        sess.quit()
    return {"total": total, "fail": fail}





# ============================================================
# v15 多进程批量池（每个 worker 一个独立 Excel 进程）
# ============================================================

def _move_output(src, dst_dir):
    """把 worker 临时输出移到最终目录；同名自动 name(1)、name(2)... 递增，绝不覆盖。"""
    os.makedirs(dst_dir, exist_ok=True)
    base = os.path.basename(src)
    target = os.path.join(dst_dir, base)
    n = 0
    while os.path.exists(target):
        n += 1
        target = _alt_path(os.path.join(dst_dir, base), n)
    shutil.move(src, target)
    return target


def run_worker_mode(out_dir=None):
    """子进程 worker 主循环：stdin 逐行收 JSON 任务 {file}，stdout 逐行回 JSON 事件。

    事件：ready(含 excel_pid) / phase(文件内进度) / done(ok+result 或 error)。
    Excel 进程 PID 上报给父进程，供超时杀进程。
    """
    import win32process as _wp
    if win32com is not None:
        try:
            import pythoncom
            pythoncom.CoInitialize()
        except Exception:
            pass
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    try:
        sys.stdin.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

    def emit(obj):
        sys.stdout.write(json.dumps(obj, ensure_ascii=False, default=str) + "\n")
        sys.stdout.flush()

    sess = ExcelSession()
    excel_pid = None
    app_name = ""
    app_version = ""
    try:
        xl = sess.get()
        try:
            h = xl.Hwnd
            if h:
                _, excel_pid = _wp.GetWindowThreadProcessId(h)
        except Exception:
            pass
        try:
            app_name = str(xl.Name or "")
            app_version = str(xl.Version or "")
        except Exception:
            pass
    except Exception:
        pass
    emit({"event": "ready", "excel_pid": excel_pid,
          "app_name": app_name, "app_version": app_version})
    try:
        for line in sys.stdin:
            line = line.strip()
            if not line:
                continue
            try:
                task = json.loads(line)
            except Exception:
                continue
            f = task.get("file")
            if not f:
                continue
            try:
                def cb(pct, msg, _f=f):
                    emit({"event": "phase", "file": _f, "pct": pct, "msg": msg})

                res = audit_file(f, out_dir, on_progress=cb, session=sess,
                                 pre=task.get("pre"))
                emit({"event": "done", "file": f, "ok": True, "result": res})
            except Exception as e:  # noqa: BLE001
                emit({"event": "done", "file": f, "ok": False, "error": str(e)})
    except Exception:  # noqa: BLE001  —— 主循环整体兜底：绝不向 windowed 进程外抛异常（防弹窗）
        pass
    finally:
        try:
            sess.quit()
        except Exception:
            pass
    return 0


def precheck(path):
    """主进程拖入预检：可读性 + 布局计划（拆分/归位/底色统计），供 worker 复用。

    返回 {"ok": bool, "error": str, "plans": [...], "shifts": [...], "fill_counts": {...}}
    .xls 无法用 openpyxl 预检，返回 ok=True（由 worker COM 处理并兜底）。
    openpyxl 预读失败（如文件被 Excel/WPS 占用、特殊结构）不判定损坏：
    降级为"无预检数据"，文件照常进池，由完整审计流程处理（openpyxl 失败自动走 COM 兜底）。
    仅"文件不存在"这类硬错误才 ok=False。
    """
    out = {"ok": True, "error": "", "plans": [], "shifts": [], "fill_counts": {}}
    try:
        if not os.path.isfile(path):
            out["ok"] = False
            out["error"] = "文件不存在"
            return out
        _base, ext = os.path.splitext(os.path.basename(path))
        if ext.lower() not in (".xlsx", ".xlsm"):
            return out
        wbp = _load_openpyxl(path, use_cache=False)
        if wbp is None:
            out["ok"] = True
            out["error"] = "openpyxl 预读失败，已按常规流程处理"
            return out
        try:
            plans, shifts = _plan_splits(wbp)
        except Exception:
            plans, shifts = [], []
        out["plans"] = plans
        out["shifts"] = shifts
        if not plans:
            try:
                out["fill_counts"] = _count_fills_openpyxl(wbp) or {}
            except Exception:
                out["fill_counts"] = {}
        try:
            wbp.close()
        except Exception:
            pass
        return out
    except Exception as e:
        out["ok"] = False
        out["error"] = str(e)[:200]
        return out


class BatchPool:
    """多进程 worker 池：每个 worker 一个独立 Excel，单文件超时杀进程重试。

    - 自动并发：逻辑核 >= 8 用 3，否则 2（客户机 4 核 8 线程 → 3）。
    - 单文件超时（默认 25s）→ 终止该 worker（含其 Excel 进程）→ 重启 worker
      → 该文件重试；最多重试 max_retry 次（默认 2），仍失败标记"需人工处理"，
      其余文件照常，绝不整批中断。
      （10s 在低配 4 核机上对复杂文件太紧，且占用类文件已用 ReadOnly 打开规避挂起）
    - 每个 worker 输出到私有临时目录，完成后主进程统一 move 到最终目录并
      同名递增命名（跨进程无竞态、不覆盖、不留半成品）。
    - 进度以"每份文件最终完成"为权威事件，逐份 +1，单调不跳。

    run(files, out_dir, on_event=None) 阻塞至全部完成，返回 results（对齐
    audit_file 返回结构，处理失败项含 error）。
    on_event(e) 事件：{"type":"progress","done":d,"total":t}
                     {"type":"done","file":f,"ok":bool,"result":dict}
                     {"type":"retry","file":f,"attempt":a,"max":m,"why":str}
    """

    def __init__(self, max_workers=2, timeout=30.0, max_retry=1):
        # WPS环境实测并行2路比串行快1.6倍；保留RPC自愈和超时重建。
        self.max_workers = max_workers or 2
        self.timeout = float(timeout)
        self.max_retry = int(max_retry)
        self.workers = []
        self.events = None
        self._pending = []
        self._attempts = {}
        self._total = 0
        self._wps = False  # 检测到 WPS/非标准 Excel 后切串行

    def _worker_cmd(self, tmp_out):
        if getattr(sys, "frozen", False):
            return [sys.executable, "--worker-mode", "--out", tmp_out]
        return [sys.executable, "-m", "audit_fail", "--worker-mode", "--out", tmp_out]

    def _spawn(self):
        wid = len(self.workers)
        tmp = tempfile.mkdtemp(prefix="auditw_%d_" % os.getpid())
        _flags = subprocess.CREATE_NO_WINDOW if hasattr(subprocess, "CREATE_NO_WINDOW") else 0
        proc = subprocess.Popen(
            self._worker_cmd(tmp),
            stdin=subprocess.PIPE, stdout=subprocess.PIPE,
            stderr=subprocess.DEVNULL, encoding="utf-8", errors="replace",
            creationflags=_flags,
        )
        w = {"id": wid, "proc": proc, "tmp_out": tmp,
             "busy": None, "excel_pid": None, "dead": False, "ready": False}
        self.workers.append(w)

        def _reader(w_=w):
            try:
                for line in proc.stdout:
                    line = line.strip()
                    if not line:
                        continue
                    try:
                        obj = json.loads(line)
                    except Exception:
                        continue
                    self.events.put(("msg", w_["id"], obj))
            except Exception:
                pass
            self.events.put(("eof", w_["id"], None))

        t = threading.Thread(target=_reader, daemon=True)
        t.start()
        w["reader"] = t
        return w

    def _send(self, w, obj):
        try:
            w["proc"].stdin.write(json.dumps(obj, ensure_ascii=False) + "\n")
            w["proc"].stdin.flush()
        except Exception:
            pass

    def _kill(self, w):
        if w["excel_pid"]:
            try:
                subprocess.run(
                    ["taskkill", "/F", "/PID", str(w["excel_pid"])],
                    capture_output=True, timeout=10)
            except Exception:
                pass
        if w["proc"] and w["proc"].poll() is None:
            try:
                w["proc"].kill()
            except Exception:
                pass
        w["dead"] = True

    def _cleanup_worker(self, w):
        self._kill(w)
        try:
            w["proc"].stdin.close()
        except Exception:
            pass
        shutil.rmtree(w["tmp_out"], ignore_errors=True)

    def run(self, files, out_dir, on_event=None, pres=None):
        self.workers = []
        self.events = queue.Queue()
        self._on_event = on_event or (lambda e: None)
        self._pending = list(files)
        self._attempts = {}
        self._pres = pres or {}
        results = {}
        out_dir = os.path.abspath(out_dir) if out_dir else None
        total = len(files)
        self._total = total

        # 预启动全部 worker（Excel 冷启动不占单文件超时）
        for _ in range(self.max_workers):
            self._spawn()

        def _busy():
            return any(w["busy"] is not None for w in self.workers)

        try:
            while self._pending or _busy():
                # 1) 事件
                try:
                    while True:
                        kind, wid, obj = self.events.get_nowait()
                        w = self.workers[wid]
                        if kind == "eof":
                            if w["busy"]:
                                self._requeue(w, "worker 异常退出", results)
                            else:
                                self._cleanup_worker(w)
                                self._spawn()
                            continue
                        if obj.get("event") == "ready":
                            w["ready"] = True
                            w["excel_pid"] = obj.get("excel_pid")
                            # WPS/非标准 Excel 环境检测：真 Excel 主版本 >= 12（2007+）；
                            # WPS 表格(ET) 模拟版常报 6.0，多进程并发易崩溃。
                            if not self._wps:
                                vmj = 0
                                try:
                                    vmj = float(str(obj.get("app_version") or "0").split(".")[0])
                                except Exception:
                                    vmj = 0
                                nm = (obj.get("app_name") or "").lower()
                                if ("wps" in nm) or ("kingsoft" in nm) or ("et" in nm) or (0 < vmj < 12):
                                    self._wps = True
                                    # WPS实测并行2路比串行快1.6倍，不降并发；
                                    # 崩溃由RPC自动重试重建兜底。
                                    try:
                                        self._on_event({"type": "notice",
                                                        "msg": "WPS环境并行运行（崩溃自动重试）"})
                                    except Exception:
                                        pass
                            if w["busy"] is None and self._pending and not w["dead"]:
                                self._dispatch(w)
                            continue
                        if obj.get("event") == "phase":
                            continue
                        if obj.get("event") == "done":
                            self._on_done(w, obj, results, out_dir)
                            continue
                except queue.Empty:
                    pass

                # 2) 超时检查
                now = time.time()
                for w in list(self.workers):
                    if w["busy"] and not w["dead"] and (now - w["busy"]["start"]) > self.timeout:
                        self._requeue(w, "处理超时（%.0fs）" % self.timeout, results)

                # 3) 派发（只派给已 ready 的 worker）
                for w in list(self.workers):
                    if w["ready"] and w["busy"] is None and not w["dead"] and self._pending:
                        self._dispatch(w)

                time.sleep(0.02)
        finally:
            # 优雅收尾：先关 stdin 触发 worker EOF（正常 sess.quit 退出），
            # 最多等 3s，仍存活再强杀兜底。避免 TerminateProcess 打断 Excel 清理
            # 导致残留进程/弹窗。
            for w in list(self.workers):
                try:
                    w["proc"].stdin.close()
                except Exception:
                    pass
            _deadline = time.time() + 3.0
            while time.time() < _deadline:
                if all(w["proc"].poll() is not None for w in self.workers):
                    break
                time.sleep(0.1)
            for w in list(self.workers):
                if w["proc"].poll() is None:
                    self._kill(w)
                shutil.rmtree(w["tmp_out"], ignore_errors=True)

        return [results[f] for f in files if f in results]

    def _dispatch(self, w):
        f = self._pending.pop(0)
        w["busy"] = {"file": f, "start": time.time()}
        self._send(w, {"file": f, "pre": self._pres.get(f)})

    def _on_done(self, w, obj, results, out_dir):
        f = obj.get("file")
        if not f:
            return
        if obj.get("ok"):
            res = obj.get("result") or {}
            src = res.get("output")
            if src and os.path.isfile(src) and out_dir:
                try:
                    src = _move_output(src, out_dir)
                except Exception:
                    src = None
            results[f] = {"file": f, "ok": True,
                          "fail": bool(res.get("is_fail")),
                          "output": src,
                          "reason": res.get("reason", ""),
                          "marked": bool(res.get("marked")),
                          "norm": res.get("norm", {}),
                          "log": res.get("log", []),
                          "error": None}
            self._on_event({"type": "done", "file": f, "ok": True, "result": results[f]})
        else:
            err = obj.get("error") or "未知错误"
            # Excel COM 连接类错误（RPC 服务器不可用 / 调用被拒绝 / 连接断开）：
            # 该 worker 的 Excel 进程可能已崩溃（WPS 环境尤甚），死对象不能再用。
            # 走 _requeue：杀掉死 worker、spawn 全新 worker（新 Excel）重试，最多 max_retry 次。
            if any(k in err for k in ("RPC", "服务器不可用", "Call was rejected",
                                       "连接断开", "0x800706BA", "-2147023174")):
                self._requeue(w, "Excel 连接断开（RPC）", results)
                return
            results[f] = {"file": f, "ok": False, "fail": False, "output": None,
                          "error": err, "reason": "", "marked": False,
                          "norm": {}, "log": []}
            self._on_event({"type": "done", "file": f, "ok": False, "result": results[f]})
        self._on_event({"type": "progress", "done": len(results), "total": self._total})
        w["busy"] = None
        if not w["dead"] and self._pending:
            self._dispatch(w)

    def _requeue(self, w, why, results):
        f = w["busy"]["file"]
        attempt = self._attempts.get(f, 0) + 1
        self._attempts[f] = attempt
        self._kill(w)
        w["busy"] = None
        if attempt <= self.max_retry:
            self._pending.append(f)
            self._spawn()
            self._on_event({"type": "retry", "file": f, "attempt": attempt,
                            "max": self.max_retry, "why": why})
        else:
            results[f] = {"file": f, "ok": False, "fail": False, "output": None,
                          "error": "%s，已重试%d次，请人工检查" % (why, self.max_retry),
                          "reason": "", "marked": False, "norm": {}, "log": []}
            self._on_event({"type": "done", "file": f, "ok": False, "result": results[f]})
            self._on_event({"type": "progress", "done": len(results), "total": self._total})

if __name__ == "__main__":
    if "--worker-mode" in sys.argv:
        _out = None
        if "--out" in sys.argv:
            _i = sys.argv.index("--out")
            if _i + 1 < len(sys.argv):
                _out = sys.argv[_i + 1]
        sys.exit(run_worker_mode(_out))

    if len(sys.argv) < 2:
        print("用法: python audit_fail.py <文件或目录> [输出目录]")
        sys.exit(1)

    src = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) > 2 else None
    res = audit_file(src, out_dir=out)
    print(res)
