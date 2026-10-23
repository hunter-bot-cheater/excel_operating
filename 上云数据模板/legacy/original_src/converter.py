#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
检验报告格式转换核心模块 (converter)
====================================

把检验报告转换为新格式 .xlsx（与 93010FSS 模板一致）。
支持输入格式：.xls / .xlsx / .csv，统一输出 .xlsx。

- .xls  → openpyxl + xlrd 重建（内容变换全部生效：多件多列、图号清理、
          日期增强、底部检具代码、去 .0、上下限 F 红等；不保留源内嵌图片，
          因为输出为标准化模板重建）
- .xlsx → openpyxl 读取 + 变换 + 输出 xlsx
- .csv  → csv 模块解析 + openpyxl 构建 xlsx

对外提供统一 API：
    convert_file(input_path, out_dir=None, on_progress=None) -> dict

本模块不依赖任何 GUI 框架，便于在命令行、GUI、服务中复用。
"""

import os
import re
import sys
from datetime import datetime, timedelta
import datetime as _dt
import subprocess
import tempfile
import csv as _csv

# 支持的输入扩展名
SUPPORTED_INPUT_EXTS = (".xls", ".xlsx", ".csv")

# 合格判定关键字（不区分大小写）—— ok/合格 统一映射为 pass
_PASS_WORDS = ("OK", "合格", "PASS", "P")
# 需要转换为 pass 的原始关键词
_TO_PASS_MAP = {"ok": "pass", "合格": "pass"}


# --------------------------------------------------------------------------- #
# 依赖管理
# --------------------------------------------------------------------------- #
def ensure_deps():
    if getattr(sys, "frozen", False):
        return
    try:
        import xlrd  # noqa: F401
        import openpyxl  # noqa: F401
    except ImportError:
        print("正在安装依赖 (xlrd, openpyxl)...")
        subprocess.run(
            [sys.executable, "-m", "pip", "install", "xlrd", "openpyxl", "-q"],
            check=True,
        )


# --------------------------------------------------------------------------- #
# 图号/命名清理工具
# --------------------------------------------------------------------------- #
def _clean_dwg(raw):
    """清理图号：去掉前缀、括号注释等，保留核心图号（含可选 _prt 等下划线后缀）。

    示例:
      "JH-YCL-93010PHE-K01（按图加工）" → "93010PHE-K01"
      "JH-YCL-93021PVD_prt"             → "93021PVD_prt"
      "93023AUY"                          → "93023AUY"
    """
    if not raw:
        return ""
    s = str(raw).strip()
    # 去掉括号注释（中英文）
    s = re.sub(r"[（(][^）)]*[）)]", "", s)
    s = re.sub(r"\([^)]*\)", "", s).strip()
    # 纯数字带 .0 浮点后缀（如 xlrd 读出 92140042.0）→ 去掉 .0
    if re.fullmatch(r"\d+\.0+", s):
        s = re.sub(r"\.0+$", "", s)
    # 匹配：数字开头 + 字母 + 可选连字符/数字 + 可选下划线后缀
    m = re.search(
        r"([0-9]+[A-Z][A-Z0-9_-]*(?:_[A-Za-z0-9]+)*)",
        s, re.IGNORECASE)
    if m:
        return m.group(1)
    # 回退：去常见前缀后再次去 .0
    for prefix in ("JH-YCL-", "JH-", "YCL-"):
        if s.upper().startswith(prefix.upper()):
            rest = s[len(prefix):].strip()
            if rest:
                if re.fullmatch(r"\d+\.0+", rest):
                    rest = re.sub(r"\.0+$", "", rest)
                return rest
    return s.strip()


def _clean_filename_part(s):
    """清理文件名中的非法字符和多余内容（括号、特殊符号、首尾点等）。"""
    if not s:
        return ""
    s = re.sub(r"[（(][^）)]*[）)]", "", s)
    s = re.sub(r"\([^)]*\)", "", s)
    s = re.sub(r'[\\/:*?"<>|]', "", s)
    s = re.sub(r"\s+", " ", s).strip()
    s = s.strip(". ")  # 去掉首尾多余的点与空格
    return s


def _header_val(sv, nrows, ncols, label_kws, max_off=3):
    """在表头区查找含 label_kws 的标签单元格，返回其右侧首个非空值。"""
    for r in range(min(nrows, 10)):
        for c in range(ncols):
            v = sv(r, c)
            if v and any(k in v.lower() for k in label_kws):
                for off in range(1, max_off + 1):
                    try:
                        vv = sv(r, c + off)
                    except Exception:
                        vv = ""
                    if vv and str(vv).strip():
                        return str(vv).strip()
    return ""


def _header_inline(sv, nrows, ncols, label_kws):
    """返回含 label_kws 的整格文本（标签与值同格，如 'Inspected By:李能'）。"""
    for r in range(min(nrows, 10)):
        for c in range(ncols):
            v = sv(r, c)
            if v and any(k in v for k in label_kws):
                return str(v).strip()
    return ""


def _excel_serial_to_date_str(serial):
    """把 Excel 序列值（如 46156.0）转为 "YY.M.D"。

    Excel base: 1899-12-30（处理 1900 伪闰年）。
    serial 范围约 40000-50000（2009-2036），5 位整数。
    """
    try:
        s = float(serial)
    except (ValueError, TypeError):
        return ""
    # Excel 日期序列值范围
    if s < 10000 or s > 60000:
        return ""
    base = datetime(1899, 12, 30)
    try:
        d = base + timedelta(days=int(s))
        return f"{d.year % 100}.{d.month}.{d.day}"
    except Exception:
        return ""


def _extract_date_from_text(text):
    """从文本提取日期，适配多种形式。

    支持格式（按优先级）：
      1. 完整年份: 2026.5.22 / 2026/5/22 / 2026-5-22 → 转 "26.5.22"
      2. 短年份  : 26.5.22 / 26/5/22 → "26.5.22"
      3. 仅月日  : 5.13 / 5/13 → "5.13"（调用方应用默认年份）
    返回标准化字符串，找不到返回 ""。
    """
    if not text:
        return ""
    s = str(text).strip()
    # 完整 4 位年份（2026.5.22 / 2026/5/22 / 2026-5-22）
    m = re.search(r"(\d{4})[/.\-](\d{1,2})[/.\-](\d{1,2})", s)
    if m:
        yr, mo, dy = m.group(1), m.group(2), m.group(3)
        return f"{yr[2:]}.{mo}.{dy}"
    # 短年份 26.5.22（避免误识别 Excel 序列值如 46156.0 → "56.0"）
    m = re.search(r"(\d{2})\.(\d{1,2})\.(\d{1,2})", s)
    if m:
        return f"{m.group(1)}.{m.group(2)}.{m.group(3)}"
    # 短年份 26/5/22
    m = re.search(r"(\d{2})[/](\d{1,2})[/](\d{1,2})", s)
    if m:
        return f"{m.group(1)}.{m.group(2)}.{m.group(3)}"
    # 仅月日 5.13 或 5/13
    m = re.search(r"(\d{1,2})\.(\d{1,2})(?![.\d])", s)
    if m:
        return f"{m.group(1)}.{m.group(2)}"
    m = re.search(r"(\d{1,2})[/\-](\d{1,2})(?![.\d])", s)
    if m:
        return f"{m.group(1)}.{m.group(2)}"
    return ""


def _apply_default_year(date_str, default_year):
    """如果 date_str 是 2 段格式（仅 月.日）且 default_year 是 4 位数字，
    拼上年份变成 "YY.M.D"。

    输入:
      date_str: "6.27" → "26.6.27"（假设 default_year=2026）
      date_str: "26.6.27" → "26.6.27"（已经有年份，不变）
      date_str: "2026.6.27" → "26.6.27"（标准化为短年份）
      date_str: "" → ""
    """
    if not date_str:
        return ""
    if not default_year:
        return date_str
    try:
        year_int = int(str(default_year).strip())
    except (ValueError, TypeError):
        return date_str
    # 已有完整年份
    m4 = re.match(r"^(\d{4})\.(\d{1,2})\.(\d{1,2})$", date_str)
    if m4:
        return f"{m4.group(1)[2:]}.{m4.group(2)}.{m4.group(3)}"
    # 短年份（YY.M.D）
    m2 = re.match(r"^(\d{2})\.(\d{1,2})\.(\d{1,2})$", date_str)
    if m2:
        return date_str  # 已有年份，不动
    # 仅 M.D（无年份）→ 拼上当前默认年份
    m_md = re.match(r"^(\d{1,2})\.(\d{1,2})$", date_str)
    if m_md:
        return f"{str(year_int)[2:]}.{m_md.group(1)}.{m_md.group(2)}"
    return date_str


# --------------------------------------------------------------------------- #
# 源文件头部位置提取（模板统一，用固定位置最稳，避免关键词与标题碰撞）
# --------------------------------------------------------------------------- #
def _header_positions(sv, nrows, ncols):
    """按固定模板位置提取头部字段。

    源文件（xls/xlsx）模板统一：
      R3: A=图号标签 B=图号值  D=名称标签 F=名称值  H=客户标签 J=客户值
      R4: A=PO标签  B=PO值    D=版本标签 F=版本值  H=温度标签 J=温度值
      R5: A=检验(含签名名)     D=审核(含签名名)
    sv(r, c) 统一为 0-based。
    """
    def g(r, c):
        if r < 0 or c < 0 or r >= nrows or c >= ncols:
            return ""
        return sv(r, c) or ""

    return {
        "title": g(0, 0),
        "dwg_no_raw": g(2, 1),
        "part_name": g(2, 5),
        "customer_code": g(2, 9),
        "po": g(3, 1),
        "rev": g(3, 5),
        "temp": g(3, 9),
        "inspected_by": g(4, 0),
        "reviewed_by": g(4, 3),
        # 最终判定区域：直接从源文件平移（含复选框符号如 þ/■/□）
        "passed_text": g(5, 6) or "Passed/合格",
        "failed_text": g(6, 6) or "Failed/不合格",
        "uai_text": g(7, 6) or " UAI   让步",
        # 签名日期（源文件 R5C0/R5C3 通常包含 "...签名日期）x.xx" 格式）
        "inspected_date_raw": g(5, 0),  # 如 '(Sign/Employee No./Date签名日期）5.13'
        "reviewed_date_raw": g(5, 3),
    }


def _name_from_filename(base, dwg_clean):
    """从文件名 图号_名称_x件_日期 提取名称段（dwg 之后、件数段之前）。"""
    import re
    b = re.sub(r'-\d+$', '', base)
    segs = [s for s in b.split("_") if s]
    for i, s in enumerate(segs):
        if "件" in s:
            if i > 0:
                return segs[i - 1]
            break
    for i, s in enumerate(segs):
        if dwg_clean and dwg_clean in s:
            if i + 1 < len(segs):
                return segs[i + 1]
            break
    return ""


# --------------------------------------------------------------------------- #
# 数值格式化
# --------------------------------------------------------------------------- #
def _fmt_num(v):
    """整数去掉 .0 后缀（8.0 → 8），非数值原样返回。"""
    try:
        f = float(str(v).strip())
        if f == int(f) and abs(f) < 1e15:
            return str(int(f))
        return str(v)
    except (ValueError, TypeError):
        return str(v) if v is not None else ""


def _is_pass(v):
    return (v or "").strip().upper() in _PASS_WORDS


# --------------------------------------------------------------------------- #
# 不合格判定与文件名前缀
# --------------------------------------------------------------------------- #
def _fnum(x):
    """数值解析：去除 '°' 等单位符号后转 float；无法解析返回 None。
    与 COM 路径超差标红用的 _parse_float 规则一致。"""
    if x is None:
        return None
    s = str(x).replace("°", "").strip()
    if not s:
        return None
    try:
        return float(s)
    except (ValueError, TypeError):
        return None


def _has_failed(data_rows):
    """任一实测值超出 [下限, 上限] 闭区间（边界算合格）即判不合格。

    规则与 COM 路径超差标红完全一致：
      - 该行无上下限（技术要求/纯文本框）跳过；
      - 读数 OK/合格/PASS/P/空 跳过；
      - 读数无法转数值（带单位文本等）跳过；
      - 仅当 min 或 max 有数值且读数数值越界时判不合格。
    """
    for row_info in data_rows:
        minv = _fnum(row_info.get("min", ""))
        maxv = _fnum(row_info.get("max", ""))
        if minv is None and maxv is None:
            continue  # 无公差，不参与判定
        for reading in row_info.get("readings", []):
            if not reading:
                continue
            rs = str(reading).strip()
            if rs.upper() in ("OK", "合格", "PASS", "P", ""):
                continue
            v = _fnum(reading)
            if v is None:
                continue
            if (minv is not None and v < minv) or (maxv is not None and v > maxv):
                return True
    return False


def _with_fail_prefix(out):
    """在输出文件名最前面加 '不合格_'，已有则不加。"""
    d, f = os.path.split(out)
    if f.startswith("不合格_"):
        return out
    return os.path.join(d, "不合格_" + f)


# --------------------------------------------------------------------------- #
# 输出文件名计算
# --------------------------------------------------------------------------- #
def _read_source_meta(input_path, default_year=None):
    """从源文件解析图号/名称/日期/版本号等元信息。返回 dict: {dwg, part, date_str, rev}。

    default_year: 用户在 UI 指定的默认年份（4 位数字字符串，如 "2026"）。
    无年份的短日期（如 "6.27"）会拼上此年份。
    """
    meta = {"dwg": "", "part": "", "date_str": "", "rev": ""}
    ext = os.path.splitext(input_path)[1].lower()
    try:
        if ext in (".xls", ".xlsx"):
            if ext == ".xls":
                import xlrd
                wb = xlrd.open_workbook(input_path, formatting_info=True)
                s = wb.sheet_by_index(0)
                nrows, ncols = s.nrows, s.ncols

                def sv(r, c):
                    try:
                        v = s.cell_value(r, c)
                        return "" if v is None else str(v).strip()
                    except Exception:
                        return ""
            else:
                from openpyxl import load_workbook
                wb = load_workbook(input_path, read_only=True, data_only=True)
                ws = wb.active
                nrows, ncols = ws.max_row, ws.max_column

                def sv(r, c):  # 0-based
                    try:
                        v = ws.cell(row=r + 1, column=c + 1).value
                        return "" if v is None else str(v).strip()
                    except Exception:
                        return ""

            pos = _header_positions(sv, nrows, ncols)
            meta["dwg"] = _clean_dwg(pos["dwg_no_raw"])
            meta["part"] = pos["part_name"] or _name_from_filename(
                os.path.basename(input_path), meta["dwg"])
            meta["rev"] = pos.get("rev", "")
            # 日期识别：优先 Inspected 下方签名日期，再 Reviewed 下方，最后文件名
            # 源文件典型结构：R5=Inspected By（签名名），R6=签名日期（如 "(Sign/...)5.13"）
            md = _extract_date_from_text(pos.get("inspected_date_raw", ""))
            if not md:
                md = _extract_date_from_text(pos.get("reviewed_date_raw", ""))
            if not md:
                # 尝试从 R7-R9 各列读 Excel 序列值或文本日期
                # 源文件典型结构：R7=Inspected 签名日期值，R8=Reviewed 签名日期值
                for r0 in range(6, 9):
                    for c0 in range(min(ncols, 12)):
                        cell = sv(r0, c0)
                        if not cell:
                            continue
                        md = _excel_serial_to_date_str(cell)
                        if md:
                            break
                        md = _extract_date_from_text(cell)
                        if md:
                            break
                    if md:
                        break
            if not md:
                sig_text = " ".join([pos["inspected_by"], pos["reviewed_by"]])
                md = _extract_date_from_text(sig_text)
            if not md:
                md = _extract_date_from_text(os.path.basename(input_path))
            # 应用默认年份（无年份短日期 → 拼上当前系统年份）
            if md:
                md = _apply_default_year(md, default_year)
            meta["date_str"] = md
            if ext == ".xlsx":
                wb.close()
        elif ext == ".csv":
            md = _extract_date_from_text(os.path.basename(input_path))
            if md:
                md = _apply_default_year(md, default_year)
            meta["date_str"] = md
    except Exception:
        pass
    return meta


def _format_rev(rev):
    """版本号格式化：A → A01, A.2 → A02, A-1 → A01, E → E01

    支持源文件中 rev 字段写法不一致的情况：
    - 纯字母 "A" → "A01"
    - 字母+分隔符+数字 "A.2" / "A-1" → "A02" / "A01"
    - 多余空格或斜杠等 "A. 1" → "A01"
    - 多个字母 "AB" → "AB01"
    """
    if not rev:
        return ""
    s = str(rev).strip()
    if not s:
        return ""
    # 提取：字母部分 + 数字部分
    m = re.search(r'([A-Za-z]+)[.\s\-_/]*(\d+)', s)
    if m:
        letter = m.group(1).upper()
        num = int(m.group(2))
        return f"{letter}{num:02d}"
    # 只有字母（如 "A", "B", "AB"）
    m2 = re.search(r'([A-Za-z]+)', s)
    if m2:
        return f"{m2.group(1).upper()}01"
    # 纯数字
    m3 = re.search(r'(\d+)', s)
    if m3:
        return f"V{int(m3.group(1)):02d}"
    return s


def compute_output_path(input_path, out_dir=None, total_qty=1, default_year=None):
    """计算最终输出 .xlsx 的落盘路径（含防覆盖递增序号）。

    文件名格式：【上云数据】【尺寸】图号_版本号-日期.xlsx
    例： 【上云数据】【尺寸】93021PVD_prt_A01-26.6.27.xlsx

    default_year: 4 位年份字符串（如 "2026"），用于无年份的短日期补全年份。
    """
    meta = _read_source_meta(input_path, default_year=default_year)
    raw_base = os.path.splitext(os.path.basename(input_path))[0]
    base_name = re.sub(r'-\d+$', '', raw_base)
    dwg = meta["dwg"] or base_name
    # 版本号：A → A01, A.2 → A02
    rev = _format_rev(meta.get("rev", ""))
    # 日期优先取文件名（更精确），否则用签名区解析出的日期
    # [关闭日期后缀 2026-09-03] 命名只到版本号 A01，不再接日期。保留原逻辑以便日后恢复。
    # dn = _extract_date_from_text(base_name)
    # if not dn:
    #     dn = meta.get("date_str", "")
    # if dn and default_year:
    #     dn = _apply_default_year(dn, default_year)
    # dn = ""  # 关闭日期后缀后视作空
    dn = ""  # [关闭日期后缀] 命名只到版本号 A01，不再接日期
    if rev:
        # cn = f"【上云数据】【尺寸】{dwg}_{rev}-{dn}"   # 原格式（含日期后缀）
        cn = f"【上云数据】【尺寸】{dwg}_{rev}"
    else:
        # cn = f"【上云数据】【尺寸】{dwg}-{dn}"          # 原格式（含日期后缀）
        cn = f"【上云数据】【尺寸】{dwg}"
    cn = _clean_filename_part(cn) or cn

    od = os.path.abspath(os.path.normpath(out_dir)) if out_dir else os.path.dirname(
        os.path.abspath(input_path)
    )
    os.makedirs(od, exist_ok=True)

    op = os.path.join(od, cn + ".xlsx")
    ct = 1
    while os.path.exists(op):
        op = os.path.join(od, f"{cn} ({ct}).xlsx")
        ct += 1
    return op


# --------------------------------------------------------------------------- #
# 表格布局探测（兼容不同源文件列偏移）
# --------------------------------------------------------------------------- #
def _detect_layout(sv, nrows, ncols):
    """探测表头/子表头/数据起始行与各关键列。

    返回 dict，或 None（未识别）。sv(row, col) 使用调用方各自索引习惯。
    """
    # 1) 表头行：含 "item no"
    header_row = None
    for r in range(min(nrows, 25)):
        for c in range(ncols):
            v = sv(r, c)
            if v and "item no" in v.lower():
                header_row = r
                break
        if header_row is not None:
            break
    if header_row is None:
        return None

    def find_col(row, *keywords):
        for c in range(ncols):
            v = sv(row, c)
            if v:
                vl = v.lower()
                if any(k in vl for k in keywords):
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
    reading_candidates = []
    if instr_col is not None and result_col is not None:
        for c in range(instr_col + 1, result_col):
            lab = sv(sub_row, c) or ""
            if lab and "最大" not in lab and "最小" not in lab:
                reading_candidates.append(c)
    elif instr_col is not None:
        for c in range(instr_col + 1, instr_col + 1 + max(total_qty, 1)):
            reading_candidates.append(c)

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
        "reading_candidates": reading_candidates,
        "data_start": data_start,
    }


def _parse_source(sv, nrows, ncols):
    """探测布局并解析数据行。返回 (data_rows, total_qty, layout)。

    data_rows 每项含 readings: 每件测量值列表。
    """
    layout = _detect_layout(sv, nrows, ncols)
    if layout is None:
        raise ValueError("未找到 Item No 表头，无法识别文件格式")
    item_col = layout["item_col"]
    spec_col = layout["spec_col"]
    max_col = layout["max_col"]
    min_col = layout["min_col"]
    instr_col = layout["instr_col"]
    cand = layout["reading_candidates"]
    data_start = layout["data_start"]

    # 有效实测列：在任意数据行有非空值
    eff_reading = []
    for c in cand:
        has = False
        for r in range(data_start, nrows):
            if (sv(r, c) or "").strip():
                has = True
                break
        if has:
            eff_reading.append(c)
    total_qty = max(1, len(eff_reading))
    reading_cols = eff_reading if eff_reading else (cand[:1] or [instr_col + 1])

    def gv(r, c):
        try:
            return (sv(r, c) or "").strip()
        except Exception:
            return ""

    # 实测列对应的“件号/序列号”表头：取数据行上一行（sub_row）中实测列标签，
    # 源文件若是序列号（如 2640012/A01）则原样平移，保持源文件格式。
    piece_headers = [gv(data_start - 1, c) for c in reading_cols]

    data_rows = []
    last_item = ""
    last_spec = ""
    _FOOTER_KW = ("检具", "代码", "备注", "结论", "检验", "审核", "日期",
                  "report", "inspect", "review", "fixture", "signature")
    for r in range(data_start, nrows):
        raw_item = gv(r, item_col) if item_col is not None else ""
        raw_spec = gv(r, spec_col) if spec_col is not None else ""
        readings = [gv(r, c) for c in reading_cols]
        # 空行（原始 item/spec 与读数全空）：跳过且不沿用上一行，
        # 否则会把页脚/分隔空行错误继承成前一个 item 的重复数据行。
        if not raw_item and not raw_spec and not any(readings):
            continue
        # 页脚/非数据行（item 像长文本或含关键词）跳过。
        if raw_item and (len(raw_item) > 12
                         or any(k in raw_item.lower() for k in _FOOTER_KW)):
            continue
        item_no = raw_item or last_item
        # 字母开头（A, B, C, D, a1, b2 等）：前面加 "0-" 变为 "0-A", "0-a1"
        if item_no and re.match(r'^[A-Za-z]', item_no):
            if not item_no.startswith('0-'):
                item_no = '0-' + item_no
        spec = raw_spec or last_spec
        if item_no:
            last_item = item_no
        if spec:
            last_spec = spec
        max_v = gv(r, max_col) if max_col is not None else ""
        min_v = gv(r, min_col) if min_col is not None else ""
        # 智能识别：源文件偶尔把 max/min 填反（max<min），自动交换保证区间合法
        try:
            max_n = float(str(max_v).strip()) if max_v else None
            min_n = float(str(min_v).strip()) if min_v else None
            if max_n is not None and min_n is not None and max_n < min_n:
                max_v, min_v = min_v, max_v
        except (ValueError, TypeError):
            pass
        data_rows.append({
            "src_row": r,
            "item_no": item_no,
            "spec": spec,
            "max": max_v,
            "min": min_v,
            "instr": gv(r, instr_col) if instr_col is not None else "目测",
            "readings": readings,
        })
    if not data_rows:
        raise ValueError("未找到数据行，请确认文件格式")

    # 分类：技术要求 / 尺寸（首个非技术要求之后均为尺寸）
    def classify(row):
        nonempty = [x for x in row["readings"] if x]
        if nonempty and all(_is_pass(x) for x in nonempty):
            return "技术要求"
        if nonempty and all(_is_pass(x) for x in nonempty) and \
                not any(ch.isdigit() for ch in row["spec"]):
            return "技术要求"
        return "尺寸"

    boundary = False
    for row in data_rows:
        if not boundary:
            if classify(row) == "尺寸":
                boundary = True
                row["insp"] = "尺寸"
            else:
                row["insp"] = "技术要求"
        else:
            row["insp"] = "尺寸"
    return data_rows, total_qty, layout, piece_headers


# --------------------------------------------------------------------------- #
# 源文件头部信息提取（用于原样保留到输出）
# --------------------------------------------------------------------------- #
def _extract_header_info(input_path):
    """从源文件提取头部元数据（标题/图号/名称/PO/签名等），用于原样写入输出。"""
    info = {
        "title": "", "dwg_no_raw": "", "dwg_no": "",
        "part_name": "", "po": "", "rev": "", "temp": "",
        "inspected_by": "", "reviewed_by": "",
        "inspected_date": "", "reviewed_date": "",
        "customer_code": "", "remark": "",
        "form_no": "", "total_qty": 1, "sample_qty": 1,
    }
    ext = os.path.splitext(input_path)[1].lower()
    try:
        if ext == ".xls":
            import xlrd
            wb = xlrd.open_workbook(input_path, formatting_info=True)
            s = wb.sheet_by_index(0)
            nrows, ncols = s.nrows, s.ncols

            def sv(r, c):
                try:
                    v = s.cell_value(r, c)
                    return "" if v is None else str(v).strip()
                except Exception:
                    return ""
        elif ext == ".xlsx":
            from openpyxl import load_workbook
            wb = load_workbook(input_path, read_only=True, data_only=True)
            ws = wb.active
            nrows, ncols = ws.max_row, ws.max_column

            def sv(r, c):  # 统一 0-based
                try:
                    v = ws.cell(row=r + 1, column=c + 1).value
                    return "" if v is None else str(v).strip()
                except Exception:
                    return ""
        else:
            return info

        pos = _header_positions(sv, nrows, ncols)
        info["title"] = pos["title"]
        info["dwg_no_raw"] = pos["dwg_no_raw"]
        info["dwg_no"] = _clean_dwg(pos["dwg_no_raw"])
        info["part_name"] = pos["part_name"] or _name_from_filename(
            os.path.basename(input_path), info["dwg_no"])
        info["po"] = pos["po"]
        info["rev"] = pos["rev"]
        info["temp"] = pos["temp"]
        info["inspected_by"] = pos["inspected_by"]
        info["reviewed_by"] = pos["reviewed_by"]
        info["customer_code"] = pos["customer_code"]
        # 最终判定区域文本（含源文件复选框符号 ■□þ 等，直接平移）
        info["passed_text"] = pos.get("passed_text", "Passed/合格")
        info["failed_text"] = pos.get("failed_text", "Failed/不合格")
        info["uai_text"] = pos.get("uai_text", " UAI   让步")
        # 签名日期原始文本（用于提取日期值写入签名区）
        info["inspected_date_raw"] = pos.get("inspected_date_raw", "")
        info["reviewed_date_raw"] = pos.get("reviewed_date_raw", "")

        # Total / Sample Qty —— 在表头区查找 "Total Qty" 标签，取其右侧数值
        tq = 1
        for r in range(min(nrows, 12)):
            found = False
            for c in range(min(ncols, 12)):
                v = sv(r, c)
                if v and "total qty" in v.lower():
                    for off in (1, 2, 3):
                        tv = sv(r, c + off) if c + off < ncols else ""
                        if tv:
                            try:
                                tq = int(float(tv))
                                found = True
                                break
                            except (ValueError, TypeError):
                                pass
                if found:
                    break
            if found:
                break
        info["total_qty"] = tq
        info["sample_qty"] = tq

        if ext == ".xlsx":
            wb.close()
    except Exception:
        pass
    return info


# --------------------------------------------------------------------------- #
# 数值范围检查（模块级，COM 和 openpyxl 共用）
# --------------------------------------------------------------------------- #
def _in_range(val_str, min_str, max_str):
    """检查数值是否在 [min, max] 范围内。"""
    try:
        v = float(str(val_str).replace("°", "").strip())
        lo = float(str(min_str).strip()) if min_str and str(min_str).strip() not in ("/", "") else None
        hi = float(str(max_str).strip()) if max_str and str(max_str).strip() not in ("/", "") else None
        if lo is not None and hi is not None:
            return lo <= v <= hi
        if lo is not None:
            return v >= lo
        if hi is not None:
            return v <= hi
        return True
    except (ValueError, TypeError):
        return True


# --------------------------------------------------------------------------- #
# 通用输出构建器（所有格式共用）
# --------------------------------------------------------------------------- #
def _build_output_xlsx(data_rows, item_no_merge_groups, input_path,
                       out_dir, on_progress, sv=None,
                       header_info=None, total_qty=1, piece_headers=None,
                       default_year=None):
    """根据解析好的 data_rows 构建标准 .xlsx 输出文件。

    布局严格匹配参考输出（93023AVE_视窗压盖_1件_26.5.04.xlsx）：
      - 12 列 A-L，固定不变
      - R1-2: 标题合并 A1:L2（Times New Roman 22pt 加粗）
      - R3:   图号/零件名/客户代码(固定0998)
      - R4:   PO/版本/温度
      - R5:   检验/审核/最终判定/备注
      - R6-8: 签名日期区 + 合格/不合格/让步
      - R9-11: 表头（Actual Readings 固定 G-K 五列）
      - R12+: 数据行
      - 末行: 检具代码固定行（紧贴数据后）
    """
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
    from openpyxl.utils import get_column_letter

    if on_progress:
        on_progress(50, "构建输出文件...")

    if header_info is None:
        header_info = _extract_header_info(input_path)
    total_qty = max(total_qty, header_info.get("total_qty", 1))
    total_qty = max(total_qty, 1)

    wb = Workbook()
    ws = wb.active
    ws.title = "检验报告"

    # ========== 样式（匹配参考输出） ==========
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    f_main = Font(name="Times New Roman", size=10)
    f_title = Font(name="Times New Roman", size=22, bold=True)
    f_title_cn = Font(name="宋体", size=14, bold=True)  # 检验报告 14号字（与源文件一致）
    f_cn = Font(name="宋体", size=10)
    f_cn_item = Font(name="宋体", size=12)
    f_result = Font(name="宋体", size=11)
    f_spec_small = Font(name="宋体", size=6)
    f_sign = Font(name="Times New Roman", size=9)
    f_red = Font(name="宋体", size=11, color="FFFFFF", bold=True)

    ac = Alignment(horizontal="center", vertical="center", wrap_text=True)
    al = Alignment(horizontal="left", vertical="center", wrap_text=True)
    at = Alignment(horizontal="left", vertical="top", wrap_text=True)
    act = Alignment(horizontal="center", vertical="top", wrap_text=True)
    acl = Alignment(horizontal="center", vertical="center", wrap_text=False)

    red_fill = PatternFill(start_color="FF0000", end_color="FF0000",
                           fill_type="solid")

    def set_cell(r, c, val, font=f_main, align=ac, fill=None):
        cell = ws.cell(row=r, column=c, value=val)
        cell.font = font
        cell.alignment = align
        if fill:
            cell.fill = fill
        return cell

    def merge_set(r1, c1, r2, c2, val, font=f_main, align=ac):
        ws.merge_cells(f"{get_column_letter(c1)}{r1}:{get_column_letter(c2)}{r2}")
        return set_cell(r1, c1, val, font, align)

    # ========== 列宽（匹配参考输出） ==========
    col_widths = {1: 7.94, 2: 12.31, 3: 13.45, 4: 8.78, 5: 8.92,
                  6: 9.28, 7: 7.85, 8: 8.35, 9: 7.40, 10: 7.20,
                  11: 13.00, 12: 13.56}
    for ci, w in col_widths.items():
        ws.column_dimensions[get_column_letter(ci)].width = w

    # ========== 行高（表头/表单区，匹配参考输出） ==========
    row_heights = {1: 22.2, 2: 24.0, 3: 31.0, 4: 31.0, 5: 18.0,
                   9: 28.0, 10: 26.0, 11: 31.0}
    for r, h in row_heights.items():
        ws.row_dimensions[r].height = h

    # ===== 第 1-2 行：标题（拆分字体：英文22号 / 中文14号，与源文件一致） =====
    _title_en = "Inspection Report"
    _title_cn = "检验报告"
    # 如果源文件标题单元格含换行，尝试分别提取英文/中文行
    _raw_title = header_info.get("title") or ""
    if '\n' in str(_raw_title):
        _lines = str(_raw_title).split('\n')
        for line in _lines:
            stripped = line.strip()
            if not stripped:
                continue
            if re.search(r'[a-zA-Z]', stripped) and not re.search(r'[\u4e00-\u9fff]', stripped):
                _title_en = stripped
            elif re.search(r'[\u4e00-\u9fff]', stripped):
                _title_cn = stripped
    # 合并 R1:R2（A1:L2）消除中间横线，与参考文件一致；
    # 值 = "Inspection Report\n检验报告"（22pt TNR bold）
    ws.merge_cells("A1:L2")
    _title_cell = ws.cell(1, 1)
    _title_cell.value = f"{_title_en}\n{_title_cn}"
    _title_cell.font = f_title
    _title_cell.alignment = ac
    # 清除标题区边框，避免内部出现横线
    _nb = Border()
    for _rr in (1, 2):
        for _cc in range(1, 13):
            ws.cell(_rr, _cc).border = _nb

    # ===== 第 3 行：图号 / 零件名称 / 客户代码(固定0998) =====
    dwg_display = header_info.get("dwg_no") or header_info.get("dwg_no_raw") or ""
    part_display = header_info.get("part_name") or ""
    set_cell(3, 1, "DWG No:\n图号", f_main, ac)                # A3
    merge_set(3, 2, 3, 4, dwg_display, f_main, ac)            # B3:D3
    merge_set(3, 5, 3, 6, "Part Description:\n零件名称:", f_main, ac)  # E3:F3
    merge_set(3, 7, 3, 8, part_display, f_cn, ac)             # G3:H3
    merge_set(3, 9, 3, 10, "Customer Code：\n客户代码", f_main, ac)  # I3:J3
    merge_set(3, 11, 3, 12, "0998", f_main, ac)               # K3:L3 ← 固定 0998

    # ===== 第 4 行：PO / 版本 / 温度 =====
    po_val = header_info.get("po") or ""
    rev_val = header_info.get("rev") or ""
    temp_val = header_info.get("temp") or ""
    set_cell(4, 1, "PO\n工单号:", f_main, ac)                  # A4
    merge_set(4, 2, 4, 4, po_val, f_main, ac)                # B4:D4
    merge_set(4, 5, 4, 6, "Rev.\n版本:", f_main, ac)          # E4:F4
    merge_set(4, 7, 4, 8, rev_val, f_cn, ac)                 # G4:H4
    merge_set(4, 9, 4, 10, "Tempreture\n温度:", f_main, ac)   # I4:J4
    merge_set(4, 11, 4, 12, temp_val, f_main, ac)            # K4:L4

    # ===== 第 5 行：检验 / 审核 / 最终判定 / 备注 =====
    insp_by = header_info.get("inspected_by") or ""
    rev_by = header_info.get("reviewed_by") or ""
    merge_set(5, 1, 5, 3, insp_by, f_main, al)                # A5:C5
    merge_set(5, 4, 5, 6, rev_by, f_main, al)                 # D5:F5
    merge_set(5, 7, 5, 9, "Final Disposition（最终判定）", f_main, ac)  # G5:I5
    merge_set(5, 10, 5, 12, "Remark（备注）:", f_main, ac)     # J5:L5

    # ===== 第 6-8 行：签名区域（标签框，含从源文件提取的日期值） =====
    # 从源文件签名区原始文本提取日期（如 '(Sign/...签名日期）5.13' → '5.13'）
    _insp_date_raw = header_info.get("inspected_date_raw", "") or ""
    _rev_date_raw = header_info.get("reviewed_date_raw", "") or ""
    # 提取日期：寻找文本末尾的数字日期（如 5.13 / 26.5.13 / 2024.5.13 / 2026.4.20）
    _date_match = re.search(r'(\d{1,2}[\.\-／/]\d{1,2}(?:[\.\-／/]\d{2,4})?)\s*$', _insp_date_raw) or \
                  re.search(r'(\d{1,2}[\.\-／/]\d{1,2}(?:[\.\-／/]\d{2,4})?)\s*$', _rev_date_raw)
    _sign_date = _date_match.group(1) if _date_match else ""
    # 如果源文件没找到日期，尝试从文件名提取
    if not _sign_date:
        _fn_date = _extract_date_from_text(os.path.basename(input_path))
        if _fn_date:
            _sign_date = _fn_date
    # 构建两行式签名标签（与源文件 R5C3 格式一致：标签一行，日期下一行）
    _sign_label = "(Sign/Employee No./Date签名日期)"
    if _sign_date:
        _sign_label_with_date = f"{_sign_label}\n{_sign_date}"
    else:
        _sign_label_with_date = _sign_label

    merge_set(6, 1, 8, 3, _sign_label_with_date, f_sign, ac)  # A6:C8 检验日期
    merge_set(6, 4, 8, 6, _sign_label_with_date, f_sign, ac)  # D6:F8 审核日期
    # 最终判定：标签文本与参考一致（纯文本，单元格内不含 ■□þ 字符；
    # 复选框 ■□ 由 _copy_shapes_via_com 从源文件复制的 Shape 提供）
    def _strip_cb(s):
        return str(s or "").lstrip('þ■☑✓□☐ ').strip()
    _passed = _strip_cb(header_info.get("passed_text")) or "Passed/合格"
    _failed = _strip_cb(header_info.get("failed_text")) or "Failed/不合格"
    _uai = _strip_cb(header_info.get("uai_text")) or "UAI   让步"

    merge_set(6, 7, 6, 9, _passed, f_main, ac)          # G6:I6 合格
    merge_set(7, 7, 7, 9, _failed, f_main, ac)            # G7:I7 不合格
    merge_set(8, 7, 8, 9, _uai, f_main, ac)               # G8:I8 让步
    merge_set(6, 10, 8, 12, "此处填写让步申请编号", f_main, ac)  # J6:L8

    # ===== 表头区（第 9-11 行） =====
    hdr_row = 9
    sub_hdr_row = 10
    dim_hdr_row = 11

    merge_set(hdr_row, 1, dim_hdr_row, 1, "Item No\n（编号）", f_main, ac)   # A9:A11
    merge_set(hdr_row, 2, dim_hdr_row, 2, "检验项目", f_cn, ac)              # B9:B11
    merge_set(hdr_row, 3, hdr_row + 1, 5,
              "Specified Parameter or Dimension & Tolerance on Drawing\n（要求与尺寸）",
              f_main, ac)                                                     # C9:E10
    merge_set(hdr_row, 6, dim_hdr_row, 6,
              "Instrument used\n（检具）", f_main, ac)                        # F9:F11

    # Total Qty / Sample Qty（按用户要求：总数G列/H列值，检查数I列/J列值，K空）
    ar_col_start = 7   # G 列起
    sample_qty = header_info.get("sample_qty", total_qty)
    set_cell(hdr_row, 7, "Total Qty\n（总数）:", f_main, ac)          # G9 标签（仅 G 列）
    set_cell(hdr_row, 8, str(total_qty), f_main, acl)                 # H9 总数值
    set_cell(hdr_row, 9, "Sample Qty\n（检查数）:", f_main, ac)       # I9 标签（仅 I 列）
    set_cell(hdr_row, 10, str(sample_qty), f_main, acl)               # J9 检查数值
    set_cell(hdr_row, 11, "", f_main, ac)                             # K9 空

    # Result 列表头：合并 L9:L10，L11 保留子注
    merge_set(hdr_row, 12, sub_hdr_row, 12, "Result（结果）", f_main, ac)  # L9:L10
    set_cell(dim_hdr_row, 12, "(P合格 /F不合格)", f_main, ac)        # L11

    # Actual Readings 表头（G10:K10 固定 5 列，与参考一致）
    ar_col_end = 11
    merge_set(sub_hdr_row, ar_col_start, sub_hdr_row, ar_col_end,
              "Actual Readings（测量数据）", f_main, ac)
    for cc in range(ar_col_start, ar_col_end + 1):
        ws.cell(row=sub_hdr_row, column=cc).border = border

    # 子表头
    set_cell(dim_hdr_row, 3, "公称尺寸\nDimension", f_main, ac)      # C11
    set_cell(dim_hdr_row, 4, "上限值\nMax", f_main, ac)             # D11
    set_cell(dim_hdr_row, 5, "下限值\nMin", f_main, ac)             # E11
    for qi in range(total_qty):
        col = ar_col_start + qi
        if col <= 11:
            # 件号表头：优先用源文件实测列标签（可能为序列号），否则用序号 1,2,...
            raw = piece_headers[qi] if (piece_headers and qi < len(piece_headers)
                                        and piece_headers[qi]) else str(qi + 1)
            label = _fmt_num(raw)  # 数值去 .0（如 xls 的 1.0 → 1），文本序列号原样
            set_cell(dim_hdr_row, col, label, f_main, acl)    # 件号/序列号

    # ===== 数据行（从 R12 开始） =====
    first_data_row = dim_hdr_row + 1  # 12
    src_to_out_row = {}

    for i, row in enumerate(data_rows):
        out_r = first_data_row + i
        src_to_out_row[row["src_row"]] = out_r

        set_cell(out_r, 1, row["item_no"], f_cn_item, ac)           # A 宋体12
        set_cell(out_r, 2, row["insp"], f_cn, ac)                   # B 宋体10
        spec_text = row["spec"]
        if len(spec_text) > 20:
            set_cell(out_r, 3, spec_text, f_spec_small, ac)
        else:
            set_cell(out_r, 3, spec_text, f_cn, ac)
        set_cell(out_r, 4, _fmt_num(row["max"]), f_main, ac)        # D Max
        set_cell(out_r, 5, _fmt_num(row["min"]), f_main, ac)        # E Min
        set_cell(out_r, 6, row["instr"], f_cn, ac)                 # F 检具

        for qi, reading in enumerate(row["readings"]):
            col = ar_col_start + qi
            if col > 11:
                break
            # 实测值处理：
            #   1) ok/合格 统一转换为 pass（用户要求）
            #   2) 其他合格类文字原样平移（PASS/P等保持源文件格式）
            #   3) 数值则去 .0 后缀
            _reading_lower = str(reading).strip().lower()
            if _reading_lower in _TO_PASS_MAP:
                disp = _TO_PASS_MAP[_reading_lower]
            elif _is_pass(reading):
                disp = reading if reading else ""
            else:
                disp = _fmt_num(reading)
            set_cell(out_r, col, disp, f_main, ac)

        minv = row["min"]
        maxv = row["max"]
        out_of_range = False
        for reading in row["readings"]:
            if _is_pass(reading):
                continue
            try:
                float(str(reading).replace("°", "").strip())
            except (ValueError, TypeError):
                continue
            if not _in_range(reading, minv, maxv):
                out_of_range = True
                break

        nonempty = [x for x in row["readings"] if x]
        if out_of_range:
            set_cell(out_r, 12, "F", f_red, ac, fill=red_fill)
        elif nonempty and all(_is_pass(x) or str(x).strip().lower() in _TO_PASS_MAP for x in nonempty):
            set_cell(out_r, 12, "P", f_result, ac)       # 结果列：合格显示 P（不改为pass）
        elif not nonempty:
            set_cell(out_r, 12, "/", f_main, ac)
        else:
            set_cell(out_r, 12, "P", f_result, ac)        # 结果列：默认 P

        # 行高（内容自适应，近似参考）
        if len(spec_text) > 30:
            ws.row_dimensions[out_r].height = 46
        elif len(spec_text) > 18:
            ws.row_dimensions[out_r].height = 34
        else:
            ws.row_dimensions[out_r].height = 23

    # Item No 合并（相同编号跨多行）
    for anchor_src_row, group_src_rows in item_no_merge_groups.items():
        out_rows = [src_to_out_row[sr] for sr in group_src_rows
                    if sr in src_to_out_row]
        if len(out_rows) >= 2:
            ws.merge_cells(f"A{out_rows[0]}:A{out_rows[-1]}")

    # ===== 底部固定区域：检具代码信息行（跨2行合并，与参考 A50:L51 一致） =====
    footer_row = first_data_row + len(data_rows)
    _FIXTURE_TEXT = (
        "检具代码：游标卡尺YK、深度尺SC、内径千分尺NJ、外径千分尺WJ、"
        "投影仪TY、温度仪WD、千分表QF、百分表BF、角度尺JD、硬度仪YD、"
        "粗糙度仪CCD、卷尺JC、塞规SG、环规HG、膜厚仪MHY、"
        "三坐标CMM、治具ZJ"
    )
    # 跨 2 行合并 A:L + 自动换行，避免文字被截断（用户要求占两行）
    _acl_gauge = Alignment(horizontal="left", vertical="center", wrap_text=True)
    set_cell(footer_row, 1, _FIXTURE_TEXT, f_main, _acl_gauge)
    ws.merge_cells(f"A{footer_row}:L{footer_row + 1}")
    ws.row_dimensions[footer_row].height = 22
    ws.row_dimensions[footer_row + 1].height = 22

    # ===== 全表细边框（确保无缺失/淡线，满足边框要求） =====
    last_row = footer_row + 1
    for r in range(1, last_row + 1):
        for c in range(1, 13):
            ws.cell(row=r, column=c).border = border

    out = compute_output_path(input_path, out_dir, total_qty=total_qty, default_year=default_year)
    if _has_failed(data_rows):
        out = _with_fail_prefix(out)
    wb.save(out)

    tech = sum(1 for x in data_rows if x["insp"] == "技术要求")
    dim = sum(1 for x in data_rows if x["insp"] == "尺寸")
    if on_progress:
        on_progress(95, "保存完成")
    return out, tech, dim


# --------------------------------------------------------------------------- #
# 单文件转换入口（GUI / CLI 调用）
# --------------------------------------------------------------------------- #
def _inspect_label(row):
    """判断检验项目标签：OK/合格→技术要求，数字→尺寸"""
    readings = row.get("readings", [])
    if not readings:
        return "技术要求"
    r0 = str(readings[0]).strip()
    if not r0 or r0.upper() in ("OK", "合格", "PASS", "P", "F"):
        return "技术要求"
    try:
        float(r0.replace("°", ""))
        return "尺寸"
    except (ValueError, TypeError):
        return "技术要求"


def _extend_merges_and_add_inspection_column(ws2, data_rows, src_merges=None):
    """保留源结构 + Insert B 列 + 写检验项目 + 关键合并扩展。"""
    import time as _pt
    _PROBE = [("start", _pt.perf_counter())]
    # 0. 过滤 data_rows 中的空行（_convert_via_com_inplace 在调本函数前未过滤）
    #    空行特征：item 纯数字 + 无公差(max/min)+ 无 readings（典型如 R36-R43 纯编号行）
    def _row_is_empty(row):
        readings = row.get("readings", [])
        if any(str(r).strip() and r is not None for r in readings):
            return False
        mx = str(row.get("max", "") or "").strip()
        mn = str(row.get("min", "") or "").strip()
        if (mx and mx not in ("/",)) or (mn and mn not in ("/",)):
            return False
        item_no = str(row.get("item_no", "") or "").strip()
        item_digit = item_no.replace('.', '').replace('-', '')
        if item_digit.isdigit():
            return True
        spec = str(row.get("spec", "") or "").strip()
        if spec and spec not in ("/",):
            return False
        # 完全没 item/spec/max/min/readings 才算空
        return not item_no and not spec

    data_rows = [r for r in data_rows if not _row_is_empty(r)]
    # 1. Insert B 列（让原 B 内容自动右移到 C，依等）

    try:
        ws2.Columns("B:B").Insert(Shift=-4161)  # xlShiftToRight
    except Exception:
        pass


    # 2. v34 头部重排：先 unMerge + 收集 cell 内容到 dict
    def _col_to_num(s):
        n = 0
        for ch in s:
            n = n * 26 + (ord(ch) - 64)
        return n

    def _num_to_col(n):
        s = ""
        while n:
            n, rem = divmod(n - 1, 26)
            s = chr(65 + rem) + s
        return s

    cell_vals = {}
    try:
        # 批量读值：1 次 COM 调用读取 R3:R12 × A:S 全部值（替代 ~200 次逐格调用）
        # 范围覆盖 R9-R11 汇总区：其合并区（如源 G9:K9→插入后 H9:L9）也在此统一解合并，
        # 避免 R9-11 段/检具段对已解区域做慢速逐格检查（~100ms/次）
        _rng = None
        try:
            _rng = ws2.Range("A3:S12")
            _vals = _rng.Value
            if _vals:
                if not isinstance(_vals, tuple):
                    _vals = ((_vals,),)
                elif not isinstance(_vals[0], tuple):
                    _vals = (_vals,)
                for _ri, _rv in enumerate(_vals):
                    for _ci, _v in enumerate(_rv):
                        cell_vals[(3 + _ri, 1 + _ci)] = _v
        except Exception:
            # 批量读失败兜底：逐格
            for r in range(3, 13):
                for c in range(1, 20):
                    try:
                        cell_vals[(r, c)] = ws2.Cells(r, c).Value
                    except Exception:
                        pass
        # 合并区处理：使用调用方在 Excel 打开源文件前预读的合并区清单（src_merges），
        # 映射到插入 B 列后的坐标，仅解除 R3:R11 × A:S 范围内的合并区
        # （数据行 R12+ 的源合并保留不动，避免破坏 item_no 等跨行合并）。
        # 注意：不能用 Range.MergeCells 数组探测——win32com 对部分合并区域
        # 返回 None/False（实测大范围与单行均如此），会静默跳过全部解合并；
        # 也不能在 Excel 打开后 openpyxl 重读（文件被 COM 独占）。
        try:
            for _a in (src_merges or []):
                _mm = re.match(r"([A-Z]+)(\d+):([A-Z]+)(\d+)$", _a)
                if not _mm:
                    continue
                # B 列前插入 1 列：原列号 >= 2 的右移 1 列
                _c1 = _col_to_num(_mm.group(1)) + 1
                _r1 = int(_mm.group(2))
                _c2 = _col_to_num(_mm.group(3)) + 1
                _r2 = int(_mm.group(4))
                if _r1 >= 3 and _r1 <= 11:
                    _addr = f"{_num_to_col(_c1)}{_r1}:{_num_to_col(_c2)}{_r2}"
                    try:
                        ws2.Range(_addr).UnMerge()
                    except Exception:
                        pass
        except Exception:
            pass
    except Exception:
        pass

    def _cv(r, c):
        v = cell_vals.get((r, c), None)
        return "" if v is None else str(v)

    def _cv_merge_block(anchor_c, ncols=3, rows=(6, 7, 8)):
        """扫描 (rows × [anchor_c..anchor_c+ncols-1]) 范围内所有非空子单元格值，
        按顺序去重拼接为 \\n 分隔的多行内容（用于签名/审核/让步合并区标签+日期平移）。
        原理：步骤 2 已把每个合并区域所有子格的 .Value 收集到 cell_vals。
        """
        parts = []
        seen = set()
        for rr in rows:
            for cc in range(anchor_c, anchor_c + ncols):
                v = _cv(rr, cc)
                if v and v.strip() and v not in seen:
                    parts.append(v)
                    seen.add(v)
        return "\n".join(parts)

    # 3. R3 重新合并
    # 3a. 图号值：B3:C3（2列）
    dwg_val = ""
    for cc in (2, 3, 4):
        v = _cv(3, cc)
        if v and v.strip() and v.strip() != "0":
            dwg_val = v
            break
    try:
        # 先写值/格式（非合并时快），再 Merge——避免对合并区左上角设 NumberFormat 的慢路径
        cell = ws2.Cells(3, 2)
        cell.NumberFormat = "@"
        cell.Value = dwg_val
        cell.HorizontalAlignment = -4108
        cell.VerticalAlignment = -4108
        ws2.Range("B3:C3").Merge()
    except Exception:
        pass
    # 3b. Part 标签：D3 1列
    try:
        part_label = _cv(3, 5) or "Part Description:\n零件名称"
        ws2.Cells(3, 4).Value = part_label
    except Exception:
        pass
    # 3c. 名称值：E3:G3（3列）
    part_val = _cv(3, 7) or _cv(3, 8) or _cv(3, 6) or ""
    try:
        # 先写值/对齐，再 Merge（避免合并区左上角慢路径）
        _c = ws2.Cells(3, 5)
        _c.Value = part_val
        _c.HorizontalAlignment = -4108
        _c.VerticalAlignment = -4108
        ws2.Range("E3:G3").Merge()
    except Exception:
        pass
    # 3d. Cust Code 标签：H3 1列
    try:
        cc_label = _cv(3, 9) or "Customer Code：\n客户代码"
        ws2.Cells(3, 8).Value = cc_label
    except Exception:
        pass
    # 3e. 代码值：I3:L3（4列）
    cc_val = _cv(3, 12) or _cv(3, 11) or _cv(3, 10) or "0998"
    try:
        # 先写值/格式，再 Merge
        cell = ws2.Cells(3, 9)
        cell.NumberFormat = "@"
        cell.Value = cc_val
        cell.HorizontalAlignment = -4108
        cell.VerticalAlignment = -4108
        ws2.Range("I3:L3").Merge()
    except Exception:
        pass

    # 4. R4
    po_val = ""
    for cc in (2, 3, 4):
        v = _cv(4, cc)
        if v and v.strip():
            po_val = v
            break
    try:
        # 先写值/格式，再 Merge（避免合并区左上角 NumberFormat 慢路径）
        cell = ws2.Cells(4, 2)
        cell.NumberFormat = "@"
        cell.Value = po_val if po_val else " "
        cell.HorizontalAlignment = -4108
        cell.VerticalAlignment = -4108
        ws2.Range("B4:C4").Merge()
    except Exception:
        pass
    try:
        rev_label = _cv(4, 5) or "Rev:\n版本"
        ws2.Cells(4, 4).Value = rev_label
    except Exception:
        pass
    rev_val = _cv(4, 7) or _cv(4, 8) or _cv(4, 6) or ""
    try:
        # 先写值/对齐，再 Merge
        _c = ws2.Cells(4, 5)
        _c.Value = rev_val
        _c.HorizontalAlignment = -4108
        _c.VerticalAlignment = -4108
        ws2.Range("E4:G4").Merge()
    except Exception:
        pass
    try:
        temp_label = _cv(4, 9) or "Tempreture\n温度:"
        ws2.Cells(4, 8).Value = temp_label
    except Exception:
        pass
    temp_val = _cv(4, 12) or _cv(4, 11) or _cv(4, 10) or ""
    try:
        # 先写值/格式，再 Merge
        cell = ws2.Cells(4, 9)
        cell.NumberFormat = "@"
        cell.Value = temp_val
        cell.HorizontalAlignment = -4108
        cell.VerticalAlignment = -4108
        ws2.Range("I4:L4").Merge()
    except Exception:
        pass

    # 5. R5 签
    for addr, anchor, default_v in (
        ("A5:C5", 1, "Inspected  By（检验）:杨梦丽"),
        ("D5:F5", 4, "Reviewed  By（审核）:张瑛"),
        ("G5:I5", 7, "Final Disposition（最终判定）"),
        ("J5:L5", 10, "Remark（备注）:"),
    ):
        try:
            # 先写值/对齐（非合并时快），再 Merge（有内容 Merge 快，空历史区 Merge 慢）
            v = _cv(5, anchor) or default_v
            _c5 = ws2.Cells(5, anchor)
            _c5.Value = v
            _c5.HorizontalAlignment = -4108
            _c5.VerticalAlignment = -4108
            ws2.Range(addr).Merge()
        except Exception:
            pass

    # 6. R6-R8 签名日期 / 让步编号：合并 A6:C8 / D6:F8 / J6:L8（各占 6-8 三行），
    #    值从源左上角（A6/D6/J6）原样平移（源含 \n 换行的两行内容一并保留），水平垂直居中 + 自动换行。
    #    最终判定：合格/不合格/让步 分别合并 G:I（三列一格）对应第 6/7/8 行
    def _unmerge_any(addr):
        try:
            rng = ws2.Range(addr)
            r0, c0 = rng.Row, rng.Column
            r1 = r0 + rng.Rows.Count - 1
            c1 = c0 + rng.Columns.Count - 1
            for cc in range(c0, c1 + 1):
                c = ws2.Cells(r0, cc)
                if c.MergeCells:
                    c.MergeArea.UnMerge()
            for rr in range(r0, r1 + 1):
                c = ws2.Cells(rr, c0)
                if c.MergeCells:
                    c.MergeArea.UnMerge()
        except Exception:
            try:
                ws2.Range(addr).UnMerge()
            except Exception:
                pass
    def _unmerge2(addr):
        """安全解合并：MergeCells 数组批量检查，仅对合并格取 MergeArea 整体解。
        避免对非合并区/合并区子区域直接 UnMerge 触发 COM 异常（异常处理极慢，累计可达秒级）。"""
        try:
            rng = ws2.Range(addr)
            _mc = rng.MergeCells
            if not _mc:
                return
            if not isinstance(_mc, tuple):
                _mc = ((_mc,),)
            elif not isinstance(_mc[0], tuple):
                _mc = (_mc,)
            r0, c0 = rng.Row, rng.Column
            for _ri, _rm in enumerate(_mc):
                for _ci, _m in enumerate(_rm or ()):
                    if _m:
                        _c = ws2.Cells(r0 + _ri, c0 + _ci)
                        try:
                            if _c.MergeCells:
                                _c.MergeArea.UnMerge()
                        except Exception:
                            pass
        except Exception:
            try:
                ws2.Range(addr).UnMerge()
            except Exception:
                pass
    for _addr, _anchor, _defv in (
        ("A6:C8", 1, "(Sign/Employee No./Date签名日期)"),
        ("D6:F8", 4, "(Sign/Employee No./Date签名日期)"),
        ("J6:L8", 10, "此处填写让步申请编号"),
    ):
        # 合并区已在 cell_vals 段统一解除（A3:S12 范围），无需再次检查
        try:
            # 先写值（非合并时快），再 Merge（避免合并区左上角写值/设格式的慢路径）
            v = _cv_merge_block(_anchor, ncols=3)
            if not v:
                v = _defv
            ws2.Cells(6, _anchor).Value = v
            ws2.Range(_addr).Merge()
        except Exception:
            pass
    # 最终判定：合格/不合格/让步 分别合并 G:I（三列一格）对应第 6/7/8 行
    for _r, _def in ((6, "Passed/合格"), (7, "Failed/不合格"), (8, " UAI   让步")):
        # 合并区已在 cell_vals 段统一解除（A3:S12 范围）
        try:
            # 先写值，再 Merge（避免合并区左上角写值慢路径）
            v = _cv(_r, 7)
            if v is None or v == "":
                v = _def
            ws2.Cells(_r, 7).Value = v
            ws2.Range(f"G{_r}:I{_r}").Merge()
        except Exception:
            pass

    # 6.5 R9-11 汇总区：总数 G9 标签/H9 值，检查数 I9 标签/J9 值，K9 空（与 J9 中间竖线分隔），
    # Result 合并 L9:L11（三行一格，L11 子注被合并覆盖）
    # 原值从 cell_vals 字典读（合并区已在 cell_vals 段统一解除，子格值被清空；
    # 合并区情况取其左上角值——插入 B 列后源 G9:K9 合并区为 H9:L9，左上角 H9）
    _tot_val = _cv(9, 9) or _cv(9, 8)
    _smp_val = _cv(9, 12) or _cv(9, 8)
    try:
        ws2.Cells(9, 7).Value = "Total Qty\n（总数）:"
        ws2.Cells(9, 8).Value = _tot_val
        ws2.Cells(9, 9).Value = "Sample Qty\n（检查数）:"
        ws2.Cells(9, 10).Value = _smp_val
        ws2.Cells(9, 11).Value = ""
        for _cc in (7, 8, 9, 10, 11):
            ws2.Cells(9, _cc).HorizontalAlignment = -4108
            ws2.Cells(9, _cc).VerticalAlignment = -4108
    except Exception:
        pass
    try:
        # K9 保留为空独立单元格：J9 右框 + K9 左框形成中间竖线分隔
        ws2.Cells(9, 11).Borders(7).LineStyle = 1    # xlEdgeLeft continuous
        ws2.Cells(9, 11).Borders(7).Weight = 2       # xlThin
        # 先写值/格式（非合并时快），再 Merge L9:L11
        _lc = ws2.Cells(9, 12)
        _lc.Value = "Result（结果）\n(P合格 /F不合格)"
        _lc.WrapText = True
        _lc.HorizontalAlignment = -4108
        _lc.VerticalAlignment = -4108
        ws2.Range("L9:L11").Merge()
    except Exception:
        pass

    # 7. B 列表头：B9:B11（3 行）
    try:
        # 先写值/格式（非合并时快），再 Merge（规避合并区左上角慢路径）
        b_header = ws2.Cells(9, 2)
        b_header.Value = "检验项目"
        b_header.Font.Bold = False
        b_header.Font.Size = 10
        b_header.Font.Name = "宋体"
        b_header.HorizontalAlignment = -4108
        b_header.VerticalAlignment = -4108
        ws2.Range("B9:B11").Merge()
    except Exception:
        pass

    # 5. 找第一个数字读数所在行索引
    first_num_idx = None
    for i, row in enumerate(data_rows):
        readings = row.get("readings", [])
        if readings:
            r0 = str(readings[0]).strip()
            try:
                float(r0.replace("°", ""))
                first_num_idx = i
                break
            except (ValueError, TypeError):
                pass

    # 6. B12+ 数据行：跨行合并 B 列 + 写 A 列 item_no（字母编号加 0- 前缀）
    # 性能：先按组收集，再按连续行号分段批量数组化写入（替代逐行逐格 COM 调用）

    def _seg_batches(rows, vals):
        """把 (row, value) 列表按连续行号分段，返回 [(start, end, sub_vals)]"""
        segs = []
        if not rows:
            return segs
        s = e = rows[0]
        sub = [vals[0]]
        for k in range(1, len(rows)):
            if rows[k] == e + 1:
                e = rows[k]
                sub.append(vals[k])
            else:
                segs.append((s, e, sub))
                s = e = rows[k]
                sub = [vals[k]]
        segs.append((s, e, sub))
        return segs

    i = 0
    n = len(data_rows)
    a_rows, a_vals = [], []
    b_rows, b_vals = [], []
    merge_groups = []
    while i < n:
        row = data_rows[i]
        start_row = row["src_row"] + 1
        if start_row < 12:
            i += 1
            continue
        j = i
        while j + 1 < n and data_rows[j + 1]["item_no"] == row["item_no"] \
                and data_rows[j + 1]["src_row"] == data_rows[j]["src_row"] + 1:
            j += 1
        end_row = data_rows[j]["src_row"] + 1
        if first_num_idx is not None and i >= first_num_idx:
            label = "尺寸"
        else:
            label = "技术要求"
        # A 列 Item No：字母开头（A/a1 等）→ 0-A / 0-a1
        item_no = str(row.get("item_no", "") or "").strip()
        if item_no and not item_no.startswith('0-') and item_no[0].isalpha():
            item_no = '0-' + item_no
        if item_no:
            a_rows.append(start_row)
            a_vals.append(item_no)
        # B 列：组内所有行都写入 label（合并后仅左上角保留，等效原逻辑）
        for rr in range(start_row, end_row + 1):
            b_rows.append(rr)
            b_vals.append(label)
        if start_row == end_row:
            try:
                cell = ws2.Cells(start_row, 2)
                if cell.MergeCells:
                    cell.UnMerge()
            except Exception:
                pass
        else:
            merge_groups.append((start_row, end_row))
        i = j + 1

    # 批量写入 A 列（item_no）
    for (s, e, sub) in _seg_batches(a_rows, a_vals):
        try:
            _rng = ws2.Range(ws2.Cells(s, 1), ws2.Cells(e, 1))
            _rng.Value = [[v] for v in sub]
            _rng.HorizontalAlignment = -4108
            _rng.VerticalAlignment = -4108
        except Exception:
            pass
    # 批量写入 B 列（检验项目 label + 字体/对齐）
    for (s, e, sub) in _seg_batches(b_rows, b_vals):
        try:
            _rng = ws2.Range(ws2.Cells(s, 2), ws2.Cells(e, 2))
            _rng.Value = [[v] for v in sub]
            _rng.HorizontalAlignment = -4108
            _rng.VerticalAlignment = -4108
            _rng.Font.Bold = False
            _rng.Font.Size = 10
            _rng.Font.Name = "宋体"
        except Exception:
            pass
    # 跨行合并 B 列（写在批量赋值之后，合并后仅左上角保留 label）
    for (sr, er) in merge_groups:
        try:
            ws2.Range(f"B{sr}:B{er}").Merge()
        except Exception:
            pass


    # 7. 检具代码行（最后一行）：按值所在行定位，合并 A:L 跨两行 + 自动换行 + 两行行高
    try:
        used = ws2.UsedRange
        last_row = used.Row + used.Rows.Count - 1
        # 定位检具代码值所在行（从末行向上找第一个非空 A 列单元格）
        g_row = last_row
        while g_row > 1:
            try:
                if str(ws2.Cells(g_row, 1).Value or "").strip():
                    break
            except Exception:
                pass
            g_row -= 1
        # 解掉该行及下一行已有的合并（精确解 MergeArea，兼容 A:K / A:L 等任意列数）
        for rr in (g_row, g_row + 1):
            try:
                c = ws2.Cells(rr, 1)
                if c.MergeCells:
                    c.MergeArea.UnMerge()
            except Exception:
                pass
        _gv = ws2.Cells(g_row, 1).Value
        ws2.Range(f"A{g_row}:L{g_row + 1}").Merge()
        _gc = ws2.Cells(g_row, 1)
        _gc.Value = _gv
        _gc.WrapText = True
        _gc.HorizontalAlignment = -4108   # xlCenter（水平居中）
        _gc.VerticalAlignment = -4108     # xlCenter
        ws2.Rows(g_row).RowHeight = 22
        ws2.Rows(g_row + 1).RowHeight = 22
    except: pass

    # 7.5 外框补齐：所有新合并区域与最外围边框必须完整（thin continuous），
    #     包括 R1-R5 标题/表头外框、R6-R8 签名/判定/让步、R9-R11 汇总表头、检具代码
    try:
        for _addr in ("A1:L1", "A2:L2", "A3:L3", "A4:L4", "A5:L5",
                      "A6:C8", "D6:F8", "J6:L8", "G6:I6", "G7:I7", "G8:I8",
                      "A9:L11",
                      f"A{g_row}:L{g_row + 1}"):
            try:
                ws2.Range(_addr).BorderAround(LineStyle=1, Weight=2)  # xlContinuous, xlThin
            except Exception:
                pass
    except Exception:
        pass


    # 8. 删除完全空行（无任何值，包括纯空白行）
    # 跳过含合并的行（如检具代码 2 行合并的子行），避免误删

    # 8. 删除完全空行（无任何值，包括纯空白行）
    # 跳过含合并的行（如检具代码 2 行合并的子行），避免误删
    # 性能优化：用 range.Value 一次性读整列数据，避免逐 cell 调 COM

    try:
        used = ws2.UsedRange
        last_row = used.Row + used.Rows.Count - 1
        # 限制扫描范围：data_rows 末尾 src_row + 20（数据行之外的多余空行）
        if data_rows:
            last_row = min(last_row, max(r['src_row'] for r in data_rows) + 20)
        # 一次性读 12-43 列所有数据（matrix 形式返回 tuple of tuples）
        scan_end = min(last_row, max(used.Row, 11) + 1)  # 至少扫描 R12 起
        scan_range = ws2.Range(ws2.Cells(12, 1), ws2.Cells(last_row, used.Columns.Count))
        values = scan_range.Value  # matrix
        if not values:
            return
        # values 可能是 tuple of tuples（一行一格）
        if not isinstance(values, tuple):
            values = ((values,),)
        if values and not isinstance(values[0], tuple):
            values = (values,)
        empty_rows = []
        for ridx, row_vals in enumerate(values):
            r = 12 + ridx
            in_merge = False
            has_value = False
            # 检查值
            for cidx, v in enumerate(row_vals):
                if v is not None and str(v).strip():
                    has_value = True
                    break
            if has_value:
                continue
            # 跳过合并区域内的行（如检具代码两行合并的子行），避免破坏合并
            try:
                if ws2.Cells(r, 1).MergeCells:
                    continue
            except Exception:
                pass
            empty_rows.append(r)
        for r in sorted(empty_rows, reverse=True):
            try:
                ws2.Rows(r).Delete()
            except: pass
    except Exception:
        pass




def _convert_via_com_inplace(input_path, out_dir, on_progress, default_year=None, xl=None):
    # 预读源文件合并区清单（必须在此处、Excel 打开文件之前读：
    # COM 打开后文件被独占，openpyxl 无法再读取）
    _src_merges = []
    try:
        from openpyxl import load_workbook as _lwb0
        _sw0 = _lwb0(os.path.abspath(input_path), data_only=True)
        _src_merges = [str(_m) for _m in _sw0.active.merged_cells.ranges]
        _sw0.close()
    except Exception:
        pass
    """COM 原地修改方案：保留源文件所有格式（列宽/行高/图片/文本框/字体/复选框符号等）。

    仅做最小必要的改动：
      1. 测量数据列 OK/合格 → pass
      2. 越界读数 → Result 列写 F 并单元格标红
      3. 删除尾部无意义空行
      4. 保存为 .xlsx（统一输出格式）

    其他一切（标题字体、检验项目列、复选框字符 ■□þ、图片、文本框/表单编号、
    单元格大小、列宽行高、文字换行）完全保留源文件状态，不硬编码。
    返回 dict（与 convert_file 格式一致）或 None（COM 不可用时）。
    """
    try:
        import win32com.client
    except ImportError:
        return None

    ext = os.path.splitext(input_path)[1].lower()

    # ---- 第一步：用 xlrd/openpyxl 解析源文件结构（仅用于定位，不修改源） ----
    if ext == ".xls":
        import xlrd as _xlrd
        _wb = _xlrd.open_workbook(input_path, formatting_info=True)
        _s = _wb.sheet_by_index(0)
        _nrows, _ncols = _s.nrows, _s.ncols

        def _sv(r, c):
            try:
                v = _s.cell_value(r, c)
                return "" if v is None else str(v).strip()
            except Exception:
                return ""
    elif ext == ".xlsx":
        from openpyxl import load_workbook
        _wb = load_workbook(input_path, read_only=True, data_only=True)
        _ws = _wb.active
        _nrows, _ncols = _ws.max_row, _ws.max_column

        def _sv(r, c):
            try:
                v = _ws.cell(row=r + 1, column=c + 1).value
                return "" if v is None else str(v).strip()
            except Exception:
                return ""
    else:
        return None

    data_rows, total_qty, _layout, piece_headers = _parse_source(_sv, _nrows, _ncols)

    # 计算空行（1-based 行号）
    def _is_empty(row):
        spec = str(row.get("spec", "") or "").strip()
        mx = str(row.get("max", "") or "").strip()
        mn = str(row.get("min", "") or "").strip()
        readings = row.get("readings", [])
        item_no = str(row.get("item_no", "") or "").strip()
        item_digit = item_no.replace('.', '').replace('-', '')
        if item_digit.isdigit() and not (mx and mx not in ('/',)) and not (mn and mn not in ('/',)) and not any(str(r).strip() and r is not None for r in readings):
            return True
        if spec and spec not in ('/',):
            return False
        if (mx and mx not in ('/',)) or (mn and mn not in ('/',)):
            return False
        if any(str(r).strip() and r is not None for r in readings):
            return False
        return True

    # 仅计算 data_rows 中的空行（保留非数据行：检具代码、尾注、签名等）
    empty_rows_1b = []
    for row_info in data_rows:
        r_0 = row_info.get('src_row')
        if r_0 is not None and _is_empty(row_info):
            empty_rows_1b.append(r_0 + 1)  # 1-based

    # ---- 第二步：COM 打开源文件，做最小改动 ----
    # xl 可外部传入（批处理复用同一 Excel 实例）；未传入则本函数自行创建并在结束时退出
    own_xl = xl is None
    src_wb = None
    try:
        if own_xl:
            try:
                xl = win32com.client.DispatchEx("Excel.Application")
            except Exception:
                xl = win32com.client.Dispatch("Excel.Application")
        try:
            xl.Visible = False
        except Exception:
            pass
        try:
            xl.DisplayAlerts = False
        except Exception:
            pass
        # 性能优化：关闭屏幕刷新/自动重算/事件，避免逐格改动触发重绘与重算（提速关键）
        try:
            xl.ScreenUpdating = False
        except Exception:
            pass
        try:
            xl.Calculation = -4135  # xlCalculationManual
        except Exception:
            pass
        try:
            xl.EnableEvents = False
        except Exception:
            pass

        if on_progress:
            on_progress(30, "COM 打开源文件（保留原格式）...")

        src_wb = xl.Workbooks.Open(os.path.abspath(input_path))
        ws = src_wb.Sheets(1)

        # 读取 used range 的列数（用于按列扫描找读数列/结果列）
        used = ws.UsedRange
        n_cols = used.Column + used.Columns.Count - 1

        # ---- 预读源列宽/行高到缓存，供第 2 阶段平移（避免再次打开源文件，提速） ----
        # 注：字体无需逐格平移——SaveAs 成 xlsx 时已完整保留源字体（仅标题行与新增 B 列由程序单独设置）。
        try:
            _su = ws.UsedRange
            _src_last_r = _su.Row + _su.Rows.Count - 1
            _src_last_c = _su.Column + _su.Columns.Count - 1
            _SRC_COLW = {}
            _SRC_ROWH = {}
            for _r in range(1, _src_last_r + 1):
                try:
                    _SRC_ROWH[_r] = ws.Rows(_r).RowHeight
                except Exception:
                    pass
            for _c in range(1, _src_last_c + 1):
                try:
                    _SRC_COLW[_c] = ws.Columns(_c).ColumnWidth
                except Exception:
                    pass
        except Exception:
            _SRC_COLW, _SRC_ROWH = {}, {}
            _src_last_r = _src_last_c = 0

        # 探测读数列 / 结果列：扫描 R9-R11 的表头文本
        reading_cols_1b = []  # 1-based 列号列表
        result_col_1b = None

        for r in range(9, 12):
            for c in range(1, n_cols + 1):
                v = ws.Cells(r, c).Value
                if v is None:
                    continue
                v_str = str(v)
                if result_col_1b is None and ('Result' in v_str or '结果' in v_str):
                    result_col_1b = c
                if ('Actual Readings' in v_str or '测量数据' in v_str) and not reading_cols_1b:
                    # 找出该单元格的合并区域 → 读数列范围
                    try:
                        ma = ws.Cells(r, c).MergeArea
                        if ma and ma.Columns.Count > 1:
                            start_col = ma.Column
                            end_col = start_col + ma.Columns.Count - 1
                            if result_col_1b and end_col >= result_col_1b:
                                end_col = result_col_1b - 1
                            reading_cols_1b = list(range(start_col, end_col + 1))
                        else:
                            reading_cols_1b = [c]
                    except Exception:
                        reading_cols_1b = [c]

        if not reading_cols_1b:
            reading_cols_1b = list(range(7, 12))
        if result_col_1b is None:
            result_col_1b = (reading_cols_1b[-1] + 1) if reading_cols_1b else 12

        if on_progress:
            on_progress(40, "标题保留源格式（SaveAs 后单独处理字号）...")

        # 标题区暂不动：完全保留源合并、字体、间距、R3 上边框
        # SaveAs xlsx 后再重新打开用 GetCharacters 设置"检验报告"为 14pt

        if on_progress:
            on_progress(50, "修改 ok/合格 → pass（保留 ■□þ 复选框原样）...")

        # 1. 读数列规则：
        #    - 含 OK/合格 → pass（不区分大小写）
        #    - 单个 P / p → pass
        #    - 单个 / → NA
        for row_info in data_rows:
            ws_row = row_info['src_row'] + 1  # 0-based → 1-based
            for qi, reading in enumerate(row_info['readings']):
                if qi >= len(reading_cols_1b):
                    break
                col = reading_cols_1b[qi]
                cell = ws.Cells(ws_row, col)
                v = cell.Value
                if v is None:
                    continue
                v_str = str(v).strip()
                if not v_str:
                    continue
                v_lower = v_str.lower()
                # / → NA（精确等于）
                if v_str == '/':
                    cell.Value = 'NA'
                    continue
                # OK/合格 → pass（含整个单元格）
                if 'ok' in v_lower or '合格' in v_str:
                    cell.Value = 'pass'
                    continue
                # 单个 P/p → pass（避免误伤含 P 的复合词如 PG/PM）
                if v_lower == 'p':
                    cell.Value = 'pass'

        # 1b. 修正上下限填反（源单元格 max<min → 交换写入，让输出文件也修正）
        try:
            max_c = _layout.get("max_col")
            min_c = _layout.get("min_col")
            if max_c is not None and min_c is not None:
                for row_info in data_rows:
                    ws_row = row_info['src_row'] + 1
                    # 读源单元格原始值（未交换）
                    try:
                        src_mx = ws.Cells(ws_row, max_c + 1).Value
                        src_mn = ws.Cells(ws_row, min_c + 1).Value
                    except Exception:
                        continue
                    try:
                        mx_n = float(str(src_mx).strip()) if src_mx not in (None, "", "/") else None
                        mn_n = float(str(src_mn).strip()) if src_mn not in (None, "", "/") else None
                    except (ValueError, TypeError):
                        continue
                    if mx_n is None or mn_n is None or mx_n >= mn_n:
                        continue  # 正常区间，不动
                    # 源单元格填反了：交换写入（max 列写大值，min 列写小值）
                    try:
                        ws.Cells(ws_row, max_c + 1).Value = mn_n
                        ws.Cells(ws_row, min_c + 1).Value = mx_n
                    except Exception:
                        pass
        except Exception:
            pass

        if on_progress:
            on_progress(55, "清理头部图号/名称...")

        # 2. 清理头部信息（R3-R5）的图号/名称：去掉前缀、（）注释
        try:
            label_keywords = {
                'dwg_no_raw': ['DWG No', '图号'],
                'part_name': ['Part Description', '零件名称'],
            }
            for fld, kws in label_keywords.items():
                found_label_col = None
                found_label_row = None
                for tr0 in range(2, 5):
                    for tc0 in range(n_cols):
                        v = _sv(tr0, tc0)
                        if v and any(k in v for k in kws):
                            found_label_col = tc0
                            found_label_row = tr0
                            break
                    if found_label_col is not None:
                        break
                if found_label_col is None:
                    continue
                tr0 = found_label_row
                # 找值列：标签右侧第一个格
                vc0 = found_label_col + 1
                if vc0 >= n_cols:
                    continue
                v = _sv(tr0, vc0)
                if not v:
                    continue
                cleaned = _clean_dwg(v) if fld == 'dwg_no_raw' else None
                if cleaned is None:
                    # 通用清理：去括号、去前缀
                    tmp = re.sub(r"[（(][^）)]*[）)]", "", v).strip()
                    tmp = re.sub(r"\([^)]*\)", "", tmp).strip()
                    for prefix in ("JH-YCL-", "JH-", "YCL-"):
                        if tmp.upper().startswith(prefix.upper()):
                            tmp = tmp[len(prefix):].strip()
                            break
                    cleaned = tmp if tmp and tmp != v else None
                if cleaned and cleaned != v:
                    # .xls 模式不支持 Characters API，直接用 Value
                    # 但 Value 会自动转数字，丢失文本"92140042"格式 → 强制设文本格式
                    try:
                        cell = ws.Cells(tr0 + 1, vc0 + 1)
                        try:
                            cell.NumberFormat = "@"  # 文本格式
                        except Exception:
                            pass
                        cell.Value = cleaned
                    except Exception:
                        pass
        except Exception:
            pass

        # 3. 客户代码固定为 0998（扫描 R3-R5 找客户代码标签，写入右侧值）
        try:
            for tr0 in range(2, 5):
                for tc0 in range(n_cols):
                    v = _sv(tr0, tc0)
                    if v and ('Customer Code' in v or '客户代码' in v):
                        # 找右侧第一个非空格列（跳过中间的合并空格列）
                        target_col = None
                        for vc0 in range(tc0 + 1, min(tc0 + 6, n_cols)):
                            target = _sv(tr0, vc0)
                            if target is not None and target.strip() != "":
                                target_col = vc0
                                break
                        # 右侧都空就用标签下一列
                        if target_col is None:
                            target_col = min(tc0 + 1, n_cols - 1)
                        cell = ws.Cells(tr0 + 1, target_col + 1)
                        try:
                            cell.NumberFormat = "@"  # 文本格式，避免 0998→998.0
                        except Exception:
                            pass
                        cell.Value = "0998"
                        break  # 只处理第一个找到的标签
        except Exception:
            pass

        if on_progress:
            on_progress(70, "检查越界 → F + 红色...")

        # 2. 越界判定：读数不在 [min, max] → Result 列写 F 并红底
        def _parse_float(s):
            try:
                return float(str(s).replace('°', '').strip())
            except (ValueError, TypeError):
                return None

        for row_info in data_rows:
            minv = _parse_float(row_info.get('min', ''))
            maxv = _parse_float(row_info.get('max', ''))
            oor = False
            # 只有当有数值公差时，才检查越界
            for reading in row_info['readings']:
                if not reading:
                    continue
                rs = str(reading).strip()
                if rs.upper() in ('OK', '合格', 'PASS', 'P', ''):
                    continue
                v = _parse_float(reading)
                if v is None:
                    continue
                if (minv is not None and v < minv) or (maxv is not None and v > maxv):
                    oor = True
                    break
            if result_col_1b:
                ws_row = row_info['src_row'] + 1
                cell = ws.Cells(ws_row, result_col_1b)
                # 覆盖：先清公式，再写值
                try:
                    cell.Formula = ""
                except Exception:
                    pass
                try:
                    cell.ClearContents()
                except Exception:
                    pass
                if oor:
                    cell.Value = 'F'
                    cell.Interior.Color = 255  # 红色
                else:
                    cell.Value = 'P'

        if on_progress:
            on_progress(85, "删除空行（保留所有列宽/行高）...")

        # 3. 删除空行（从下往上，避免索引错乱）
        for row in sorted(set(empty_rows_1b), reverse=True):
            try:
                ws.Rows(row).Delete()
            except Exception:
                continue

        if on_progress:
            on_progress(95, "保存为 xlsx（保留源标题/列宽/行高/图片）...")

        out = compute_output_path(input_path, out_dir, total_qty=total_qty, default_year=default_year)
        if _has_failed(data_rows):
            out = _with_fail_prefix(out)
        src_wb.SaveAs(os.path.abspath(out), FileFormat=51)  # 51 = xlOpenXMLWorkbook (.xlsx)

        if on_progress:
            on_progress(97, "设置标题（拆分为两行 + 14pt + 中间横线）...")

        # 第 2 阶段：不重开文件，直接在同一工作簿会话上继续（省一次 Workbooks.Open）
        # 标题拆为两行（R1=Inspection Report / R2=检验报告），都 14pt，中间有粗横线
        try:
            ws2 = src_wb.Sheets(1)
            r1 = ws2.Cells(1, 1)
            cur_val = str(r1.Value or "")
            # 找英文/中文分隔
            en_part = ""
            cn_part = ""
            for _ln in cur_val.split('\n'):
                _s = _ln.strip()
                if _s and any(ch.isascii() and ch.isalpha() for ch in _s):
                    en_part = _s
                elif _s and any('\u4e00' <= ch <= '\u9fff' for ch in _s):
                    cn_part = _s
            if not en_part:
                en_part = "Inspection Report"
            if not cn_part:
                cn_part = "检验报告"

            # 取消原合并（如 A1:K2），分别合并 A1:K1 和 A2:K2
            try:
                if r1.MergeCells:
                    r1.MergeArea.UnMerge()
            except Exception:
                pass
            # 确保 R1、R2 都是独立合并行
            try:
                ws2.Range("A1:K1").Merge()
                ws2.Range("A2:K2").Merge()
            except Exception:
                pass

            # R1: 标题英文行（内容/对齐保留源拆分值，字体交下方“源字体平移”照搬源）
            ws2.Cells(1, 1).Value = en_part
            ws2.Cells(1, 1).HorizontalAlignment = -4108
            ws2.Cells(1, 1).VerticalAlignment = -4108

            # R2: 标题中文行（内容/对齐保留源拆分值，字体交下方“源字体平移”照搬源）
            ws2.Cells(2, 1).Value = cn_part
            ws2.Cells(2, 1).HorizontalAlignment = -4108
            ws2.Cells(2, 1).VerticalAlignment = -4108

            # 标题字体：统一 14pt（与原 v38「保持源文件字号」一致），不照搬源（源为 22pt）
            try:
                r1.Font.Name = "Times New Roman"
                r1.Font.Size = 14
                r1.Font.Bold = True
            except Exception:
                pass
            try:
                r2 = ws2.Cells(2, 1)
                r2.Font.Name = "宋体"
                r2.Font.Size = 14
                r2.Font.Bold = True
            except Exception:
                pass

            # 中间横线分隔：R1 底部 + R2 顶部 都设 medium 线
            # xlEdgeBottom=9, xlEdgeTop=8
            # xlLineStyleContinuous=1, xlMedium=-4138 (Weight)
            # 区域级设置（4 次 COM 调用替代 22 次逐格）
            try:
                _r1rng = ws2.Range("A1:K1")
                _r2rng = ws2.Range("A2:K2")
                _r1rng.Borders(9).LineStyle = 1    # xlContinuous
                _r1rng.Borders(9).Weight = -4138   # xlMedium
                _r2rng.Borders(8).LineStyle = 1    # xlContinuous
                _r2rng.Borders(8).Weight = -4138   # xlMedium
            except Exception:
                pass

            # 合并扩展（图号/PO 2→3，Inspected By/签名日期 3→4）+ 添加 B 列检验项目
            try:
                _extend_merges_and_add_inspection_column(ws2, data_rows, src_merges=_src_merges)
            except Exception:
                pass

            # ---- 第 2 阶段：列宽/行高平移（直接使用第 1 阶段预读的缓存，不再重复打开源文件） ----
            # 字体无需逐格平移：SaveAs 成 xlsx 时已完整保留源字体，仅标题行(1-2)与新增 B 列由程序单独设置。
            try:
                _oU = ws2.UsedRange
                _oLast_r = _oU.Row + _oU.Rows.Count - 1
                _oLast_c = _oU.Column + _oU.Columns.Count - 1

                # 1) 字体：SaveAs 已完整保留源字体，无需逐格平移（提速关键，且输出与源逐格一致）
                #    仅标题行(1-2)与新增 B 列由程序单独设置字体。

                # 2) 列宽显式复制（B 列插入映射：源 s → 输出 s(=1)/s+1(≥2)）
                #    数组化赋值：1 次 COM 调用替代逐列循环
                try:
                    _tgt_cols = []
                    _tgt_widths = []
                    for _sc in range(1, _src_last_c + 1):
                        _oc = _sc if _sc == 1 else _sc + 1
                        if _oc > _oLast_c:
                            continue
                        _w = _SRC_COLW.get(_sc)
                        if _w is None:
                            continue
                        _tgt_cols.append(_oc)
                        _tgt_widths.append(_w)
                    if _tgt_cols:
                        _rng = ws2.Range(
                            ws2.Cells(1, _tgt_cols[0]),
                            ws2.Cells(1, _tgt_cols[-1]))
                        try:
                            _rng.ColumnWidth = _tgt_widths
                        except Exception:
                            for _oc, _w in zip(_tgt_cols, _tgt_widths):
                                try:
                                    ws2.Columns(_oc).ColumnWidth = _w
                                except Exception:
                                    pass
                except Exception:
                    pass

                # 3) 行高显式复制（行号不变，B 列插入不影响行）
                try:
                    _tgt_rows = []
                    _tgt_heights = []
                    for _r in range(1, _src_last_r + 1):
                        _h = _SRC_ROWH.get(_r)
                        if _h is None:
                            continue
                        _tgt_rows.append(_r)
                        _tgt_heights.append(_h)
                    if _tgt_rows:
                        _rng = ws2.Range(
                            ws2.Cells(_tgt_rows[0], 1),
                            ws2.Cells(_tgt_rows[-1], 1))
                        try:
                            _rng.RowHeight = _tgt_heights
                        except Exception:
                            for _r, _h in zip(_tgt_rows, _tgt_heights):
                                try:
                                    ws2.Rows(_r).RowHeight = _h
                                except Exception:
                                    pass
                except Exception:
                    pass
            except Exception:
                pass

            src_wb.Save()
            src_wb.Close(SaveChanges=False)
        except Exception:
            pass

        tech = sum(1 for x in data_rows if x['insp'] == '技术要求')
        dim = sum(1 for x in data_rows if x['insp'] == '尺寸')

        if on_progress:
            on_progress(100, "完成")

        return {
            "output": out,
            "total_qty": total_qty,
            "tech": tech,
            "dim": dim,
            "input": input_path,
        }
    except Exception as e:
        try:
            if src_wb is not None:
                src_wb.Close(SaveChanges=False)
        except Exception:
            pass
        return None
    finally:
        try:
            # 仅自建实例才恢复设置并退出；外部复用实例由调用方统一管理
            if own_xl and xl is not None:
                try: xl.ScreenUpdating = True
                except Exception: pass
                try: xl.Calculation = -4105  # xlCalculationAutomatic
                except Exception: pass
                try: xl.EnableEvents = True
                except Exception: pass
                xl.Quit()
        except Exception:
            pass


def convert_file(input_path, out_dir=None, on_progress=None, default_year=None, xl=None):
    """转换单个检验报告文件，返回结果 dict。

    支持 .xls / .xlsx / .csv。
    优先使用 COM 原地修改（保留源文件全部格式：列宽/行高/图片/文本框/复选框字符等），
    失败时回退到 openpyxl 重建。

    default_year: 4 位年份字符串（如 "2026"），用于无年份短日期补全。
    xl: 外部传入的 Excel COM 实例（批处理复用，避免每文件重复启动 Excel）；
        为 None 时内部自行创建并在结束后退出。
    """
    ext = os.path.splitext(input_path)[1].lower()
    if on_progress:
        on_progress(5, "读取源文件...")

    # ---- .xls / .xlsx 优先尝试 COM 原地修改（保留源格式） ----
    if ext in (".xls", ".xlsx"):
        if on_progress:
            on_progress(10, "COM 原地修改（保留原列宽/图片/文本框）...")
        com_result = _convert_via_com_inplace(input_path, out_dir, on_progress, default_year=default_year, xl=xl)
        if com_result is not None:
            return com_result
        if on_progress:
            on_progress(10, "COM 不可用，回退 openpyxl 重建...")

    if ext == ".xls":
        import xlrd
        wb = xlrd.open_workbook(input_path, formatting_info=True)
        s = wb.sheet_by_index(0)
        nrows, ncols = s.nrows, s.ncols

        def sv(r, c):
            try:
                v = s.cell_value(r, c)
                return "" if v is None else str(v).strip()
            except Exception:
                return ""
        src_sv = sv
    elif ext == ".xlsx":
        from openpyxl import load_workbook
        wb = load_workbook(input_path, read_only=True, data_only=True)
        ws = wb.active
        nrows, ncols = ws.max_row, ws.max_column

        def sv(r, c):
            try:
                v = ws.cell(row=r + 1, column=c + 1).value
                return "" if v is None else str(v).strip()
            except Exception:
                return ""
        src_sv = sv
    elif ext == ".csv":
        import csv
        rows = []
        with open(input_path, "r", encoding="utf-8-sig", newline="") as fh:
            for row in csv.reader(fh):
                rows.append(row)
        nrows = len(rows)
        ncols = max((len(r) for r in rows), default=0)

        def sv(r, c):
            try:
                return (rows[r][c] or "").strip()
            except (IndexError, TypeError):
                return ""
        src_sv = sv
    else:
        raise ValueError("不支持的文件格式: " + ext)

    data_rows, total_qty, _layout, piece_headers = _parse_source(src_sv, nrows, ncols)

    # 删除尾部无意义行（只有编号或完全空白，无实际检验内容）
    def _is_empty_row(row):
        """判断是否为无意义的空行（无规格、无公差、无实测值）。

        加强版：若一行仅有 item_no(纯数字)、insp(尺寸/技术要求)、
        instr(简单数字) 而无 spec/max/min/任何有效读数，也视为空行。
        """
        spec = str(row.get("spec", "") or "").strip()
        mx = str(row.get("max", "") or "").strip()
        mn = str(row.get("min", "") or "").strip()
        readings = row.get("readings", [])
        instr = str(row.get("instr", "") or "").strip()
        insp = str(row.get("insp", "") or "").strip()
        item_no = str(row.get("item_no", "") or "").strip()

        spec_ok = spec and spec not in ("/",)
        num_ok = (mx and mx not in ("/",)) or (mn and mn not in ("/",))
        read_ok = any(str(r).strip() and r is not None for r in readings)

        # 加强检查2（优先）：编号为纯数字且无公差/无读数 → 视为空行
        # （如 MWA 尾部 14/15/…/21 行：只有 编号+尺寸+spec='6'，无上下限无读数）
        # 必须放在 spec_ok 早退之前，否则 spec='6' 会误判为有效行
        item_digit = item_no.replace('.', '').replace('-', '')
        if item_digit.isdigit() and not num_ok and not read_ok:
            return True

        # 加强检查1：若仅有 编号(纯数字)+检验项目(尺寸/技术要求)+检具(纯数字)
        # 而没有任何实际检验内容，视为空行（如截图中的 14/尺寸/6 行）
        has_only_meta = (
            item_no and item_no.isdigit()
            and insp in ("尺寸", "技术要求")
            and instr.isdigit()
            and not spec_ok and not num_ok and not read_ok
        )
        if has_only_meta:
            return True

        # 基本检查：有规格或公差或读数 → 非空
        if spec_ok or num_ok or read_ok:
            return False

        return True

    # 从尾部逐个删除空行
    while data_rows and _is_empty_row(data_rows[-1]):
        data_rows.pop()

    # 同时清除数据区中间可能混入的纯元数据行（仅有编号+尺寸+数字）
    data_rows = [row for row in data_rows if not _is_empty_row(row)]

    # 构建 Item No 合并分组（相同编号跨多行）
    item_no_merge_groups = {}
    cur_item = None
    cur_group = None
    for row in data_rows:
        it = row["item_no"]
        if not it:
            if cur_group:
                item_no_merge_groups[cur_group[0]] = cur_group
                cur_group = None
            cur_item = None
            continue
        if it != cur_item:
            if cur_group:
                item_no_merge_groups[cur_group[0]] = cur_group
            cur_item = it
            cur_group = [row["src_row"]]
        else:
            cur_group.append(row["src_row"])
    if cur_group:
        item_no_merge_groups[cur_group[0]] = cur_group
    item_no_merge_groups = {a: g for a, g in item_no_merge_groups.items()
                             if len(g) >= 2}

    header_info = _extract_header_info(input_path)

    out, tech, dim = _build_output_xlsx(
        data_rows, item_no_merge_groups, input_path, out_dir,
        on_progress, sv=src_sv, header_info=header_info,
        total_qty=total_qty, piece_headers=piece_headers)

    if ext == ".xlsx":
        try:
            wb.close()
        except Exception:
            pass

    # ---- 保留内嵌图片/Shape：用 Excel COM 从源文件复制到输出（原坐标） ----
    # CSV 无任何内嵌对象，直接跳过（避免每次白付 ~5s 的 Excel 启动/打开成本）；
    # .xlsx 先用 openpyxl 探测有无图片/图表，无则同样跳过。
    if ext != ".csv" and (ext == ".xls" or _has_shapes(input_path)):
        if on_progress:
            on_progress(92, "保留内嵌图片/Logo...")
        try:
            _copy_shapes_via_com(input_path, out, on_progress, xl=xl)
        except Exception:
            pass

    return {
        "output": out,
        "total_qty": total_qty,
        "tech": tech,
        "dim": dim,
        "input": input_path,
    }


def _has_shapes(path):
    """探测 xlsx 是否含内嵌图片/图表（决定是否需要 COM 复制 Shape）。

    用 openpyxl 普通模式加载，检查 _images / _charts；检测失败时保守返回 True
    （交给 COM 侧兜底）。CSV 无此问题（由调用方直接跳过）。
    """
    try:
        from openpyxl import load_workbook
        wb = load_workbook(path, read_only=False, data_only=False)
        try:
            ws = wb.active
            return bool(getattr(ws, "_images", None) or getattr(ws, "_charts", None))
        finally:
            wb.close()
    except Exception:
        return True


# --------------------------------------------------------------------------- #
# Excel COM 转换（用于 .xls 文件，保留内嵌图片/Shape/Logo/表单编号）
# --------------------------------------------------------------------------- #
def _copy_shapes_via_com(src_path, out_path, on_progress=None, xl=None):
    """用 Excel COM 把源文件的所有 Shape（图片/Logo/文本框/表单编号）按原始坐标
    复制到 openpyxl 生成的输出 xlsx 中，从而保留全部内嵌对象。

    对齐原理：输出与源文件列布局一致（均含插入的 B 列），复制列宽/行高后，
    以原始 Top/Left 坐标粘贴即可精确还原位置。返回复制的 Shape 数量；失败返回 0。
    xl: 外部传入的 Excel COM 实例（批处理复用）；为 None 时自行创建并在结束后退出。
    """
    try:
        import win32com.client
    except ImportError:
        return 0
    own_xl = xl is None
    try:
        if own_xl:
            try:
                xl = win32com.client.DispatchEx("Excel.Application")
            except Exception:
                xl = win32com.client.Dispatch("Excel.Application")
        try:
            xl.Visible = False
        except Exception:
            pass
        try:
            xl.DisplayAlerts = False
        except Exception:
            pass

        src = xl.Workbooks.Open(os.path.abspath(src_path))
        out = xl.Workbooks.Open(os.path.abspath(out_path))
        sWs = src.Sheets(1)
        oWs = out.Sheets(1)
        n = sWs.Shapes.Count
        if n == 0:
            src.Close(SaveChanges=False)
            out.Close(SaveChanges=False)
            return 0

        # 复制列宽/行高，保证像素坐标对齐
        try:
            maxC = sWs.UsedRange.Columns.Count
            for c in range(1, maxC + 1):
                oWs.Columns(c).ColumnWidth = sWs.Columns(c).ColumnWidth
            maxR = sWs.UsedRange.Rows.Count
            for r in range(1, maxR + 1):
                oWs.Rows(r).RowHeight = sWs.Rows(r).RowHeight
        except Exception:
            pass

        copied = 0
        for i in range(1, n + 1):
            sp = sWs.Shapes(i)
            top, left = sp.Top, sp.Left
            ok = False
            for _attempt in range(2):  # 剪贴板时序不稳，失败重试一次
                try:
                    sp.Copy()
                    oWs.Paste(oWs.Cells(1, 1))
                    ns = oWs.Shapes(oWs.Shapes.Count)
                    ns.Left = left
                    ns.Top = top
                    try:
                        xl.CutCopyMode = False  # 清空复制模式，避免下次粘贴错乱
                    except Exception:
                        pass
                    ok = True
                    break
                except Exception:
                    try:
                        xl.CutCopyMode = False
                    except Exception:
                        pass
            if ok:
                copied += 1

        out.Save()
        src.Close(SaveChanges=False)
        out.Close(SaveChanges=False)
        if on_progress:
            on_progress(95, f"已保留 {copied} 个图片/对象")
        return copied
    except Exception:
        return 0
    finally:
        try:
            # 仅自建实例才退出；外部复用实例由调用方统一管理
            if own_xl and xl is not None:
                xl.Quit()
        except Exception:
            pass
