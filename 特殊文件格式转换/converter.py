#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
检验报告模板转换核心模块
========================
把「原始检验数据」(源文件) 转换为内置模板 (template.xlsx, 即 LJG-FH56-1(1) 格式)。

原则:
  - 模板只提供格式(布局/合并/列宽行高/字体边框/固定文字)，逐格与模板一致;
  - 数据全部取自源文件; 源文件没有的字段 → 留空(不保留模板示例数据);
  - 数据区由内容驱动: 按「检验项目/公称尺寸」内容匹配行类型，不写死行号;
  - 不合格判定参考 audit_fail: 结果列为 F / 数值测量超出上下限 → 结果格标红底并写 F;
  - 输出文件名 = 源文件图号; 不合格文件加「不合格_」前缀。

写入路径:
  - 主路径: Excel COM 打开模板写入并另存(完整保留模板格式，参考 converter.py);
  - 兜底:   无 Excel 时用 openpyxl 重建(尽力保留样式)。
"""
import os
import re
import shutil
import sys
import tempfile
from copy import copy

from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

TEMPLATE_NAME = "template.xlsx"          # 内置模板文件名(与本模块同目录 / 打包进 exe)
SHEET_TITLE = "检验报告"                  # 展示用
RED_FILL_RGB = "FF0000"                  # 标红底色(与 audit_fail 一致)

# ---------------------------------------------------------------- #
# 文本判定工具
# ---------------------------------------------------------------- #
_PASS_TOKENS = {"p", "pass", "ok", "合格", "good", "accept", "y", "yes",
                "√", "✓", "合格品", "pass合格", "p合格"}
_FAIL_TOKENS = {"f", "fail", "ng", "不合格", "不良", "bad", "f不合格"}


def _norm(v):
    return "" if v is None else str(v).strip()


def _is_pass(v):
    s = _strip_symbols(_norm(v)).lower()
    return not s or s in _PASS_TOKENS or s.startswith("pass")


def _is_fail(v):
    s = _strip_symbols(_norm(v)).lower()
    return s in _FAIL_TOKENS or "不合格" in s or s.startswith("fail")


def _strip_symbols(t):
    return re.sub(r"[\sþ■☑✓□☐√·\-—–~～/\\（）()]", "", str(t or ""))


def _fnum(x):
    """尽力转 float; 转不了返回 None。"""
    if x is None:
        return None
    s = str(x).strip().replace("°", "").replace(",", "").replace("％", "")
    if not s:
        return None
    try:
        return float(s)
    except (ValueError, TypeError):
        return None


def _clean_filename(s):
    s = re.sub(r'[\\/:*?"<>|]', "_", str(s or "")).strip()
    return s or "未命名"


def _fmt_num(x):
    """数值转显示串: 3.10 → '3.1', 0.8 → '0.8' (去尾零)。"""
    s = ("%.6f" % x).rstrip("0").rstrip(".")
    return s if s else "0"


def _parse_tolerances(dim_text):
    """从尺寸文本内嵌公差解析 (上限值, 下限值)；解析不出返回 (None, None)。
      - 'C0.6±0.2'  → (0.8, 0.4)   对称公差
      - 'R0.8（0，-0.2）' → (0.8, 0.6) 括号偏差(上偏差, 下偏差)
    """
    s = str(dim_text or "").strip()
    if not s:
        return None, None
    # 对称公差: ± ／ ＋-
    m = re.search(r"(-?\d+(?:\.\d+)?)\s*[±]\s*(-?\d+(?:\.\d+)?)", s)
    if not m:
        m = re.search(r"(-?\d+(?:\.\d+)?)\s*[+\-]\s*(-?\d+(?:\.\d+)?)\s*$", s)
    if m:
        base = float(m.group(1))
        tol = abs(float(m.group(2)))
        return _fmt_num(base + tol), _fmt_num(base - tol)
    # 括号内嵌偏差: 'R0.8（0，-0.2）'
    m = re.search(r"(-?\d+(?:\.\d+)?)\s*[（(]\s*(-?\d+(?:\.\d+)?)\s*[,，]\s*(-?\d+(?:\.\d+)?)\s*[)）]", s)
    if m:
        base = float(m.group(1))
        up_dev = float(m.group(2))
        lo_dev = float(m.group(3))
        return _fmt_num(base + up_dev), _fmt_num(base + lo_dev)
    return None, None


def _fmt_cell_value(v):
    """数值去尾 .0(如 40.0 → 40)，文本原样。"""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    if v is None:
        return ""
    return str(v).strip()


# ---------------------------------------------------------------- #
# 模板定位
# ---------------------------------------------------------------- #
def _exe_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def find_template_path(custom=None):
    """定位内置模板文件: 自定义 > 打包进 exe(_MEIPASS) > 程序目录。"""
    if custom and os.path.isfile(custom):
        return custom
    _meipass = getattr(sys, "_MEIPASS", None)
    cands = []
    if _meipass:
        cands.append(os.path.join(_meipass, TEMPLATE_NAME))
    cands.append(os.path.join(_exe_dir(), TEMPLATE_NAME))
    for p in cands:
        if os.path.isfile(p):
            return p
    raise FileNotFoundError(f"找不到内置模板文件 {TEMPLATE_NAME}，请将模板放回程序目录")


# ---------------------------------------------------------------- #
# 源文件解析(openpyxl 快速读取，只读数据)
# ---------------------------------------------------------------- #
def _find_header_row(ws, ncols, max_scan=12):
    """找同时含 Item No 与 检验项目 的表头行(1-based)。"""
    for r in range(1, max_scan + 1):
        vals = [_norm(ws.cell(r, c).value).lower() for c in range(1, ncols + 1)]
        if any("item no" in v for v in vals) and any("检验项目" in v for v in vals):
            return r
    return None


def _col_by_keywords(ws, row, ncols, keywords):
    for c in range(1, ncols + 1):
        v = _norm(ws.cell(row, c).value)
        if v and any(k.lower() in v.lower() for k in keywords):
            return c
    return None


def detect_layout(ws):
    """按内容识别源表布局。返回 dict(列号均为 1-based)。"""
    nrows, ncols = ws.max_row, ws.max_column
    hdr = _find_header_row(ws, ncols)
    if hdr is None:
        raise ValueError("未识别到表头(找不到 Item No / 检验项目)，请检查源文件格式")

    # 尺寸行: 含「公称尺寸」的行(表头下方 1~3 行内)
    dim_row = None
    for r in range(hdr, min(hdr + 5, nrows + 1)):
        vals = [_norm(ws.cell(r, c).value) for c in range(1, ncols + 1)]
        if any("公称尺寸" in v for v in vals):
            dim_row = r
            break
    if dim_row is None:
        dim_row = hdr + 2

    item_no_col = _col_by_keywords(ws, hdr, ncols, ["item no", "编号"])
    insp_col = _col_by_keywords(ws, hdr, ncols, ["检验项目"])
    dim_col = _col_by_keywords(ws, dim_row, ncols, ["公称尺寸", "dimension"])
    max_col = _col_by_keywords(ws, dim_row, ncols, ["上限值", "max"])
    min_col = _col_by_keywords(ws, dim_row, ncols, ["下限值", "min"])
    instr_col = _col_by_keywords(ws, hdr, ncols, ["instrument", "检具"])
    # 结果列: Result 表头可能位于表头行(9)到尺寸行(11)之间 → 逐行找
    result_col = None
    for r in range(hdr, dim_row + 1):
        result_col = _col_by_keywords(ws, r, ncols, ["result"])
        if result_col:
            break
    if not result_col:
        result_col = _col_by_keywords(ws, dim_row, ncols,
                                      ["p合格", "f不合格", "p/f", "合格 /f"])

    # 实测列: 检具列(或下限列)右侧 ~ 结果列之前、数据区有值的全部列。
    # 新源文件 G10:K10 是合并表头、只有锚点 G10 有值，按"子表头列"取会漏掉后面几十列；
    # 直接取区间内任一数据行有值的列。
    reading_cols = []
    if result_col:
        data_start_est = dim_row + 1
        lo = max(instr_col or 0, min_col or 0) + 1
        hi = result_col
        for c in range(lo, hi):
            has_data = any(_norm(ws.cell(r, c).value)
                           for r in range(data_start_est, min(data_start_est + 45, nrows + 1)))
            if has_data:
                reading_cols.append(c)

    return {
        "nrows": nrows, "ncols": ncols,
        "hdr_row": hdr, "dim_row": dim_row,
        "item_no_col": item_no_col, "insp_col": insp_col, "dim_col": dim_col,
        "max_col": max_col, "min_col": min_col, "instr_col": instr_col,
        "result_col": result_col, "reading_cols": reading_cols,
        "data_start": dim_row + 1,
    }


_LABEL_BLACKLIST = (
    "disposition", "remark", "passed", "failed", "uai", "reviewed",
    "inspected", "temperature", "tempreture", "customer", "sample qty",
    "total qty", "签", "日期", "dwg", "part description", "po", "qty",
    "编号", "no.", "item no", "检验项目", "公称", "上限", "下限",
    "要求", "结果", "measuring", "备注",
)


def _looks_like_label(v):
    s = _norm(v).lower()
    if not s:
        return False
    return any(k in s for k in _LABEL_BLACKLIST)


def _value_right_of(ws, row, col, ncols, skip_cols=()):
    """标签格右侧(同一行)第一个非空值。跳过指定列与相邻标签格。"""
    for c in range(col + 1, min(col + 4, ncols + 1)):
        if c in skip_cols:
            continue
        v = ws.cell(row, c).value
        if v is None or not _norm(v):
            continue
        if _looks_like_label(v):   # 相邻格仍是标签(合并区结束后的下一个标签)，不是值
            continue
        return v
    return None


def _find_label(ws, max_row, ncols, keywords):
    """在顶部区域找含关键词的标签格，返回 (row, col) 或 None。大小写不敏感。"""
    for r in range(1, max_row + 1):
        for c in range(1, ncols + 1):
            v = _norm(ws.cell(r, c).value)
            if v and any(k.lower() in v.lower() for k in keywords):
                return r, c
    return None


def _label_value(ws, pos, ncols, keywords, skip_cols=()):
    """标签格取值: 优先右侧值; 右侧空时解析标签格内嵌值(如「Inspected By(检验):徐立中」)。"""
    r, c = pos
    v = _value_right_of(ws, r, c, ncols, skip_cols)
    if v:
        return v
    text = _norm(ws.cell(r, c).value)
    # 内嵌值: 取最后一个冒号后的内容(如「Inspected By (检验)：徐立中」)
    m = re.search(r"[：:]\s*([^：:\n]{1,30})$", text)
    if m:
        cand = m.group(1).strip()
        if cand and "/" not in cand and "版本" not in cand \
                and not any(k.lower() in cand.lower() for k in keywords) \
                and not any(lab in cand for lab in ("客户代码", "温度", "审核", "检验")):
            return cand
    for k in keywords:
        idx = text.lower().find(k.lower())
        if idx < 0:
            continue
        tail = text[idx + len(k):]
        tail = re.sub(r"^[\s:：\-—]+", "", tail)
        if not tail:
            continue
        # 尾串里再出现标签关键词(如「Customer Code：\n客户代码」)或含换行 → 仍是标签不是值
        if "\n" in tail:
            continue
        if any(k.lower() in tail.lower() for k in keywords):
            continue
        # 内嵌值须是纯数值，或简短文本且不含版本号/斜杠/冒号等标签残留
        if _fnum(tail) is not None:
            return tail
        if "/" not in tail and "版本" not in tail and "rev" not in tail.lower() \
                and ":" not in tail and "：" not in tail and len(tail) <= 20:
            return tail
    return ""


def extract_header(ws, layout):
    nrows, ncols = layout["nrows"], layout["ncols"]

    def right(keywords, row_lim=9, skip_cols=()):
        pos = _find_label(ws, row_lim, ncols, keywords)
        if not pos:
            return ""
        return _fmt_cell_value(_label_value(ws, pos, ncols, keywords, skip_cols))

    dwg = right(["dwg no", "图号"])
    part = right(["part description", "零件名称"])
    cust = right(["customer code", "客户代码"])
    po = right(["工单号", "\npo", "po\n"])
    if not po:
        po = right(["po"])
    temp = right(["tempreture", "温度"])
    insp_by = right(["inspected by", "inspected"])
    rev_by = right(["reviewed by", "reviewed"])
    total_qty = right(["total qty", "总数"])
    sample_qty = right(["sample qty", "检查数"])

    # 报告编号: 顶部区域含「报告编号」的单元格原文
    report_no = ""
    for r in range(1, 4):
        for c in range(1, ncols + 1):
            v = _norm(ws.cell(r, c).value)
            if "报告编号" in v:
                report_no = v
                break
        if report_no:
            break

    # 签名区日期: 在签名/日期区(rows 5~9)找日期时间文本
    date_re = re.compile(
        r"(\d{4}[-/.]\d{1,2}[-/.]\d{1,2}(?:[ T]\d{1,2}:\d{2}(?::\d{2})?)?"
        r"|\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}(?:[ T]\d{1,2}:\d{2}(?::\d{2})?)?)")
    # 签名区日期: 以「Reviewed By(审核)」标签所在列为界，左边=检验日期，右边=审核日期
    rev_label = _find_label(ws, 8, ncols, ["reviewed by", "审核"])
    rev_col = rev_label[1] if rev_label else 5
    insp_date, rev_date = "", ""
    for r in range(5, min(9, nrows) + 1):
        for c in range(1, ncols + 1):
            v = _norm(ws.cell(r, c).value)
            if not v:
                continue
            m = date_re.search(v)
            if not m:
                continue
            d = m.group(1)
            if c < rev_col and not insp_date:   # 检验签名区
                insp_date = d
            elif c >= rev_col and not rev_date:  # 审核签名区
                rev_date = d
    if not insp_date and not rev_date:
        # 兜底: 扫描整个表头区(1~11)取第一个日期
        for r in range(1, 11):
            for c in range(1, ncols + 1):
                m = date_re.search(_norm(ws.cell(r, c).value))
                if m:
                    insp_date = m.group(1)
                    break
            if insp_date:
                break

    return {
        "dwg_no": dwg, "part_name": part, "cust_code": cust,
        "po": po, "temp": temp,
        "inspected_by": insp_by, "reviewed_by": rev_by,
        "total_qty": total_qty, "sample_qty": sample_qty,
        "report_no": report_no,
        "insp_date": insp_date, "rev_date": rev_date,
    }


def extract_items(ws, layout):
    """读取数据行(rows >= data_start，到 检具代码/空行 为止)。"""
    nrows, ncols = layout["nrows"], layout["ncols"]
    reading_cols = layout["reading_cols"]
    items = []
    for r in range(layout["data_start"], nrows + 1):
        row_vals = [_norm(ws.cell(r, c).value) for c in range(1, ncols + 1)]
        if any("检具代码" in v for v in row_vals):
            break
        if not any(row_vals):
            break
        no = _fmt_cell_value(ws.cell(r, layout["item_no_col"]).value) if layout["item_no_col"] else ""
        insp = _norm(ws.cell(r, layout["insp_col"]).value) if layout["insp_col"] else ""
        dim = _norm(ws.cell(r, layout["dim_col"]).value) if layout["dim_col"] else ""
        if not no and not insp and not dim:
            continue   # 件号子表头行(1,2,3,4,5)等无项目行
        mx = _norm(ws.cell(r, layout["max_col"]).value) if layout["max_col"] else ""
        mn = _norm(ws.cell(r, layout["min_col"]).value) if layout["min_col"] else ""
        # 源 D/E 空 → 从尺寸文本内嵌公差补全(如 R0.8（0，-0.2）/ C0.6±0.2)
        if not mx or not mn:
            emx, emn = _parse_tolerances(dim)
            if not mx and emx:
                mx = emx
            if not mn and emn:
                mn = emn
        readings = [_norm(ws.cell(r, c).value) for c in reading_cols]
        result = _norm(ws.cell(r, layout["result_col"]).value) if layout["result_col"] else ""
        if not (no or insp or dim or mx or mn or any(readings) or result):
            continue
        items.append({
            "no": no, "insp": insp, "dim": dim,
            "max": mx, "min": mn, "readings": readings, "result": result,
            "src_row": r,
        })
    return items


# ---------------------------------------------------------------- #
# 行类型与转换规则
# ---------------------------------------------------------------- #
def classify_row(insp, dim):
    """按内容匹配行类型(不写死行号):
      - 含「零件编号」 → other_dims(其他尺寸符合图纸要求)
      - 含「量检具」   → gauge_info(量检具信息：)
      - 其余技术要求行(材质/版本等中间行) → delete(整行删除)
      - 其余           → normal(测量行)
    """
    if "零件编号" in insp or "零件号" in insp or "零件编号" in dim:
        return "other_dims"
    if "量检具" in insp or "量检具" in dim:
        return "gauge_info"
    if "技术要求" in insp or insp.strip() == "要求":
        return "delete"
    return "normal"


def measure_text(readings):
    """把多个实测值合并成显示文本:
      - 全部合格/空 → 合格
      - 单个数值   → 该值原样
      - 多个数值   → 最小值～最大值(全角～)
    """
    nums = []  # (float, 原字符串)
    for rd in readings:
        s = _norm(rd)
        if not s or _is_pass(s):
            continue
        f = _fnum(s)
        if f is not None:
            nums.append((f, s))
    if not nums:
        return "合格"
    if len(nums) == 1:
        return nums[0][1]
    lo = min(nums, key=lambda x: x[0])
    hi = max(nums, key=lambda x: x[0])
    if lo[0] == hi[0]:
        return lo[1]
    return f"{lo[1]}～{hi[1]}"


def row_is_fail(item):
    """判定一行是否不合格(对齐 audit_fail._scan_core):
      - 源结果列为 F/不合格 → 不合格
      - 数值读数超出上下限(任一侧) → 不合格; max<min 自动交换
    """
    if _is_fail(item["result"]):
        return True
    minv, maxv = _fnum(item["min"]), _fnum(item["max"])
    if minv is None and maxv is None:
        return False
    if minv is not None and maxv is not None and maxv < minv:
        minv, maxv = maxv, minv
    for rd in item["readings"]:
        s = _norm(rd)
        if not s or _is_pass(s):
            continue
        v = _fnum(s)
        if v is None:
            continue
        if (minv is not None and v < minv) or (maxv is not None and v > maxv):
            return True
    return False


def prepare_rows(items):
    """为每条源数据行生成输出行。返回 (rows, has_fail, deleted_count)。"""
    rows = []
    has_fail = False
    deleted = 0
    for it in items:
        rtype = classify_row(it["insp"], it["dim"])
        if rtype == "delete":
            deleted += 1
            continue
        fail = row_is_fail(it) if rtype == "normal" else False
        if fail:
            has_fail = True
        row = {
            "no": it["no"], "insp": it["insp"], "dim": it["dim"],
            "max": it["max"], "min": it["min"],
            "type": rtype, "fail": fail,
            "result": "F" if fail else "P",
        }
        if rtype == "normal":
            row["measure"] = measure_text(it["readings"])
        rows.append(row)
    return rows, has_fail, deleted


# ---------------------------------------------------------------- #
# 输出路径
# ---------------------------------------------------------------- #
def compute_output_path(parsed, input_path, out_dir=None):
    base = _clean_filename(parsed["header"].get("dwg_no") or "")
    if not base:
        base = _clean_filename(os.path.splitext(os.path.basename(input_path))[0])
    od = os.path.abspath(out_dir) if out_dir else os.path.dirname(os.path.abspath(input_path))
    os.makedirs(od, exist_ok=True)
    out = os.path.join(od, base + ".xlsx")
    if parsed["has_fail"] and not os.path.basename(out).startswith("不合格_"):
        out = os.path.join(od, "不合格_" + base + ".xlsx")
    return out


# ---------------------------------------------------------------- #
# Excel COM 写入(主路径: 完整保留模板格式)
# ---------------------------------------------------------------- #
def _com_write(xl, template_path, parsed, out_path, on_progress=None):
    """用 COM 打开模板 → 写值 → 重建数据区 → 另存。"""
    wb = xl.Workbooks.Open(template_path)
    try:
        ws = wb.Worksheets(1)
        hdr = parsed["header"]
        rows = parsed["rows"]
        n = len(rows)

        # 先捕获模板页脚(检具代码)原文，再清数据区
        footer_text = _norm(ws.Range("A28").Value)
        if not footer_text or "检具代码" not in footer_text:
            footer_text = ("检具代码：游标卡尺YK、深度尺SC、内径千分尺NJ、外径千分尺WJ、投影仪TY、"
                           "温度仪WD、千分表QF、百分表BF、角度尺JD、硬度仪YD、粗糙度仪CCD、卷尺JC、"
                           "塞规SG、环规HG、膜厚仪MHY、三坐标CMM、治具ZJ")

        # ---------- 头部数据(源文件有则写，没有则留空) ----------
        # 报告编号(模板 I1:L1): 源文件有则原样写入，没有则清空(不留模板示例)
        ws.Range("I1").Value = hdr.get("report_no") or ""
        ws.Range("C3").Value = hdr.get("dwg_no") or ""          # 图号(不带版本)
        ws.Range("G3").Value = hdr.get("part_name") or ""       # 零件名称
        ws.Range("K3").Value = hdr.get("cust_code") or ""       # 客户代码(源无→空)
        ws.Range("G4").NumberFormat = "@"                       # PO 按文本写入(防前导零丢失)
        ws.Range("G4").Value = hdr.get("po") or ""              # PO 工单号
        ws.Range("K4").NumberFormat = "@"                       # 温度按文本写入, 保留源文本
        ws.Range("K4").Value = hdr.get("temp") or ""            # 温度(源无→空)
        ws.Range("C5").Value = hdr.get("inspected_by") or ""    # 检验人
        # E5:G5 模板自带「Reviewed By(审核):」标签, 不写入也不清空, 保持原样
        # 签名格: 标签 + 源文件日期
        _sign_label = "(Sign/Employee No./Date签名日期)"
        insp_sign = f"{_sign_label}\n{hdr['insp_date']}" if hdr.get("insp_date") else _sign_label + "\n\n"
        rev_sign = f"{_sign_label}\n{hdr['rev_date']}" if hdr.get("rev_date") else _sign_label + "\n\n"
        ws.Range("A6").Value = insp_sign
        ws.Range("E6").Value = rev_sign
        # 总数/检查数
        ws.Range("I9").Value = hdr.get("total_qty") or 2
        ws.Range("L9").Value = hdr.get("sample_qty") or 2

        # ---------- 重建数据区(行 12 起，动态) ----------
        # 拆掉模板数据区与页脚的合并(避免残留)
        try:
            ws.Range("A12:L29").UnMerge()
        except Exception:
            for r in range(12, 30):
                try:
                    ws.Range(f"A{r}:L{r}").UnMerge()
                except Exception:
                    pass
        # 清掉模板示例数据
        ws.Range("A12:L29").ClearContents()

        # 行数对齐: 少则删、多则插(插入行从下行复制格式 = 正常测量行)
        if n < 16:
            ws.Range(f"A{12 + n}:A27").EntireRow.Delete()
        elif n > 16:
            for _ in range(n - 16):
                ws.Range("A12:A12").EntireRow.Insert(
                    Shift=-4121, CopyOrigin=0)  # xlDown, xlFormatFromRightOrBelow

        # 页脚(检具代码)行: 放到数据区之后
        footer_r = 12 + n
        ws.Range(f"A{footer_r}").Value = footer_text
        ws.Range(f"A{footer_r}:L{footer_r + 1}").Merge()

        # 数据行
        for i, row in enumerate(rows):
            r = 12 + i
            rtype = row["type"]
            # 统一先拆该行可能的合并，再按类型合并
            try:
                ws.Range(f"A{r}:L{r}").UnMerge()
            except Exception:
                pass
            ws.Range(f"A{r}").Value = row["no"]
            ws.Range(f"B{r}").Value = row["insp"]
            if rtype == "other_dims":
                ws.Range(f"B{r}").Value = "其他尺寸符合图纸要求"
                ws.Range(f"B{r}:L{r}").Merge()
                ws.Range(f"B{r}").WrapText = True
            elif rtype == "gauge_info":
                ws.Range(f"B{r}").Value = "量检具信息："
                ws.Range(f"C{r}:L{r}").Merge()
                ws.Range(f"B{r}").WrapText = True
            else:  # normal
                ws.Range(f"C{r}").Value = row["dim"]
                ws.Range(f"D{r}").Value = row["max"] or "/"
                ws.Range(f"E{r}").Value = row["min"] or "/"
                ws.Range(f"F{r}").Value = row["measure"]
                ws.Range(f"F{r}:K{r}").Merge()
                ws.Range(f"F{r}").WrapText = True
                ws.Range(f"L{r}").Value = row["result"]
                ws.Range(f"L{r}").HorizontalAlignment = -4108  # xlCenter 居中
                if row["fail"]:
                    ws.Range(f"L{r}").Interior.Color = 0x0000FF  # 红底(字体不变)

        # 行高: 正常 33.8 / 其他尺寸 22 / 量检具 36 / 页脚 14.3
        for i, row in enumerate(rows):
            r = 12 + i
            if row["type"] == "other_dims":
                ws.Rows(r).RowHeight = 22.0
            elif row["type"] == "gauge_info":
                ws.Rows(r).RowHeight = 36.0
            else:
                ws.Rows(r).RowHeight = 33.8
        ws.Rows(footer_r).RowHeight = 14.3
        ws.Rows(footer_r + 1).RowHeight = 14.3

        # 另存为 .xlsx(覆盖已存在)
        if os.path.exists(out_path):
            os.remove(out_path)
        wb.SaveAs(out_path, FileFormat=51)  # xlOpenXMLWorkbook
        return out_path
    finally:
        wb.Close(SaveChanges=False)


# ---------------------------------------------------------------- #
# openpyxl 兜底写入(无 Excel 时, 尽力保留格式)
# ---------------------------------------------------------------- #
def _openpyxl_write(template_path, parsed, out_path, on_progress=None):
    from openpyxl.styles import Border, Side

    wb = load_workbook(template_path)
    ws = wb.active
    hdr = parsed["header"]
    rows = parsed["rows"]
    n = len(rows)
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # 拆数据区合并
    for m in list(ws.merged_cells.ranges):
        if m.min_row >= 12:
            ws.unmerge_cells(str(m))
    # 清数据
    for r in range(12, 30):
        for c in range(1, 13):
            ws.cell(row=r, column=c).value = None

    def _set(r, c, v):
        ws.cell(row=r, column=c, value=v)

    _set(1, 9, hdr.get("report_no") or "")
    _set(3, 3, hdr.get("dwg_no") or "")
    _set(3, 7, hdr.get("part_name") or "")
    _set(3, 11, hdr.get("cust_code") or "")
    _set(4, 7, hdr.get("po") or "")
    ws.cell(row=4, column=7).number_format = "@"
    _set(4, 11, hdr.get("temp") or "")
    ws.cell(row=4, column=11).number_format = "@"
    _set(5, 3, hdr.get("inspected_by") or "")
    # E5:G5 模板自带「Reviewed By(审核):」标签, 不清空
    _sign_label = "(Sign/Employee No./Date签名日期)"
    _set(6, 1, f"{_sign_label}\n{hdr['insp_date']}" if hdr.get("insp_date") else _sign_label + "\n\n")
    _set(6, 5, f"{_sign_label}\n{hdr['rev_date']}" if hdr.get("rev_date") else _sign_label + "\n\n")
    _set(9, 9, hdr.get("total_qty") or 2)
    _set(9, 12, hdr.get("sample_qty") or 2)

    # 数据行样式原型: 模板第 12 行
    proto = {}
    for c in range(1, 13):
        sc = ws.cell(row=12, column=c)
        proto[c] = copy(sc._style) if sc.has_style else None

    def _apply_style(r, cols):
        for c in cols:
            cell = ws.cell(row=r, column=c)
            if proto.get(c) is not None:
                cell._style = copy(proto[c])
            cell.border = border

    for i, row in enumerate(rows):
        r = 12 + i
        rtype = row["type"]
        _apply_style(r, range(1, 13))
        _set(r, 1, row["no"])
        _set(r, 2, row["insp"])
        if rtype == "other_dims":
            _set(r, 2, "其他尺寸符合图纸要求")
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=12)
        elif rtype == "gauge_info":
            _set(r, 2, "量检具信息：")
            ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=12)
        else:
            _set(r, 3, row["dim"])
            _set(r, 4, row["max"] or "/")
            _set(r, 5, row["min"] or "/")
            _set(r, 6, row["measure"])
            ws.merge_cells(start_row=r, start_column=6, end_row=r, end_column=11)
            _set(r, 12, row["result"])
            if row["fail"]:
                ws.cell(row=r, column=12).fill = PatternFill(
                    patternType="solid", fgColor=RED_FILL_RGB)
        ws.row_dimensions[r].height = (22.0 if rtype == "other_dims"
                                       else 36.0 if rtype == "gauge_info" else 33.8)

    footer_r = 12 + n
    _apply_style(footer_r, range(1, 13))
    _apply_style(footer_r + 1, range(1, 13))
    _set(footer_r, 1, "检具代码：游标卡尺YK、深度尺SC、内径千分尺NJ、外径千分尺WJ、投影仪TY、"
                      "温度仪WD、千分表QF、百分表BF、角度尺JD、硬度仪YD、粗糙度仪CCD、卷尺JC、"
                      "塞规SG、环规HG、膜厚仪MHY、三坐标CMM、治具ZJ")
    ws.merge_cells(start_row=footer_r, start_column=1, end_row=footer_r + 1, end_column=12)
    ws.row_dimensions[footer_r].height = 14.3
    ws.row_dimensions[footer_r + 1].height = 14.3

    # 数据区之后的模板残留行(如 29)清空行高
    wb.save(out_path)
    return out_path


def _xls_to_xlsx(xl, input_path):
    """.xls → 临时 .xlsx(用 Excel COM 打开另存), 返回临时文件路径。"""
    tmp = os.path.join(tempfile.gettempdir(),
                       os.path.splitext(os.path.basename(input_path))[0] + "_tmp.xlsx")
    if os.path.exists(tmp):
        os.remove(tmp)
    wb = xl.Workbooks.Open(os.path.abspath(input_path))
    try:
        wb.SaveAs(tmp, FileFormat=51)  # xlOpenXMLWorkbook
    finally:
        wb.Close(SaveChanges=False)
    return tmp


# ---------------------------------------------------------------- #
# 单文件转换入口
# ---------------------------------------------------------------- #
def convert_file(input_path, out_dir=None, on_progress=None, xl=None,
                 template_path=None, force_openpyxl=False):
    """转换单个源文件。xl: 可复用的 Excel COM 实例(GUI 批量时传入)。"""
    if on_progress:
        on_progress(5, "读取源文件...")

    # .xls 老格式 → 先用 COM 转成 .xlsx 再解析
    ext = os.path.splitext(input_path)[1].lower()
    orig_path = input_path
    own_xl = None
    if ext == ".xls":
        if xl is None:
            try:
                import pythoncom
                import win32com.client
                pythoncom.CoInitialize()
                own_xl = win32com.client.DispatchEx("Excel.Application")
                own_xl.Visible = False
                own_xl.DisplayAlerts = False
                xl = own_xl
            except Exception:
                xl = None
        if xl is None:
            raise ValueError(".xls 文件需本机安装 Excel(COM) 才能转换")
        input_path = _xls_to_xlsx(xl, input_path)

    try:
        tpl = find_template_path(template_path)

        # 1) openpyxl 快速解析(只读数据，不碰格式)
        wb = load_workbook(input_path, data_only=True)
        try:
            ws = wb.active
            layout = detect_layout(ws)
            header = extract_header(ws, layout)
            items = extract_items(ws, layout)
        finally:
            wb.close()
        if not items:
            raise ValueError("源文件没有数据行(表头之后无内容)")

        rows, has_fail, deleted = prepare_rows(items)
        parsed = {"header": header, "rows": rows, "has_fail": has_fail}
        out_path = compute_output_path(parsed, orig_path, out_dir)

        if on_progress:
            on_progress(40, f"构建输出 {os.path.basename(out_path)}...")
        # 2) 写入: COM 优先(保留模板格式)，失败降级 openpyxl
        if not force_openpyxl and xl is not None:
            try:
                _com_write(xl, tpl, parsed, out_path, on_progress)
            except Exception:
                if on_progress:
                    on_progress(60, "COM 写入失败，改用 openpyxl 兜底...")
                _openpyxl_write(tpl, parsed, out_path, on_progress)
        else:
            _openpyxl_write(tpl, parsed, out_path, on_progress)

        if on_progress:
            on_progress(100, f"{'不合格' if has_fail else '合格'} -> {os.path.basename(out_path)}")

        # ---------- 操作日志(供 GUI 详情展示) ----------
        log = []
        h = header
        if h.get("dwg_no"):
            log.append({"level": "info", "op": "图号",
                        "msg": f"写入图号 {h['dwg_no']}"})
        else:
            log.append({"level": "warn", "op": "图号",
                        "msg": "源文件未识别到图号，输出文件名采用源文件名"})
        log.append({"level": "info", "op": "报告编号",
                    "msg": f"源文件{'有' if h.get('report_no') else '无'} → 输出{'写入' if h.get('report_no') else '留空'}"})
        log.append({"level": "info", "op": "温度",
                    "msg": f"源文件{'有' if h.get('temp') else '无'} → 输出{'写入 ' + str(h['temp']) if h.get('temp') else '留空'}"})
        log.append({"level": "info", "op": "检验人",
                    "msg": f"源文件{'有' if h.get('inspected_by') else '无'} → 输出{'写入 ' + str(h['inspected_by']) if h.get('inspected_by') else '留空'}"})
        slash = sum(1 for r in rows if r["type"] == "normal" and (not r["max"] or not r["min"]))
        if slash:
            log.append({"level": "info", "op": "上下限",
                        "msg": f"空白上下限 {slash} 处自动填 /"})
        merged = sum(1 for r in rows if r["type"] == "normal")
        if merged:
            log.append({"level": "info", "op": "测量数据",
                        "msg": f"{merged} 行测量值合并为范围(原多列读数合并到 F:K 大格)"})
        for i, r in enumerate(rows):
            if r["fail"]:
                log.append({"level": "err", "op": "不合格",
                            "msg": f"第 {i + 1} 项「{r['insp']}」超差/源结果F → 结果列标红并写 F"})
        log.append({"level": "info", "op": "文件",
                    "msg": f"输出 {'不合格_' if has_fail else ''}{os.path.basename(out_path)}"})
        if deleted:
            log.append({"level": "info", "op": "技术要求行",
                        "msg": f"删除中间技术要求行 {deleted} 行(材质/版本等)"})
        return {"input": orig_path, "output": out_path, "is_fail": has_fail,
                "rows": len(rows), "tech": sum(1 for r in rows if r["type"] != "normal"),
                "slash_filled": slash, "log": log}
    finally:
        # 清理 .xls 转换产生的临时文件与自建的 COM 实例
        if input_path != orig_path and os.path.exists(input_path):
            try:
                os.remove(input_path)
            except Exception:
                pass
        if own_xl is not None:
            try:
                own_xl.Quit()
            except Exception:
                pass
            try:
                import pythoncom
                pythoncom.CoUninitialize()
            except Exception:
                pass


# ---------------------------------------------------------------- #
# CLI
# ---------------------------------------------------------------- #
def main():
    import argparse
    ap = argparse.ArgumentParser(description="检验报告模板转换(命令行)")
    ap.add_argument("inputs", nargs="+", help="源文件(.xls/.xlsx/.xlsm)")
    ap.add_argument("-o", "--out-dir", default=None, help="输出目录(默认源文件目录)")
    ap.add_argument("--no-com", action="store_true", help="禁用 Excel COM，强制 openpyxl")
    args = ap.parse_args()

    xl = None
    if not args.no_com:
        try:
            import pythoncom
            import win32com.client
            pythoncom.CoInitialize()
            xl = win32com.client.DispatchEx("Excel.Application")
            xl.Visible = False
            xl.DisplayAlerts = False
        except Exception:
            xl = None

    try:
        for f in args.inputs:
            res = convert_file(f, args.out_dir, xl=xl, force_openpyxl=args.no_com)
            print(f"[{'不合格' if res['is_fail'] else '合格'}] {f} -> {res['output']}")
    finally:
        if xl is not None:
            try:
                xl.Quit()
            except Exception:
                pass


# ---------------------------------------------------------------- #
# 多进程 worker 批处理池(仿 audit_fail: 独立子进程+独立 Excel、超时重试、逐份进度)
# ---------------------------------------------------------------- #
import json as _json
import threading as _threading
import queue as _queue
import time as _time
import subprocess as _subprocess
import uuid as _uuid


class ExcelSession:
    """复用的 Excel COM 进程(每个 worker 子进程一个独立实例)。"""

    def __init__(self):
        self._xl = None
        self.xl_pid = None

    def get(self):
        if self._xl is None:
            import pythoncom
            import win32com.client
            pythoncom.CoInitialize()
            xl = win32com.client.DispatchEx("Excel.Application")
            xl.Visible = False
            xl.DisplayAlerts = False
            xl.AskToUpdateLinks = False
            try:
                xl.AutomationSecurity = 3  # 禁宏自动运行
            except Exception:
                pass
            self._xl = xl
            try:
                import win32process
                _, self.xl_pid = win32process.GetWindowThreadProcessId(xl.Hwnd)
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


def precheck(path):
    """拖入即预检: openpyxl 快读验证可读性。失败不拒收(降级由 COM 兜底)。"""
    out = {"ok": True, "error": ""}
    try:
        if not os.path.isfile(path):
            out["ok"] = False
            out["error"] = "文件不存在"
            return out
        ext = os.path.splitext(path)[1].lower()
        if ext not in (".xlsx", ".xlsm"):
            return out
        wb = load_workbook(path, data_only=True, read_only=True)
        try:
            ws = wb.active
            detect_layout(ws)   # 能识别布局即视为可读
        finally:
            wb.close()
    except Exception as e:
        out["error"] = str(e)[:120]   # 不 ok=False, 照常进池由 worker 处理
    return out


def run_worker_mode(out_dir=None):
    """子进程 worker 主循环: stdin 逐行收 {file} 任务, stdout 逐行回 JSON 事件。"""
    if win32com_available():
        try:
            import pythoncom
            pythoncom.CoInitialize()
        except Exception:
            pass
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stdin.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

    def emit(obj):
        sys.stdout.write(_json.dumps(obj, ensure_ascii=False, default=str) + "\n")
        sys.stdout.flush()

    sess = ExcelSession()
    excel_pid = None
    app_name = app_ver = ""
    try:
        xl = sess.get()
        try:
            import win32process
            _, excel_pid = win32process.GetWindowThreadProcessId(xl.Hwnd)
            app_name = str(xl.Name or "")
            app_ver = str(xl.Version or "")
        except Exception:
            pass
    except Exception:
        pass
    emit({"event": "ready", "excel_pid": excel_pid,
          "app_name": app_name, "app_version": app_ver})

    try:
        for line in sys.stdin:
            line = line.strip()
            if not line:
                continue
            try:
                task = _json.loads(line)
            except Exception:
                continue
            f = task.get("file")
            if not f:
                continue
            try:
                def cb(pct, msg, _f=f):
                    emit({"event": "phase", "file": _f, "pct": pct, "msg": msg})
                res = convert_file(f, out_dir=out_dir, on_progress=cb, xl=sess.get())
                emit({"event": "done", "file": f, "ok": True, "result": res})
            except Exception as e:  # noqa: BLE001
                emit({"event": "done", "file": f, "ok": False, "error": str(e)})
    except Exception:
        pass
    finally:
        try:
            sess.quit()
        except Exception:
            pass
    return 0


def win32com_available():
    try:
        import win32com.client  # noqa: F401
        return True
    except Exception:
        return False


class BatchPool:
    """多进程 worker 池: 每 worker 独立 Excel, 单文件超时杀进程重试, 逐份完成回调。"""

    def __init__(self, max_workers=2, timeout=30.0, max_retry=1):
        self.max_workers = max_workers
        self.timeout = float(timeout)
        self.max_retry = int(max_retry)
        self.workers = []
        self.events = None

    def _worker_cmd(self, out_dir):
        if getattr(sys, "frozen", False):
            return [sys.executable, "--worker-mode", "--out", out_dir]
        return [sys.executable, "-m", "converter", "--worker-mode", "--out", out_dir]

    def _spawn(self):
        wid = len(self.workers)
        flags = getattr(_subprocess, "CREATE_NO_WINDOW", 0)
        proc = _subprocess.Popen(
            self._worker_cmd(self._out_dir),
            stdin=_subprocess.PIPE, stdout=_subprocess.PIPE,
            stderr=_subprocess.DEVNULL, encoding="utf-8", errors="replace",
            creationflags=flags,
        )
        w = {"id": wid, "proc": proc, "tmp_out": None, "busy": None,
             "excel_pid": None, "dead": False, "ready": False}
        self.workers.append(w)

        def _reader(w_=w):
            try:
                for line in proc.stdout:
                    line = line.strip()
                    if not line:
                        continue
                    try:
                        obj = _json.loads(line)
                    except Exception:
                        continue
                    self.events.put(("msg", w_["id"], obj))
            except Exception:
                pass
            self.events.put(("eof", w_["id"], None))

        t = _threading.Thread(target=_reader, daemon=True)
        t.start()
        w["reader"] = t
        return w

    def _send(self, w, obj):
        try:
            w["proc"].stdin.write(_json.dumps(obj, ensure_ascii=False) + "\n")
            w["proc"].stdin.flush()
        except Exception:
            pass

    def _kill(self, w):
        if w["excel_pid"]:
            try:
                _subprocess.run(["taskkill", "/F", "/PID", str(w["excel_pid"])],
                               capture_output=True, timeout=10)
            except Exception:
                pass
        if w["proc"] and w["proc"].poll() is None:
            try:
                w["proc"].kill()
            except Exception:
                pass
        w["dead"] = True

    def run(self, files, out_dir, on_event=None, pres=None):
        self.workers = []
        self.events = _queue.Queue()
        self._on_event = on_event or (lambda e: None)
        pending = list(files)
        attempts = {}
        results = {}
        out_dir = os.path.abspath(out_dir) if out_dir else None
        self._out_dir = out_dir
        total = len(files)

        for _ in range(self.max_workers):
            self._spawn()

        def _busy():
            return any(w["busy"] is not None for w in self.workers)

        def _requeue(w, why):
            f = w["busy"]["file"]
            attempts[f] = attempts.get(f, 0) + 1
            self._kill(w)
            w["busy"] = None
            if attempts[f] <= self.max_retry:
                pending.append(f)
                self._spawn()
                self._on_event({"type": "retry", "file": f,
                                "attempt": attempts[f], "max": self.max_retry, "why": why})
            else:
                results[f] = {"file": f, "ok": False, "fail": False, "output": None,
                              "error": "%s，已重试%d次，请人工检查" % (why, self.max_retry),
                              "reason": "", "marked": False, "norm": {}, "log": []}
                self._on_event({"type": "done", "file": f, "ok": False, "result": results[f]})
                self._on_event({"type": "progress", "done": len(results), "total": total})

        def _dispatch(w):
            f = pending.pop(0)
            w["busy"] = {"file": f, "start": _time.time()}
            self._send(w, {"file": f})

        def _on_done(w, obj):
            f = obj.get("file")
            if not f:
                return
            if obj.get("ok"):
                res = obj.get("result") or {}
                results[f] = {"file": f, "ok": True,
                              "fail": bool(res.get("is_fail")),
                              "output": res.get("output"),
                              "reason": "", "marked": res.get("is_fail"),
                              "norm": {"slash_filled": res.get("slash_filled", 0)},
                              "log": res.get("log", []), "error": None}
                self._on_event({"type": "done", "file": f, "ok": True, "result": results[f]})
            else:
                err = obj.get("error") or "未知错误"
                if any(k in err for k in ("RPC", "服务器不可用", "Call was rejected",
                                          "连接断开", "0x800706BA")):
                    _requeue(w, "Excel 连接断开(RPC)")
                    return
                results[f] = {"file": f, "ok": False, "fail": False, "output": None,
                              "error": err, "reason": "", "marked": False,
                              "norm": {}, "log": []}
                self._on_event({"type": "done", "file": f, "ok": False, "result": results[f]})
            self._on_event({"type": "progress", "done": len(results), "total": total})
            w["busy"] = None
            if not w["dead"] and pending:
                _dispatch(w)

        try:
            while pending or _busy():
                try:
                    while True:
                        kind, wid, obj = self.events.get_nowait()
                        w = self.workers[wid]
                        if kind == "eof":
                            if w["busy"]:
                                _requeue(w, "worker 异常退出")
                            else:
                                self._kill(w)
                                if w["tmp_out"]:
                                    shutil.rmtree(w["tmp_out"], ignore_errors=True)
                                self._spawn()
                            continue
                        if obj.get("event") == "ready":
                            w["ready"] = True
                            w["excel_pid"] = obj.get("excel_pid")
                            if w["busy"] is None and pending and not w["dead"]:
                                _dispatch(w)
                            continue
                        if obj.get("event") == "done":
                            _on_done(w, obj)
                            continue
                except _queue.Empty:
                    pass

                now = _time.time()
                for w in list(self.workers):
                    if w["busy"] and not w["dead"] and (now - w["busy"]["start"]) > self.timeout:
                        _requeue(w, "处理超时(%.0fs)" % self.timeout)
                for w in list(self.workers):
                    if w["ready"] and w["busy"] is None and not w["dead"] and pending:
                        _dispatch(w)
                _time.sleep(0.02)
        finally:
            for w in list(self.workers):
                try:
                    w["proc"].stdin.close()
                except Exception:
                    pass
            deadline = _time.time() + 3.0
            while _time.time() < deadline:
                if all(w["proc"].poll() is not None for w in self.workers):
                    break
                _time.sleep(0.1)
            for w in list(self.workers):
                if w["proc"].poll() is None:
                    self._kill(w)
                if w["tmp_out"]:
                    shutil.rmtree(w["tmp_out"], ignore_errors=True)

        return [results[f] for f in files if f in results]


if __name__ == "__main__":
    if "--worker-mode" in sys.argv:
        _out = None
        if "--out" in sys.argv:
            _i = sys.argv.index("--out")
            if _i + 1 < len(sys.argv):
                _out = sys.argv[_i + 1]
        sys.exit(run_worker_mode(_out))
    main()
