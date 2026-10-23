# -*- coding: utf-8 -*-
"""报告布局探测（模板无关，基于表头文字定位）。

探测内容（每张表）：
  - DWG No（图号）
  - 测量数据起始列（"Actual Readings / 测量数据" 标签所在列）
  - 测量序号行（含 "公称尺寸 / Dimension" 的表头行）
  - 测量数据列（起始列 -> 结果列之前）
  - 结果列（含 "结果 / P合格"）、最大值列（"最大值"）、最小值列（"最小值"）
  - 总数 / 检查数 数值单元格
  - 数据区起止行
并建立 全局件号 -> (表名, 局部列号) 的映射，供抽取使用。
"""
import os
import shutil
import tempfile

import openpyxl

from .errors import ExtractError
from .filename_parser import parse_report_filename
from .models import ReportInfo, SheetLayout


_XLS_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"  # OLE2 复合文档（老 .xls/BIFF）签名


def _is_fake_xlsx(path: str) -> bool:
    """扩展名是 .xlsx/.xlsm、内容却是老 .xls(BIFF/OLE2) 的「假 xlsx」。

    WPS 按老 .xls 格式写盘却套 .xlsx 名时会出现：Excel 签名校验失败拒绝打开，
    这里按文件头魔数识别，供读取前自动转换修复。
    """
    if not path.lower().endswith((".xlsx", ".xlsm")):
        return False
    try:
        with open(path, "rb") as f:
            return f.read(8) == _XLS_MAGIC
    except Exception:  # noqa: BLE001
        return False


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
        return v.strip().lstrip("-").isdigit()
    return False


def _grid_of(ws, cap_r=220, cap_c=70):
    data = ws.iter_rows(min_row=1, max_row=cap_r, min_col=1, max_col=cap_c, values_only=True)
    grid = {}
    for i, row in enumerate(data):
        r = i + 1
        for j, v in enumerate(row):
            if v is not None and str(v).strip() != "":
                grid[(r, j + 1)] = v
    return grid


def _find(grid, substr, cap_r, cap_c):
    """返回所有含 substr 的 (r, c)。"""
    sub = substr.lower()
    out = []
    for (r, c), v in grid.items():
        if r > cap_r or c > cap_c:
            continue
        if sub in str(v).lower():
            out.append((r, c))
    return out


def _detect_sheet(ws, cap_r=400, cap_c=600):
    # cap_c 只是安全上界（不扫满 Excel 的 16384 列）；真正的右边界是"结果列"，
    # 一旦定位到结果列，测量列自动收在它前面，不依赖任何写死的件数上限。
    # cap_r 只用于「布局探测」（表头/序号行/结果列/测量列都在前 400 行内）；
    # 数据区终点不依赖 cap_r：以「检具代码」行为界，找不到则扫到最后一个有内容的行。
    grid = _grid_of(ws, cap_r, cap_c)
    name = ws.title

    # DWG No
    dwg = ""
    dwg_hits = _find(grid, "DWG No", cap_r, cap_c)
    if dwg_hits:
        r, c = dwg_hits[0]
        dwg = _norm(grid.get((r, c + 1)) or grid.get((r, c + 2)))
    if not dwg:
        dwg = _norm(grid.get((3, 2)))  # 退而求其次：B3

    # Part Description（零件名称）——与 DWG No 同排，取标签右侧值
    part = ""
    part_hits = _find(grid, "Part Description", cap_r, cap_c)
    if part_hits:
        r, c = part_hits[0]
        part = _norm(grid.get((r, c + 1)) or grid.get((r, c + 2)))
    if not part:
        part = _norm(grid.get((3, 6)))  # 退而求其次：F3

    # 测量数据起始列
    meas_hits = _find(grid, "Actual Readings", cap_r, cap_c) or _find(grid, "测量数据", cap_r, cap_c)
    meas_start = meas_hits[0][1] if meas_hits else 6

    # 表头行（含 公称尺寸 / Dimension）
    hdr_hits = _find(grid, "公称尺寸", cap_r, cap_c) or _find(grid, "Dimension", cap_r, cap_c)
    hdr_row = hdr_hits[0][0] if hdr_hits else (meas_hits[0][0] if meas_hits else 11)

    # 真序号行：表头行往下 0~3 行内，找"从 meas_start 列起是连续 1,2,3,4,5"的那一行。
    # 有的模板「公称尺寸」和「Actual Readings」同排，真正序号在下一行；旧模板则就在本行。
    seq_row = hdr_row
    for off in range(0, 4):
        r = hdr_row + off
        nums = []
        for c in range(meas_start, meas_start + 6):
            v = grid.get((r, c))
            nums.append(int(v) if _is_int(v) else None)
        if nums[0] == 1 and nums[1] == 2 and nums[2] == 3:
            seq_row = r
            break

    # 结果 / 最大 / 最小列：结果列在全表头区找（可能和 Total/Sample 同排，不在序号行）
    result_col = None
    max_col = None
    min_col = None
    # 先认「P合格/F不合格」这种结果列表头（它才是结果列真身），
    # 再退而求其次找「Result」——避免命中 Actual Readings 上方的 Result 标题格
    for key in ("P合格", "F不合格", "Result"):
        hits = _find(grid, key, 15, cap_c)
        if hits:
            result_col = hits[0][1]
            break
    for c in range(meas_start, cap_c + 1):
        v = grid.get((seq_row, c))
        if v is None:
            continue
        s = str(v)
        if result_col is None and ("结果" in s or "P合格" in s or "F不合格" in s or "Result" in s):
            result_col = c
        if max_col is None and "最大值" in s:
            max_col = c
        if min_col is None and "最小值" in s:
            min_col = c

    # 测量数据列 = 起始列 -> 结果列之前
    if result_col and result_col > meas_start:
        sample_cols = list(range(meas_start, result_col))
    else:
        # 结果列没找到：测量数据列取 起始列起的连续有读数块（最多 25 列）
        sample_cols = []
        for c in range(meas_start, meas_start + 25):
            # 该列在数据区有内容即视为测量列
            has = any(grid.get((r, c)) is not None for r in range(seq_row + 1, min(cap_r, seq_row + 60) + 1))
            if has:
                sample_cols.append(c)
            elif sample_cols:
                break

    # 测量序号：纯按物理列位置（从左到右第几根测量列 = 第几件）。
    # 标签里的件号可能带前缀（2640001/C01）、可能乱序、可能多行叠写，一律不解析、不依赖。
    seq_numbers = list(range(1, len(sample_cols) + 1))
    # 各测量列在源报告序号行上的原始文本（新报告照抄，保证数据列可溯源）
    seq_labels = [_norm(grid.get((seq_row, c))) for c in sample_cols]

    # 总数 / 检查数 数值单元格（标签通常在序号行上方的第 9 行附近）
    total_label = _find(grid, "Total Qty", cap_r, cap_c) or _find(grid, "总数", cap_r, cap_c)
    total_row = total_label[0][0] if total_label else (seq_row - 2)
    total_cell = _find_value_cell(grid, total_row, cap_c, ["Total Qty", "总数"])
    sample_label = _find(grid, "Sample Qty", cap_r, cap_c) or _find(grid, "检查数", cap_r, cap_c)
    sample_row = sample_label[0][0] if sample_label else (seq_row - 2)
    sample_cell = _find_value_cell(grid, sample_row, cap_c, ["Sample Qty", "检查数"])

    # 数据区起止：末尾以「检具代码」行为界（客户行数不定，禁止用固定行数截断）。
    # 单次流式扫描完成三件事（read_only 模式下同一 sheet 只能有一个活跃迭代器）：
    #   1) 命中「检具代码/检具编号/Gauge Code/...」行 -> 数据区终点 = 该行 - 1；
    #   2) 未命中 -> 回退为「最后一个有内容的行」（扫到 _SAFE_CAP，不再受 400 行限制）；
    #   3) 顺带收集各测量列的读数（col_values），最后按 data_last_row 截断。
    data_first_row = seq_row + 1
    right_limit = max([result_col or 0, max_col or 0, min_col or 0,
                       sample_cols[-1] if sample_cols else meas_start])
    _GAUGE_KWS = ("检具代码", "检具编号", "gauge code", "fixture code",
                  "gauge no", "fixture no")
    _SAFE_CAP = 50000
    _scan_max = min(int(getattr(ws, "max_row", 0) or 0) or _SAFE_CAP, _SAFE_CAP)
    _scan_cols = min(max(right_limit + 8, sample_cols[-1] if sample_cols else meas_start),
                     cap_c)
    gauge_row = None
    last_data_row = data_first_row - 1
    raw = {c: [] for c in sample_cols}
    _it = ws.iter_rows(min_row=data_first_row, max_row=_scan_max,
                       min_col=1, max_col=_scan_cols, values_only=True)
    for i, row_vals in enumerate(_it):
        r = data_first_row + i
        joined = " ".join(str(v) for v in row_vals if v is not None)
        if any(k in joined.lower() for k in _GAUGE_KWS):
            gauge_row = r
            break
        if any(v is not None for v in row_vals):
            last_data_row = r
        for c in sample_cols:
            raw[c].append(row_vals[c - 1])
    data_last_row = (gauge_row - 1) if gauge_row is not None else last_data_row

    # 测量列原始读数（探测时一次性读出，供生成新报告直接使用，避免逐格 COM 读取）
    _keep = max(data_last_row - data_first_row + 1, 0)
    col_values = {c: raw[c][:_keep] for c in sample_cols}

    return SheetLayout(
        name=name,
        dwg=_norm(dwg),
        part=_norm(part),
        seq_row=seq_row,
        sample_cols=sample_cols,
        seq_numbers=seq_numbers,
        seq_labels=seq_labels,
        result_col=result_col,
        max_col=max_col,
        min_col=min_col,
        total_cell=total_cell,
        sample_cell=sample_cell,
        data_first_row=data_first_row,
        data_last_row=data_last_row,
        col_values=col_values,
    )


def _find_value_cell(grid, near_row, cap_c, labels):
    """在 near_row 行（或邻近 1 行）找含 labels 的标签，取其右侧第一个数值单元格。"""
    for row_window in (near_row, near_row - 1, near_row + 1):
        for c in range(1, cap_c + 1):
            v = grid.get((row_window, c))
            if v is None:
                continue
            s = str(v)
            if any(lbl.lower() in s.lower() for lbl in labels):
                for cc in range(c + 1, min(cap_c, c + 10) + 1):
                    w = grid.get((row_window, cc))
                    if _is_int(w):
                        return (row_window, cc)
    return None


def extract_identity(path: str):
    """轻量读取报告内部身份（DWG No 图号 + Part Description 零件名）。

    供匹配阶段做多信号交叉验证；读取失败 / 旧版 .xls（BIFF）等无法快速读取的
    场景返回 ("", "")，由调用方回退到「仅文件名」匹配，绝不因此阻塞流程。
    注意：扩展名 .xlsx 但内容实为老 .xls 的「假 xlsx」（WPS 产物）也会自动
    COM 转换后再读，不直接判失败。
    返回 (dwg, part)，取自第一张含 DWG No 标签的工作表。
    """
    if not path.lower().endswith((".xlsx", ".xlsm")):
        return "", ""  # .xls 需要 COM 转换，匹配阶段不启动 Excel
    conv = None
    try:
        try:
            wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        except Exception:  # noqa: BLE001
            # 假 .xlsx：老 .xls 内容套 .xlsx 名，openpyxl 读不了 -> COM 转临时再读
            if not _is_fake_xlsx(path):
                return "", ""
            conv = _com_convert_to_xlsx(path)
            try:
                wb = openpyxl.load_workbook(conv, data_only=True, read_only=True)
            except Exception:  # noqa: BLE001
                return "", ""
        try:
            for ws in wb.worksheets:
                # 直接扫描单元格文本（read_only 下逐格取）
                dwg = part = ""
                for row in ws.iter_rows(min_row=1, max_row=12, max_col=40):
                    for cell in row:
                        if cell.value is None:
                            continue
                        s = str(cell.value)
                        if "DWG No" in s and not dwg:
                            v = ws.cell(cell.row, cell.column + 1).value or \
                                ws.cell(cell.row, cell.column + 2).value
                            dwg = _norm(v)
                        if "Part Description" in s and not part:
                            v = ws.cell(cell.row, cell.column + 1).value or \
                                ws.cell(cell.row, cell.column + 2).value
                            part = _norm(v)
                if dwg or part:
                    return dwg, part
        except Exception:  # noqa: BLE001
            return "", ""
        finally:
            try:
                wb.close()
            except Exception:  # noqa: BLE001
                pass
        return "", ""
    finally:
        if conv and os.path.exists(conv):
            try:
                os.remove(conv)
            except Exception:  # noqa: BLE001
                pass


def _com_convert_to_xlsx(path: str) -> str:
    """用 Excel COM 把 openpyxl 读不了的旧格式（真 .xls / 假 .xlsx）转成临时 .xlsx。

    返回临时 .xlsx 绝对路径（调用方负责删除）；转换失败抛 ExtractError(FILE_READ_ERROR)。
    """
    import pythoncom
    import win32com.client

    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False, prefix="xls2xlsx_")
    tmp.close()
    out = os.path.abspath(tmp.name)
    # Excel 也按扩展名判定格式：假 .xlsx（OLE2 内容）必须先换 .xls 名才能打开
    open_path = path
    tmp_copy = None
    if _is_fake_xlsx(path):
        fd, tmp_copy = tempfile.mkstemp(suffix=".xls", prefix="xls2xlsx_")
        os.close(fd)
        shutil.copyfile(path, tmp_copy)
        open_path = tmp_copy
    try:
        pythoncom.CoInitialize()
        xl = win32com.client.DispatchEx("Excel.Application")
        xl.Visible = False
        xl.DisplayAlerts = False
        xl.AskToUpdateLinks = False
        try:
            wb = xl.Workbooks.Open(os.path.abspath(open_path), ReadOnly=True, UpdateLinks=0)
            try:
                # 51 = xlOpenXMLWorkbook (.xlsx，去宏)。布局探测不需要宏。
                wb.SaveAs(out, FileFormat=51)
            finally:
                wb.Close(SaveChanges=False)
        finally:
            xl.Quit()
    except ExtractError:
        raise
    except Exception as e:  # noqa: BLE001
        if os.path.exists(out):
            try:
                os.remove(out)
            except Exception:  # noqa: BLE001
                pass
        raise ExtractError(
            "FILE_READ_ERROR",
            f"旧版/异常 .xls 格式转换失败（需本机安装 Excel）：{e}",
            file=path,
        )
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:  # noqa: BLE001
            pass
        if tmp_copy and os.path.exists(tmp_copy):
            try:
                os.remove(tmp_copy)
            except Exception:  # noqa: BLE001
                pass
    return out


def _maybe_convert_xls(path: str) -> "Optional[str]":
    """openpyxl 读不了的旧格式统一用 Excel COM 转成临时 .xlsx 再读。

    触发条件：扩展名为 .xls（BIFF），或扩展名为 .xlsx/.xlsm 但内容实为
    老 .xls 的「假 xlsx」（WPS 产物，Excel 会拒绝打开）。
    返回临时 .xlsx 路径（调用方负责删除）；本身可读的 .xlsx 返回 None。
    转换失败抛 ExtractError(FILE_READ_ERROR)。
    """
    if not (path.lower().endswith(".xls") or _is_fake_xlsx(path)):
        return None
    return _com_convert_to_xlsx(path)


def detect_layout(path: str) -> ReportInfo:
    parse = parse_report_filename(path)  # 先校验文件名
    # .xls 老格式 openpyxl 读不了，先用 Excel 转成临时 .xlsx
    conv_path = _maybe_convert_xls(path)
    read_path = conv_path or path
    try:
        try:
            wb = openpyxl.load_workbook(read_path, data_only=True, read_only=True)
        except Exception as e:  # noqa: BLE001
            raise ExtractError("FILE_READ_ERROR", f"报告文件读取失败：{e}", file=path)

        sheets = []
        sample_map = {}
        values = {}
        try:
            next_gi = 1   # 下一个期望全局件号（跨表累加）
            for ws in wb.worksheets:
                # 隐藏表不参与：它往往是模板残留，会抢占全局件号并导致复制报错
                if getattr(ws, "sheet_state", "visible") != "visible":
                    continue
                sl = _detect_sheet(ws)
                sheets.append(sl)
                nums = sl.seq_numbers
                if not nums:
                    continue
                # 序号跨表规则：
                #  - 本表序号从 1 重新开始（局部编号，如第二页标 1-5 实为全局 6-10）：
                #    整体加偏移，使第一个 1 落到 next_gi。
                #  - 本表直接接着上一张往下排（第一个号 > 1，如 6/11/15）：那就是全局号，
                #    直接用，不再偏移。
                if nums[0] == 1:
                    off = next_gi - 1
                    global_nums = [n + off for n in nums]
                else:
                    global_nums = list(nums)
                for gi, col in zip(global_nums, sl.sample_cols):
                    sample_map[gi] = (sl.name, col)
                    values[gi] = sl.col_values.get(col, [])
                next_gi = max(global_nums) + 1
        finally:
            wb.close()

        if not sheets:
            raise ExtractError("FILE_READ_ERROR", "报告中未找到任何工作表", file=path)
    finally:
        # 转换产生的临时 .xlsx 用完即删（绝不改动用户原文件）
        if conv_path and os.path.exists(conv_path):
            try:
                os.remove(conv_path)
            except Exception:  # noqa: BLE001
                pass

    return ReportInfo(path=path, parse=parse, sheets=sheets,
                      sample_map=sample_map, values=values)
