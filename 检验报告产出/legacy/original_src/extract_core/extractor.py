# -*- coding: utf-8 -*-
"""核心编排（单任务，事务回滚；支持多来源拼接 + K>5 多表）。

单任务 run_one_task_multi(segments)：
  把 segments（同一图号的多份来源，按顺序）拼接为本次 K 件 →
  从各来源抽取其「第 used+1 .. used+count 件」→ 生成一份独立小报告 →
  就地把各来源文件已用量后缀推进为 -(used+count)（用完则 -0）。

  典型场景：某 BOM 行需要 12 件，但一份基准报告只剩 5 件（其余已在上几套用掉），
  则先用满这 5 件，再从另一份基准报告取 7 件，合成一份 12 件的新报告。

事务保证：
  - 仅在所有来源改名成功「且」新报告成功生成后才保留改名；
  - 任一步失败（改名/COM/保存），恢复全部来源原名、不生成新文件。

所有写盘均走 Excel COM（win32com），以「原样复制」保留格式/合并单元格/形状；
布局探测由 report_reader 用 openpyxl 完成，COM 仅负责读取数值与写入。
"""
import os
import re
import shutil
import time

from .errors import ExtractError
from .filename_parser import parse_report_filename
from .bom_reader import canon
from .report_reader import detect_layout, _com_convert_to_xlsx


SHEET_MAX = 5  # 单张输出工作表最多承载的「测量件数」列数；K>5 时自动拆成多张表
# ceil(K / SHEET_MAX) 张表。例如 K=12 -> 3 张表（5+5+2），最后一张尾部列留空。

_DEBUG = bool(os.environ.get("EXTRACT_DEBUG"))


def _dbg(label, t0):
    """EXTRACT_DEBUG=1 时打印分段耗时（性能排查用）。"""
    if _DEBUG:
        print(f"[extract] {label}: {time.perf_counter() - t0:.2f}s", flush=True)
    return time.perf_counter()


# --------------------------------------------------------------------------- #
# Excel 实例管理（批次内复用，避免每任务重启 Excel）
# --------------------------------------------------------------------------- #
def create_excel_app():
    """创建并配置一个 Excel COM 实例（供整批任务复用）。

    调用方负责在批次结束时调用 quit_excel_app(xl)。
    """
    try:
        import pythoncom
        pythoncom.CoInitialize()
    except Exception:  # noqa: BLE001
        pass
    import win32com.client
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AskToUpdateLinks = False
    xl.EnableEvents = False
    try:
        xl.ScreenUpdating = False
    except Exception:  # noqa: BLE001
        pass
    try:
        xl.Calculation = -4135  # 手动重算，写大量单元格时显著提速
    except Exception:  # noqa: BLE001
        pass
    return xl


def _kill_pid(pid, wait=0.5):
    """Quit 后给 wait 秒自行退出，还在就强杀（只杀工具自己启动的那个 PID，
    不碰用户手动开的 Excel）。"""
    if not pid:
        return
    import subprocess
    import time as _t
    _t.sleep(wait)
    try:
        # GUI(无控制台)程序里 subprocess 会闪黑窗，CREATE_NO_WINDOW 抑制
        subprocess.run(["taskkill", "/PID", str(pid), "/T", "/F"],
                       capture_output=True, text=True,
                       creationflags=0x08000000)
    except Exception:  # noqa: BLE001
        pass


def quit_excel_app(xl):
    """退出并释放 Excel 实例（快速收尾：先关簿、再 Quit、释放 COM 引用）。"""
    if xl is None:
        return
    # 不再恢复 Calculation——进程随即退出，省一次 COM 往返。
    try:
        xl.DisplayAlerts = False
    except Exception:  # noqa: BLE001
        pass
    try:
        # 先关掉所有打开的工作簿（不保存），避免残留簿拖慢 Quit
        xl.Workbooks.Close()
    except Exception:  # noqa: BLE001
        pass
    try:
        xl.Quit()
    except Exception:  # noqa: BLE001
        pass
    # 释放 Python 侧 COM 引用并强制回收，Excel 进程才会真正退出（否则常挂残进程、
    # 下次启动慢、且 UI 要等很久按钮才恢复）。
    # 记录本实例 PID（此时 xl 仍存活，可拿主窗口句柄），用于 Quit 后退不干净时兜底强杀
    pid = None
    try:
        import win32process
        pid = win32process.GetWindowThreadProcessId(xl.hwnd)[1]
    except Exception:  # noqa: BLE001
        pid = None
    try:
        import gc
        del xl
        gc.collect()
    except Exception:  # noqa: BLE001
        pass
    # 兜底强杀本实例进程（Quit 退不干净时），避免僵尸 Excel 越堆越多、越用越慢
    if pid:
        try:
            _kill_pid(pid)
        except Exception:  # noqa: BLE001
            pass
    try:
        import pythoncom
        pythoncom.CoUninitialize()
    except Exception:  # noqa: BLE001
        pass


# --------------------------------------------------------------------------- #
# 工具函数
# --------------------------------------------------------------------------- #
def _norm_name_no_ext(path):
    base = os.path.basename(path)
    name, _ext = os.path.splitext(base)
    return re.sub(r"\s*\(\d+\)\s*$", "", name).strip()


def _unique(path):
    if not os.path.exists(path):
        return path
    d, name = os.path.split(path)
    base, ext = os.path.splitext(name)
    i = 2
    while os.path.exists(os.path.join(d, f"{base}({i}){ext}")):
        i += 1
    return os.path.join(d, f"{base}({i}){ext}")


def _sanitize_sheet_name(s):
    bad = [':', '\\', '/', '?', '*', '[', ']']
    for b in bad:
        s = s.replace(b, '_')
    return s[:31]


def _sheet_token_match(front, dwg):
    """front（图号_描述整体）与表 DWG 的宽松匹配。

    front 现在是整体（含描述），不能直接子串比较；规则：
      - DWG 归一化后出现在 front 中，或 front 出现在 DWG 中；或
      - front 按 _ 分段后，任一长度≥4 的段是 DWG 的子串
        （如段 93019282-dm 命中 DWG jh-ycl-93019282-dm(按图加工)）。
    """
    c1, c2 = canon(front), canon(dwg)
    if not c2:
        return False
    if c2 in c1 or c1 in c2:
        return True
    return any(len(tok) >= 4 and tok in c2 for tok in c1.split("_"))


def _choose_target_sheet(info):
    """选目标表：直接用第一张主表，不做图号二次校验。

    基准报告里每张表都是同一模板，只有序号/数据列不同；BOM 匹配这一关已确认
    这份文件就是目标零件，故直接复制第一张表当模板，表内印错的图号不影响、
    也绝不报 NO_SHEET_MATCH。数据按 sample_map 取、序号在生成时重排为 1..K。
    """
    return info.sheets[0], False


def _rename_source_path(source_path, used, suffix):
    """构造来源文件改名后的路径：后缀由 -used 改为 -suffix（用完 suffix='0'）。

    注意 used 为「当前任务的已用量」（调用方显式传入），不依赖 parse 缓存，
    以便同一文件连续多任务时后缀正确递推。
    """
    d = os.path.dirname(os.path.abspath(source_path))
    name_no_ext = _norm_name_no_ext(source_path)
    if used > 0:
        new_name = re.sub(rf"-{re.escape(str(used))}$", f"-{suffix}", name_no_ext)
    else:
        new_name = name_no_ext + f"-{suffix}"
    ext = os.path.splitext(source_path)[1]
    return os.path.join(d, new_name + ext)


# --------------------------------------------------------------------------- #
# COM 生成新报告（多来源拼接 + K>5 多表）
# --------------------------------------------------------------------------- #
def _range_of(ws, r1, c1, r2, c2):
    """返回 ws 上的矩形 Range 对象。"""
    return ws.Range(ws.Cells(r1, c1), ws.Cells(r2, c2))


def _read_col_values(src_wb, info, sbn, gi):
    """读取某件（全局件号 gi）的测量数据列（优先 info.values 缓存，否则 COM 整块取回）。"""
    cached = info.values.get(gi) if info.values else None
    if cached is not None:
        return list(cached)
    sname, col = info.sample_map[gi]
    sl = sbn[sname]
    try:
        block = _range_of(src_wb.Sheets(sname), sl.data_first_row, col,
                          sl.data_last_row, col).Value
    except Exception:  # noqa: BLE001
        block = None
    if block is None:
        return []
    if isinstance(block, tuple) and block and isinstance(block[0], tuple):
        return [row[0] for row in block]
    return [block]  # 单行区域返回标量


def _generate_report(segments, out_dir, xl=None):
    """按 segments（多份来源，按顺序拼成本次 K 件）生成新报告；K>5 自动拆多张表。

    segments: [{"path","info","used","count"}, ...]（同一图号；path 为「已改名后」的来源路径）
      - 第 1 段既作模板也作命名/日期基准；各段依次贡献其「第 used+1 .. used+count 件」，
        输出报告的测量序号 1..K 按段顺序连续拼接。
    表数量：n_out = max(1, ceil(K/SHEET_MAX))；每张最多 SHEET_MAX 列；序号 1..K 连续；
    每张表「总数/检查数」= K；最后一张未用到的列留空。

    任何 COM/保存异常抛 ExtractError。
    """
    import math
    seg0 = segments[0]
    info0 = seg0["info"]
    parse = info0.parse
    target, _switched = _choose_target_sheet(info0)
    K = sum(int(s["count"]) for s in segments)

    out_dir = out_dir or os.path.dirname(os.path.abspath(seg0["path"]))
    os.makedirs(out_dir, exist_ok=True)
    # parse.bh 即文件名 front（图号_描述整体），新报告 = front_K件_日期
    new_base = f"{parse.bh}_{K}件_{parse.date}.xlsx"
    final_path = _unique(os.path.join(out_dir, new_base))

    local_xl = xl is None
    _t = time.perf_counter()
    if local_xl:
        try:
            xl = create_excel_app()
        except Exception as e:  # noqa: BLE001
            raise ExtractError("COM_ERROR", f"Excel 启动失败：{e}", file=seg0["path"])
    _dbg("Excel启动", _t)

    n_out = max(1, int(math.ceil(K / float(SHEET_MAX))))
    opened = []       # [(seg, wb, sheet_by_name)]
    new_wb = None
    try:
        # 打开各来源（只读）
        _t = time.perf_counter()
        for seg in segments:
            try:
                wb = xl.Workbooks.Open(os.path.abspath(seg["path"]),
                                       ReadOnly=True, UpdateLinks=0)
            except Exception as e:  # noqa: BLE001
                raise ExtractError("COM_ERROR",
                                   f"Excel 打开报告失败（可能被占用）：{e}",
                                   file=seg["path"])
            opened.append((seg, wb, {s.name: s for s in seg["info"].sheets}))
        _dbg("打开源簿", _t)

        # 复制模板表为独立新簿；K>5 再多复制若干张（位置参数，具名 After= 会另开新簿）
        # 注意：Sheet.Copy 依赖系统剪贴板，剪贴板被其它程序占用时会抛
        # RPC_E_SERVERFAULT（-2147352567，Excel「发生意外」）——瞬时占用，重试即可。
        _t = time.perf_counter()
        first_wb = opened[0][1]

        def _copy_sheet_once(use_before=None):
            sh = first_wb.Sheets(target.name)
            if use_before is None:
                sh.Copy()
            else:
                # 单参数 Copy(Before=表)：复制到该表之前（同簿内）。
                # 不能用 Copy(None, After=表)：win32com 动态模式把 None 当作有效参数
                # 传给 Before，Excel 会为每次调用新建一个工作簿（实测多出 2 个残留簿，
                # 且残留簿处于坏状态，后续访问 Sheets 报 -2147352567「发生意外」）。
                sh.Copy(use_before)
            return xl.ActiveWorkbook if use_before is None else None

        def _copy_sheet_with_retry(use_before=None):
            _last = None
            for _attempt in range(3):
                try:
                    return _copy_sheet_once(use_before)
                except Exception as e:  # noqa: BLE001
                    _last = e
                    time.sleep(1.0 + 0.5 * _attempt)
            raise _last

        try:
            _copy_sheet_with_retry()
        except Exception as e:  # noqa: BLE001
            raise ExtractError("COM_ERROR", f"Excel 复制工作表失败：{e}", file=seg0["path"])
        new_wb = xl.ActiveWorkbook
        for _ in range(n_out - 1):
            try:
                _copy_sheet_with_retry(new_wb.Sheets(1))
            except Exception as e:  # noqa: BLE001
                raise ExtractError("COM_ERROR",
                                   f"Excel 复制工作表（多表）失败：{e}", file=seg0["path"])
        _dbg("复制工作表", _t)

        try:
            xl.Calculation = -4135  # 写入期间手动重算
        except Exception:  # noqa: BLE001
            pass

        # 组装：输出序号 1..K -> (源簿, info, 该源 sheet 索引, 全局件号 gi)
        _t = time.perf_counter()
        pieces = []
        max_last = target.data_last_row
        for (seg, wb, sbn) in opened:
            info = seg["info"]
            used = int(seg["used"])
            for t in range(int(seg["count"])):
                gi = used + t + 1
                sname, _c = info.sample_map[gi]
                max_last = max(max_last, sbn[sname].data_last_row)
                pieces.append((wb, info, sbn, gi))
        if len(pieces) != K:
            raise ExtractError("COM_ERROR", "内部错误：分片件数与 K 不一致", file=seg0["path"])
        values_cache = {}
        for (wb, info, sbn, gi) in pieces:
            ck = (id(info), gi)
            if ck not in values_cache:
                values_cache[ck] = _read_col_values(wb, info, sbn, gi)
        _dbg("取数", _t)

        # 逐张输出表：清空 -> 填数 -> 序号 -> 总数/检查数 -> 表名
        _t = time.perf_counter()
        sample_cols = target.sample_cols
        for s in range(n_out):
            ws = new_wb.Sheets(s + 1)
            n_here = min(SHEET_MAX, K - s * SHEET_MAX)  # 本表承载件数（最后一张可能不足）

            # (1) 清空测量列数据区（整块一次，保留格式/列宽）
            if sample_cols and max_last >= target.data_first_row:
                c1, c2 = sample_cols[0], sample_cols[-1]
                n_rows = max_last - target.data_first_row + 1
                blank = tuple(tuple(None for _ in range(c2 - c1 + 1))
                              for _ in range(n_rows))
                try:
                    _range_of(ws, target.data_first_row, c1, max_last, c2).Value = blank
                except Exception:  # noqa: BLE001
                    for c in sample_cols:
                        for r in range(target.data_first_row, max_last + 1):
                            try:
                                ws.Cells(r, c).Value = None
                            except Exception:  # noqa: BLE001
                                pass

            # (2) 序号行整行先清空（模板自带序号）
            if sample_cols:
                c1, c2 = sample_cols[0], sample_cols[-1]
                try:
                    _range_of(ws, target.seq_row, c1, target.seq_row, c2).Value = \
                        (tuple(None for _ in sample_cols),)
                except Exception:  # noqa: BLE001
                    for c in sample_cols:
                        try:
                            ws.Cells(target.seq_row, c).Value = None
                        except Exception:  # noqa: BLE001
                            pass

            # (3) 逐件填数：本地序号 serial=s*SHEET_MAX+j+1 -> pieces[serial-1]
            for j in range(n_here):
                serial = s * SHEET_MAX + j + 1
                wb, info, sbn, gi = pieces[serial - 1]
                local_col = sample_cols[j]
                sname, _c = info.sample_map[gi]
                sl = sbn[sname]
                vals = values_cache.get((id(info), gi), [])
                r1 = sl.data_first_row
                try:
                    _range_of(ws, r1, local_col,
                              r1 + max(len(vals) - 1, 0), local_col).Value = \
                        tuple((v,) for v in vals)
                except Exception:  # noqa: BLE001
                    for off, v in enumerate(vals):
                        try:
                            ws.Cells(r1 + off, local_col).Value = v
                        except Exception:  # noqa: BLE001
                            pass
                try:
                    # 序号名照抄源报告该列的原始标签（可溯源），不自行编号
                    src_col = info.sample_map[gi][1]
                    src_idx = sl.sample_cols.index(src_col)
                    ws.Cells(target.seq_row, local_col).Value = sl.seq_labels[src_idx]
                except Exception:  # noqa: BLE001
                    try:
                        ws.Cells(target.seq_row, local_col).Value = serial
                    except Exception:  # noqa: BLE001
                        pass

            # (4) 总数 / 检查数 = K（每张表都写 K）
            if target.total_cell:
                try:
                    ws.Cells(target.total_cell[0], target.total_cell[1]).Value = K
                except Exception:  # noqa: BLE001
                    pass
            if target.sample_cell:
                try:
                    ws.Cells(target.sample_cell[0], target.sample_cell[1]).Value = K
                except Exception:  # noqa: BLE001
                    pass

            # (5) 表名：单表=日期；多表=日期-1/日期-2/...
            try:
                ws.Name = _sanitize_sheet_name(
                    parse.date if n_out == 1 else f"{parse.date}-{s + 1}")
            except Exception:  # noqa: BLE001
                pass
        _dbg("清空+填数+序号", _t)

        try:
            xl.Calculation = -4105  # 恢复自动重算
        except Exception:  # noqa: BLE001
            pass

        _t = time.perf_counter()
        try:
            # 显式 FileFormat=51（xlOpenXMLWorkbook）强制写真 .xlsx：
            # 不指定时 Excel 按「默认保存格式」写盘，若默认是老 .xls(97-2003)，
            # 写 .xlsx 名会产出 OLE2 内容的假 xlsx（WPS 生成、Excel 报格式无效）。
            new_wb.SaveAs(os.path.abspath(final_path), FileFormat=51)
        except Exception as e:  # noqa: BLE001
            raise ExtractError("SAVE_ERROR",
                               f"保存新报告失败（可能被占用/无权限）：{e}",
                               file=final_path)
        _dbg("SaveAs", _t)
        # 自校验：确认落盘是真 OOXML(.xlsx)。若 COM 层仍产出老格式
        # （OLE2 头，如默认格式被 WPS/旧设置劫持），自动重转后再覆盖。
        try:
            with open(final_path, "rb") as _f:
                _magic = _f.read(8)
            if _magic != b"PK\x03\x04":
                _conv = _com_convert_to_xlsx(final_path)
                if _conv and os.path.exists(_conv):
                    shutil.copyfile(_conv, final_path)
                    os.remove(_conv)
        except Exception:  # noqa: BLE001
            pass
        # 保险：输出文件确保可写（Excel 新建文件本不带只读，但若目标路径/上级目录
        # 继承到只读位，显式去掉，保证客户拿到后可直接打开、可另存）
        try:
            os.chmod(final_path, 0o666)
            if hasattr(os.stat(final_path), "st_file_attributes"):
                import ctypes
                attrs = os.stat(final_path).st_file_attributes
                if attrs & 0x01:  # FILE_ATTRIBUTE_READONLY
                    ctypes.windll.kernel32.SetFileAttributesW(
                        final_path, attrs & ~0x01)
        except Exception:  # noqa: BLE001
            pass
    finally:
        try:
            if new_wb is not None:
                new_wb.Close(SaveChanges=False)
        except Exception:  # noqa: BLE001
            pass
        for (_seg, wb, _sbn) in opened:
            try:
                wb.Close(SaveChanges=False)
            except Exception:  # noqa: BLE001
                pass
        if local_xl:
            quit_excel_app(xl)

    return final_path


# --------------------------------------------------------------------------- #
# 单任务（多来源，事务回滚）
# --------------------------------------------------------------------------- #
def run_one_task_multi(segments, out_dir=None, on_progress=None, xl=None):
    """多来源单任务：把 segments 依次拼接为本次 K 件，生成一份新报告，并推进各来源已用量。

    segments: [{"path","info"|None,"used","count"}, ...]（同一图号，按顺序拼接为 K 件）
    事务：所有来源改名成功后才生成新报告；任一步失败 -> 全部来源恢复原名、不生成新文件。
    """
    def prog(p, msg):
        if on_progress:
            on_progress(p, msg)

    segments = list(segments)
    K = sum(int(s["count"]) for s in segments)
    if not segments or K <= 0:
        raise ExtractError("NO_TASK", "空任务（无来源或件数为 0）",
                           file=segments[0]["path"] if segments else "")

    prog(5, "探测布局")
    _t = time.perf_counter()
    for s in segments:
        if s.get("info") is None:
            s["info"] = detect_layout(s["path"])   # ParseResult + 布局
    parse0 = segments[0]["info"].parse
    _dbg("探测布局", _t)

    # ---- 校验 ----
    for s in segments:
        info = s["info"]
        parse = info.parse
        used = int(s["used"])
        cnt = int(s["count"])
        if used + cnt > parse.total:
            raise ExtractError(
                "K_EXCEEDS_REMAINING",
                f"图号 {parse.bh} 本来源抽取后需 {used + cnt} 件，但报告总件数仅 {parse.total} 件",
                file=s["path"],
            )
        for t in range(1, cnt + 1):
            gi = used + t
            if gi not in info.sample_map:
                max_rec = max(info.sample_map.keys()) if info.sample_map else 0
                raise ExtractError(
                    "SAMPLE_NOT_FOUND",
                    f"需抽取第 {gi} 件，但报告仅记录到第 {max_rec} 件（尚未录入）",
                    file=s["path"],
                )

    prog(15, "定位目标工作表")
    target0, switched0 = _choose_target_sheet(segments[0]["info"])

    # ---- 事务开始：推进各来源已用量（就地改名）----
    prog(25, "推进基准报告已用量（事务开始）")
    renamed = []   # [(seg, orig, new)]
    try:
        for s in segments:
            info = s["info"]
            parse = info.parse
            used = int(s["used"])
            cnt = int(s["count"])
            new_used = used + cnt
            suffix = "0" if new_used >= parse.total else str(new_used)
            new_source = _rename_source_path(s["path"], used, suffix)
            os.replace(s["path"], new_source)
            renamed.append((s, s["path"], new_source))
    except Exception as e:  # noqa: BLE001
        for (_s, orig, new) in renamed:
            try:
                os.replace(new, orig)
            except Exception:  # noqa: BLE001
                pass
        raise ExtractError("RENAME_ERROR",
                           f"更新基准报告文件名失败（可能被占用）：{e}",
                           file=segments[0]["path"])

    for (s, orig, new) in renamed:
        s["_orig"] = orig
        s["path_new"] = new

    # ---- 生成新报告（失败则回滚全部来源）----
    try:
        output = _generate_report(
            [{"path": s["path_new"], "info": s["info"],
              "used": int(s["used"]), "count": int(s["count"])} for s in segments],
            out_dir, xl=xl,
        )
    except ExtractError:
        for (_s, orig, new) in renamed:
            try:
                os.replace(new, orig)
            except Exception:  # noqa: BLE001
                pass
        raise
    except Exception as e:  # noqa: BLE001
        for (_s, orig, new) in renamed:
            try:
                os.replace(new, orig)
            except Exception:  # noqa: BLE001
                pass
        raise ExtractError("COM_ERROR", f"生成新报告异常：{e}", file=segments[0]["path"])

    prog(100, "完成")
    # 汇总结果
    items = []
    seg_out = []
    for s in segments:
        info = s["info"]
        parse = info.parse
        used = int(s["used"])
        cnt = int(s["count"])
        items += list(range(used + 1, used + cnt + 1))
        seg_out.append({
            "path": s.get("_orig", s["path"]),
            "source_after": s.get("path_new"),
            "bh": parse.bh,
            "used_from": used,
            "used_to": used + cnt,
            "count": cnt,
            "total": parse.total,
        })
    return {
        "ok": True,
        "bh": parse0.bh,
        "desc": parse0.desc,
        "qty": K,
        "used_from": int(segments[0]["used"]),
        "used_to": int(segments[0]["used"]) + int(segments[0]["count"]),
        "total": parse0.total,
        "date": parse0.date,
        "source_before": segments[0].get("_orig", segments[0]["path"]),
        "source_after": segments[0].get("path_new"),
        "output": output,
        "switched_sheet": (target0.name if switched0 else None),
        "items": items,
        "segments": seg_out,
        "multi_source": len(segments) > 1,
    }


def run_one_task(source_path, K, used, out_dir=None, on_progress=None,
                 info=None, xl=None):
    """兼容包装：单来源单任务（从该来源抽第 used+1..used+K 件）。"""
    return run_one_task_multi(
        [{"path": source_path, "info": info, "used": used, "count": K}],
        out_dir=out_dir, on_progress=on_progress, xl=xl,
    )
