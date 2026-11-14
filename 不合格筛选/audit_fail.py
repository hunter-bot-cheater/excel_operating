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
import traceback
import uuid

try:
    import win32com.client
except Exception:
    win32com = None

try:
    import winreg
except Exception:
    winreg = None

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

# Excel.Application 的固定 CLSID。WPS 抢的是 "Excel.Application" 这个 ProgID
# 的映射，一般不会去改 Microsoft 自己的 CLSID，所以按 CLSID 创建能绕开劫持。
EXCEL_CLSID = "{00024500-0000-0000-C000-000000000046}"


def _engine_pref():
    """引擎偏好。

    默认 wps：走 ProgID（"Excel.Application"），即客户机的 WPS 环境，与旧版行为一致。
    可选 excel：只认真 Microsoft Excel，找不到直接报错。
    可选 auto：优先真 Excel，找不到再退回 ProgID。
    客户机确认没有装 Excel，所以默认必须保持在 WPS 侧，不能自动切引擎。
    """
    _v = os.environ.get("AUDITFAIL_ENGINE", "")
    _v = str(_v).strip().lower()
    if _v == "excel":
        return "excel"
    if _v == "auto":
        return "auto"
    return "wps"


def _is_wps_app(xl):
    """判断拿到的是不是 WPS 冒充的 Excel。

    只看安装路径。Name 不可靠：WPS 表格会把自己报成 "Microsoft Excel"，
    Version 也会报 12.0。Path 指向 office6 才是真相。
    """
    _p = ""
    try:
        _p = str(xl.Path).lower()
    except Exception:  # noqa: BLE001
        _p = ""
    if ("wps" in _p) or ("kingsoft" in _p):
        return True
    _n = ""
    try:
        _n = str(xl.Name).lower()
    except Exception:  # noqa: BLE001
        _n = ""
    if ("wps" in _n) or ("kingsoft" in _n):
        return True
    return False


def _split_server_path(raw):
    """把注册表里的 '"C:\\...\\EXCEL.EXE" /automation' 拆成纯 exe 路径。"""
    if not raw:
        return None
    _s = str(raw).strip()
    if not _s:
        return None
    if _s.startswith('"'):
        _end = _s.find('"', 1)
        if _end > 0:
            return _s[1:_end]
        return None
    _i = _s.lower().find(".exe")
    if _i > 0:
        return _s[:_i + 4].strip()
    return _s


def _find_real_excel_exe():
    """从注册表找 Microsoft Excel 的 EXCEL.EXE；找不到或只找到 WPS 的返回 None。"""
    if winreg is None:
        return None
    _cands = []
    for _sub in ("CLSID\\%s\\LocalServer32" % EXCEL_CLSID,
                 "Wow6432Node\\CLSID\\%s\\LocalServer32" % EXCEL_CLSID):
        try:
            _k = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, _sub)
            try:
                _cands.append(winreg.QueryValue(_k, None))
            finally:
                _k.Close()
        except Exception:  # noqa: BLE001
            pass
    for _hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        try:
            _k = winreg.OpenKey(
                _hive,
                "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\excel.exe")
            try:
                _cands.append(winreg.QueryValue(_k, None))
            finally:
                _k.Close()
        except Exception:  # noqa: BLE001
            pass
    for _raw in _cands:
        _p = _split_server_path(_raw)
        if not _p:
            continue
        _low = _p.lower()
        if "excel.exe" not in _low:
            continue
        if ("wps" in _low) or ("kingsoft" in _low):
            continue
        if os.path.exists(_p):
            return _p
    return None


def _spawn_excel_and_attach(exe_path, wait=15.0):
    """用 EXCEL.EXE /automation 起一个真 Excel，再从 ROT 取回对象。

    只认 PID 与我们刚启动进程一致的那个实例。拿到别人的实例一律不碰：
    用户可能开着未保存的文档，Quit 会毁数据。
    """
    try:
        import win32process
    except Exception:  # noqa: BLE001
        return None
    try:
        _proc = subprocess.Popen([exe_path, "/automation"], close_fds=True,
                                 creationflags=0x08000000)  # CREATE_NO_WINDOW
    except Exception as _e:  # noqa: BLE001
        _dbg("启动 EXCEL.EXE 失败: %r" % (_e,))
        return None
    _deadline = time.time() + wait
    while time.time() < _deadline:
        try:
            _xl = win32com.client.GetActiveObject("Excel.Application")
        except Exception:  # noqa: BLE001
            time.sleep(0.4)
            continue
        _pid = None
        try:
            _, _pid = win32process.GetWindowThreadProcessId(_xl.Hwnd)
        except Exception:  # noqa: BLE001
            _pid = None
        if _pid == _proc.pid:
            _dbg("EXCEL.EXE 就绪 pid=%s" % (_pid,))
            return _xl
        time.sleep(0.4)
    _dbg("EXCEL.EXE 启动后 %.0fs 内未拿到自己的实例，放弃该路径" % (wait,))
    try:
        _proc.terminate()
    except Exception:  # noqa: BLE001
        pass
    return None


def _make_xl():
    """新建一个隐藏的 Excel COM 实例。

    默认走 "Excel.Application" 这个 ProgID，在客户机上就是 WPS 表格
    （WPS 装好后会把这个 ProgID 抢注到自己身上），与其 Ctrl 逻辑保持一致。

    另外留了两条可选路径，给装了真 Office 的机器用，由 AUDITFAIL_ENGINE 控制：
      auto  ：先按 Microsoft 的固定 CLSID 创建（绕开 WPS 对 ProgID 的劫持），
              取到的是 WPS 就弃用，再拉注册表里的 EXCEL.EXE 接管，全失败退回 ProgID
      excel ：只认真 Excel，找不到直接报错，不静默退回 WPS
    """
    if win32com is None:
        raise RuntimeError("缺少 win32com 模块，无法读取 Excel")
    try:  # COM 必须在线程内初始化（GUI worker / API worker 均需）
        import pythoncom
        pythoncom.CoInitialize()
        _dbg("CoInitialize OK")
    except Exception as _e:  # noqa: BLE001
        _dbg("CoInitialize already init or failed: %r" % (_e,))

    _pref = _engine_pref()
    xl = None
    _used = ""

    if _pref == "auto" or _pref == "excel":
        # 1) 按 CLSID 直连：绕开 WPS 对 ProgID 的劫持
        try:
            _cand = win32com.client.DispatchEx(EXCEL_CLSID)
            if _is_wps_app(_cand):
                _dbg("CLSID 实例仍是 WPS（Path 指向 office6），弃用")
                try:
                    _cand.Quit()
                except Exception:  # noqa: BLE001
                    pass
            else:
                xl = _cand
                _used = "excel-clsid"
                _dbg("按 CLSID 取到 Microsoft Excel")
        except Exception as _e:  # noqa: BLE001
            _dbg("按 CLSID 创建失败: %r" % (_e,))
        # 2) CLSID 也失效时，直接拉 EXCEL.EXE 再接管
        if xl is None:
            _exe = _find_real_excel_exe()
            if _exe:
                _dbg("注册表找到 EXCEL.EXE: %s" % (_exe,))
                xl = _spawn_excel_and_attach(_exe)
                if xl is not None:
                    _used = "excel-exe"
            else:
                _dbg("注册表里没有非 WPS 的 EXCEL.EXE")
        if xl is None and _pref == "excel":
            raise RuntimeError("配置要求只用 Microsoft Excel，但本机没找到可用的 EXCEL.EXE")
    elif _pref == "wps":
        _dbg("默认走 ProgID（WPS 环境）")

    if xl is None:
        # 3) 旧行为：走 ProgID，可能落到 WPS
        try:
            xl = win32com.client.DispatchEx("Excel.Application")
            _dbg("DispatchEx Excel.Application OK")
            _used = "progid"
        except Exception as _e:  # noqa: BLE001
            _dbg("DispatchEx failed, fallback to Dispatch: %r" % (_e,))
            xl = win32com.client.Dispatch("Excel.Application")
            _dbg("Dispatch fallback OK")
            _used = "progid-dispatch"

    try:
        _dbg("引擎=%s Path=%r Version=%s Name=%s"
             % (_used, xl.Path, xl.Version, xl.Name))
    except Exception as _e:  # noqa: BLE001
        _dbg("read Excel.Version failed: %r" % (_e,))
    for _attr, _val in (("Visible", False), ("DisplayAlerts", False),
                        ("ScreenUpdating", False), ("EnableEvents", False),
                        ("AskToUpdateLinks", False)):
        try:
            setattr(xl, _attr, _val)
            _dbg("set %s=%s OK" % (_attr, _val))
        except Exception as e:  # noqa: BLE001
            # 属性可能被 WPS/旧版 Excel 拒绝；不影响读取，但留痕便于排查"为什么弹窗/慢"
            _dbg("_make_xl set %s=%r failed: %r" % (_attr, _val, e))
    # 宏安全：禁用宏自动运行，避免 WPS 弹"启用宏"对话框（不影响数据读取）
    try:
        xl.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
        _dbg("set AutomationSecurity=3 OK")
    except Exception as e:  # noqa: BLE001
        _dbg("_make_xl set AutomationSecurity failed: %r" % (e,))
    # 进程一建好就把环境拍照（仅在日志开启时真正执行）
    try:
        _diag_header()
        _diag_excel(xl)
    except Exception as _e:  # noqa: BLE001
        _dbg("[DIAG] excel env snapshot failed: %r" % (_e,))
    return xl


def _recalc(xl):
    """强制全表重算（保证公式最新值）。

    优先 CalculateFull：Calculate 只重算依赖树里被标记"脏"的单元格，
    程序生成/从未在本机打开计算过的公式文件可能没有脏标记，Calculate
    会跳过 → 结果列公式读到空/旧值 → B4 漏判。CalculateFull 无条件重算全部。
    """
    try:
        xl.CalculateFull()
    except Exception:
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


_DBG_FILE = {"path": None}
# 日志默认完全关闭：客户机上不生成任何文件。
# 排查时才显式开：AUDITFAIL_DEBUG=1 用默认位置，AUDITFAIL_DEBUG=<完整路径> 写指定文件。
_DBG_RAW = (os.environ.get("AUDITFAIL_DEBUG") or "").strip()
_DBG_ON = bool(_DBG_RAW)


def _dbg_candidates():
    """开启日志时的候选落盘位置：显式路径 > 桌面 > 桌面(中文名) > 用户目录 > 临时目录。

    兜到临时目录是为了保证「桌面不可写 / 桌面被重定向」时日志仍然落得下来。
    """
    _out = []
    if _DBG_RAW and _DBG_RAW not in ("1", "on", "ON", "true", "TRUE"):
        _out.append(_DBG_RAW)
    try:
        _home = os.path.expanduser("~")
    except Exception:
        _home = None
    if _home:
        _out.append(os.path.join(_home, "Desktop", "audit_debug.log"))
        _out.append(os.path.join(_home, "桌面", "audit_debug.log"))
        _out.append(os.path.join(_home, "audit_debug.log"))
    _tmp = os.environ.get("TEMP")
    if not _tmp:
        try:
            _tmp = tempfile.gettempdir()
        except Exception:
            _tmp = None
    if _tmp:
        _out.append(os.path.join(_tmp, "audit_debug.log"))
    return _out


def _dbg_path():
    """当前实际使用的日志路径（第一次调用时按候选顺序落定）。"""
    if _DBG_FILE["path"]:
        return _DBG_FILE["path"]
    for _p in _dbg_candidates():
        try:
            _d = os.path.dirname(_p)
            if _d and os.path.isdir(_d):
                _DBG_FILE["path"] = _p
                return _p
        except Exception:
            continue
    return None


def _dbg(msg):
    """诊断日志。默认关闭，客户机上不产生任何文件。

    排查时才开：设 AUDITFAIL_DEBUG=1 写到默认位置（桌面，失败退临时目录），
    或 AUDITFAIL_DEBUG=<完整路径> 写到指定文件。
    """
    if not _DBG_ON:
        return
    if msg.startswith("[DIAG]") or msg.startswith("=== DIAG"):
        if _DIAG_FORCED == "0":
            return
    import datetime as _dt
    _paths = []
    _cur = _dbg_path()
    if _cur:
        _paths.append(_cur)
    for _p in _dbg_candidates():
        if _p not in _paths:
            _paths.append(_p)
    try:
        _ts = _dt.datetime.now().strftime("%H:%M:%S.%f")[:-3]
        _line = "[%s][pid=%d] %s\n" % (_ts, os.getpid(), msg)
    except Exception:
        return
    for _p in _paths:
        try:
            with open(_p, "a", encoding="utf-8") as _f:
                _f.write(_line)
            _DBG_FILE["path"] = _p
            return
        except Exception:
            continue


# ============================================================
# 现场诊断（客户机 SaveAs 报 0x80010105 时定位用）
#
# 只写日志，不改任何业务行为；所有探针都吞异常，并用 ERR:... 把失败本身
# 记下来（探针全失败本身就是"COM 通道已断"的证据）。
# 用 grep [DIAG] 就能把现场信息单独拎出来。
# ============================================================
_DIAG_ONCE = {"header": False, "events": False}
# 诊断探针（COM 属性探测、事件日志读取等）只在日志开启时做。
# AUDITFAIL_DIAG=0 可以在开了日志的前提下再单独关掉探针，只留基础流程日志。
_DIAG_FORCED = (os.environ.get("AUDITFAIL_DIAG") or "").strip()


def _diag_off():
    """诊断探针是否跳过：日志没开就一律跳过，显式 AUDITFAIL_DIAG=0 也跳过。"""
    if not _DBG_ON:
        return True
    return (_DIAG_FORCED == "0")


def _safe(fn, label=""):
    """跑一个 COM 探针，成功返回取值，失败返回 'ERR:...' 字符串，不抛。"""
    try:
        return fn()
    except Exception as _e:  # noqa: BLE001
        return "ERR[%s]:%r" % (label, _e)


def _safe_size(path):
    """取文件大小，不存在或取不到返回 None / 'ERR'。"""
    try:
        if not os.path.exists(path):
            return None
        return os.path.getsize(path)
    except Exception:
        return "ERR"


def _diag_hr(exc):
    """从 com_error 里取 (hresult, excepinfo 里的 scode)。

    外层 hresult 常常是 0x80020009（DISP_E_EXCEPTION），真正的 Excel 内部
    错误码在 excepinfo 第 6 个元素里，两个都要看才不会漏判。
    """
    _hr = None
    _sc = None
    try:
        _hr = exc.hresult
    except Exception:
        pass
    try:
        _info = exc.args[2]
        if _info is not None:
            _sc = _info[5]
    except Exception:
        pass
    return _hr, _sc


def _dbg_app_version():
    """读 VERSION（exe 同目录 > 打包内 > 源码目录），返回 (版本, 来源路径)。"""
    _cands = []
    try:
        _cands.append(os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), "VERSION"))
    except Exception:
        pass
    _meipass = getattr(sys, "_MEIPASS", None)
    if _meipass:
        _cands.append(os.path.join(_meipass, "VERSION"))
    try:
        _cands.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), "VERSION"))
    except Exception:
        pass
    for _p in _cands:
        try:
            with open(_p, "r", encoding="utf-8-sig") as _f:
                _v = _f.read().strip()
            if _v:
                return _v, _p
        except Exception:
            continue
    return "?", None


def _diag_header():
    """本次运行的环境头，每个进程只写一次。"""
    if _DIAG_ONCE["header"]:
        return
    _DIAG_ONCE["header"] = True
    _dbg("=== DIAG SESSION ===")
    _dbg("[DIAG] frozen=%r argv=%r" % (getattr(sys, "frozen", False), sys.argv))
    try:
        _v, _vp = _dbg_app_version()
        _dbg("[DIAG] app VERSION=%s from=%s" % (_v, _vp))
    except Exception as _e:  # noqa: BLE001
        _dbg("[DIAG] version probe failed: %r" % (_e,))
    try:
        _exe = sys.executable
        _mt = None
        try:
            _mt = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(os.path.getmtime(_exe)))
        except Exception:
            pass
        _dbg("[DIAG] exe=%s size=%r mtime=%s" % (_exe, _safe_size(_exe), _mt))
    except Exception as _e:  # noqa: BLE001
        _dbg("[DIAG] exe probe failed: %r" % (_e,))
    try:
        import platform
        _dbg("[DIAG] python=%s arch=%s user=%s cwd=%s" % (
            platform.python_version(), platform.machine(),
            os.environ.get("USERNAME"), os.getcwd()))
        _dbg("[DIAG] os=%s" % (platform.platform(),))
    except Exception as _e:  # noqa: BLE001
        _dbg("[DIAG] platform probe failed: %r" % (_e,))
    _dbg("[DIAG] log=%s" % (_dbg_path(),))
    try:
        _dbg("[DIAG] temp=%s" % (os.environ.get("TEMP"),))
        _dbg("[DIAG] sessionname=%s" % (os.environ.get("SESSIONNAME"),))
    except Exception:
        pass


def _diag_pid_of_hwnd(hwnd):
    """由窗口句柄取进程 PID，取不到返回 None。"""
    if not isinstance(hwnd, int) or hwnd == 0:
        return None
    try:
        import ctypes
        from ctypes import wintypes
        _pid = wintypes.DWORD(0)
        ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(_pid))
        return _pid.value
    except Exception:
        return None


def _diag_windows_of_pid(pid, tag=""):
    """列出该进程的可见顶层窗口（ctypes 实现，不依赖 pywin32 子模块）。

    弹窗卡住时这里能看到类名 #32770 的对话框和它的标题；进程已崩则一个窗口都
    没有。这一条正好用来区分「被模态框挡住」和「进程真的挂了」。
    """
    if not pid:
        _dbg("[DIAG] windows dump skipped (pid=%r) %s" % (pid, tag))
        return
    try:
        import ctypes
        from ctypes import wintypes
    except Exception as _e:  # noqa: BLE001
        _dbg("[DIAG] ctypes unavailable: %r" % (_e,))
        return
    _u = ctypes.windll.user32
    _rows = []

    def _cb(hwnd, _lp):
        try:
            _pid = wintypes.DWORD(0)
            _u.GetWindowThreadProcessId(hwnd, ctypes.byref(_pid))
            if _pid.value != pid:
                return True
            if not _u.IsWindowVisible(hwnd):
                return True
            _t = ctypes.create_unicode_buffer(512)
            _u.GetWindowTextW(hwnd, _t, 512)
            _c = ctypes.create_unicode_buffer(256)
            _u.GetClassNameW(hwnd, _c, 256)
            _rows.append((hwnd, _c.value, _t.value))
        except Exception:
            pass
        return True

    try:
        _proc = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)(_cb)
        _u.EnumWindows(_proc, 0)
    except Exception as _e:  # noqa: BLE001
        _dbg("[DIAG] EnumWindows failed: %r" % (_e,))
        return
    _dbg("[DIAG] excel pid=%r visible windows=%d %s" % (pid, len(_rows), tag))
    for _h, _c, _t in _rows[:10]:
        _dbg("[DIAG] window hwnd=%r class=%r title=%r" % (_h, _c, _t))


_DIAG_SAVE_TIMEOUT = 8.0


def _diag_watchdog_fire(pid, n):
    """SaveAs 迟迟不返回时的采样点。

    被模态框挡住时 COM 调用是一声不响地卡着，不会抛异常，日志就停在
    「--- SaveAs attempt N ---」后面什么都没有。这个看门狗在超时后从独立
    线程用纯 ctypes 抓窗口，不碰 COM，不会被卡住的调用连累。
    """
    _dbg("[DIAG] SaveAs attempt %d 超过 %.0f 秒未返回，疑似被弹窗阻塞，抓窗口" % (
        n, _DIAG_SAVE_TIMEOUT))
    _diag_windows_of_pid(pid, "while SaveAs hanging")


def _diag_dir(out_dir):
    """输出目录侧：路径长度、UNC/OneDrive、磁盘空间、写探针。"""
    if _diag_off():
        return
    try:
        out_dir = os.path.abspath(out_dir)
    except Exception as _e:  # noqa: BLE001
        _dbg("[DIAG] out_dir abspath failed: %r raw=%r" % (_e, out_dir))
        return
    _dbg("[DIAG] out_dir=%s len=%d" % (out_dir, len(out_dir)))
    _dbg("[DIAG] out_dir exists=%s isdir=%s unc=%s" % (
        os.path.exists(out_dir), os.path.isdir(out_dir),
        out_dir.startswith("\\\\")))
    _low = out_dir.lower()
    if "onedrive" in _low:
        _dbg("[DIAG] out_dir WARNING: 在 OneDrive 同步目录下，同步锁可能挡住 SaveAs")
    try:
        _du = shutil.disk_usage(out_dir)
        _dbg("[DIAG] disk free=%.2fGB total=%.2fGB" % (
            _du.free / 1073741824.0, _du.total / 1073741824.0))
    except Exception as _e:  # noqa: BLE001
        _dbg("[DIAG] disk_usage failed: %r" % (_e,))
    _probe = os.path.join(out_dir, ".audit_probe_%d.tmp" % (os.getpid(),))
    try:
        with open(_probe, "w", encoding="utf-8") as _f:
            _f.write("probe")
        os.remove(_probe)
        _dbg("[DIAG] write probe OK")
    except Exception as _e:  # noqa: BLE001
        _dbg("[DIAG] write probe FAILED: %r" % (_e,))
        try:
            if os.path.exists(_probe):
                os.remove(_probe)
        except Exception:
            pass
    try:
        _locks = []
        for _n in os.listdir(out_dir):
            if _n.startswith("~$"):
                _locks.append(_n)
            if len(_locks) >= 5:
                break
        _dbg("[DIAG] out_dir ~$lock files=%r" % (_locks,))
    except Exception as _e:  # noqa: BLE001
        _dbg("[DIAG] out_dir listdir failed: %r" % (_e,))


def _diag_excel_env(xl):
    """Excel 进程级：版本、打印机、加载项。每进程只要记一次。"""
    _dbg("[DIAG] excel Version=%r Build=%r OS=%r" % (
        _safe(lambda: xl.Version, "Version"),
        _safe(lambda: xl.Build, "Build"),
        _safe(lambda: xl.OperatingSystem, "OS")))
    _dbg("[DIAG] excel Visible=%r DisplayAlerts=%r Interactive=%r AskToUpdateLinks=%r" % (
        _safe(lambda: xl.Visible, "Visible"),
        _safe(lambda: xl.DisplayAlerts, "DisplayAlerts"),
        _safe(lambda: xl.Interactive, "Interactive"),
        _safe(lambda: xl.AskToUpdateLinks, "AskToUpdateLinks")))
    _dbg("[DIAG] excel ActivePrinter=%r DefaultFilePath=%r" % (
        _safe(lambda: xl.ActivePrinter, "ActivePrinter"),
        _safe(lambda: xl.DefaultFilePath, "DefaultFilePath")))
    _dbg("[DIAG] excel Path=%r StartupPath=%r" % (
        _safe(lambda: xl.Path, "Path"),
        _safe(lambda: xl.StartupPath, "StartupPath")))
    _dbg("[DIAG] excel Ready=%r Hwnd=%r" % (
        _safe(lambda: xl.Ready, "Ready"),
        _safe(lambda: xl.Hwnd, "Hwnd")))
    _dbg("[DIAG] excel Workbooks.Count=%r Calculation=%r AutomationSecurity=%r" % (
        _safe(lambda: xl.Workbooks.Count, "Workbooks"),
        _safe(lambda: xl.Calculation, "Calculation"),
        _safe(lambda: xl.AutomationSecurity, "AutomationSecurity")))
    try:
        _n = xl.AddIns.Count
        _dbg("[DIAG] addins count=%r" % (_n,))
        _lim = _n
        if _lim > 20:
            _lim = 20
        for _i in range(1, _lim + 1):
            try:
                _ai = xl.AddIns.Item(_i)
                _dbg("[DIAG] addin[%d] name=%r installed=%r" % (
                    _i, _ai.Name, _ai.Installed))
            except Exception as _e:  # noqa: BLE001
                _dbg("[DIAG] addin[%d] read failed: %r" % (_i, _e))
    except Exception as _e:  # noqa: BLE001
        _dbg("[DIAG] addins probe failed: %r" % (_e,))


def _diag_excel(xl, wb=None, env=True):
    """Excel 进程侧 + 工作簿特性。

    env=True 才打进程级信息（版本/打印机/加载项）。进程级每进程记一次就够，
    逐文件重复打只会把日志撑大。
    """
    if _diag_off():
        return
    if env:
        _diag_excel_env(xl)
    if wb is None:
        return
    _dbg("[DIAG] wb Name=%r" % (_safe(lambda: wb.Name, "Name"),))
    _dbg("[DIAG] wb FullName=%r" % (_safe(lambda: wb.FullName, "FullName"),))
    _dbg("[DIAG] wb ReadOnly=%r Saved=%r FileFormat=%r CheckCompatibility=%r" % (
        _safe(lambda: wb.ReadOnly, "ReadOnly"),
        _safe(lambda: wb.Saved, "Saved"),
        _safe(lambda: wb.FileFormat, "FileFormat"),
        _safe(lambda: wb.CheckCompatibility, "CheckCompatibility")))
    _dbg("[DIAG] wb Sheets=%r Styles=%r ProtectStructure=%r HasPassword=%r" % (
        _safe(lambda: wb.Sheets.Count, "Sheets"),
        _safe(lambda: wb.Styles.Count, "Styles"),
        _safe(lambda: wb.ProtectStructure, "ProtectStructure"),
        _safe(lambda: wb.HasPassword, "HasPassword")))
    try:
        _n = wb.Sheets.Count
        _lim = _n
        if _lim > 5:
            _lim = 5
        for _i in range(1, _lim + 1):
            _sh = wb.Sheets(_i)
            _dbg("[DIAG] sheet[%d] name=%r visible=%r protect=%r used=%r merged=%r" % (
                _i,
                _safe(lambda: _sh.Name, "Name"),
                _safe(lambda: _sh.Visible, "Visible"),
                _safe(lambda: _sh.ProtectContents, "ProtectContents"),
                _safe(lambda: _sh.UsedRange.Address, "UsedRange"),
                _safe(lambda: _sh.UsedRange.MergeCells, "MergeCells")))
    except Exception as _e:  # noqa: BLE001
        _dbg("[DIAG] sheet scan failed: %r" % (_e,))


def _diag_events_wevtutil():
    """用系统自带 wevtutil 抓最近的应用错误/挂起事件（不依赖 pywin32 子模块）。

    返回 True 表示命令本身跑通了（有没有命中另说）。
    """
    _q = ("*[System[Provider[@Name='Application Error' or @Name='Application Hang' "
          "or @Name='Windows Error Reporting']]]")
    _cmd = ["wevtutil", "qe", "Application", "/c:25", "/rd:true", "/f:text",
            "/q:" + _q]
    try:
        _r = subprocess.run(_cmd, stdin=subprocess.DEVNULL,
                            capture_output=True, timeout=20,
                            creationflags=0x08000000)  # CREATE_NO_WINDOW
    except Exception as _e:  # noqa: BLE001
        _dbg("[DIAG] wevtutil run failed: %r" % (_e,))
        return False
    _raw = _r.stdout or b""
    _txt = ""
    try:
        if b"\x00" in _raw[:200]:
            _txt = _raw.decode("utf-16-le", "replace")   # wevtutil 在部分系统输出 UTF-16
        else:
            _txt = _raw.decode("mbcs", "replace")
    except Exception:
        _txt = _raw.decode("utf-8", "replace")
    _buf = []
    for _line in _txt.splitlines():
        _s = _line.strip()
        if not _s:
            continue
        if "EXCEL" in _s.upper():
            _buf.append(_s)
    _dbg("[DIAG] wevtutil rc=%r bytes=%d excelLines=%d" % (
        _r.returncode, len(_raw), len(_buf)))
    for _s in _buf[:12]:
        _dbg("[DIAG] eventlog %s" % (_s[:300],))
    return True


def _diag_events_win32():
    """pywin32 读 Application 日志（wevtutil 失败时的退路）。"""
    try:
        import win32evtlog
    except Exception as _e:  # noqa: BLE001
        _dbg("[DIAG] win32evtlog unavailable: %r" % (_e,))
        return
    _h = None
    try:
        _h = win32evtlog.OpenEventLog(None, "Application")
        _flags = win32evtlog.EVENTLOG_BACKWARDS_READ | win32evtlog.EVENTLOG_SEQUENTIAL_READ
        _read = 0
        _hit = 0
        while _read < 400 and _hit < 5:
            _batch = win32evtlog.ReadEventLog(_h, _flags, 0)
            if not _batch:
                break
            for _ev in _batch:
                _read += 1
                _src = _ev.SourceName or ""
                if "Application Error" not in _src and "Application Hang" not in _src \
                        and "Windows Error Reporting" not in _src:
                    continue
                _data = ""
                try:
                    _data = " ".join([str(_x) for _x in (_ev.StringInserts or [])])
                except Exception:
                    pass
                _blob = (_src + " " + _data).upper()
                if "EXCEL" not in _blob:
                    continue
                _hit += 1
                _dbg("[DIAG] eventlog %s src=%s data=%s" % (
                    _ev.TimeGenerated, _src, _data[:400]))
            if _read >= 400:
                break
        _dbg("[DIAG] win32evtlog scanned=%d excelHits=%d" % (_read, _hit))
    except Exception as _e:  # noqa: BLE001
        _dbg("[DIAG] eventlog read failed: %r" % (_e,))
    finally:
        if _h is not None:
            try:
                win32evtlog.CloseEventLog(_h)
            except Exception:
                pass


def _diag_events():
    """取证 Excel 崩溃记录：先用系统自带 wevtutil，失败再走 win32evtlog。

    优先 wevtutil 是故意的：win32evtlog 是 pywin32 的 win32 子模块，
    打进 onefile 包不一定被 PyInstaller 收集到，wevtutil 是系统组件不会缺。
    """
    if _DIAG_ONCE["events"]:
        return
    _DIAG_ONCE["events"] = True
    if _diag_events_wevtutil():
        return
    _diag_events_win32()


def _save_with_fallback(wb, target, fmt, is_wps=False, same_fmt=False):
    """保存：同格式优先走 SaveCopyAs，WPS 走降级链。

    真 Excel 上 SaveAs(FileFormat) 实测 31ms，SaveCopyAs 只要 10ms —— 后者
    不重建格式定义、不切换工作簿的 Path/Saved 状态，只按当前内存状态写一份
    文件。已逐格验证两者产出的值、字体、填充、数字格式、合并区、公式缓存值
    （data_only 读）完全一致；zip 内唯一差异是 workbook.xml 里的 absPath 记录
    与随机 documentId，无业务含义。

    只在「打开的文件与输出同扩展名」时才可用：SaveCopyAs 不能指定 FileFormat，
    跨格式（如 .xlsx 中间件要存成 .xls）必须走 SaveAs，否则输出格式会错。

    WPS 上不用 SaveCopyAs（其兼容层实现未知），继续走原来的降级链：
      1. SaveAs(FileFormat=fmt)：标准写法
      2. SaveAs(无 FileFormat)：让宿主按扩展名自己定格式
      3. SaveAs(FileFormat=fmt, Local=True)：走本地化格式入口
      4. SaveCopyAs：不改变宿主文档状态，另存一份副本
    全部失败时抛第一级的原始异常，保持调用方看到的错误码不变。
    """
    _abs = os.path.abspath(target)
    if same_fmt and not is_wps:
        try:
            wb.SaveCopyAs(_abs)
            # 副本必须真的落盘且非空，否则宁可退回 SaveAs，不能让调用方拿到空壳文件
            if os.path.exists(_abs) and os.path.getsize(_abs) > 0:
                return "copy-first"
            _dbg("SaveCopyAs 产出异常（不存在或为空），回退 SaveAs")
            try:
                os.remove(_abs)
            except Exception:  # noqa: BLE001
                pass
        except Exception as _ec:  # noqa: BLE001
            _dbg("SaveCopyAs 首选失败，回退 SaveAs: %r" % (_ec,))
    try:
        wb.SaveAs(_abs, FileFormat=fmt)
        return "fmt"
    except Exception as _e0:  # noqa: BLE001
        if not is_wps:
            raise
        _first = _e0
    _dbg("[DIAG] SaveAs 第一级失败，进入降级链（WPS）")
    try:
        wb.SaveAs(_abs)
        _dbg("[DIAG] SaveAs 第二级（无 FileFormat）成功")
        return "nofmt"
    except Exception as _e:  # noqa: BLE001
        _dbg("[DIAG] SaveAs 第二级失败: %r" % (_e,))
    try:
        wb.SaveAs(_abs, FileFormat=fmt, Local=True)
        _dbg("[DIAG] SaveAs 第三级（Local=True）成功")
        return "local"
    except Exception as _e:  # noqa: BLE001
        _dbg("[DIAG] SaveAs 第三级失败: %r" % (_e,))
    try:
        wb.SaveCopyAs(_abs)
        _dbg("[DIAG] SaveAs 第四级（SaveCopyAs）成功")
        return "copy"
    except Exception as _e:  # noqa: BLE001
        _dbg("[DIAG] SaveAs 第四级失败: %r" % (_e,))
    raise _first


def _diag_save_as(wb, xl, target, fmt, n, first, is_wps=False, same_fmt=False):
    """逐次记录 SaveAs 的入参、耗时、异常、残留与 Excel 存活状态。"""
    if _diag_off():
        # 诊断关闭时退化成裸 SaveAs（带 WPS 降级链）：不探测、不 stat、不起看门狗
        _save_with_fallback(wb, target, fmt, is_wps, same_fmt)
        return
    _dbg("[DIAG] --- SaveAs attempt %d ---" % (n,))
    _dbg("[DIAG] SaveAs target=%s len=%d exists=%s size=%r" % (
        target, len(target), os.path.exists(target), _safe_size(target)))
    _dbg("[DIAG] SaveAs FileFormat=%r" % (fmt,))
    if first:
        _diag_dir(os.path.dirname(target))
        _dbg("[DIAG] wb.Saved(before)=%r wb.CheckCompatibility=%r" % (
            _safe(lambda: wb.Saved, "Saved"),
            _safe(lambda: wb.CheckCompatibility, "CheckCompatibility")))
        _dbg("[DIAG] wb.Name(before)=%r" % (_safe(lambda: wb.Name, "Name"),))
    _t0 = time.time()
    _wpid = _diag_pid_of_hwnd(_safe(lambda: xl.Hwnd, "Hwnd"))
    _wd = None
    if _wpid:
        try:
            _wd = threading.Timer(_DIAG_SAVE_TIMEOUT, _diag_watchdog_fire,
                                  args=(_wpid, n))
            _wd.daemon = True
            _wd.start()
        except Exception as _e:  # noqa: BLE001
            _dbg("[DIAG] watchdog start failed: %r" % (_e,))
            _wd = None
    try:
        _save_with_fallback(wb, target, fmt, is_wps, same_fmt)
    except Exception as _e:  # noqa: BLE001
        _dt = time.time() - _t0
        _hr, _sc = _diag_hr(_e)
        _hex = "?"
        if isinstance(_hr, int):
            _hex = hex(_hr & 0xFFFFFFFF)
        _shex = "?"
        if isinstance(_sc, int):
            _shex = hex(_sc & 0xFFFFFFFF)
        _dbg("[DIAG] SaveAs FAILED in %.2fs type=%s hresult=%r(%s) scode=%r(%s)" % (
            _dt, type(_e).__name__, _hr, _hex, _sc, _shex))
        _dbg("[DIAG] SaveAs exc str=%r args=%r" % (str(_e), getattr(_e, "args", None)))
        _dbg("[DIAG] SaveAs target after fail exists=%s size=%r" % (
            os.path.exists(target), _safe_size(target)))
        _dbg("[DIAG] excel alive probe Version=%r Workbooks=%r" % (
            _safe(lambda: xl.Version, "Version"),
            _safe(lambda: xl.Workbooks.Count, "Workbooks")))
        _dbg("[DIAG] wb alive probe Saved=%r Name=%r" % (
            _safe(lambda: wb.Saved, "Saved"),
            _safe(lambda: wb.Name, "Name")))
        # 窗口快照：有 #32770 对话框说明被弹窗挡住，一个窗口都没有说明进程已崩
        _diag_windows_of_pid(_diag_pid_of_hwnd(_safe(lambda: xl.Hwnd, "Hwnd")),
                             "after SaveAs failure")
        if first:
            _diag_events()
        raise
    finally:
        if _wd is not None:
            try:
                _wd.cancel()
            except Exception:
                pass
    _dbg("[DIAG] SaveAs OK in %.2fs size=%r" % (time.time() - _t0, _safe_size(target)))


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

    def _fetch_pid(self):
        self.xl_pid = None
        try:
            h = self._xl.Hwnd
            if h:
                import win32process
                _, self.xl_pid = win32process.GetWindowThreadProcessId(h)
        except Exception:
            self.xl_pid = None

    def get(self):
        if self._xl is not None:
            try:
                _ = self._xl.Application.Version
                _dbg("ExcelSession.get alive, pid=%s" % (self.xl_pid,))
            except Exception as _e:
                _dbg("ExcelSession.get probe failed, rebuild: %r" % (_e,))
                self._xl = None
                self.xl_pid = None
        if self._xl is None:
            _dbg("ExcelSession.get creating new Excel...")
            self._xl = _make_xl()
            self._fetch_pid()
            _dbg("ExcelSession.get new Excel pid=%s" % (self.xl_pid,))
        return self._xl

    def quit(self):
        if self._xl is not None:
            pid = self.xl_pid
            try:
                self._xl.Quit()
            except Exception:
                pass
            self._xl = None
            # Quit 偶发不干净会留下无窗口 EXCEL（锁文件、拖慢下次启动），
            # 只强杀本会话自己启动的那个 PID。
            if pid:
                try:
                    import time as _t
                    _t.sleep(0.4)
                    subprocess.run(["taskkill", "/F", "/PID", str(pid), "/T"],
                                   capture_output=True,
                                   creationflags=0x08000000)
                except Exception:
                    pass


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

    提速：逐行访问 Hidden 是每行一次 RPC，102 行就是 102 次，实测占单份
    耗时的 18%。改成一次 SpecialCells(xlCellTypeVisible) 拿到全部可见区域，
    再用 Area 的起止行反推隐藏行，COM 往返从 N 次降到 1 次。
    """
    if nrows <= 0:
        return set()
    try:
        rng = ws.Range(ws.Cells(1, 1), ws.Cells(nrows, 1))
        vis = rng.SpecialCells(12)          # xlCellTypeVisible
        seen = set()
        for area in vis.Areas:
            r1 = int(area.Row)
            r2 = r1 + int(area.Rows.Count) - 1
            if r2 > nrows:
                r2 = nrows
            r = r1
            while r <= r2:
                seen.add(r - 1)
                r += 1
        if seen:
            hidden = set()
            for r in range(nrows):
                if r not in seen:
                    hidden.add(r)
            return hidden
        # 整表都不可见（极端情况）→ 视为没有隐藏信息，退回逐行判定
    except Exception:
        pass
    hidden = set()
    for r in range(1, nrows + 1):
        try:
            if ws.Rows(r).Hidden:
                hidden.add(r - 1)
        except Exception:
            pass
    return hidden


def _used_bounds(ws, hard_cap=4096):
    """工作表的真实使用范围 (max_row, max_col)，只认「真正有值」的单元格。

    不能用 ws.max_column：它统计的是所有单元格对象，包括 Excel 另存时补出来的
    大量空白带样式单元（常见到 16384 列）。按它建网格，一张 60 行的表会变成
    60×16384 ≈ 100 万格，几十张表叠加直接把进程拖死（客户机上表现就是进度停在
    50%、永远等不到完成、也不触发重试）。
    """
    try:
        cells = getattr(ws, "_cells", None)
        if cells:
            vr = vc = 0
            for (r, c), cell in cells.items():
                try:
                    if cell.value is not None:
                        if r > vr:
                            vr = r
                        if c > vc:
                            vc = c
                except Exception:  # noqa: BLE001
                    continue
            if vr and vc:
                return vr, min(vc, hard_cap)
            # 全空表：退回单元格对象范围
            return max(k[0] for k in cells), min(max(k[1] for k in cells), hard_cap)
    except Exception:  # noqa: BLE001
        pass
    return ws.max_row, min(ws.max_column, hard_cap)


def _openpyxl_grid(ws):
    """从 openpyxl 工作表读取全部值，规整成 (grid, nrows, ncols, has_formula)。

    与 COM 版同构；顺带检测公式单元格（含公式必须降级 COM 强制重算，
    否则快路径可能读到旧缓存值或把公式文本当数值误判）。
    """
    nrows, ncols = _used_bounds(ws)
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

    # JO-WL-PZ-010 中微模板：表头行写「基准值 / 最大值 / 最小值」，子表头行只有
    # 「Actual Readings / Result」，三个规格列在 sub_row 全部为空 → 上下限判定整体失效。
    # 只在 max/min 双双未识别时回退到表头行定位；旧模板 sub_row 能命中，不走这里。
    if max_col is None and min_col is None:
        for rr in (sub_row, header_row):
            if max_col is None:
                max_col = find_col(rr, "上限", "max", "最大值")
            if min_col is None:
                min_col = find_col(rr, "下限", "min", "最小值")
            if spec_col is None:
                spec_col = find_col(rr, "公称", "dimension", "基准值", "nominal")
            if max_col is not None and min_col is not None:
                break

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
        kept = []
        for c in reading_cols:
            if sv(sub_row, c) or (header_row is not None and sv(header_row, c)):
                kept.append(c)
        # JO-WL-PZ-010 中微模板：读数列序号 1..N 单独放在表头下方第二行（r7），
        # 既不在 sub_row 也不在 header_row 上；而 header_row 的 Total/Sample Qty
        # 反而会被误当成读数列标签。仅当标签过滤会丢列、且能找到连续 1..N 的
        # 序号行时，采信 instr+1..result-1 的全部候选列。旧模板过滤不丢列时不动。
        if len(kept) < len(reading_cols) and item_col is not None:
            seq_ok = False
            for rr in range(header_row, min(header_row + 6, nrows)):
                if sv(rr, item_col):
                    continue
                ints = []
                good = True
                for c in reading_cols:
                    v = sv(rr, c)
                    try:
                        fv = float(v)
                    except (ValueError, TypeError):
                        good = False
                        break
                    if fv != int(fv) or fv < 1 or fv > 99:
                        good = False
                        break
                    ints.append(int(fv))
                if good and ints:
                    expect = list(range(1, len(ints) + 1))
                    if ints == expect:
                        seq_ok = True
                        break
            if seq_ok:
                kept = list(reading_cols)
        reading_cols = kept

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

# A1.7/T6：统一两套判定口径的 15 项完整 token（A2 验收口径，覆盖 _is_fail_text 与
# _is_fail_token 认过的全部词，少一个都是能力回退）。整格等值匹配，strict=True 不认裸 f/x。
_FAIL_TOKENS = ("f", "fail", "failed", "ng", "nogo", "reject", "rejected",
                "不合格", "不通过", "不良", "超差", "超上限", "超下限", "拒收", "x")

# 标红时"该改写成 F"的通过态取值。openpyxl 与 COM 两条标注路径共用一份，
# 避免两处各自维护后漂移。
_PASS_TOKENS = ("p", "pass", "ok", "合格", "good", "accept", "y", "yes")


def _only_cjk(s):
    """字符串是否全由汉字组成（空串返回 False）。"""
    if not s:
        return False
    for ch in s:
        if not ("\u4e00" <= ch <= "\u9fff"):
            return False
    return True


def _is_fail_token(v, strict=False):
    """结果列单元格是否为「不合格」判定（原文件检验员已标 F，必须尊重）。

    A18/T2（方案 B，已与 converter.py 逐字符对齐）：规范化顺序与 converter._is_fail_token
    完全一致——strip → 折叠/删除内部空白 → 转小写 → 去标点但保留数字
    （正则 [^\\w\\u4e00-\\u9fff]，同 converter._PUNCT_RE）→ 整格等值匹配 15 项。
    "F." → "f" 命中；"F.1" → "f1" 不命中（不再撞裸 f 兜底）；FQC / FPI / OK-F 全不命中。
    strict=True 时不认裸 "f" 与 "x"。（仅 analysis/a18 的验收脚本会传 True；
    主流程恒用默认 False，判定口径与 converter._is_fail_token 完全一致。）
    """
    if v is None:
        return False
    s = str(v).strip()
    if not s:
        return False
    s = re.sub(r"\s+", "", s)  # 折叠/删除内部空白（含换行、制表符）
    s = s.lower()
    s = re.sub(r"[^\w\u4e00-\u9fff]", "", s)  # 去标点符号，保留数字（converter._PUNCT_RE 同款）
    if strict and s in ("f", "x"):
        return False
    if s in _FAIL_TOKENS:
        return True
    # 检验员常写 "F(略)"、"F 超差" 这类「F + 中文备注」，规范化后是 f略 / f超差，
    # 必须判不合格（这是 F 的变体，不是别的词）。
    # 但 FQC → fqc、FPI → fpi、F.1 → f1 不能一起误判：
    # 只有「f 后面全是中文」才算 F 的备注，后面是字母/数字的一律不算。
    if s.startswith("f") and len(s) > 1 and _only_cjk(s[1:]):
        return True
    return False


def _is_unknown_result(v):
    """结果列的写法是不是「程序没认出来的」。只用于提醒，不参与合格/不合格判定。

    「认识」的口径 = 15 项不合格 token ∪ 8 项合格 token ∪ 空格与常见占位符。
    返回 True 表示这一格程序既没判成合格、也没判成不合格，而是直接跳过了它——
    这类格子如果不说一声，就等于静默放行，必须提示操作的人自己看一眼。
    """
    if v is None:
        return False
    s = str(v).strip()
    if not s:
        return False
    s2 = re.sub(r"[^\w\u4e00-\u9fff]", "", re.sub(r"\s+", "", s.lower()))
    if not s2:
        return False
    if s2 in _FAIL_TOKENS or s2 in _PASS_TOKENS:
        return False
    if s2 in ("", "/", "-", "—", "无"):
        return False
    return True


def _row_cause(grid, r, reading_cols, min_col, max_col):
    """判断某“结果列=F”行的失败原因（供 _fail_prefix 区分命名前缀）：

    - ``oob``     ：数据列写出了真实数值，但超出 [下限, 上限]（数据确实不在上下限范围内）
    - ``missing`` ：数据列写的是 ``/`` 或空白（没写数据），仅凭此判不合格
    - ``unknown`` ：数据列写出了合格值 / 文本（数据已写），但区间内却被标 F（归为数据不合格）

    仅当整行所有数据列都为 ``/`` 或空白时才判 ``missing``；
    只要任一数据列写了真实数据（数值 / 合格标记 / 文本），即视为“数据已写”。
    """
    minv = _fnum(_sv(grid, r, min_col)) if min_col is not None else None
    maxv = _fnum(_sv(grid, r, max_col)) if max_col is not None else None
    if minv is not None and maxv is not None and maxv < minv:
        minv, maxv = maxv, minv
    oob = False
    has_data = False
    for c in reading_cols:
        rs = _sv(grid, r, c)
        if rs is None:
            continue
        s = str(rs).strip()
        if s == "" or s == "/":
            continue
        if s.upper() in ("OK", "合格", "PASS", "P"):
            has_data = True
            continue
        v = _fnum(rs)
        if v is None:
            has_data = True
            continue
        has_data = True
        if (minv is not None and v < minv) or (maxv is not None and v > maxv):
            oob = True
    if oob:
        return "oob"
    if has_data:
        return "unknown"
    return "missing"


def _scan_core(grid, nrows, ncols, min_col, max_col, result_col, reading_cols,
               data_start, result_is_layout, hidden=None, warns=None):
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
            # A18/T1：结果列判定只用 15 项 token 整格精确匹配（_is_fail_text 包含式已退役），
            # 不做「超差已放行/让步/特采」类放行逻辑——真实结果列只有 P/F（A18 第 0 节口径）。
            if _is_fail_token(rv):
                fails.append({"r": r + 1, "c": result_col + 1,
                              "value": rv, "kind": "reading", "is_result": True,
                              "cause": _row_cause(grid, r, reading_cols, min_col, max_col)})
                notes.append(f"第{r + 1}行 原文件结果列={rv}")
            else:
                # A18/T1：非标准值（既不在 15 项 token，也不属明确合格值 P/合格/OK//空）。
                # 不改判定、不新增 fails 字段，只记一条提醒（界面显示黄色「结果待核」），
                # 让操作的人知道「这一格我没认出来」，不被静默当成合格。
                if _is_unknown_result(rv):
                    _dbg("结果列非标准值 第%d行 第%d列 值=%r" % (r + 1, result_col + 1, rv))
                    if warns is not None:
                        warns.append({"r": r + 1, "c": result_col + 1,
                                      "value": str(rv).strip()})
    # 不信任检验员结果列：无论结果列是否已标 F，程序都必须独立扫描实测数据一遍，
    # 把「数据超出上下限、但检验员没标 F」的行也抓出来（之前信任结果列会漏判这类行）。
    if not reading_cols:
        if fails:
            return fails, "；".join(notes)
        return [], "未识别到实测列"
    if min_col is None and max_col is None:
        if fails:
            return fails, "；".join(notes)
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
                # 这一行若已有结果列的 fail（检验员标过 F），直接把原因升级为 oob，
                # 避免与独立扫描重复计数；否则新增一条实测超差 fail。
                hit = False
                for f in fails:
                    if f["r"] == r + 1 and f.get("c") == result_col + 1:
                        f["cause"] = "oob"
                        hit = True
                        break
                if not hit:
                    fails.append({"r": r + 1, "c": result_col + 1,
                                  "value": _sv(grid, r, result_col),
                                  "kind": "reading", "is_result": True, "cause": "oob"})
            else:
                fails.append({"r": r + 1, "c": oob_col + 1, "value": oob_val,
                              "kind": "reading", "is_result": False, "cause": "oob"})
            notes.append(f"第{r + 1}行 测量值 {oob_val} 超出公差 [{lo_str}, {hi_str}]")

    if notes:
        return fails, "；".join(notes)
    return fails, "未检出不合格（逐行按上下限校验，无越界）"


def _scan_one_sheet(grid, nrows, ncols, hidden=None, warns=None):
    """对单张表扫描越界，返回 (fails, reason)。hidden: 被筛选/隐藏的 0-based 行集合。

    warns: 可选列表，收集「结果列写法没认出来」的行（只做黄色提醒，不改判定）。
    """
    layout = _detect_layout(grid, nrows, ncols)
    if layout is not None:
        max_col = layout["max_col"]
        min_col = layout["min_col"]
        result_col = layout["result_col"]
        reading_cols = layout["reading_cols"]
        data_start = layout["data_start"]

        fails, reason = _scan_core(grid, nrows, ncols, min_col, max_col,
                                   result_col, reading_cols, data_start,
                                   result_is_layout=True, hidden=hidden,
                                   warns=warns)
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

    # 兜底路径的有效数据行区间：跳过表头标签行（result/结果/判定/结论/P合格 等），
    # 从第一个非表头标签的非空结果列行开始做 F 判定（B3 区间约束）。
    data_start = 0
    if result_col is not None:
        for r in range(nrows):
            v = _sv(grid, r, result_col)
            if not v:
                continue
            vl = v.lower()
            if any(k in vl for k in ("result", "结果", "判定", "结论", "p合格")):
                continue
            data_start = r
            break

    fails, reason = _scan_core(grid, nrows, ncols, min_col, max_col,
                               result_col, reading_cols, data_start,
                               result_is_layout=False, hidden=hidden,
                               warns=warns)
    return fails, reason


def _scan_workbook(xl, wb):
    """扫描已打开的 COM 工作簿的「可见」工作表，返回 (fails, reason, norm)。

    隐藏/深度隐藏工作表跳过；被筛选或手动隐藏的行不判断。
    """
    all_fails = []
    all_reasons = []
    all_warns = []
    grids = []
    sheet_count = wb.Sheets.Count
    _dbg("scan_workbook Sheets.Count=%s" % (sheet_count,))
    for i in range(1, sheet_count + 1):
        try:
            ws = wb.Sheets(i)
            _vis = ws.Visible
            try:
                _vis_i = int(_vis)
            except Exception:
                _vis_i = "?"
            _dbg("scan_workbook sheet %s name=%r Visible=%r type=%s int=%s"
                 % (i, str(ws.Name), _vis, type(_vis).__name__, _vis_i))
            if int(ws.Visible) != -1:
                continue  # 隐藏工作表：用户看不到，不判断
        except Exception as e:
            _dbg("scan_workbook sheet %s visible-probe ERR %r" % (i, e))
            continue
        sheet_idx = len(grids) + 1  # 可见表序号
        try:
            grid, nrows, ncols = _sheet_grid(ws)
            _dbg("scan_workbook sheet %s grid (nrows=%s, ncols=%s)" % (i, nrows, ncols))
            if not grid or nrows == 0:
                _dbg("scan_workbook sheet %s grid EMPTY -> continue" % (i,))
                continue
            hidden = _hidden_rows_com(ws, nrows)
            grids.append((grid, nrows, ncols, hidden))
            sheet_warns = []
            fails, reason = _scan_one_sheet(grid, nrows, ncols, hidden,
                                            warns=sheet_warns)
            for w in sheet_warns:
                w["sheet"] = sheet_idx
                all_warns.append(w)
            if fails:
                for f in fails:
                    f["sheet"] = sheet_idx
                all_fails.extend(fails)
                all_reasons.append(f"工作表 {sheet_idx}（{ws.Name}）: {reason}")
        except Exception as e:
            _dbg("scan_workbook sheet %s scan ERR: %s" % (i, traceback.format_exc()))
            # 图表工作表等无 UsedRange 的对象：跳过并记录，不让整文件报错
            all_reasons.append(f"工作表 {sheet_idx}: 读取失败（可能为图表工作表），已跳过")

    norm = _collect_norm(grids)
    # 剔除合并单元格（表头/页脚结构性区域，不填 /）
    _drop_merged_slash_com(norm, wb)
    # A1.6：读取失败不得谎报成「已扫描全部」。读取失败项单独统计，reason 如实说明，
    # 判定口径不变（不升级为不合格），只把失败事实带出。
    read_fail = [r for r in all_reasons if "读取失败" in r]
    if all_fails:
        reason = "；".join(all_reasons)
    elif read_fail:
        reason = "未检出不合格，但有 %d 张表读取失败；%s" % (len(read_fail), read_fail[0])
    else:
        reason = "未检出不合格（已扫描全部可见工作表，无越界）"
    if read_fail:
        norm["scan_errors"] = {"count": len(read_fail), "items": read_fail[:3]}
    if all_warns:
        norm["result_warns"] = all_warns
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
    # 兼容回退：JO-WL-PZ-010 中微模板用「Part No / 料号」，没有 图号 / WG No 标签。
    # 仅在上面完全找不到图号时启用，旧模板一律走原路径。
    for r in range(min(nrows, 12)):
        for c in range(ncols):
            t = _sv(grid, r, c).lower()
            if not t:
                continue
            if "part no" in t or "料号" in t:
                if "description" in t or "零件名称" in t:
                    continue
                for cc in range(c + 1, min(c + 6, ncols)):
                    v = _sv(grid, r, cc)
                    if v and not _is_cust_label(v):
                        first = v.split("\n")[0].strip()
                        return (r, cc, first)
    return None


def _is_item_no(s):
    """A 列是否为检验项编号（1 / 13_1 / A / B）。

    用于区分数据行与尾部签名/说明区：页脚常见「Inspected By (检验)：」
    「(Sign/Employee No./Date签名日期)」，它们 A 列有长文本、结果列为空。
    """
    s = (s or "").strip()
    if not s or len(s) > 24:
        return False
    for ch in s:
        if ch.isdigit() or ch in "_-.":
            continue
        if "a" <= ch.lower() <= "z":
            continue
        return False
    return True


MATERIAL_DEFAULT = "6061-T6"


def _find_material_label(grid, nrows, ncols):
    """定位表头区「Material / 材质」标签格，返回 (r, c)（0-based）。

    只在前 8 行找，避开数据区里「材质：6061-T6」这类检验项文字。
    """
    for r in range(min(nrows, 8)):
        for c in range(ncols):
            t = _sv(grid, r, c).lower()
            if not t:
                continue
            if "material" in t or "材质" in t:
                if "description" in t or "零件名称" in t:
                    continue
                return (r, c)
    return None


def _material_value_rc_openpyxl(ws, r1, c1):
    """标签合并区右边界 +1 列即值格；值格若属于合并区，取左上角。

    JO-WL-PZ-010 中微模板：标签 E4:G4 合并，值格 H4:I4 合并。
    """
    end_col = c1
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= r1 <= rng.max_row and rng.min_col <= c1 <= rng.max_col:
            end_col = rng.max_col
            break
    tc = end_col + 1
    if tc > ws.max_column:
        return None
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= r1 <= rng.max_row and rng.min_col <= tc <= rng.max_col:
            tc = rng.min_col
            break
    return (r1, tc)


def _material_value_rc_com(sh, r1, c1):
    """COM 版材质值格定位，逻辑同 _material_value_rc_openpyxl。"""
    end_col = c1
    try:
        ma = sh.Cells(r1, c1).MergeArea
        end_col = int(ma.Column) + int(ma.Columns.Count) - 1
    except Exception:
        pass
    tc = end_col + 1
    try:
        tc = int(sh.Cells(r1, tc).MergeArea.Column)
    except Exception:
        pass
    if tc >= 16384:
        return None
    return (r1, tc)


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
           "label_erase": [], "n_meas": 0, "material": None}

    # 1) 客户代码：标签右侧第一个非空值单元格（跳过嵌套标签）
    cc = _find_customer_code(grid, nrows, ncols)
    if cc is not None:
        out["customer_code"] = {"sheet": sheet_idx, "r": cc[0], "c": cc[1],
                                "old": _sv(grid, cc[0], cc[1])}

    # 1b) 材质标签：为空时统一补 6061-T6（只记录标签位置，值格在写回阶段按合并区定位）
    ml = _find_material_label(grid, nrows, ncols)
    if ml is not None:
        out["material"] = {"sheet": sheet_idx, "r": ml[0], "c": ml[1]}

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

            # 尾部签名/说明区保护（JO-WL-PZ-010 中微模板）：页脚的 A 列有长文本、
            # 结果列为空，但 I/L 列恰好落在测量列范围内，会被误算成「有数据的读数列」
            # （总数虚高、结果列被补 P、序号标签删错）。
            # 仅当最后一个结果值之后的行 A 列全部不是编号时，才把数据区截到那里；
            # 尾部若有真正的编号行（结果列缺失、v11 需自动补 P/F）则保持原行为。
            if result_col is not None and item_col is not None and data_end > data_start:
                last_res = None
                for r in range(data_start, data_end):
                    if hidden and r in hidden:
                        continue
                    if _sv(grid, r, result_col):
                        last_res = r + 1
                if last_res is not None and last_res < data_end:
                    tail_footer = True
                    for r in range(last_res, data_end):
                        if hidden and r in hidden:
                            continue
                        if _is_item_no(_sv(grid, r, item_col)):
                            tail_footer = False
                            break
                    if tail_footer:
                        data_end = last_res
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
            "label_erase": [], "n_meas_total": 0, "split": [],
            "material": [], "material_done": 0}
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
        if si.get("material") is not None:
            norm["material"].append(si["material"])
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
        "material_filled": int(norm.get("material_done") or 0),
    }


def _scan_errors_summary(norm):
    """从 norm 提取 scan_errors 对外摘要（数量 + 前 3 条），缺省返回空结构。

    A1.6：GUI 与 worker 的 result 均需携带该字段，如实反映表读取失败。
    """
    se = norm.get("scan_errors")
    if not se:
        return {"count": 0, "items": []}
    return {"count": int(se.get("count", 0) or 0),
            "items": list((se.get("items") or [])[:3])}


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
    # 材质为空 → 补 6061-T6。只在值格确实为空时写，已有材质一律不动。
    for mt in norm.get("material", []):
        try:
            ws = vis[mt["sheet"] - 1]
            rc = _material_value_rc_openpyxl(ws, mt["r"] + 1, mt["c"] + 1)
            if rc is None:
                continue
            cell = ws.cell(rc[0], rc[1])
            cur = "" if cell.value is None else str(cell.value).strip()
            if cur == "":
                cell.value = MATERIAL_DEFAULT
                norm["material_done"] = int(norm.get("material_done") or 0) + 1
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


def _batch_write_value(ws, cells, value, max_chars=200):
    """把同一格里的一批不相邻单元格一次性写成同一个值。

    Excel 单格写入实测 5ms（Value 的 PropertyLet 要走依赖树与撤销栈），
    读只要 0.7ms。斜杠填充在部分客户表里有几十上百个格子，逐格写就是
    几百毫秒。ws.Range("A1,C5,E9").Value = "/" 可一次完成，实测可用。

    地址串长度上限取 200 字符（实测 230 可用、310 报 0x800A03EC）。
    任何一批失败都退回逐格，保证不漏填。返回实际写入的格数。
    """
    if not cells:
        return 0
    done = 0
    buf = []
    buf_len = 0

    def _flush():
        nonlocal done, buf, buf_len
        if not buf:
            return
        try:
            ws.Range(",".join(buf)).Value = value
            done += len(buf)
        except Exception:  # noqa: BLE001
            for addr in buf:
                try:
                    ws.Range(addr).Value = value
                    done += 1
                except Exception:  # noqa: BLE001
                    pass
        buf = []
        buf_len = 0

    for r, c in cells:
        addr = "%s%d" % (_col_letter(c), r)
        if buf and buf_len + len(addr) + 1 > max_chars:
            _flush()
        buf.append(addr)
        buf_len += len(addr) + 1
    _flush()
    return done


def _apply_norm_com(wb, norm):
    """把规范化应用到 COM 工作簿（全表）。

    客户代码以文本格式写入（NumberFormat="@"），保证前导 0 不被省略；
    若目标单元格在合并区域内，写合并区域左上角。
    """
    if not norm:
        return
    # 同一张表的重复查找只做一次：_com_visible_sheet 每次都要遍历 Sheets 集合
    # 读 Visible，单份文件里被调用 6 次，实测十几毫秒。
    _shc = {}

    def _sh(idx):
        if idx in _shc:
            return _shc[idx]
        _s = _com_visible_sheet(wb, idx)
        _shc[idx] = _s
        return _s

    n_clear = _clear_fills_com(wb, norm.get("fill_counts"))
    if n_clear:
        norm["fill_cleared"] = n_clear
    cc = norm.get("customer_code")
    if cc:
        try:
            sh = _sh(cc["sheet"])
            if sh is None:
                raise Exception("no visible sheet")
            cell = sh.Cells(cc["r"] + 1, cc["c"] + 1)
            # 已经是目标值就别再写：单格写入 5ms，读只要 0.7ms。
            # 批量跑存量文件时"客户代码本来就对"是常态，这一步能整段跳过。
            try:
                _cur = cell.Value
            except Exception:  # noqa: BLE001
                _cur = None
            _cur_s = "" if _cur is None else str(_cur).strip()
            if _cur_s == CUSTOMER_CODE or _cur_s == str(CUSTOMER_CODE):
                cc = None
        except Exception:
            pass
    if cc:
        try:
            sh = _sh(cc["sheet"])
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
    # 材质为空 → 补 6061-T6（与 openpyxl 路径同口径）
    for mt in norm.get("material", []):
        try:
            sh = _sh(mt["sheet"])
            if sh is None:
                continue
            rc = _material_value_rc_com(sh, mt["r"] + 1, mt["c"] + 1)
            if rc is None:
                continue
            cell = sh.Cells(rc[0], rc[1])
            try:
                cur = cell.Value
            except Exception:  # noqa: BLE001
                cur = None
            if cur is None or str(cur).strip() == "":
                cell.Value = MATERIAL_DEFAULT
                norm["material_done"] = int(norm.get("material_done") or 0) + 1
        except Exception:
            pass
    # 斜杠填充：按表分组批量写，逐格写会按 5ms/格线性增长
    _slash_by_sheet = {}
    for sheet_idx, r, c in norm.get("slash_cells", []):
        _slash_by_sheet.setdefault(sheet_idx, []).append((r + 1, c + 1))
    for sheet_idx, _cells in _slash_by_sheet.items():
        try:
            sh = _sh(sheet_idx)
            if sh is not None:
                _batch_write_value(sh, _cells, "/")
        except Exception:
            pass
    # v11：结果列缺失自动补 P/F（黄底）
    for rf in norm.get("result_fill", []):
        try:
            sh = _sh(rf["sheet"])
            if sh is not None:
                cell = sh.Cells(rf["r"] + 1, rf["c"] + 1)
                try:
                    _cur = cell.Value
                except Exception:  # noqa: BLE001
                    _cur = None
                if _cur is None or str(_cur).strip() != str(rf["val"]).strip():
                    cell.Value = rf["val"]
                # 底色已经是黄色就不必再设：设 Pattern + Color 是两次 5ms 写入
                try:
                    _need_fill = (int(cell.Interior.Color) != 0x00FFFF
                                  or int(cell.Interior.Pattern) != 1)
                except Exception:  # noqa: BLE001
                    _need_fill = True
                if _need_fill:
                    cell.Interior.Pattern = 1
                    cell.Interior.Color = 0x00FFFF  # BGR 黄色
        except Exception:
            pass
    # v11：总数/检样数更正
    for t in norm.get("total_qty", []):
        try:
            sh = _sh(t["sheet"])
            if sh is not None:
                cell = sh.Cells(t["r"] + 1, t["c"] + 1)
                try:
                    _cur = cell.Value
                except Exception:  # noqa: BLE001
                    _cur = None
                if _cur is None or str(_cur).strip() != str(t["new"]).strip():
                    cell.Value = t["new"]
        except Exception:
            pass
    for t in norm.get("sample_qty", []):
        try:
            sh = _sh(t["sheet"])
            if sh is not None:
                cell = sh.Cells(t["r"] + 1, t["c"] + 1)
                try:
                    _cur = cell.Value
                except Exception:  # noqa: BLE001
                    _cur = None
                if _cur is None or str(_cur).strip() != str(t["new"]).strip():
                    cell.Value = t["new"]
        except Exception:
            pass
    # v11.2：删多余序号标签（清空单元格）
    for le in norm.get("label_erase", []):
        try:
            sh = _sh(le["sheet"])
            if sh is not None:
                cell = sh.Cells(le["r"] + 1, le["c"] + 1)
                try:
                    _cur = cell.Value
                except Exception:  # noqa: BLE001
                    _cur = None
                if _cur is not None and str(_cur).strip() != "":
                    cell.Value = ""
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
    except Exception as e:
        _dbg("load_openpyxl fail %r" % (e,))
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


def _audit_openpyxl(input_path, wb=None, on_progress=None):
    """openpyxl 快路径：扫描全部工作表并探测规范化信息。

    返回 (fails, reason, norm)；若文件含公式或 openpyxl 无法保真的复杂元素，返回 None
    （调用方应降级到 COM 路径）。
    wb 传入时使用外部已加载的工作簿（不负责关闭）；否则内部加载并在结束时关闭。
    on_progress(pct, msg)：中途进度回调，让界面在大文件上也能看到动。
    """
    own = wb is None

    def _prog(pct, msg):
        if on_progress:
            try:
                on_progress(pct, msg)
            except Exception:  # noqa: BLE001
                pass

    if own:
        _prog(3, "读取工作簿")
        wb = _load_openpyxl(input_path)
    _dbg("audit_openpyxl entry input=%s wb_is_none=%s" % (os.path.basename(input_path), wb is None))
    if wb is None:
        return None
    try:
        risky = _openpyxl_risky(wb)
        _dbg("audit_openpyxl risky=%s" % (risky,))
        if risky:
            return None
        all_fails = []
        all_reasons = []
        all_warns = []
        grids = []
        formula_in_result = False
        vis = [w for w in wb.worksheets if w.sheet_state == "visible"]
        nvis = max(len(vis), 1)
        for idx, ws in enumerate(vis):
            # 扫描阶段占 5%-80%，逐表推进；单表很大时（几百列）也能看出在动
            _prog(5 + int(75 * idx / nvis), "扫描第 %d/%d 张表" % (idx + 1, nvis))
            grid, nrows, ncols, has_formula = _openpyxl_grid(ws)
            _dbg("audit_openpyxl sheet %r max_row=%s max_col=%s grid=(%s,%s,%s)"
                 % (ws.title, ws.max_row, ws.max_column, nrows, ncols, has_formula))
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
            sheet_warns = []
            fails, reason = _scan_one_sheet(grid, nrows, ncols, hidden,
                                            warns=sheet_warns)
            for w in sheet_warns:
                w["sheet"] = len(grids)
                all_warns.append(w)
            if fails:
                sidx = len(grids)
                for f in fails:
                    f["sheet"] = sidx
                all_fails.extend(fails)
                all_reasons.append(f"工作表 {sidx}（{ws.title}）: {reason}")
            elif not formula_in_result:
                # B4 兜底：该 sheet 判合格，但结果列存在公式（openpyxl 读到公式文本，
                # 无缓存值）→ 不能据此判合格，交给 COM 重算复核
                lay = _detect_layout(grid, nrows, ncols)
                rc = (lay or {}).get("result_col")
                if rc is not None:
                    for rr in range(nrows):
                        if _sv(grid, rr, rc).startswith("="):
                            formula_in_result = True
                            break
        if not all_fails and formula_in_result:
            return None  # 结果列是公式且无缓存值 → 交给 COM 重算复核（B4）
        _prog(82, "汇总判定结果")
        norm = _collect_norm(grids)
        # 剔除合并单元格（表头/页脚结构性区域，不填 /）
        _drop_merged_slash_openpyxl(norm, wb)
        if all_fails:
            reason = "；".join(all_reasons)
        else:
            reason = "未检出不合格（openpyxl 快路径，已扫描全部工作表）"
        if all_warns:
            norm["result_warns"] = all_warns
        return all_fails, reason, norm
    finally:
        if own:
            try:
                wb.close()
            except Exception:
                pass


def _mark_openpyxl(ws, fails):
    """对 openpyxl 工作表中的目标单元格应用红底色 + 标 F（字体保持不变）。"""
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
        except Exception as e:  # noqa: BLE001
            _dbg("_mark_openpyxl r%s c%s failed: %r" % (r, c, e))


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


# 单块读数达到这个数量就不再「宽表拆分」。
# 背景：.xls 最多 256 列，件数超过约 250 时制表人只能把余下件折到下面另起一块。
# 我们把两块合成一块后是 300-400 列，如果仍按每组 5 列拆，会拆成 60-80 张表，
# 客户明确不接受（"公司不允许这样分两列"）。240 留出余量：正常模板表远达不到。
WIDE_SKIP_SPLIT_M = 240


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
        if M >= WIDE_SKIP_SPLIT_M:
            # 单块 240 列以上：这是「件数多到 .xls 放不下」的大表（合并后一块到底），
            # 按 5 列一组拆会变成几十张表，客户不接受。整表直接审核即可。
            _dbg("plan_splits skip %r M=%d (>=%d) 不拆分"
                 % (ws.title, M, WIDE_SKIP_SPLIT_M))
            continue
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
                    pending_shifts=None, on_saved=None):
    """单次打开完成「重算 → 扫描 → 归位 → 规范化 → 标红 → SaveAs」（v14 合并，省去二次打开）。

    返回 (fails, reason, norm, output, marked, copied)。
    扫描为纯读；规范化/标红在同一工作簿完成，最后 SaveAs 到输出目录（不覆盖、同名递增）。
    on_saved: 可选回调，SaveAs 成功后立即调用 on_saved(output)，供上层在后续环节
    （图号前缀改名等）抛 RPC 异常时能清掉本轮已写出的半成品输出（A1.7/T2）。
    """
    tmp_path = None
    wb = None
    prev = None
    try:
        ext_ok = ext or ".xlsx"
        tmp_path = _make_tmp("_audit_", ext_ok)
        shutil.copy2(os.path.abspath(input_path), tmp_path)
        prev = _com_speed_on(xl)
        import time as _t0
        _t_start = _t0.time()
        _dbg("Workbooks.Open start: tmp=%s input=%s" % (tmp_path, input_path))
        wb = xl.Workbooks.Open(os.path.abspath(tmp_path), ReadOnly=False, UpdateLinks=0)
        _dbg("Workbooks.Open OK in %.2fs, sheets=%d" % (_t0.time() - _t_start, wb.Sheets.Count))
        # 打开后就拍照：工作簿特性 + 各表状态（进程级信息在 _make_xl 已记过）
        _diag_header()
        _diag_excel(xl, wb, env=False)
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
        prefix = _fail_prefix(all_fails, norm)
        out_path = os.path.join(out_dir, base + ext)
        if prefix:
            out_path = _with_prefix(out_path, prefix)

        # 同一工作簿内：空列归位 → 规范化（清底色/补值/黄底）→ 标红（红覆盖黄）
        # 每步记耗时，万一崩在 SaveAs 之前也能看出是哪一步把 Excel 弄挂的
        _t_p = _t0.time()
        _apply_gap_move_com(wb, norm, all_fails)
        _dbg("[DIAG] apply_gap_move done in %.2fs" % (_t0.time() - _t_p,))
        _t_p = _t0.time()
        _apply_norm_com(wb, norm)
        _dbg("[DIAG] apply_norm done in %.2fs" % (_t0.time() - _t_p,))
        by_sheet = {}
        for f in all_fails:
            by_sheet.setdefault(f.get("sheet", 1), []).append(f)
        _t_p = _t0.time()
        for sheet_idx, sheet_fails in sorted(by_sheet.items()):
            sh = _com_visible_sheet(wb, sheet_idx)
            if sh is not None:
                _mark_fails(sh, sheet_fails)
        _dbg("[DIAG] mark_fails done in %.2fs, fails=%d" % (
            _t0.time() - _t_p, len(all_fails)))

        fmt = _EXT_FMT.get(ext_ok.lower(), 51)
        _dbg("[DIAG] ext=%r fmt=%r out_path=%s len=%d" % (
            ext_ok, fmt, out_path, len(out_path)))
        # Excel 2007 保存新版 .xlsx 时会弹"兼容性检查"对话框，不吃 DisplayAlerts，
        # 必须显式关 CheckCompatibility，否则 SaveAs 被模态框阻塞 → RPC_E_CALL_REJECTED。
        try:
            wb.CheckCompatibility = False
            _dbg("set wb.CheckCompatibility=False OK, readback=%r" % (
                _safe(lambda: wb.CheckCompatibility, "CheckCompatibility"),))
        except Exception as _e:
            _dbg("set wb.CheckCompatibility=False failed (old Excel?): %r" % (_e,))

        _save_state = {"n": 0, "dead": False, "first_exc": None}
        # RPC 通道死亡的四个 hr：重试没有任何意义，只会白等 20 次
        _dead_hrs = (-2147417851, -2147418111, -2147417846, -2147023174)
        # WPS 判定：它的 Excel 兼容层对标准 SaveAs 会内部失败，需要走降级链。
        # 用安装路径判断最可靠（WPS 会把自己伪装成 "Microsoft Excel" 12.0）。
        _is_wps = False
        try:
            _xp = str(xl.Path).lower()
            if ("wps" in _xp) or ("kingsoft" in _xp):
                _is_wps = True
        except Exception:  # noqa: BLE001
            _is_wps = False
        # SaveCopyAs 不能指定 FileFormat，只有「打开的文件」与「输出」同扩展名时
        # 才能用它（实测比 SaveAs 快 20ms，产出内容逐格一致）。跨格式必须 SaveAs。
        # 注意要按 input_path 判断：tmp 副本是按输出扩展名命名的，内容格式才是真相。
        _same_fmt = False
        try:
            _same_fmt = (os.path.splitext(str(input_path))[1].lower()
                         == ext_ok.lower())
        except Exception:  # noqa: BLE001
            _same_fmt = False

        def _save_fn(t):
            _save_state["n"] += 1
            _n = _save_state["n"]
            if _save_state["dead"]:
                if _n == 2:
                    _dbg("[DIAG] Excel 通道已判定死亡，后续重试立即返回原异常（不再逐个刷日志）")
                raise _save_state["first_exc"]
            try:
                _diag_save_as(wb, xl, t, fmt, _n, _n == 1, _is_wps, _same_fmt)
            except Exception as _e:  # noqa: BLE001
                _save_state["first_exc"] = _e
                _hr, _sc = _diag_hr(_e)
                if _hr in _dead_hrs or _sc in _dead_hrs:
                    _save_state["dead"] = True
                raise

        output = _write_output(_save_fn, out_path)
        if on_saved is not None:
            try:
                on_saved(output)
            except Exception:
                pass
        return all_fails, reason, norm, output, is_fail, (not is_fail)
    finally:
        try:
            if wb is not None:
                wb.Close(False)
        except Exception as _e:  # noqa: BLE001
            # 收尾都关不掉，说明 Excel 进程已经不可用，这条日志是重要线索
            _dbg("[DIAG] wb.Close(False) failed in cleanup: %r" % (_e,))
        if prev is not None:
            _com_speed_off(xl, prev)
        if tmp_path and os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except Exception:
                pass
        _dbg("[DIAG] _audit_com_file exit")


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

def _col_letter(n):
    """列号转字母（1→A，27→AA）。避免为这一处引入 openpyxl 依赖。"""
    s = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return s


def _batch_colored(ws, cells, color, pattern=1, max_chars=200):
    """把一批不相邻单元格一次性铺底色。

    逐格写要 2 次 COM 往返/格（Pattern + Color）；实测 100 个失败格光这一步
    就要几十毫秒。Excel 允许 ws.Range("A1,C5,E9") 这种多区域对象一次性
    设置 Interior，于是 N 格压到 ceil(N/批) 次往返。

    地址串有长度上限：实测 230 字符可用、310 字符报 0x800A03EC，取 200 保守值。
    某一批失败时退回逐格写，保证不因批量失败而漏标。

    返回实际标记的格数。
    """
    if not cells:
        return 0
    done = 0
    buf = []
    buf_len = 0

    def _flush():
        nonlocal done, buf, buf_len
        if not buf:
            return
        try:
            rng = ws.Range(",".join(buf))
            rng.Interior.Pattern = pattern
            rng.Interior.Color = color
            done += len(buf)
        except Exception:  # noqa: BLE001
            for addr in buf:
                try:
                    c = ws.Range(addr)
                    c.Interior.Pattern = pattern
                    c.Interior.Color = color
                    done += 1
                except Exception:  # noqa: BLE001
                    pass
        buf = []
        buf_len = 0

    for r, c in cells:
        addr = "%s%d" % (_col_letter(c), r)
        if buf and buf_len + len(addr) + 1 > max_chars:
            _flush()
        buf.append(addr)
        buf_len += len(addr) + 1
    _flush()
    return done


def _mark_fails(ws, fails):
    """对 fails 中的【目标单元格】应用红底色 + 标 F（字体保持不变）。

    分两步走：先把所有失败格一次性铺红（批量 RPC），再单独处理需要在结果列
    写 F 的那些格（数量远少于红格，且必须先问 HasFormula 才能覆盖）。
    原实现每格 3 到 4 次往返，失败格多时这里是纯 RPC 堆积。
    """
    cells = []
    for f in fails:
        try:
            cells.append((int(f["r"]), int(f["c"])))
        except Exception:  # noqa: BLE001
            continue
    _batch_colored(ws, cells, RED_FILL)

    for f in fails:
        r, c = f["r"], f["c"]
        if f.get("kind") != "reading" or not f.get("is_result"):
            continue
        try:
            cell = ws.Cells(r, c)
            if bool(cell.HasFormula):
                continue
            # 先用扫描阶段已拿到的原值判断，省掉一次 cell.Value 往返
            if "value" in f:
                raw = f.get("value")
            else:
                raw = cell.Value
            s = "" if raw is None else str(raw).strip()
            t2 = _strip_symbols(s.lower())
            if t2 in _PASS_TOKENS or s == "":
                cell.Value = "F"
        except Exception as e:  # noqa: BLE001
            _dbg("_mark_fails r%s c%s failed: %r" % (r, c, e))


def _mark_and_save(input_path, fails, out_path):
    """兼容 API：整簿标红后另存。

    主流程已改为 `_audit_openpyxl` / `_audit_com_file` 内联标注（同一次打开里
    完成判定+标红+保存），此函数仅保留给外部脚本与旧测试调用，主流程不再走它。
    """
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
    return _with_prefix(out_path, "图号错误_")


# 图号核对：字段命中所需的最小字符数（依据见 _dwg_match 的实测说明）
DWG_MIN_FIELD = 6


def _dwg_bad(norm, base):
    """图号与文件名完全对不上 → True（需加 图号错误_ 前缀）。

    放宽规则（客户 2026-10-24 要求）：只要有任一「字段」与文件名相同，就不算图号错误，
    只给黄色提示「图号不完全一致」。判定见 _dwg_match。
    """
    dn = (norm or {}).get("drawing_no")
    if not dn:
        return False
    dwg = str(dn.get("value", "")).strip()
    if not dwg:
        return False
    return _dwg_match(dwg, base)[0] == "bad"


def _dwg_match(dwg, base):
    """图号与文件名的匹配级别：("ok"|"partial"|"bad", 命中的字段)。

    ok      : 表内图号整串就是文件名的一部分（老规则）
    partial : 整串对不上，但有某个字段对上了（新规则：不算图号错误，只黄色提示）
    bad     : 一个字段都对不上（仍然加「图号错误_」前缀）

    字段长度下限 6：实测真实文件里能对上的字段最短是 8 个字符（93022011），
    整串对不上时最长的巧合只有 2 个字符（'30'）。取 6 既能容忍少写几位，
    也不会被 '30'、'93' 这类数字巧合蒙过去。
    """
    dwg = (dwg or "").strip()
    base = base or ""
    if not dwg:
        return "ok", None
    d, b = dwg.lower(), base.lower()
    if d in b:
        return "ok", dwg
    best = ""
    for tok in re.split(r"[-_/\s（）()【】\[\]，,。.．:：;；]+", dwg):
        t = tok.strip().lower()
        if len(t) >= DWG_MIN_FIELD and t in b:
            if len(t) > len(best):
                best = tok.strip()
    if best:
        return "partial", best
    # 没有完整字段命中：再用最长公共子串兜底（图号与文件名分段写法不一致时常见）
    m = _common_substr_len(d, b)
    if m[0] >= DWG_MIN_FIELD + 2:
        return "partial", m[1]
    return "bad", None


def _common_substr_len(a, b):
    """最长公共子串 (长度, 文本)。仅用于图号核对，长度都很短，O(n*m) 可接受。"""
    if not a or not b:
        return 0, ""
    best = 0
    best_end = 0
    prev = [0] * (len(b) + 1)
    for i in range(1, len(a) + 1):
        cur = [0] * (len(b) + 1)
        ai = a[i - 1]
        for j in range(1, len(b) + 1):
            if ai == b[j - 1]:
                cur[j] = prev[j - 1] + 1
                if cur[j] > best:
                    best = cur[j]
                    best_end = i
        prev = cur
    return best, a[best_end - best:best_end]


def _v11_apply_result_f(fails, reason, norm):
    """结果列缺失自动补出的 F → 并入 fails（判不合格并标红）。

    去重：扫描阶段（_scan_core 已改为不信任结果列、始终独立扫描）已对实测超差行
    独立判定并加入 fails；此处只对「扫描未覆盖」的自动补 F 行（如扫描因读数位未识别
    而漏判的极少数情况）追加，避免同一行被计两次。
    """
    rf = (norm or {}).get("result_fill", [])
    add = [x for x in rf if x["val"] == "F"]
    if not add:
        return fails, reason, norm
    # 已存在的 fail 位置集合（1-based r/c，含 sheet），用于去重
    existing = set()
    for f in fails:
        existing.add((f.get("sheet", 1), f["r"], f["c"]))
    new_fails = list(fails)
    for x in add:
        if (x.get("sheet", 1), x["r"] + 1, x["c"] + 1) in existing:
            continue  # 扫描已覆盖该行，跳过
        new_fails.append({
            "r": x["r"] + 1, "c": x["c"] + 1, "value": "F",
            "kind": "reading", "is_result": True, "sheet": x["sheet"],
            "cause": "oob",
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
    n_mat = int(norm.get("material_done") or 0)
    if n_mat:
        log.append({"level": "info", "op": "材质",
                    "msg": "%d 处为空，已补 %s" % (n_mat, MATERIAL_DEFAULT)})
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
        lvl, tok = _dwg_match(dwg, base)
        if lvl == "ok":
            log.append({"level": "info", "op": "图号",
                        "msg": "图号 %s 在文件名中，一致" % dwg})
        elif lvl == "partial":
            # 有字段相同：不改名、不判图号错误，只提醒核对（客户要求）
            log.append({"level": "warn", "op": "图号",
                        "msg": "图号不完全一致：表内「%s」与文件名「%s」只有字段「%s」相同，"
                               "请核对图号（未加「图号错误_」前缀）"
                               % (dwg or "(空)", base, tok)})
        else:
            log.append({"level": "err", "op": "图号",
                        "msg": "图号 %s 不在文件名「%s」中，已加「图号错误_」前缀"
                               % (dwg or "(空)", base)})
    if norm.get("label_erase"):
        log.append({"level": "info", "op": "序号标签",
                    "msg": "删除多余序号标签 %d 个（实际有数据列较少，空白数据列保留）" % len(norm["label_erase"])})
    for g in norm.get("gap_move", []):
        log.append({"level": "warn", "op": "空列整理",
                    "msg": "结果列与数据列之间删除 %d 个空列，结果列移至第 %d 列" % (g.get("gap"), g.get("c_new") + 1)})
    if norm.get("wide_fixed"):
        log.append({"level": "warn", "op": "格式统一",
                    "msg": "已把第 251 列以后的 %d 列改为与前面同样的边框和列宽（数据未改动）"
                           % norm["wide_fixed"]})
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
    rw = norm.get("result_warns") or []
    if rw:
        # 结果列有程序不认识的写法：不改判定，只提醒（界面黄色「结果待核」）
        sample = []
        for w in rw[:3]:
            sample.append("第%d行「%s」" % (w.get("r", 0), str(w.get("value", ""))[:12]))
        more = "" if len(rw) <= 3 else " 等"
        log.append({"level": "warn", "op": "结果列待核", "n": len(rw),
                    "msg": "结果列有 %d 个格子不是程序认识的写法（P/F/合格/不合格等），"
                           "已按未判定处理：%s%s 请人工核对"
                           % (len(rw), "、".join(sample), more)})
    if not log:
        log.append({"level": "info", "op": "结论", "msg": "未发现异常"})
    return log


def _with_fail_prefix(out_path):
    """在文件名最前加 不合格_（已有则不加）。"""
    return _with_prefix(out_path, "不合格_")


def _fail_prefix(fails, norm):
    """根据不合格原因返回文件名前缀（最前面）：

    - 只要有“数据写出但超出上下限”（``cause == 'oob'``）或“数据写出但区间内被标 F”
      （``cause == 'unknown'``）→ ``数据不合格_``；
    - 仅因“数据列没写数据、写了 /”（``cause == 'missing'``）而判不合格 → ``不合格_``；
    - 两者都有（部分行超差、部分行缺数据）→ 仍是 ``数据不合格_``；
    - 都没有 → ``''``（合格，不加前缀）。

    上下限空白自动填 ``/`` 不算不合格，不加前缀（见 ``slash_cells``，不在此处理）。
    """
    has_data_fail = False
    has_missing = False
    for f in (fails or []):
        c = f.get("cause")
        if c is None:
            c = "unknown"   # 防御：未分类视作“数据已写”，避免漏标不合格
        if c == "missing":
            has_missing = True
        else:
            has_data_fail = True
    if has_data_fail:
        return "数据不合格_"
    if has_missing:
        return "不合格_"
    return ""


def _with_prefix(out_path, prefix):
    """在文件名最前加前缀（已有则不加）。不合格_ / 图号错误_ 共用一套。"""
    d, f = os.path.split(out_path)
    if f.startswith(prefix):
        return out_path
    return os.path.join(d, prefix + f)


def _alt_path(out_path, n):
    """目标被占用时的备选名：name(2).ext / name(3).ext ...（Windows 惯例）。"""
    d, f = os.path.split(out_path)
    base, ext = os.path.splitext(f)
    return os.path.join(d, f"{base}({n}){ext}")


def _unique_target(path, max_try=999, exclude=None):
    """返回第一个未被占用的路径（只算名字，不创建文件）。

    exclude: 允许返回该路径本身（即使已存在）——图号错误改名时原输出文件
    本来就在盘上，不能被当成"占用了"。

    实现说明：原来逐个候选名做 os.path.exists，目录里堆了 N 个同名文件就要
    N 次 stat（实测批量跑同一文件到同一目录时，单份为此花掉 37ms）。改为
    一次 os.listdir 拿目录快照，之后在 set 里查，开销与已占数量无关。
    """
    d, f = os.path.split(path)
    base, ext = os.path.splitext(f)
    try:
        names = set(os.listdir(d or "."))
    except Exception:  # noqa: BLE001
        names = None     # 目录读不到（不存在/无权限）→ 退回逐个探测的老路径

    if names is not None:
        if f not in names:
            return path
        if exclude is not None and os.path.abspath(path) == os.path.abspath(exclude):
            return path
        for n in range(1, max_try + 1):
            cand_name = "%s(%d)%s" % (base, n, ext)
            cand = os.path.join(d, cand_name) if d else cand_name
            if cand_name not in names:
                return cand
            if exclude is not None and os.path.abspath(cand) == os.path.abspath(exclude):
                return cand
        raise RuntimeError("同名文件过多（已尝试 %d 次）: %s" % (max_try, path))

    if not os.path.exists(path):
        return path
    if exclude is not None and os.path.abspath(path) == os.path.abspath(exclude):
        return path
    for n in range(1, max_try + 1):
        cand = _alt_path(path, n)
        if not os.path.exists(cand):
            return cand
        if exclude is not None and os.path.abspath(cand) == os.path.abspath(exclude):
            return cand
    raise RuntimeError("同名文件过多（已尝试 %d 次）: %s" % (max_try, path))


def _write_output(write_fn, first_target, max_try=20):
    """写输出：目标已存在（含源文件本身）时不覆盖，自动换 name(1)、name(2)... 后写入。

    用户要求：不管合不合格都不覆盖原文件，只输出到指定目录；同名文件在名字后加 (1) 区分。
    write_fn(target) 成功返回 target；全部尝试失败抛最后一次异常；
    若 20 个候选名全部被占用（没有任何一次真正尝试写入），抛 RuntimeError。
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
            # A1.7/T2：SaveAs 抛异常（如 Excel RPC 断连）可能已在磁盘留下损坏/半成品，
            # 立即清理，避免残留文件干扰重试后的命名（自动 (1) 自增）与两侧对照。
            try:
                if os.path.exists(target):
                    os.remove(target)
            except Exception:
                pass
            last_err = e
    if last_err is None:
        # 20 个候选名全被占用（每次都是 os.path.exists 命中后 continue），
        # 此时 last_err 仍是 None，直接 raise 会抛 "exceptions must derive from BaseException"。
        raise RuntimeError("无法生成不冲突的输出名（已尝试 %d 次）: %s"
                           % (max_try, first_target))
    raise last_err


# ============================================================
# 主流程
# ============================================================

def _col_width(ws, idx):
    """取第 idx 列（1-based）的列宽；没有显式定义返回 None。"""
    try:
        from openpyxl.utils import get_column_letter
    except Exception:  # noqa: BLE001
        return None
    try:
        d = ws.column_dimensions.get(get_column_letter(idx))
        if d is not None and d.width:
            return d.width
    except Exception:  # noqa: BLE001
        pass
    try:
        for d in ws.column_dimensions.values():
            lo, hi = getattr(d, "min", None), getattr(d, "max", None)
            if lo and hi and lo <= idx <= hi and d.width:
                return d.width
    except Exception:  # noqa: BLE001
        pass
    return None


def _copy_column_widths(ws, src_col, dst_first, dst_n):
    """给新接出来的列统一套上第一块「末尾读数列」的宽度。

    只抄单元格样式不抄列宽，新列会退回默认 8.43，和左边 8.25 的列并排就显出一宽
    一窄（客户看到的「每列宽度和前面不一样」）。这里统一取末列宽度，接缝处无跳变。
    """
    try:
        from openpyxl.utils import get_column_letter
    except Exception:  # noqa: BLE001
        return
    w = _col_width(ws, src_col)
    if not w:
        return
    for i in range(dst_n):
        try:
            ws.column_dimensions[get_column_letter(dst_first + i)].width = w
        except Exception:  # noqa: BLE001
            pass


def _trim_far_cells(ws, max_col):
    """删掉 max_col 之外那些「没有值、只带样式」的单元格。

    Excel 生成的 xlsx 常把格式铺到第 16384 列，openpyxl 会为它们建对象；
    留着的话，之后每次读取都要按 16384 列 × 行数建网格。
    """
    try:
        dead = [k for k, c in ws._cells.items()
                if k[1] > max_col and c.value is None]
        for k in dead:
            del ws._cells[k]
    except Exception:  # noqa: BLE001
        pass


def _normalize_column_dims(ws, max_col):
    """把 `A`、`J..IV`、`IV..XFD` 这种区间列宽定义拆成逐列定义，且不超过 max_col。

    源文件里 `<col min=256 max=16384 width=13.5>` 会让新接出来的列全部变成
    13.5 宽（和左边不一致），也让工作表看起来有 16384 列。
    """
    try:
        from openpyxl.utils import get_column_letter
    except Exception:  # noqa: BLE001
        return
    widths = {}
    try:
        for d in list(ws.column_dimensions.values()):
            lo, hi = getattr(d, "min", None), getattr(d, "max", None)
            if not lo or not hi or not d.width:
                continue
            for i in range(max(1, lo), min(hi, max_col) + 1):
                widths[i] = d.width
    except Exception:  # noqa: BLE001
        return
    try:
        ws.column_dimensions.clear()
    except Exception:  # noqa: BLE001
        return
    for i in range(1, max_col + 1):
        w = widths.get(i)
        if not w:
            continue
        try:
            ws.column_dimensions[get_column_letter(i)].width = w
        except Exception:  # noqa: BLE001
            pass


def _clamp_merges(ws, max_col):
    """把跨到 max_col 以右的合并区截断到 max_col。

    源表标题/脚注常写成「A1:XFD1」这种跨满 16384 列的合并区，直接留着会让
    工作表一直背着 3 万多个空合并单元格，读取和保存都变慢，看起来也怪。
    """
    try:
        for rng in list(ws.merged_cells.ranges):
            if rng.max_col <= max_col:
                continue
            try:
                ws.unmerge_cells(str(rng))
            except Exception:  # noqa: BLE001
                continue
            if rng.min_col <= max_col:
                try:
                    ws.merge_cells(start_row=rng.min_row, start_column=rng.min_col,
                                   end_row=rng.max_row, end_column=max_col)
                except Exception:  # noqa: BLE001
                    pass
    except Exception:  # noqa: BLE001
        pass


def _find_number_runs(ws, c0=6):
    """找出所有「连续件号表头」行：[(row, [件号...]), ...]。"""
    max_c = _used_bounds(ws)[1]

    def _int_at(r, c):
        v = ws.cell(row=r, column=c).value
        if isinstance(v, bool):
            return None
        if isinstance(v, (int, float)) and float(v) == int(v):
            return int(v)
        if isinstance(v, str) and v.strip().isdigit():
            return int(v.strip())
        return None

    runs = []
    for r in range(1, ws.max_row + 1):
        nums = []
        for c in range(c0, max_c + 1):
            n = _int_at(r, c)
            if n is None:
                break
            nums.append(n)
        if len(nums) >= 3 and all(nums[i + 1] == nums[i] + 1
                                 for i in range(len(nums) - 1)):
            runs.append((r, nums))
    return runs


def _repair_wide_inplace(wb):
    """已经是「一块到底」的大表（单块读数 >= 251 列）→ 就地统一后半段格式。

    早期版本合并时只搬了单元格样式，没抄列宽，也没处理源表第 251 列以后
    本来就没画边框的情况，客户看到的就是「后半段没有格子线、宽窄和前面不一样」。
    这里按前 250 列循环套用同样的样式，列宽统一取末个读数列，纯修格式不动数值。

    直接在已加载的工作簿上改，避免为了修格式再整簿加载一次（普通文件零开销）。
    返回处理的列数；0 表示不适用。
    """
    from copy import copy as _copy
    if wb is None or not wb.worksheets:
        return 0
    ws = wb.worksheets[0]
    if ws.sheet_state != "visible":
        return 0
    runs = _find_number_runs(ws)
    if len(runs) != 1:
        return 0                 # 还有两块 → 交给合并流程
    h1, nums = runs[0]
    n = len(nums)
    if n < 251:
        return 0                 # 普通表，不动
    c0 = 6
    span = 250                   # .xls 一排放得下的读数个数
    last_read = c0 + n - 1
    last_col = last_read
    result_col = None
    for rr in (h1 - 1, h1):
        t = str(ws.cell(row=rr, column=last_read + 1).value or "")
        if "Result" in t or "P合格" in t:
            result_col = last_read + 1
            break
    if result_col:
        last_col = result_col

    max_row = _used_bounds(ws)[0]
    width = _col_width(ws, c0 + span - 1)
    try:
        from openpyxl.utils import get_column_letter
    except Exception:  # noqa: BLE001
        get_column_letter = None
    cols = 0
    for c in range(c0 + span, last_col + 1):
        if c == result_col:
            src_c = c0 + span          # 第一块的结果列
        else:
            src_c = c0 + span - 1      # 第一块最后一个读数格，接缝两侧一致
        if width and get_column_letter is not None:
            try:
                ws.column_dimensions[get_column_letter(c)].width = width
            except Exception:  # noqa: BLE001
                pass
        for r in range(1, max_row + 1):
            t = ws.cell(row=r, column=c)
            s = ws.cell(row=r, column=src_c)
            try:
                t._style = _copy(s._style)
            except Exception:  # noqa: BLE001
                pass
        cols += 1
    _clamp_merges(ws, last_col)
    _trim_far_cells(ws, last_col)
    _normalize_column_dims(ws, last_col)
    _dbg("repair_wide_inplace ok 读数=%d 处理列=%d 结果列=%s" % (n, cols, result_col))
    return cols


def _xlsx_has_formula_quick(path):
    """zip 级快速判断：任一工作表 XML 里是否含公式标签。

    只用来决定「要不要尝试 openpyxl 快路径」。判断方向是保守的：
    找到 <f> 就返回 True（跳过快路径，直接走 COM，结果一定正确）；
    解析失败也返回 True。永远不会把「有公式」误判成「没公式」。
    """
    try:
        import zipfile
    except Exception:  # noqa: BLE001
        return True
    try:
        zf = zipfile.ZipFile(path)
    except Exception:  # noqa: BLE001
        return True
    try:
        names = zf.namelist()
        for name in names:
            if not name.startswith("xl/worksheets/"):
                continue
            if not name.endswith(".xml"):
                continue
            try:
                data = zf.read(name)
            except Exception:  # noqa: BLE001
                continue
            if re.search(rb"<f[ >/]", data):
                return True
        return False
    except Exception:  # noqa: BLE001
        return True
    finally:
        try:
            zf.close()
        except Exception:
            pass


def _merge_continuation_blocks(path, wb=None):
    """把「上下两块」合并成一块，返回新文件路径；不需要合并时返回 None。

    为什么会有两块：.xls 最多 256 列，件数超过约 250 时，制表人只能把剩下的件
    折到下面再起一个表头（同一张表里上下两个块）。公司不接受这种分块，而且
    审核只看得到上面一块 —— 下面那块等于漏判。合并后是一张连续的表（.xlsx 能放
    16384 列）。

    wb：外部已加载的工作簿。传入就不重复加载（单份能省一次 openpyxl 解析），
    且本函数不负责关闭它；不传则内部加载并关闭。

    判定很严（宁可不合并，也绝不误改普通表）：
      1. 某行出现 >=3 个「连续整数」，认作件号表头；
      2. 第二块的第一个件号必须正好等于第一块最后一个件号 +1；
      3. 第二块行按 B 列项目名与第一块顺序对齐，逐行搬到右侧；
      4. 第一块的结果列（Result / P合格 F不合格）整体右移，腾出的位置放第二块读数。
    """
    from copy import copy as _copy
    own = wb is None
    if own:
        try:
            from openpyxl import load_workbook as _lwb
            wb = _lwb(path)
        except Exception:  # noqa: BLE001
            return None
    try:
        if not wb.worksheets:
            return None
        ws = wb.worksheets[0]
        C0 = 6   # 读数起始列

        def _int_at(r, c):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, bool):
                return None
            if isinstance(v, (int, float)) and float(v) == int(v):
                return int(v)
            if isinstance(v, str) and v.strip().isdigit():
                return int(v.strip())
            return None

        def _label(r):
            b = ws.cell(row=r, column=2).value
            if b is not None and str(b).strip():
                return str(b).strip()
            a = ws.cell(row=r, column=1).value
            return str(a).strip() if a is not None else ""

        max_c = _used_bounds(ws)[1]
        runs = []
        for r in range(1, ws.max_row + 1):
            nums = []
            for c in range(C0, max_c + 1):
                n = _int_at(r, c)
                if n is None:
                    break
                nums.append(n)
            if len(nums) >= 3 and all(nums[i + 1] == nums[i] + 1
                                     for i in range(len(nums) - 1)):
                runs.append((r, nums))
        if len(runs) < 2:
            return None

        (h1, nums1) = runs[0]
        cont = None
        for (h2, nums2) in runs[1:]:
            if nums2[0] == nums1[-1] + 1:
                cont = (h2, nums2)
                break
        if cont is None:
            return None
        (h2, nums2) = cont

        n1 = len(nums1)
        n2 = len(nums2)
        last_c1 = C0 + n1 - 1

        # 第一块结果列：表头写着 Result / (P合格，紧跟最后一个读数
        result_col = None
        for rr in (h1 - 1, h1):
            for cc in (last_c1 + 1, last_c1 + 2):
                t = str(ws.cell(row=rr, column=cc).value or "")
                if "Result" in t or "P合格" in t:
                    result_col = cc
                    break
            if result_col:
                break
        shift = n2 if result_col else 0

        rows1 = [r for r in range(h1 + 1, h2) if _label(r)]
        rows2 = [r for r in range(h2 + 1, ws.max_row + 1) if _label(r)]
        used, pairs = set(), []
        for r2 in rows2:
            for r1 in rows1:
                if r1 in used:
                    continue
                if _label(r1) == _label(r2):
                    used.add(r1)
                    pairs.append((r2, r1))
                    break
        # 一行都对不上：多半不是「同一张表的续块」，不动
        if not pairs:
            return None

        # 1) 结果列整体右移
        if result_col:
            for r in range(1, h2):
                s = ws.cell(row=r, column=result_col)
                if s.value is not None or s.has_style:
                    t = ws.cell(row=r, column=result_col + shift, value=s.value)
                    try:
                        t._style = _copy(s._style)
                    except Exception:  # noqa: BLE001
                        pass
                    s.value = None

        # 2) 第二块件号表头接到右侧，格式照抄第一块最后一列读数
        #    （照抄「最后一列」而不是循环对应前面的列：接缝两侧才会完全一致）
        new_col = last_c1 + 1
        for i in range(n2):
            t = ws.cell(row=h1, column=new_col + i, value=nums2[i])
            try:
                t._style = _copy(ws.cell(row=h1, column=last_c1)._style)
            except Exception:  # noqa: BLE001
                pass

        # 3) 第二块数据搬到对应行的右侧
        #    样式一律取「第一块同一行、最后一个读数格」：源表第二块经常没画边框、
        #    字号也不统一，直接搬过去就是客户看到的「线没加上、宽窄不一致」。
        for r2, r1 in pairs:
            src_style = None
            try:
                src_style = _copy(ws.cell(row=r1, column=last_c1)._style)
            except Exception:  # noqa: BLE001
                src_style = None
            for i in range(n2):
                s = ws.cell(row=r2, column=C0 + i)
                t = ws.cell(row=r1, column=new_col + i, value=s.value)
                if src_style is None:
                    continue
                try:
                    t._style = _copy(src_style)
                except Exception:  # noqa: BLE001
                    pass

        # 3b) 数据区之外的行（外观/材质/图号/脚注等）也要补上边框，
        #     否则第二块右侧留空的那几行看起来是断的。
        #     只补到「最后一个有内容的行」，避免在末尾凭空多出空行。
        last_body = h1
        for r in range(h1, ws.max_row + 1):
            if any(ws.cell(row=r, column=c).value not in (None, "")
                   for c in range(1, 6)):
                last_body = r
        for r in range(h1, last_body + 1):
            src = ws.cell(row=r, column=last_c1)
            for i in range(n2):
                t = ws.cell(row=r, column=new_col + i)
                if t.value is not None:
                    continue
                try:
                    t._style = _copy(src._style)
                except Exception:  # noqa: BLE001
                    pass

        # 4) 删掉第二块（表头及其上方的空行一起删），第一块脚注保留
        del_from = max(h1 + 1, h2 - 2)
        ws.delete_rows(del_from, ws.max_row - del_from + 1)

        # 5) 列宽照抄第一块末尾读数那一列（接缝处看不出宽窄跳变）；
        #    合并区截到最后一列、清掉 16384 列那种空样式单元，
        #    否则后续每次读取都要按 16384 列建网格（巨慢）。
        last_col = new_col + n2 - 1 + shift
        _copy_column_widths(ws, last_c1, new_col, n2)
        _clamp_merges(ws, last_col)
        _trim_far_cells(ws, last_col)
        _normalize_column_dims(ws, last_col)

        out = _make_tmp("_merged_", ".xlsx")
        wb.save(out)
        _dbg("merge_blocks ok src=%s -> %s 块1件数=%d 块2件数=%d 对齐行=%d 结果列=%s"
             % (os.path.basename(path), os.path.basename(out), n1, n2,
                len(pairs), result_col))
        return out
    except Exception as e:  # noqa: BLE001
        _dbg("merge_blocks fail %r" % (e,))
        return None
    finally:
        if own:
            try:
                wb.close()
            except Exception:
                pass


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
    if on_progress:
        try:
            on_progress(2, "读取源文件")
        except Exception:  # noqa: BLE001
            pass
    _dbg("audit_file entry file=%r ext=%s out=%s" % (os.path.basename(input_path), ext, od))

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

    # ---- 上下两块合并（.xls 256 列上限 → 制表人把余下件数折到下面） ----
    # 必须在审核之前合并：否则只审到第一块，第二块整块漏判，总数还会被改成第一块的数量。
    # 这里先加载一次工作簿，合并判断用掉之后，下面的布局规划继续复用同一个对象，
    # 不再二次解析（单份省约 26ms，批量 100 份就是 2.6 秒）。
    merge_tmp = None
    _pre_wb = None
    if ext.lower() in (".xlsx", ".xlsm") and openpyxl is not None:
        try:
            _pre_wb = _load_openpyxl(work_path, use_cache=False)
        except Exception:  # noqa: BLE001
            _pre_wb = None
    try:
        merge_tmp = _merge_continuation_blocks(work_path, wb=_pre_wb)
        if merge_tmp:
            work_path = merge_tmp
            conv_msgs.append("检测到上下两块（.xls 256 列上限导致），已合并为一块连续表")
    except Exception:  # noqa: BLE001
        merge_tmp = None
    if merge_tmp and _pre_wb is not None:
        # 合并产出的是新文件，旧对象没用了（也不用它读新文件）
        try:
            _pre_wb.close()
        except Exception:
            pass
        _pre_wb = None

    # ---- v11.3 宽表拆分预处理 ----
    # zip 级预判有没有公式：有公式就必然走 COM，没必要再让 openpyxl 把整表
    # 遍历一遍再发现"哦有公式"。这一步只读 XML 找 <f>，几毫秒。
    formula_quick = False
    if ext.lower() in (".xlsx", ".xlsm"):
        try:
            formula_quick = _xlsx_has_formula_quick(work_path)
        except Exception:  # noqa: BLE001
            formula_quick = False
    split_tmp = None
    split_made = []
    shift_made = []
    fill_counts = None
    reuse_wb = None
    pending_shifts = None
    # A1.7/T1：plans 只在 xlsx/xlsm 分支内赋值；老 .xls(OLE2) 转换失败等路径
    # 进不了该分支，此处兜底初始化，供下方 _dbg 与后续引用，避免 UnboundLocalError。
    plans = []
    if conv_msgs and on_progress:
        try:
            on_progress(0, conv_msgs[-1])
        except Exception:
            pass
    if ext.lower() in (".xlsx", ".xlsm"):
        plans, shifts = [], []
        wbp = _pre_wb          # 复用合并判断时已加载的对象，避免二次解析
        pre = pre or {}
        if pre.get("plans") is not None:
            # 主进程已预检：直接复用布局计划（快路径候选仍需 wbp）
            plans = pre.get("plans") or []
            shifts = pre.get("shifts") or []
            fill_counts = pre.get("fill_counts")
            if not plans and not shifts and wbp is None:
                try:
                    wbp = _load_openpyxl(input_path, use_cache=False)
                except Exception:
                    wbp = None
        else:
            if wbp is None:
                try:
                    wbp = _load_openpyxl(input_path, use_cache=False)
                except Exception:
                    wbp = None
            try:
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

    if merge_tmp:
        # 合并/改格式后行列结构变了：合并前预检出的 wb / 拆分计划 / 归位计划都不能再用
        reuse_wb = None
        plans = []
        pending_shifts = None
        split_made = []

    _dbg("audit_file plans=%s pending_shifts=%s work_path==input=%s"
         % (len(plans), len(pending_shifts) if pending_shifts else 0,
            os.path.abspath(work_path) == input_path))

    # ---- A1.7/T2：文件级 RPC 自愈重试 ----
    # Excel 进程在批量连续处理时偶发退出（0x800706BA RPC 服务器不可用 / 0x80010108 对象断开
    # / 0x800706BE RPC 调用失败），死对象继续复用会让后续文件整批失败。命中时丢弃死会话
    # 并重建，整份文件重试一次；重试前清掉本轮已写出的半成品输出，不留残文件。
    # 二次仍失败抛中文业务异常。判定口径不变，仅容错增强。
    _rpc_hrs = (0x800706BA, 0x80010108, 0x800706BE, 0x80010105)
    _made_out = {"path": None}

    def _run_once():
        nonlocal sess, own_sess
        # ---- openpyxl 快路径（无公式 + 无复杂元素；有删空列归位时跳过，走 COM） ----
        if ext.lower() in (".xlsx", ".xlsm") and not pending_shifts and not formula_quick:
            wb = reuse_wb if reuse_wb is not None else _load_openpyxl(work_path)
            _dbg("audit_file FAST attempt wb_is_none=%s" % (wb is None,))
            if wb is not None:
                # 已合并过的大表：就地统一第 251 列以后的边框/列宽。
                # 普通表直接返回 0，没有额外开销（不再为判断而整簿预加载）。
                try:
                    wide_fixed = _repair_wide_inplace(wb)
                except Exception:  # noqa: BLE001
                    wide_fixed = 0
                r = _audit_openpyxl(work_path, wb=wb, on_progress=on_progress)
                if r is not None:
                    _dbg("audit_file FAST ok")
                    fails, reason, norm = r
                    fails, reason, norm = _v11_apply_result_f(fails, reason, norm)
                    norm["split_made"] = split_made
                    norm["shift_made"] = shift_made
                    if wide_fixed:
                        norm["wide_fixed"] = wide_fixed
                    is_fail = len(fails) > 0
                    prefix = _fail_prefix(fails, norm)
                    out_path = os.path.join(od, base + ext)
                    if prefix:
                        out_path = _with_prefix(out_path, prefix)
                    if _dwg_bad(norm, base):
                        out_path = _with_dwg_prefix(out_path)
                    marked = copied = False
                    if on_progress:
                        on_progress(90, "写出结果文件")
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
                        _made_out["path"] = output
                    finally:
                        try:
                            wb.close()
                        except Exception:
                            pass
                    if on_progress:
                        label = "数据不合格" if prefix == "数据不合格_" else ("不合格" if prefix else "合格")
                        on_progress(100, f"{label} -> {os.path.basename(output)}"
                                    + ("（已标红标F）" if marked else ""))
                    return {"input": input_path, "output": output, "is_fail": is_fail,
                            "reason": reason, "copied": copied, "marked": marked,
                            "norm": _norm_summary(norm),
                            "scan_errors": _scan_errors_summary(norm),
                            "log": _build_log(fails, reason, norm, base)
                                  + ([{"level": "info", "op": "格式修复", "msg": m} for m in conv_msgs] if conv_msgs else [])}
                # 快路径不适用（公式/复杂元素）→ 降级 COM，关闭已加载的 wb
                _dbg("audit_file FAST -> COM degrade")
                try:
                    wb.close()
                except Exception:
                    pass

        # ---- COM 慢路径（有公式 / .xls / 复杂元素 / 已拆分） ----
        _dbg("audit_file COM path")
        own = sess is None
        if sess is None:
            sess = ExcelSession()
            own_sess = True
        try:
            xl = sess.get()
            fails, reason, norm, output, marked, copied = _audit_com_file(
                xl, work_path, od, base, ext, fill_counts, pending_shifts,
                on_saved=lambda _o: _made_out.update(path=_o))
            fails, reason, norm = _v11_apply_result_f(fails, reason, norm)
            norm["split_made"] = split_made
            norm["shift_made"] = norm.get("shift_made") or shift_made
            is_fail = len(fails) > 0
            prefix = _fail_prefix(fails, norm)
            # 图号错误 → 输出文件加 图号错误_ 前缀（不覆盖已写文件，改名）
            if _dwg_bad(norm, base):
                newout = _unique_target(_with_dwg_prefix(output), exclude=output)
                if os.path.abspath(newout) != os.path.abspath(output):
                    shutil.move(output, newout)
                    output = newout
            _made_out["path"] = output
        finally:
            if own:
                sess.quit()

        if on_progress:
            label = "数据不合格" if prefix == "数据不合格_" else ("不合格" if prefix else "合格")
            on_progress(100, f"{label} -> {os.path.basename(output)}"
                            + ("（已标红标F）" if marked else ""))
        return {"input": input_path, "output": output, "is_fail": is_fail,
                "reason": reason, "copied": copied, "marked": marked,
                "norm": _norm_summary(norm),
                "scan_errors": _scan_errors_summary(norm),
                "log": _build_log(fails, reason, norm, base)
                      + ([{"level": "info", "op": "格式修复", "msg": m} for m in conv_msgs] if conv_msgs else [])}

    try:
        for _attempt in (1, 2):
            try:
                _dbg("=== audit_file start: %s attempt=%s ===" % (os.path.basename(input_path), _attempt))
                return _run_once()
            except Exception as _e:
                _hr = getattr(_e, "hr", None)
                if _hr is None:
                    _a = getattr(_e, "args", ()) or ()
                    if _a and isinstance(_a[0], (int, float)):
                        _hr = int(_a[0])
                import traceback as _tb
                _dbg("audit_file EXCEPTION attempt=%s hr=%r e=%r" % (_attempt, _hr, _e))
                _dbg("traceback:\n" + _tb.format_exc())
                if _hr not in _rpc_hrs or _attempt == 2:
                    raise
                _dbg("audit_file RPC retry attempt=%s hr=%s" % (_attempt, _hr))
                if _made_out["path"]:
                    try:
                        os.remove(_made_out["path"])
                    except Exception:
                        pass
                    _made_out["path"] = None
                if sess is not None:
                    try:
                        sess.quit()
                    except Exception:
                        pass
                if session is None:
                    sess = ExcelSession()
                else:
                    sess._xl = None
                    sess.xl_pid = None
        raise RuntimeError("Excel 进程异常退出，已自动重启并重试仍失败")
    finally:
        if split_tmp and os.path.exists(split_tmp):
            try:
                os.remove(split_tmp)
            except Exception:
                pass
        if merge_tmp and os.path.exists(merge_tmp):
            try:
                os.remove(merge_tmp)
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
                    b = os.path.basename(f)
                    if b.startswith("不合格_") or b.startswith("数据不合格_"):
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
    target = _unique_target(os.path.join(dst_dir, os.path.basename(src)))
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

    _emit_lock = threading.Lock()

    def emit(obj):
        # 心跳线程和主线程都会写 stdout，必须串行化，否则两行会交叉成坏 JSON
        with _emit_lock:
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

    # 心跳：每 5 秒报一次「还活着」，父进程据此区分「在跑」和「卡死」。
    # 为什么必须有心跳：audit_file 的进度回调很稀（本机跑一份 346 件的老 .xls，
    # 中间会静默 7.5 秒；客户四核 i5 上更久）。只靠业务 phase 判断，
    # 慢活会被当成超时杀掉重试 —— 大文件永远跑不完。
    _hb_stop = threading.Event()

    def _heartbeat():
        while not _hb_stop.wait(5.0):
            try:
                emit({"event": "alive"})
            except Exception:  # noqa: BLE001
                return

    threading.Thread(target=_heartbeat, daemon=True).start()

    _reported_pid = [excel_pid]

    def _sync_pid():
        """Excel 在 worker 内部可能被重建（探活失败即丢弃重建），新 PID 要回报父进程。
        否则超时强杀会打在已死的旧 PID 上，新实例变成杀不掉的孤儿进程。"""
        p = getattr(sess, "xl_pid", None)
        if p and p != _reported_pid[0]:
            _reported_pid[0] = p
            emit({"event": "pid", "excel_pid": p})

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
                _last = {"pct": 0}

                def cb(pct, msg, _f=f, _st=_last):
                    # 只许前进：合并/转换阶段会回退到 0，界面看着像重头开始
                    try:
                        p = max(_st["pct"], min(100, int(pct)))
                    except Exception:  # noqa: BLE001
                        p = _st["pct"]
                    _st["pct"] = p
                    emit({"event": "phase", "file": _f, "pct": p, "msg": msg})

                res = audit_file(f, out_dir, on_progress=cb, session=sess,
                                 pre=task.get("pre"))
                emit({"event": "done", "file": f, "ok": True, "result": res})
            except Exception as e:  # noqa: BLE001
                # A1.7/T4：错误带完整 traceback（_dbg 落日志 + done.trace 截断 2000 字符），
                # 避免 exe 侧只能看到一句话、定位成本高。
                _dbg("worker done error %s: %r\n%s"
                     % (os.path.basename(f), e, traceback.format_exc()))
                emit({"event": "done", "file": f, "ok": False, "error": str(e),
                      "trace": traceback.format_exc()[-2000:]})
            _sync_pid()
    except Exception:  # noqa: BLE001  —— 主循环整体兜底：绝不向 windowed 进程外抛异常（防弹窗）
        pass
    finally:
        _hb_stop.set()
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


def _excel_pids():
    """列出当前所有 EXCEL.EXE 的 PID（纯 Win32 API）。"""
    import ctypes
    from ctypes import wintypes

    class _PE(ctypes.Structure):
        _fields_ = [
            ("dwSize", wintypes.DWORD), ("cntUsage", wintypes.DWORD),
            ("th32ProcessID", wintypes.DWORD),
            ("th32DefaultHeapID", ctypes.POINTER(ctypes.c_ulong)),
            ("th32ModuleID", wintypes.DWORD), ("cntThreads", wintypes.DWORD),
            ("th32ParentProcessID", wintypes.DWORD),
            ("pcPriClassBase", ctypes.c_long), ("dwFlags", wintypes.DWORD),
            ("szExeFile", ctypes.c_char * 260),
        ]

    out = []
    try:
        k32 = ctypes.WinDLL("kernel32", use_last_error=True)
        k32.CreateToolhelp32Snapshot.restype = wintypes.HANDLE
        k32.CreateToolhelp32Snapshot.argtypes = [wintypes.DWORD, wintypes.DWORD]
        k32.Process32First.argtypes = [wintypes.HANDLE, ctypes.POINTER(_PE)]
        k32.Process32Next.argtypes = [wintypes.HANDLE, ctypes.POINTER(_PE)]
        snap = k32.CreateToolhelp32Snapshot(0x00000002, 0)
        if not snap or snap == ctypes.c_void_p(-1).value:
            return out
        e = _PE()
        e.dwSize = ctypes.sizeof(_PE)
        ok = k32.Process32First(snap, ctypes.byref(e))
        while ok:
            if e.szExeFile.lower() == b"excel.exe":
                out.append(e.th32ProcessID)
            ok = k32.Process32Next(snap, ctypes.byref(e))
        k32.CloseHandle(snap)
    except Exception:
        pass
    return out


def _visible_excel_pids():
    """有可见主窗口的 Excel（用户自己开的），这些绝不动。"""
    out = set()
    try:
        import win32gui
        import win32process

        def _cb(h, _):
            try:
                if win32gui.IsWindowVisible(h) and win32gui.GetClassName(h) == "XLMAIN":
                    out.add(win32process.GetWindowThreadProcessId(h)[1])
            except Exception:
                pass
        win32gui.EnumWindows(_cb, None)
    except Exception:
        pass
    return out


def sweep_orphan_excel(before):
    """清掉本批次新出现、且没有可见窗口的 EXCEL（= 工具自己起的 COM 实例）。

    Excel 经 DCOM 拉起，不在 worker 的进程树里，`taskkill /T` 收不到它；
    ready 未到时父进程也不知道它的 PID。唯一可靠的判据就是「批次前不存在的
    无窗口 EXCEL」。用户自己打开的有窗口实例会被跳过。
    """
    vis = _visible_excel_pids()
    n = 0
    for pid in _excel_pids():
        if pid in before or pid in vis:
            continue
        try:
            subprocess.run(["taskkill", "/F", "/T", "/PID", str(pid)],
                           capture_output=True, creationflags=0x08000000)
            n += 1
        except Exception:
            pass
    return n


# 并行下限（份）：少于该份数一律串行。
# 定这个值的依据是本机 24 核 + 4 核客户机的双端实测：
#   1-5 份：并行 0.87x-0.96x（更慢），6 份起开始稳定为正收益。
# 改这个值前先跑 bench_parallel.py A 1 3 5 6 8 12，别凭感觉调。
PARALLEL_MIN_FILES = 6


class BatchPool:
    """多进程 worker 池：每个 worker 一个独立 Excel，单文件超时杀进程重试。

    - 并发上限 2，且文件数 >= PARALLEL_MIN_FILES(6) 且核数 >= 4 时才真并行；
      否则退化为单 worker 串行（小批量并行反而更慢，见常量处的实测数据）。
      max_workers 参数保留给测试脚本，GUI 一律默认 2。
    - 文件数 < PARALLEL_MIN_FILES 或机器核数 < 4 时自动退化为串行（不开子进程）。
    - **超时按「无进展」算，不按总耗时**：worker 每 5s 发一次心跳（alive），
      外加业务 phase 事件，父进程收到任意一条就刷新计时。因此大文件在慢机器上
      跑 10 分钟也不会被误杀，只有「连续 timeout 秒一条事件都没有」才算卡死。
      另设 hard_cap 作单文件绝对上限，防止「有心跳但永远做不完」。
    - 单文件卡死（默认 30s 无任何事件）→ 终止该 worker（含其 Excel 进程）→
      重启 worker → 该文件重试；最多重试 max_retry 次（默认 1，重试窗口放宽 1.5 倍），
      仍失败标记"需人工处理"，其余文件照常，绝不整批中断。
    - 每个 worker 输出到私有临时目录，完成后主进程统一 move 到最终目录并
      同名递增命名（跨进程无竞态、不覆盖、不留半成品）。
    - 进度以"每份文件最终完成"为权威事件，逐份 +1，单调不跳。

    run(files, out_dir, on_event=None) 阻塞至全部完成，返回 results（对齐
    audit_file 返回结构，处理失败项含 error）。
    on_event(e) 事件：{"type":"progress","done":d,"total":t}
                     {"type":"done","file":f,"ok":bool,"result":dict}
                     {"type":"retry","file":f,"attempt":a,"max":m,"why":str}
    """

    def __init__(self, max_workers=2, timeout=30.0, max_retry=1,
                 min_files_for_parallel=PARALLEL_MIN_FILES, hard_cap=1800.0):
        # WPS环境实测并行2路比串行快1.6倍；保留RPC自愈和超时重建。
        self.max_workers = max_workers or 2
        # timeout 语义：多久「没有任何事件」算卡死（不是单文件总耗时上限）
        self.timeout = float(timeout)
        self.hard_cap = float(hard_cap)   # 单文件绝对上限（秒）
        self.max_retry = int(max_retry)
        # 份数太少时并行反而更慢：worker 进程 + Excel 冷启动约 1.5-2s，
        # 而每份处理只要 1.5s 左右，2 份的总收益抵不过多起一套实例的开销。
        self.min_files_for_parallel = int(min_files_for_parallel)
        self._eff_workers = max(1, self.max_workers)   # 本次实际并发数（run 内计算）
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
        # 源码模式下 worker 是 `python -m audit_fail`，模块必须在搜索路径里。
        # 调用方 cwd 不是程序目录时（快捷方式工作目录设错、被别的程序调起），
        # 子进程会直接退出、等不到 ready，整批看起来像卡死。这里显式给 cwd+PYTHONPATH。
        _cwd = None
        _env = None
        if not getattr(sys, "frozen", False):
            _here = os.path.dirname(os.path.abspath(__file__))
            _cwd = _here
            _env = dict(os.environ)
            _old = _env.get("PYTHONPATH") or ""
            _env["PYTHONPATH"] = _here + ((os.pathsep + _old) if _old else "")
        proc = subprocess.Popen(
            self._worker_cmd(tmp),
            stdin=subprocess.PIPE, stdout=subprocess.PIPE,
            stderr=subprocess.DEVNULL, encoding="utf-8", errors="replace",
            creationflags=_flags, cwd=_cwd, env=_env,
        )
        w = {"id": wid, "proc": proc, "tmp_out": tmp,
             "busy": None, "excel_pid": None, "dead": False, "ready": False,
             "spawn_at": time.time()}
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
        # 1) 已知 PID 的 Excel 单独收一次
        if w["excel_pid"]:
            try:
                subprocess.run(
                    ["taskkill", "/F", "/T", "/PID", str(w["excel_pid"])],
                    capture_output=True, creationflags=0x08000000)
            except Exception:
                pass
        # 2) 再按「整棵进程树」收 worker：Excel 是 worker 拉起来的子进程，
        #    但 ready 事件还没到时父进程并不知道它的 PID —— 只杀 python 会留下孤儿 Excel。
        proc = w.get("proc")
        if proc is not None and proc.poll() is None:
            try:
                subprocess.run(
                    ["taskkill", "/F", "/T", "/PID", str(proc.pid)],
                    capture_output=True, creationflags=0x08000000)
            except Exception:
                pass
            try:
                proc.kill()
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

    def cancel(self):
        """外部要求中止（用户关窗口）：杀掉所有 worker 及其 Excel，让 run 尽快返回。

        不这么做的话，关窗口时 QThread 还在跑、子进程和 Excel 都活着，
        进程退出时没人收尸，留一堆无窗口 EXCEL。
        """
        self._cancel = True
        self._pending.clear()
        for w in list(self.workers):
            try:
                self._kill(w)
            except Exception:  # noqa: BLE001
                pass
        sweep_orphan_excel(self._excel_before)

    def run(self, files, out_dir, on_event=None, pres=None):
        # 小批量或低核数机器不并行：并行收益来自"两份同时算"，
        # 核数不足时两个 Excel 会与系统争抢，反而更慢。
        files = list(files)
        # 统一成绝对路径：worker 子进程的 cwd 固定为程序目录，
        # 相对路径过去会解析错地方
        _abs = []
        for _f in files:
            _abs.append(os.path.abspath(_f))
        files = _abs
        eff = self.max_workers
        if len(files) < self.min_files_for_parallel:
            eff = 1
        elif (os.cpu_count() or 1) < 4:
            eff = 1
        # 只在本次运行内决定并发数，不改 self.max_workers：
        # 否则同一个 pool 实例第二次 run 会继承上一次的 1，永久退化为串行。
        self._eff_workers = max(1, eff)

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
        self._keep_tmp = set()      # 产物搬出失败时保留临时目录，不删
        self._cancel = False
        # 批次前快照：这些 Excel 不许动（用户自己开的 / 其它软件后台的）
        self._excel_before = set(_excel_pids())
        self._ready_fail = 0        # worker 起不来（无 Office/COM 挂起）的连续次数
        # ready 等不到就会整批卡住（进度 0%、按钮不恢复），所以给个上限：
        # Excel 冷启动最慢 10s 上下，留到 30s 已很宽松。
        ready_timeout = max(30.0, float(self.timeout))

        # 预启动全部 worker（Excel 冷启动不占单文件超时）
        for _ in range(self._eff_workers):
            self._spawn()

        def _busy():
            return any(w["busy"] is not None for w in self.workers)

        try:
            while (self._pending or _busy()) and not self._cancel:
                # 1) 事件
                try:
                    while True:
                        kind, wid, obj = self.events.get_nowait()
                        w = self.workers[wid]
                        if kind == "eof":
                            # 已被 _requeue/_kill 主动杀掉的 worker：reader 线程
                            # 随后补的这条 eof 是回声，再 spawn 一次会让并发数
                            # 超过阈值（每次重试多养一个 Excel）。
                            if w["dead"]:
                                continue
                            if w["busy"]:
                                self._requeue(w, "worker 异常退出", results)
                            else:
                                self._cleanup_worker(w)
                                self._spawn()
                            continue
                        # 任何来自该 worker 的事件（含 5s 一次的 alive 心跳）都算
                        # 「还在动」，刷新卡死计时。这就是大文件不会被误杀的保证。
                        if w["busy"] is not None:
                            w["busy"]["last"] = time.time()
                        if obj.get("event") == "pid":
                            # worker 内部重建了 Excel（原实例崩溃）：更新父进程
                            # 记录的 PID，否则超时强杀会打在已死进程上，新 Excel 成孤儿。
                            if obj.get("excel_pid"):
                                w["excel_pid"] = obj["excel_pid"]
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
                            # 把「当前文件内部的进度」转给界面。以前这里直接 continue，
                            # 池只在整份文件结束时发一次 progress，于是单份文件时进度条
                            # 从 0 直接跳 100，看起来像"没跑就完成了"。
                            self._on_event({"type": "phase",
                                            "file": obj.get("file"),
                                            "pct": obj.get("pct"),
                                            "msg": obj.get("msg")})
                            continue
                        if obj.get("event") == "done":
                            self._on_done(w, obj, results, out_dir)
                            continue
                except queue.Empty:
                    pass

                # 2) 超时检查
                now = time.time()
                for w in list(self.workers):
                    if not w["busy"] or w["dead"]:
                        continue
                    _f = w["busy"]["file"]
                    # 重试时窗口放宽：第一次「超时」常常只是机器当时在忙
                    _factor = 1.0 + 0.5 * self._attempts.get(_f, 0)
                    _idle_limit = self.timeout * _factor
                    _idle = now - w["busy"].get("last", w["busy"]["start"])
                    _elapsed = now - w["busy"]["start"]
                    if _idle > _idle_limit:
                        self._requeue(w, "处理超时（%.0fs 无任何进展）" % _idle_limit,
                                      results)
                    elif _elapsed > self.hard_cap:
                        self._requeue(w, "超过单文件上限（%.0fs）" % self.hard_cap,
                                      results)

                # 2b) ready 超时：worker 卡在建 Excel（没装 Office、COM 弹窗、实例挂起）
                # 时父进程永远等不到 ready，整批会 0% 卡死。这里杀掉重建，
                # 连续 3 次仍起不来就判定环境不可用，把剩余文件一次性标失败。
                for w in list(self.workers):
                    if (not w["ready"]) and (not w["dead"]) \
                            and (now - w["spawn_at"]) > ready_timeout:
                        self._kill(w)
                        self._ready_fail += 1
                        if self._ready_fail >= 3:
                            for f in list(self._pending):
                                results[f] = {"file": f, "ok": False, "fail": False,
                                              "output": None, "error":
                                              "Excel 无法启动（可能未安装 Excel / COM 被占用），"
                                              "已放弃剩余 %d 个文件" % len(self._pending),
                                              "reason": "", "marked": False,
                                              "norm": {}, "log": []}
                                self._on_event({"type": "done", "file": f, "ok": False,
                                                "result": results[f]})
                            self._pending.clear()
                            self._on_event({"type": "progress", "done": len(results),
                                            "total": self._total})
                        else:
                            self._spawn()

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
                # 产物搬出失败的临时目录要留着（那里面有唯一一份文件）
                if w["tmp_out"] not in self._keep_tmp:
                    shutil.rmtree(w["tmp_out"], ignore_errors=True)
            # 兜底：批次结束后不该再有任何「本批次新起、又没有窗口」的 Excel
            sweep_orphan_excel(self._excel_before)

        return [results[f] for f in files if f in results]

    def _dispatch(self, w):
        f = self._pending.pop(0)
        _now = time.time()
        w["busy"] = {"file": f, "start": _now, "last": _now}
        self._send(w, {"file": f, "pre": self._pres.get(f)})

    def _on_done(self, w, obj, results, out_dir):
        f = obj.get("file")
        if not f:
            return
        if obj.get("ok"):
            res = obj.get("result") or {}
            src = res.get("output")
            # 落盘目录：没指定输出目录时落到「这个文件自己所在的目录」。
            # 以前条件写成 `... and out_dir`，out_dir 为 None 时直接跳过搬迁，
            # 产物留在 worker 的临时目录里、收尾时被 rmtree 删掉，界面却显示成功。
            dst_dir = out_dir or os.path.dirname(os.path.abspath(f))
            move_err = None
            if not src or not os.path.isfile(src) or os.path.getsize(src) <= 0:
                move_err = "产物未落盘或为空文件"
            else:
                for _try in range(3):
                    try:
                        src = _move_output(src, dst_dir)
                        move_err = None
                        break
                    except Exception as e:  # noqa: BLE001
                        move_err = str(e)
                        time.sleep(0.3)
            if move_err:
                # 搬不出去就别删临时目录：那里是唯一一份产物，删了就真丢了
                self._keep_tmp.add(w["tmp_out"])
                _dbg("move output failed %s -> %s: %s" % (src, dst_dir, move_err))
                results[f] = {"file": f, "ok": False, "fail": False, "output": src,
                              "error": "产物搬出失败（%s），已保留在临时目录：%s"
                                       % (move_err, w["tmp_out"]),
                              "reason": "", "marked": False, "norm": {}, "log": []}
                self._on_event({"type": "done", "file": f, "ok": False,
                                "result": results[f]})
                self._on_event({"type": "progress", "done": len(results),
                                "total": self._total})
                w["busy"] = None
                if not w["dead"] and self._pending:
                    self._dispatch(w)
                return
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
            # 这个 worker 已被杀；若还有待处理的文件，必须补一个顶上，
            # 否则 eff=1 时没有活着的 worker 可派发，while 循环永远出不来（整批挂死）。
            if self._pending:
                self._spawn()

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
