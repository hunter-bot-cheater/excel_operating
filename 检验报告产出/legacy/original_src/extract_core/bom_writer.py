# -*- coding: utf-8 -*-
"""BOM 回写：把「日期/件数」写入 BOM 的「第N套」列 × 对应图号行。

规则（用户需求）：
  - 每成功转换一个报告文件，就在 BOM 中「第N套」列与该图号所在行的交叉单元格
    写入 `日期/件数`（日期 = 基准文件名里的日期，件数 = 本次抽取件数）。
  - 「第N套」列不存在时，主动插入一列：位置在所有「第X套」列最右、且
    「图纸/材质报告/尺寸报告」等无关列最左之间（见 bom_reader.find_insert_col），
    列名写为「第N套」。
  - 原地修改用户拖入的 BOM 母版；**不做备份**（用户明确要求：BOM 不需要 .bak）。
  - **多表单只认第一张有效数据表**：BOM 含多个工作表时，匹配、建任务、回写一律以
    第一张「表头有效 + 有组件数量列 + 有数据行」的表为准，其余表不参与、也不同步回写。
  - N 由调用方传入（GUI 必填）。

健壮性设计（真机 Excel 自动化调用偶发失败是常态，必须自愈而不是一次失败就全废）：
  - **自带独立 Excel 实例**：不与任务抽取共用，任务级 Excel 出错重建 / Quit 都不影响回写；
  - 每次 COM 调用带**重试**（RPC 忙/被拒绝等瞬时错误：RPC_E_CALL_REJECTED 等）；
  - 写入硬失败 -> **自动重开 BOM 并回放已写入的值**（pending），不丢数据、不重复插列
    （重开后按 live 表头重新定位列）；
  - 「第N套」列定位改为**读取 live 工作簿的表头行**，不再依赖可能过期的 openpyxl 快照；
  - 只有真正 `Save()` 成功才置 `saved=True`，调用方据此显示「已回写 / 未保存」，
    不再出现「显示已回写、其实没落盘」的误导。
"""
import os
import time

from .bom_reader import _IRRELEVANT_KWS, canon, parse_round_header
from .errors import ExtractError

_XL_TO_RIGHT = -4161  # Excel Shift:=xlToRight
# 可重试的瞬时 COM 错误（Excel 忙 / 调用被拒绝 / 服务器稍后重试 / 线程 COM 未初始化）
_TRANSIENT = {
    -2147418111,  # RPC_E_CALL_REJECTED
    -2147417846,  # RPC_E_SERVERCALL_RETRYLATER
    -2147417851,  # RPC_E_SERVERCALL_RETRYLATER(alt)
    -2147221008,  # CO_E_NOTINITIALIZED
    0x80010001,   # CALL_E_REJECTED_BY_CALLEE
}


def _fmt(e) -> str:
    s = str(e).replace("\r", " ").replace("\n", " ")
    return s[:300]


def _is_transient(e) -> bool:
    hr = getattr(e, "hresult", None)
    if hr is None:
        try:
            hr = e.args[0]
        except Exception:  # noqa: BLE001
            hr = None
    try:
        return int(hr) in _TRANSIENT
    except Exception:  # noqa: BLE001
        return False


def _insert_pos(hdrs, meta, n: int) -> int:
    """计算「第n套」列插入位置：所有「第X套」列最右、无关列（图纸等）最左。"""
    rounds = [c for c, h in hdrs.items() if parse_round_header(h) is not None]
    others = [c for c, h in hdrs.items()
              if any(k in canon(h) for k in _IRRELEVANT_KWS)]
    if rounds:
        return max(rounds) + 1
    if others:
        return min(others)
    if meta.qty_col:
        return meta.qty_col + 1
    return (max(hdrs) + 1) if hdrs else 1


class BomWriter:
    """在一次批次内保持 BOM 打开，逐任务写入，批次结束统一保存。

    用法：
        w = BomWriter(bom_path, round_no, meta)
        w.write(sheet_name, excel_row, date, qty)   # 每次成功任务
        w.close()                                   # 结束后保存并关闭（失败抛 BOM_WRITE_ERROR）

    自带独立 Excel 实例，与任务抽取的 Excel 完全隔离。
    """

    def __init__(self, bom_path, round_no, meta, xl_app=None):
        self.bom_path = os.path.abspath(bom_path)
        self.round_no = int(round_no)
        self.meta = meta
        # 允许复用调用方的 Excel 实例（与任务抽取同一个），避免再起第二个实例导致
        # 两个 Excel 并发、收尾时 Workbooks.Close/Quit 互相拖累卡十几秒。
        self._ext_xl = xl_app
        self._owns_xl = xl_app is None   # 只有自己创建的实例，close 时才负责 Quit
        self.xl = None if xl_app is None else xl_app
        self.wb = None
        self.backup_path = None     # 已废弃：不再生成 .bak，恒为 None（保留属性兼容旧诊断脚本）
        self.dirty = False
        self.saved = False          # 是否真正 Save() 成功
        self.pending = []           # [(sheet, row, text)] 已写入的值（用于重开后回放）
        self._cols = {}             # 工作表名 -> 「第N套」列号
        self._row_cache = {}        # (canon(bh), canon(desc)) -> [(sheet, row), ...] 主表定位缓存

    # ---- 生命周期 ----
    def _ensure_open(self):
        if self.wb is not None:
            return
        # 延迟 import 放在函数开头：create_excel_app / quit_excel_app 必须
        # 在本函数作用域内始终绑定（此前在 if 块内 import，复用调用方实例时
        # 走到重建分支会 UnboundLocalError）
        from .extractor import create_excel_app, quit_excel_app
        # 不做备份（用户明确要求）：直接以可写方式打开母版
        # 复用调用方传入的 Excel 实例时不另起；否则自己起一个
        if self.xl is None:
            try:
                self.xl = create_excel_app()
                self._owns_xl = True
            except Exception as e:  # noqa: BLE001
                raise ExtractError("BOM_WRITE_ERROR", f"Excel 启动失败（BOM 回写）：{_fmt(e)}",
                                   file=self.bom_path)
        # 2) 打开（可写）；失败则重建实例再试一次
        try:
            self.wb = self._call(lambda: self.xl.Workbooks.Open(self.bom_path, UpdateLinks=0))
        except Exception as e1:  # noqa: BLE001
            try:
                try:
                    quit_excel_app(self.xl)
                except Exception:  # noqa: BLE001
                    pass
                self.xl = create_excel_app()
                self.wb = self.xl.Workbooks.Open(self.bom_path, UpdateLinks=0)
            except Exception as e2:  # noqa: BLE001
                raise ExtractError(
                    "BOM_WRITE_ERROR",
                    f"打开 BOM 写入失败（可能正被 Excel 占用/只读）：{_fmt(e2)}"
                    f"（首次尝试：{_fmt(e1)}）",
                    file=self.bom_path,
                )

    def _quit(self):
        # 仅自己创建的实例才 Quit；复用调用方实例时交给调用方统一关
        if self.xl is not None and self._owns_xl:
            from .extractor import quit_excel_app
            try:
                quit_excel_app(self.xl)
            except Exception:  # noqa: BLE001
                pass
        self.xl = None

    def _call(self, fn, tries=4, delay=0.35):
        """执行 COM 调用；瞬时错误（Excel 忙）自动重试。"""
        last = None
        for i in range(tries):
            try:
                return fn()
            except Exception as e:  # noqa: BLE001
                last = e
                if not _is_transient(e) or i == tries - 1:
                    raise
                time.sleep(delay * (i + 1))
        raise last  # pragma: no cover

    # ---- 列定位 ----
    def _live_headers(self, sheet_name):
        """读 live 工作簿表头行，返回 {列号: 表头文本}（不依赖 openpyxl 快照）。"""
        ws = self.wb.Worksheets(sheet_name)
        hdr = int(self.meta.header_row or 1)
        end = (max(self.meta.headers) if self.meta.headers else 40) + 30
        vals = self._call(lambda: ws.Range(ws.Cells(hdr, 1), ws.Cells(hdr, end)).Value)
        if vals is None:
            return {}
        if not isinstance(vals[0], (tuple, list)):
            vals = (vals,)
        out = {}
        for i, v in enumerate(vals[0]):
            if v is None:
                continue
            t = str(v).strip()
            if t:
                out[i + 1] = t
        return out

    def _resolve_col(self, sheet_name: str) -> int:
        """定位（必要时插入）「第N套」列，返回 1-based 列号。按表缓存。"""
        if sheet_name in self._cols:
            return self._cols[sheet_name]
        try:
            self.wb.Worksheets(sheet_name)
        except Exception as e:  # noqa: BLE001
            raise ExtractError(
                "BOM_WRITE_ERROR", f"找不到 BOM 工作表「{sheet_name}」：{_fmt(e)}",
                file=self.bom_path,
            )

        hdrs = self._live_headers(sheet_name)
        col = None
        for c, h in hdrs.items():
            if parse_round_header(h) == self.round_no:
                col = c
                break
        if col is None:
            col = _insert_pos(hdrs, self.meta, self.round_no)
            ws = self.wb.Worksheets(sheet_name)
            try:
                self._call(lambda: ws.Columns(col).Insert(Shift=_XL_TO_RIGHT))
            except Exception as e:  # noqa: BLE001
                raise ExtractError(
                    "BOM_WRITE_ERROR",
                    f"插入「第{self.round_no}套」列失败（第 {col} 列）：{_fmt(e)}",
                    file=self.bom_path,
                )
            try:
                ws.Cells(int(self.meta.header_row or 1), col).Value = f"第{self.round_no}套"
            except Exception:  # noqa: BLE001
                pass
        self._cols[sheet_name] = col
        return col

    # ---- 写入 ----
    def _set_cell(self, sname, excel_row, text):
        self._ensure_open()
        col = self._resolve_col(sname)
        cell = self.wb.Worksheets(sname).Cells(excel_row, col)
        try:
            cell.NumberFormat = "@"   # 强制文本，避免被 Excel 误判为日期
        except Exception:  # noqa: BLE001
            pass
        cell.Value = text

    def _reopen(self):
        """重开 BOM（不保存）并回放已写入的值，用于硬失败自愈。"""
        wb, self.wb = self.wb, None
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:  # noqa: BLE001
                pass
        self._cols = {}
        self._ensure_open()
        replay = list(self.pending)
        self.pending = []
        for (s, r, t) in replay:
            self._set_cell(s, r, t)
            self.pending.append((s, r, t))

    def write(self, sheet_name: str, excel_row: int, date: str, qty: int):
        """把 `日期/件数` 写入 (工作表, 行, 第N套列)。失败抛 BOM_WRITE_ERROR。"""
        if not excel_row or excel_row <= 0:
            raise ExtractError("BOM_WRITE_ERROR", "无法定位 BOM 行号，跳过回写",
                               file=self.bom_path)
        text = f"{date}/{qty}"
        sname = sheet_name or self.meta.sheet
        try:
            self._call(lambda: self._set_cell(sname, excel_row, text))
        except Exception as e1:  # noqa: BLE001
            # 硬失败：重开 BOM + 回放，再试一次（覆盖「Excel 实例已失效 / 工作簿被关」等）
            try:
                self._reopen()
                self._call(lambda: self._set_cell(sname, excel_row, text))
            except Exception as e2:  # noqa: BLE001
                raise ExtractError(
                    "BOM_WRITE_ERROR",
                    f"写入 BOM 失败（{sname}!{excel_row}）：{_fmt(e2)}"
                    f"（首次错误：{_fmt(e1)}）",
                    file=self.bom_path,
                )
        self.pending.append((sname, excel_row, text))
        self.dirty = True

    # ---- 多表定位 + 写入（BOM 含多张数据表时一律以第一张有效表为准）----
    def write_row(self, bom_row, date, qty):
        """把 `日期/件数` 写入该图号行 —— **只写主表（第一张有效数据表）精确行**。

        BOM 含多个表单时一律以第一张有效数据表为准（匹配、建任务、回写都是它），
        其它表不再同步；同表内多行（合法重复，含完全相同行）各自写各自的精确行。
        """
        text = f"{date}/{qty}"
        self._ensure_open()
        key = (bom_row.sheet, bom_row.row, canon(bom_row.bh), canon(bom_row.desc))
        targets = self._row_cache.get(key)
        if targets is None:
            targets = [(bom_row.sheet, bom_row.row)]   # 主表精确行（唯一真相源）
            self._row_cache[key] = targets

        err = None
        for (sname, erow) in targets:
            try:
                self._call(lambda s=sname, e=erow: self._set_cell(s, e, text))
            except Exception as e1:  # noqa: BLE001
                try:
                    self._reopen()
                    self._call(lambda s=sname, e=erow: self._set_cell(s, e, text))
                except Exception as e2:  # noqa: BLE001
                    err = ExtractError(
                        "BOM_WRITE_ERROR",
                        f"写入 BOM 失败（{sname}!{erow}）：{_fmt(e2)}"
                        f"（首次错误：{_fmt(e1)}）",
                        file=self.bom_path,
                    )
                    break
            self.pending.append((sname, erow, text))
        self.dirty = True
        if err:
            raise err

    def close(self):
        """保存并关闭 BOM；dirty 且最终保存失败 -> 抛 BOM_WRITE_ERROR。"""
        if self.wb is None:
            self._quit()
            return
        err = None
        if self.dirty:
            try:
                self._call(lambda: self.wb.Save(), tries=3)
                self.saved = True
            except Exception as e1:  # noqa: BLE001
                err = _fmt(e1)
                # 自愈一次：重开 + 回放 + 再存
                try:
                    self._reopen()
                    self._call(lambda: self.wb.Save(), tries=3)
                    self.saved = True
                    err = None
                except Exception as e2:  # noqa: BLE001
                    err = f"{_fmt(e2)}（首次错误：{err}）"
        try:
            self.wb.Close(SaveChanges=False)
        except Exception:  # noqa: BLE001
            pass
        self.wb = None
        self._quit()
        if err:
            raise ExtractError(
                "BOM_WRITE_ERROR",
                f"保存 BOM 失败（本次回写内容未落盘）：{err}",
                file=self.bom_path,
            )
