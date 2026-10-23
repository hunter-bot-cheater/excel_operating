#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
特殊格式文件审核 —— PySide6 GUI
================================
仿照 D:\\文件转换不改格式 (audit_fail_gui.py):
  - 拖入即预检(openpyxl 快读)
  - 多 worker 子进程池(独立 Excel 实例、单文件超时杀进程重试、干完一份+1 实时进度)
  - 结果列表实时插入、单击查看详情、失败置顶
  - 转换核心在 converter.py
"""
import os as _os
import sys as _sys

_exe_dir = _os.path.dirname(_os.path.abspath(_sys.executable)) if getattr(_sys, "frozen", False) \
    else _os.path.dirname(_os.path.abspath(__file__))
_MEIPASS = getattr(_sys, "_MEIPASS", None)


def _load_version():
    for c in ([_os.path.join(_exe_dir, "VERSION")]
              + ([_os.path.join(_MEIPASS, "VERSION")] if _MEIPASS else [])):
        try:
            with open(c, "r", encoding="utf-8") as f:
                v = f.read().strip()
                if v:
                    return v
        except Exception:
            pass
    return "1"


VERSION = _load_version()
APP_TITLE = f"特殊格式文件审核 v{VERSION}"

import os
import sys
import subprocess

from PySide6.QtCore import Qt, QThread, Signal, QPropertyAnimation, QEasingCurve
from PySide6.QtGui import QDragEnterEvent, QDropEvent, QFont, QIcon, QColor
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QProgressBar, QFileDialog, QMessageBox, QListWidget,
    QListWidgetItem, QFrame, QAbstractItemView, QStyle,
    QGraphicsDropShadowEffect, QMenu, QDialog, QPlainTextEdit, QScrollArea,
)

import converter

C_PRIMARY = "#2C3E50"
C_ACCENT = "#3498DB"
C_ACCENT_HOVER = "#2980B9"
C_ACCENT_PRESS = "#21618C"
C_BG = "#F5F7FA"
C_CARD = "#FFFFFF"
C_TEXT = "#2C3E50"
C_TEXT_2 = "#7F8C8D"
C_BORDER = "#DCE3EA"
C_PROG_A = "#3498DB"
C_PROG_B = "#2ECC71"
C_OK = "#27AE60"
C_ERR = "#E74C3C"

SUPPORTED = (".xls", ".xlsx", ".xlsm")
FILTER_STR = "Excel 文件 (*.xls *.xlsx *.xlsm)"


def _icon(pm, size=18):
    return QApplication.style().standardIcon(pm).pixmap(size, size)


class AuditWorker(QThread):
    progress_changed = Signal(int, str)
    status_msg = Signal(str)
    file_done = Signal(object)
    all_finished = Signal(list)

    def __init__(self, files, out_dir, pres=None):
        super().__init__()
        self.files = files
        self.out_dir = out_dir
        self.pres = pres or {}

    def run(self):
        results = []
        total = len(self.files)

        def on_event(e):
            if e["type"] == "progress":
                done = e["done"]
                pct = int(done / total * 100) if total else 100
                self.progress_changed.emit(pct, f"[{done}/{total}]")
            elif e["type"] == "retry":
                self.status_msg.emit(
                    f"[重试 {e['attempt']}/{e['max']}] {os.path.basename(e['file'])}：{e['why']}")
            elif e["type"] == "done":
                r = e.get("result") or {}
                if e.get("ok"):
                    tag = "不合格" if r.get("fail") else "合格"
                    self.status_msg.emit(f"[完成] {os.path.basename(e['file'])}：{tag}")
                else:
                    self.status_msg.emit(
                        f"[失败] {os.path.basename(e['file'])}："
                        f"{str(r.get('error') or '未知错误')[:60]}")
                self.file_done.emit(r)

        try:
            if self.files:
                n_workers = 3 if (os.cpu_count() or 4) >= 8 else 2
                pool = converter.BatchPool(max_workers=n_workers, timeout=30.0, max_retry=1)
                results = pool.run(self.files, self.out_dir, on_event=on_event, pres=self.pres)
        except Exception as e:  # noqa: BLE001
            import traceback
            self.status_msg.emit(f"[严重错误] {e}")
            traceback.print_exc()
        self.progress_changed.emit(100, "全部完成")
        self.all_finished.emit(results)


class PreloadWorker(QThread):
    done = Signal(dict)

    def __init__(self, files, parent=None):
        super().__init__(parent)
        self.files = files

    def run(self):
        out = {}
        for f in self.files:
            try:
                out[f] = converter.precheck(f)
            except Exception:
                out[f] = {"ok": False, "error": "预检失败"}
        self.done.emit(out)


class DetailDialog(QDialog):
    def __init__(self, title, lines, parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setMinimumSize(620, 440)
        self.resize(720, 480)
        lay = QVBoxLayout(self)
        lay.setSpacing(10)
        txt = QPlainTextEdit()
        txt.setReadOnly(True)
        txt.setPlainText("\n".join(lines))
        txt.setStyleSheet(
            f"QPlainTextEdit{{background:{C_BG};border:1px solid {C_BORDER};"
            f"border-radius:8px;font-size:12px;color:{C_TEXT};padding:10px;}}")
        lay.addWidget(txt)
        row = QHBoxLayout()
        bcopy = QPushButton("复制全部")
        bclose = QPushButton("关闭")
        for b in (bcopy, bclose):
            b.setMinimumHeight(34)
            b.setCursor(Qt.PointingHandCursor)
        bcopy.setStyleSheet(
            "QPushButton{background:#FFFFFF;color:#1F4E79;border:1.5px solid #1F4E79;"
            "border-radius:6px;font-size:13px;font-weight:600;padding:6px 14px;}"
            "QPushButton:hover{background:#EAF2FA;}")
        bclose.setDefault(True)
        bclose.setStyleSheet(
            "QPushButton{background:#1F6FB2;color:#FFFFFF;border:1.5px solid #1F6FB2;"
            "border-radius:6px;font-size:13px;font-weight:700;padding:6px 14px;}"
            "QPushButton:hover{background:#185A92;}")
        bcopy.clicked.connect(lambda: QApplication.clipboard().setText("\n".join(lines)))
        bclose.clicked.connect(self.accept)
        row.addStretch()
        row.addWidget(bcopy)
        row.addWidget(bclose)
        lay.addLayout(row)


class DropFileList(QListWidget):
    files_dropped = Signal(list)

    _STYLE_IDLE = (
        "QListWidget{border:2px dashed #9AA7B5;border-radius:12px;background:#FFFFFF;padding:8px;}"
        "QListWidget::item{padding:6px 10px;margin:2px 4px;border-radius:6px;}"
        "QListWidget::item:hover{background:#EAF4FD;}")
    _STYLE_HOVER = (
        "QListWidget{border:2px dashed #3498DB;border-radius:12px;background:#EAF4FD;padding:8px;}"
        "QListWidget::item{padding:6px 10px;margin:2px 4px;border-radius:6px;}"
        "QListWidget::item:hover{background:#D0EBFF;}")

    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)
        self.setDragDropMode(QAbstractItemView.DropOnly)
        self.setSelectionMode(QAbstractItemView.NoSelection)
        self.setMinimumHeight(140)
        self.setMaximumHeight(200)
        self.setStyleSheet(self._STYLE_IDLE)
        self._hint = None

    def show_hint(self):
        self.clear()
        it = QListWidgetItem("  拖入 .xls / .xlsx / .xlsm 报告（或文件夹），或点击下方按钮选择")
        it.setForeground(Qt.gray)
        it.setFlags(it.flags() & ~Qt.ItemIsEnabled)
        it.setTextAlignment(Qt.AlignCenter)
        self._hint = it
        self.addItem(it)

    def add_path(self, fp):
        if self._hint is not None:
            self.takeItem(self.row(self._hint))
            self._hint = None
        it = QListWidgetItem(os.path.basename(fp))
        it.setToolTip(fp)
        it.setData(Qt.UserRole, fp)
        self.addItem(it)

    def dragEnterEvent(self, e: QDragEnterEvent):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()
            self.setStyleSheet(self._STYLE_HOVER)

    def dragMoveEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()

    def dragLeaveEvent(self, e):
        self.setStyleSheet(self._STYLE_IDLE)
        super().dragLeaveEvent(e)

    def dropEvent(self, e: QDropEvent):
        self.setStyleSheet(self._STYLE_IDLE)
        files = []
        for u in e.mimeData().urls():
            p = u.toLocalFile()
            if os.path.isdir(p):
                for root, _, fs in os.walk(p):
                    for fn in fs:
                        if fn.lower().endswith(SUPPORTED):
                            files.append(os.path.join(root, fn))
            elif p.lower().endswith(SUPPORTED):
                files.append(p)
        if files:
            self.files_dropped.emit(files)
        e.acceptProposedAction()


class Card(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("card")
        self.setStyleSheet(
            f"QFrame#card{{background:{C_CARD};border:1px solid {C_BORDER};border-radius:12px;}}")


class MainApp(QWidget):
    def __init__(self):
        super().__init__()
        self.input_files = []
        self.out_dir = ""
        self.worker = None
        self.pre_map = {}
        self.last_outputs = []
        self._preload_workers = []
        self._build_ui()

    def _build_ui(self):
        self.setWindowTitle(APP_TITLE)
        self.setMinimumSize(720, 620)
        self.resize(780, 800)
        self.setObjectName("MainWindow")
        self.setStyleSheet(
            f"QWidget#MainWindow{{background:{C_BG};}}"
            "QWidget{font-family:'Microsoft YaHei','Segoe UI',sans-serif;color:#2C3E50;}")

        root = QVBoxLayout(self)
        root.setContentsMargins(18, 18, 18, 14)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll.setStyleSheet("QScrollArea{background:transparent;border:none;}")
        card = Card()
        sh = QGraphicsDropShadowEffect(card)
        sh.setBlurRadius(24)
        sh.setColor(QColor(0, 0, 0, 60))
        sh.setOffset(0, 6)
        card.setGraphicsEffect(sh)
        cl = QVBoxLayout(card)
        cl.setSpacing(16)
        cl.setContentsMargins(22, 20, 22, 20)
        scroll.setWidget(card)
        root.addWidget(scroll)

        trow = QHBoxLayout()
        t = QLabel("特殊格式文件审核")
        t.setFont(QFont("Microsoft YaHei", 19, QFont.Weight.Bold))
        t.setStyleSheet(f"color:{C_PRIMARY};")
        trow.addWidget(t)
        trow.addStretch()
        badge = QLabel(f"v{VERSION}")
        badge.setStyleSheet(
            f"background:{C_ACCENT};color:#FFFFFF;font-size:11px;font-weight:bold;"
            "padding:3px 12px;border-radius:10px;")
        trow.addWidget(badge)
        cl.addLayout(trow)

        in_card = Card()
        il = QVBoxLayout(in_card)
        il.setSpacing(10)
        il.setContentsMargins(16, 14, 16, 14)
        h = QLabel("输入文件")
        h.setStyleSheet(f"font-weight:bold;font-size:13px;color:{C_PRIMARY};")
        il.addWidget(h)
        self.file_list = DropFileList()
        self.file_list.files_dropped.connect(self._add_files)
        il.addWidget(self.file_list)
        brow = QHBoxLayout()
        brow.addStretch()
        bsel = QPushButton("选择文件")
        bclr = QPushButton("清空列表")
        for b in (bsel, bclr):
            b.setMinimumHeight(34)
            b.setCursor(Qt.PointingHandCursor)
            b.setStyleSheet(
                f"QPushButton{{background:{C_BG};color:{C_PRIMARY};border:1px solid {C_BORDER};"
                f"font-weight:500;font-size:12px;border-radius:6px;padding:6px 16px;}}"
                f"QPushButton:hover{{background:#EAF4FD;border-color:{C_ACCENT};}}")
        bsel.clicked.connect(self._select_files)
        bclr.clicked.connect(self._clear_files)
        brow.addWidget(bsel)
        brow.addWidget(bclr)
        il.addLayout(brow)
        cl.addWidget(in_card)

        ocard = Card()
        ol = QVBoxLayout(ocard)
        ol.setSpacing(10)
        ol.setContentsMargins(16, 14, 16, 14)
        oh = QLabel("输出目录")
        oh.setStyleSheet(f"font-weight:bold;font-size:13px;color:{C_PRIMARY};")
        ol.addWidget(oh)
        orow = QHBoxLayout()
        self.out_edit = QLabel("默认：源文件所在目录")
        self.out_edit.setStyleSheet(
            f"QLabel{{background:{C_BG};border:1px solid {C_BORDER};border-radius:6px;"
            f"padding:9px 12px;font-size:12px;color:{C_TEXT};}}")
        bdir = QPushButton("选择目录")
        bdir.setMinimumHeight(34)
        bdir.setCursor(Qt.PointingHandCursor)
        bdir.setStyleSheet(
            f"QPushButton{{background:{C_BG};color:{C_PRIMARY};border:1px solid {C_BORDER};"
            f"font-weight:500;font-size:12px;border-radius:6px;padding:6px 16px;}}"
            f"QPushButton:hover{{background:#EAF4FD;border-color:{C_ACCENT};}}")
        bdir.clicked.connect(self._select_out)
        orow.addWidget(self.out_edit, 1)
        orow.addWidget(bdir)
        ol.addLayout(orow)
        cl.addWidget(ocard)

        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.progress.setMinimumHeight(22)
        self.progress.setStyleSheet(
            "QProgressBar{background:%s;border:none;border-radius:11px;font-size:11px;color:%s;text-align:center;}"
            "QProgressBar::chunk{background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 %s,stop:1 %s);border-radius:11px;}"
            % (C_BG, C_TEXT_2, C_PROG_A, C_PROG_B))
        cl.addWidget(self.progress)
        self._anim = QPropertyAnimation(self.progress, b"value")
        self._anim.setDuration(500)
        self._anim.setEasingCurve(QEasingCurve.Type.OutCubic)

        self.status = QLabel("")
        self.status.setWordWrap(True)
        self.status.setStyleSheet(f"font-size:13px;font-weight:bold;color:{C_OK};padding:6px 0;")
        self.status.hide()
        cl.addWidget(self.status)

        rcard = Card()
        rl = QVBoxLayout(rcard)
        rl.setSpacing(10)
        rl.setContentsMargins(16, 14, 16, 14)
        rh = QLabel("处理结果(单击任意行查看详情)")
        rh.setStyleSheet(f"font-weight:bold;font-size:13px;color:{C_PRIMARY};")
        rl.addWidget(rh)
        self.result_list = QListWidget()
        self.result_list.setMinimumHeight(120)
        self.result_list.setMaximumHeight(220)
        self.result_list.setSelectionMode(QAbstractItemView.NoSelection)
        self.result_list.setStyleSheet(
            f"QListWidget{{background:{C_BG};border:1px solid {C_BORDER};border-radius:8px;padding:6px;}}"
            "QListWidget::item{padding:5px 8px;margin:1px 2px;border-radius:5px;}"
            "QListWidget::item:hover{background:#EAF4FD;}")
        self.result_list.itemClicked.connect(self._on_result_clicked)
        rl.addWidget(self.result_list)
        cl.addWidget(rcard)

        brow2 = QHBoxLayout()
        brow2.setSpacing(12)
        self.btn_start = QPushButton("开始审核")
        self.btn_start.setIcon(QIcon(_icon(QStyle.SP_MediaPlay, 18)))
        self.btn_start.setCursor(Qt.PointingHandCursor)
        self.btn_start.setMinimumHeight(46)
        self.btn_start.setStyleSheet(
            f"QPushButton{{background:{C_ACCENT};color:#FFFFFF;font-weight:bold;font-size:14px;"
            f"border:none;border-radius:6px;padding:10px 20px;}}"
            f"QPushButton:hover{{background:{C_ACCENT_HOVER};}}"
            f"QPushButton:pressed{{background:{C_ACCENT_PRESS};}}"
            f"QPushButton:disabled{{background:#AEBECB;color:#F0F4F8;}}")
        self.btn_start.clicked.connect(self._start)

        self.btn_open = QPushButton("打开输出目录")
        self.btn_open.setIcon(QIcon(_icon(QStyle.SP_DirOpenIcon, 18)))
        self.btn_open.setCursor(Qt.PointingHandCursor)
        self.btn_open.setEnabled(False)
        self.btn_open.setMinimumHeight(46)
        self.btn_open.setStyleSheet(
            f"QPushButton{{background:#FFFFFF;color:{C_PRIMARY};border:1px solid {C_BORDER};"
            f"font-weight:bold;font-size:13px;border-radius:6px;padding:10px 20px;}}"
            f"QPushButton:hover{{background:#EAF4FD;border-color:{C_ACCENT};}}"
            f"QPushButton:disabled{{color:#C3CDD6;border-color:#E5EAF0;}}")
        self.btn_open.clicked.connect(self._open_dir)

        brow2.addWidget(self.btn_start, 3)
        brow2.addWidget(self.btn_open, 2)
        cl.addLayout(brow2)

        self.file_list.show_hint()
        self._update_start_state()

    def _select_files(self):
        fs, _ = QFileDialog.getOpenFileNames(self, "选择检验报告", "", FILTER_STR)
        if fs:
            self._add_files(fs)

    def _add_files(self, files):
        new = []
        for f in files:
            fp = os.path.abspath(f)
            if fp not in self.input_files:
                self.input_files.append(fp)
                new.append(fp)
                self.file_list.add_path(fp)
        if new:
            w = PreloadWorker(new)
            self._preload_workers.append(w)
            w.done.connect(self._on_preloaded)
            w.finished.connect(lambda _w=w: self._preload_workers.remove(_w) if _w in self._preload_workers else None)
            w.start()
        self._update_start_state()

    def _on_preloaded(self, out):
        self.pre_map.update(out)

    def _clear_files(self):
        self.input_files = []
        self.file_list.show_hint()
        self._update_start_state()

    def _select_out(self):
        d = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if d:
            self.out_dir = d
            self.out_edit.setText(d)

    def _update_start_state(self):
        busy = self.worker is not None and self.worker.isRunning()
        self.btn_start.setEnabled(bool(self.input_files) and not busy)

    def _start(self):
        if not self.input_files:
            QMessageBox.warning(self, "提示", "请先选择要审核的文件。")
            return
        self.btn_start.setEnabled(False)
        self.btn_start.setText("审核中…")
        self.progress.setValue(0)
        self.status.hide()
        self.result_list.clear()
        self.worker = AuditWorker(list(self.input_files), self.out_dir or None, self.pre_map)
        self.worker.progress_changed.connect(self._on_progress)
        self.worker.status_msg.connect(self._on_status)
        self.worker.file_done.connect(self._append_row)
        self.worker.all_finished.connect(self._on_finished)
        self.worker.start()

    def _on_progress(self, pct, msg):
        self._anim.stop()
        self._anim.setStartValue(self.progress.value())
        self._anim.setEndValue(pct)
        self._anim.start()

    def _on_status(self, msg):
        self.status.setStyleSheet(f"font-size:13px;color:{C_TEXT_2};padding:6px 0;")
        self.status.setText(msg)
        self.status.show()

    def _append_row(self, r):
        if not r.get("ok"):
            it = QListWidgetItem(
                f"[处理失败] {os.path.basename(r['file'])}"
                + (f"（{str(r.get('error',''))[:40]}）" if r.get("error") else ""))
            it.setForeground(QColor(C_ERR))
        else:
            tag = "不合格" if r.get("fail") else "合格"
            color = C_ERR if r.get("fail") else C_OK
            it = QListWidgetItem(f"[{tag}] {os.path.basename(r['file'])}")
            it.setForeground(QColor(color))
        it.setData(Qt.UserRole, r)
        self.result_list.addItem(it)

    def _on_finished(self, results):
        self.last_outputs = [
            r.get("output") for r in results
            if r.get("ok") and r.get("output") and os.path.isfile(r["output"])]
        self.btn_start.setText("开始审核")
        self.btn_start.setIcon(QIcon(_icon(QStyle.SP_MediaPlay, 18)))
        self._update_start_state()
        self.btn_open.setEnabled(bool(self.last_outputs))

        failed = [r for r in results if r.get("ok") and r.get("fail")]
        passed = [r for r in results if r.get("ok") and not r.get("fail")]
        errored = [r for r in results if not r.get("ok")]
        parts = []
        if failed:
            parts.append(f"不合格 {len(failed)} 个")
        if passed:
            parts.append(f"合格 {len(passed)} 个")
        if errored:
            parts.append(f"失败 {len(errored)} 个")
        self.status.setText("全部完成：" + "，".join(parts) if parts else "完成")
        self.status.setStyleSheet(
            f"font-size:13px;font-weight:bold;color:{C_ERR if failed or errored else C_OK};padding:6px 0;")
        self.status.show()

        def key(r):
            return 0 if not r.get("ok") else (1 if r.get("fail") else 2)
        self.result_list.clear()
        for r in sorted(results, key=key):
            self._append_row(r)

        self.input_files = []
        self.file_list.show_hint()

    def _on_result_clicked(self, item):
        r = item.data(Qt.UserRole)
        if not r:
            return
        lines = ["文件: %s" % os.path.basename(r["file"]),
                 "判定: %s" % ("不合格" if r.get("fail") else "合格")]
        if r.get("output"):
            lines.append("输出: %s" % r["output"])
        lines.append("-" * 40)
        log = r.get("log") or []
        if log:
            for l in log:
                m = {"info": "·", "warn": "▲", "err": "✘"}.get(l.get("level"), "·")
                lines.append("%s %s：%s" % (m, l.get("op", ""), l.get("msg", "")))
        else:
            lines.append("(无操作记录)")
        DetailDialog("详情 - %s" % os.path.basename(r["file"]), lines, self).exec()

    def _open_dir(self):
        target = (self.out_dir
                  or (os.path.dirname(self.last_outputs[0]) if self.last_outputs else None)
                  or (os.path.dirname(self.input_files[0]) if self.input_files else None))
        if target and os.path.isdir(target):
            try:
                os.startfile(target)
            except Exception:
                pass


if __name__ == "__main__":
    if "--worker-mode" in sys.argv:
        _out = None
        if "--out" in sys.argv:
            _i = sys.argv.index("--out")
            if _i + 1 < len(sys.argv):
                _out = sys.argv[_i + 1]
        import converter
        sys.exit(converter.run_worker_mode(_out))

    # 全局异常钩子: 窗口化 exe 无控制台, 崩溃写入 exe 同目录 crash.log
    def _excepthook(typ, val, tb):
        import traceback, datetime
        try:
            log = os.path.join(_exe_dir, "crash.log")
            with open(log, "a", encoding="utf-8") as f:
                f.write("=== %s ===\n" % datetime.datetime.now())
                traceback.print_exception(typ, val, tb, file=f)
        except Exception:
            pass
        sys.__excepthook__(typ, val, tb)
    sys.excepthook = _excepthook

    for a in ("AA_EnableHighDpiScaling", "AA_UseHighDpiPixmaps"):
        if hasattr(Qt, a):
            QApplication.setAttribute(getattr(Qt, a), True)
    app = QApplication(sys.argv)
    win = MainApp()
    win.show()
    sys.exit(app.exec())
