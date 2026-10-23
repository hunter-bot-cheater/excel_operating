#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
提取文件数据 —— 原生桌面 GUI (PySide6)
======================================
界面与「审核不合格数据」「检验报告格式转换工具」保持一致（卡片式布局、可拖拽、
进度动画、底部按钮）。本次重构：

    - 内容置于 QScrollArea，避免控件重叠/被裁剪
    - ① BOM 文件（单文件）② 报告文件池（文件夹/多文件，按图号匹配 BOM 行）
      ③ 输出目录 ④ 进度条+状态 ⑤ 结果列表（逐任务实时） ⑥ 底部按钮
    - 流程（报告池驱动）：报告文件图号在 BOM 中匹配到的才处理；
      匹配不到的报告单独报 NO_BOM_MATCH 失败；BOM 其余图号静默跳过；
      智能调度来源文件、逐任务实时显示；单任务失败隔离、来源文件事务回滚

核心逻辑见 extract_core 包（前后端分离，网页端可复用 api.run_batch）。
"""
import os
import sys
import subprocess
import glob

from PySide6.QtCore import Qt, QThread, Signal, QPropertyAnimation, QEasingCurve
from PySide6.QtGui import (
    QDragEnterEvent, QDropEvent, QFont, QIcon, QColor, QIntValidator, QPalette,
)
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QProgressBar, QFileDialog, QMessageBox, QListWidget,
    QListWidgetItem, QFrame, QAbstractItemView, QStyle, QLineEdit,
    QGraphicsDropShadowEffect, QFileIconProvider, QMenu, QScrollArea,
    QDialog, QPlainTextEdit,
)

import extract_core
from extract_core.api import run_batch

# --------------------------------------------------------------------------- #
# 版本管理
# --------------------------------------------------------------------------- #
_exe_dir = os.path.dirname(os.path.abspath(sys.executable)) if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__))
_MEIPASS = getattr(sys, "_MEIPASS", None)


def _load_version():
    for _c in [_exe_dir, _MEIPASS]:
        if not _c:
            continue
        try:
            with open(os.path.join(_c, "VERSION"), "r", encoding="utf-8") as _vf:
                _v = _vf.read().strip()
                if _v:
                    return _v
        except Exception:
            pass
    return "5"


VERSION = _load_version()
APP_TITLE = f"提取文件数据 v{VERSION}"

# --------------------------------------------------------------------------- #
# 配色（与转换/审核工具一致）
# --------------------------------------------------------------------------- #
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
C_WARN = "#E67E22"

SUPPORTED_EXTENSIONS = (".xls", ".xlsx", ".xlsm")
FILTER_STR = "Excel 文件 (*.xls *.xlsx *.xlsm)"
_icon_provider = QFileIconProvider()


def _style_icon(standard_pixmap, size=18):
    return QApplication.style().standardIcon(standard_pixmap).pixmap(size, size)


def _file_icon(path):
    return _icon_provider.icon(QFileIconProvider.IconType.File)


def _expand_path(p):
    p = os.path.abspath(p)
    if os.path.isdir(p):
        out = []
        for ext in (".xls", ".xlsx", ".xlsm"):
            # 通配符必须是 "*" + ext；少写 "*" 只会匹配名字恰好叫 ".xlsx" 的文件
            out.extend(glob.glob(os.path.join(p, "**", "*" + ext), recursive=True))
        return sorted(out)
    if os.path.isfile(p) and p.lower().endswith(SUPPORTED_EXTENSIONS):
        return [p]
    return []


class ExtractWorker(QThread):
    progress = Signal(int, str)
    task_done = Signal(dict)
    finished = Signal(dict)

    def __init__(self, bom_path, report_files, out_dir, round_no=None):
        super().__init__()
        self.bom_path = bom_path
        self.report_files = report_files
        self.out_dir = out_dir or None
        self.round_no = round_no

    def run(self):
        summary = run_batch(
            self.bom_path,
            self.report_files,
            self.out_dir,
            round_no=self.round_no,
            on_progress=lambda p, m: self.progress.emit(p, m),
            on_task=lambda r: self.task_done.emit(r),
        )
        self.finished.emit(summary)


class DropFileList(QListWidget):
    files_dropped = Signal(list)

    STYLE_NORMAL = (
        "QListWidget{border:2px dashed #9AA7B5;border-radius:12px;background:#FFFFFF;"
        "padding:8px;}"
        "QListWidget::item{padding:6px 10px;margin:2px 4px;border-radius:6px;}"
        "QListWidget::item:hover{background:#EAF4FD;}"
    )
    STYLE_HOVER = (
        "QListWidget{border:2px dashed #3498DB;border-radius:12px;background:#EAF4FD;"
        "padding:8px;}"
        "QListWidget::item{padding:6px 10px;margin:2px 4px;border-radius:6px;}"
        "QListWidget::item:hover{background:#D0EBFF;}"
    )

    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)
        self.setSelectionMode(QAbstractItemView.NoSelection)
        self.setMinimumHeight(96)
        self.setMaximumHeight(160)
        self.setStyleSheet(self.STYLE_NORMAL)
        self._empty_hint = None
        self._single = False

    def set_single(self, v):
        self._single = v

    def _show_empty_hint(self):
        self.clear()
        item = QListWidgetItem(self._hint_text())
        item.setForeground(Qt.gray)
        item.setFlags(item.flags() & ~Qt.ItemIsEnabled)
        item.setTextAlignment(Qt.AlignCenter)
        self._empty_hint = item
        self.addItem(item)

    def _hint_text(self):
        return "  拖入 Excel 文件（或文件夹），或点击下方按钮选择"

    def _hide_empty_hint(self):
        if self._empty_hint is not None:
            self.takeItem(self.row(self._empty_hint))
            self._empty_hint = None

    def add_file_item(self, filepath):
        self._hide_empty_hint()
        base = os.path.basename(filepath)
        item = QListWidgetItem(base)
        item.setToolTip(filepath)
        item.setData(Qt.UserRole, filepath)
        item.setIcon(_file_icon(filepath))
        ext = os.path.splitext(filepath)[1].lower()
        if ext == ".xlsx":
            item.setForeground(Qt.darkBlue)
        elif ext == ".xlsm":
            item.setForeground(Qt.darkMagenta)
        else:
            item.setForeground(Qt.black)
        self.addItem(item)

    def clear_files(self):
        self.clear()
        self._show_empty_hint()

    def files(self):
        out = []
        for i in range(self.count()):
            it = self.item(i)
            p = it.data(Qt.UserRole)
            if p:
                out.append(p)
        return out

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            self.setStyleSheet(self.STYLE_HOVER)
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragLeaveEvent(self, event):
        self.setStyleSheet(self.STYLE_NORMAL)
        super().dragLeaveEvent(event)

    def dropEvent(self, event: QDropEvent):
        self.setStyleSheet(self.STYLE_NORMAL)
        urls = event.mimeData().urls()
        files = []
        for u in urls:
            files.extend(_expand_path(u.toLocalFile()))
        if files:
            self.files_dropped.emit(files)
        event.acceptProposedAction()


class Card(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("card")
        self.setStyleSheet(
            "QFrame#card{background:" + C_CARD + ";border:1px solid " + C_BORDER +
            ";border-radius:12px;}"
        )


def _section_header(text, icon_pm=None):
    h = QHBoxLayout()
    h.setContentsMargins(0, 0, 0, 6)
    if icon_pm is not None:
        ic = QLabel()
        ic.setPixmap(_style_icon(icon_pm, 16))
        h.addWidget(ic)
    lab = QLabel(text)
    lab.setStyleSheet(f"font-weight:bold;font-size:13px;color:{C_PRIMARY};")
    h.addWidget(lab)
    h.addStretch()
    wrap = QWidget()
    wrap.setLayout(h)
    return wrap


# --------------------------------------------------------------------------- #
# 任务详情（单击结果行弹出）
# --------------------------------------------------------------------------- #
_KIND_LABEL = {
    "task": "图号任务",
    "file": "报告文件",
    "bom": "BOM 文件",
    "bom_write": "BOM 回写",
}


def _detail_text(res: dict) -> str:
    """把任务结果 dict 格式化为完整的多行详情文本。"""
    if res.get("ok"):
        status = "✅ 成功"
    elif res.get("kind") == "bom_write":
        status = "⚠️ BOM 回写失败（不影响本次抽取结果）"
    else:
        status = "❌ 失败"
    lines = [f"状态：{status}"]
    kind = res.get("kind", "")
    lines.append(f"类型：{_KIND_LABEL.get(kind, kind or '未知')}")
    if res.get("bh"):
        lines.append(f"图号：{res['bh']}")
    if res.get("qty") not in (None, "", 0):
        lines.append(f"件数：{res['qty']}")
    if res.get("error_code"):
        lines.append(f"错误码：{res['error_code']}")
    if res.get("reason"):
        lines.append("")
        lines.append("原因：")
        lines.append(str(res["reason"]))
    if res.get("file"):
        lines.append("")
        lines.append(f"相关文件：{res['file']}")
    segs = res.get("segments") or []
    if len(segs) > 1:
        lines.append("")
        lines.append(f"来源：{len(segs)} 份基准报告拼接")
        for i, s in enumerate(segs, 1):
            lines.append(
                f"  {i}. {os.path.basename(s.get('path', ''))} "
                f"取 {s.get('count')} 件（已用 {s.get('used_from')}→{s.get('used_to')}）"
                f" → {os.path.basename(s.get('source_after') or '')}"
            )
    elif res.get("source_before") or res.get("source_after"):
        lines.append("")
        lines.append(f"基准报告（前）：{res.get('source_before', '')}")
        lines.append(f"基准报告（后）：{res.get('source_after', '')}")
    if res.get("output"):
        lines.append(f"输出报告：{res['output']}")
    if res.get("used_from") is not None and res.get("used_to") is not None:
        lines.append(f"已用量：{res['used_from']} → {res['used_to']}（总件数 {res.get('total', '')}）")
    if res.get("date"):
        lines.append(f"日期：{res['date']}")
    return "\n".join(lines)


class DetailDialog(QDialog):
    """显示单个任务的完整信息（长文本可滚动 + 一键复制）。"""

    def __init__(self, res: dict, parent=None):
        super().__init__(parent)
        self.setWindowTitle("任务详情")
        self.setMinimumSize(660, 420)
        v = QVBoxLayout(self)
        v.setContentsMargins(16, 14, 16, 14)
        v.setSpacing(10)

        text = QPlainTextEdit()
        text.setReadOnly(True)
        text.setPlainText(_detail_text(res))
        text.setStyleSheet(
            "QPlainTextEdit{background:#FBFCFE;border:1px solid " + C_BORDER +
            ";border-radius:8px;padding:10px;font-family:'Microsoft YaHei','Consolas',"
            "monospace;font-size:12px;color:" + C_TEXT + ";}"
        )
        v.addWidget(text, 1)

        row = QHBoxLayout()
        row.addStretch()
        btn_copy = QPushButton("复制全部")
        btn_copy.setCursor(Qt.PointingHandCursor)
        btn_copy.clicked.connect(
            lambda: QApplication.clipboard().setText(text.toPlainText())
        )
        btn_close = QPushButton("关闭")
        btn_close.setCursor(Qt.PointingHandCursor)
        btn_close.setDefault(True)
        btn_close.clicked.connect(self.accept)
        row.addWidget(btn_copy)
        row.addWidget(btn_close)
        v.addLayout(row)



class ExtractApp(QWidget):
    def __init__(self):
        super().__init__()
        self.bom_files = []
        self.report_files = []
        self.out_dir = ""
        self.worker = None
        self.last_outputs = []
        self._build_ui()

    def _build_ui(self):
        self.setWindowTitle(APP_TITLE)
        self.setMinimumSize(720, 640)
        self.resize(780, 720)

        self.setStyleSheet(
            f"QWidget#MainWindow{{background:{C_BG};}}"
            "QWidget{font-family:'Microsoft YaHei','Segoe UI','PingFang SC',sans-serif;color:#2C3E50;}"
        )
        self.setObjectName("MainWindow")

        # 滚动区域，避免控件重叠/被裁剪
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll.setStyleSheet("QScrollArea{border:none;background:" + C_BG + ";}")
        content = QWidget()
        content.setObjectName("content")
        shadow = QGraphicsDropShadowEffect(content)
        shadow.setBlurRadius(24)
        shadow.setColor(QColor(0, 0, 0, 60))
        shadow.setOffset(0, 6)
        content.setGraphicsEffect(shadow)
        content.setStyleSheet(
            "QWidget#content{background:" + C_CARD + ";border:1px solid " + C_BORDER +
            ";border-radius:12px;}"
        )

        cl = QVBoxLayout(content)
        cl.setSpacing(14)
        cl.setContentsMargins(22, 20, 22, 20)
        root.addWidget(scroll)
        scroll.setWidget(content)

        # 标题
        title_row = QHBoxLayout()
        title = QLabel("提取文件数据")
        title.setFont(QFont("Microsoft YaHei", 19, QFont.Weight.Bold))
        title.setStyleSheet(f"color:{C_PRIMARY};")
        title_row.addWidget(title)
        title_row.addStretch()
        badge = QLabel(f"v{VERSION}")
        badge.setStyleSheet(
            f"background:{C_ACCENT};color:#FFFFFF;font-size:11px;font-weight:bold;"
            "padding:3px 12px;border-radius:10px;"
        )
        title_row.addWidget(badge)
        cl.addLayout(title_row)

        sub = QLabel("从检验大报告按 BOM 行逐件抽取测量数据生成独立小报告，就地推进基准报告的已用件数，"
                     "并把「日期/件数」回写到 BOM 的第N套列")
        sub.setStyleSheet(f"font-size:11px;color:{C_TEXT_2};margin-top:2px;")
        sub.setWordWrap(True)
        cl.addWidget(sub)

        # BOM 卡片
        bom_card = Card()
        bom_l = QVBoxLayout(bom_card)
        bom_l.setSpacing(8)
        bom_l.setContentsMargins(16, 14, 16, 14)
        bom_l.addWidget(_section_header("① BOM 文件（图号 + 数量，单文件）", QStyle.SP_FileIcon))
        self.bom_list = DropFileList()
        self.bom_list.set_single(True)
        self.bom_list.files_dropped.connect(self._add_bom)
        bom_l.addWidget(self.bom_list)
        bom_btn = QHBoxLayout()
        self.btn_bom = self._make_button("选择 BOM", primary=False, icon_pm=QStyle.SP_FileIcon)
        self.btn_bom.clicked.connect(self._select_bom)
        self.btn_bom_clear = self._make_button("清空", primary=False, icon_pm=QStyle.SP_TrashIcon)
        self.btn_bom_clear.clicked.connect(self._clear_bom)
        bom_btn.addStretch()
        bom_btn.addWidget(self.btn_bom)
        bom_btn.addWidget(self.btn_bom_clear)
        bom_l.addLayout(bom_btn)

        # 当前第几套（必填）：转换完成后把「日期/件数」回写到 BOM 的「第N套」列
        round_row = QHBoxLayout()
        round_lab = QLabel("当前第几套：")
        round_lab.setStyleSheet(f"font-size:12px;font-weight:bold;color:{C_PRIMARY};")
        self.round_edit = QLineEdit()
        self.round_edit.setPlaceholderText("必填，如1")
        self.round_edit.setValidator(QIntValidator(1, 9999, self))
        self.round_edit.setFixedWidth(110)
        self.round_edit.setMinimumHeight(32)
        self.round_edit.setStyleSheet(
            "QLineEdit{background:" + C_BG + ";border:1px solid " + C_BORDER +
            ";border-radius:6px;padding:4px 10px;font-size:12px;color:" + C_TEXT + ";}"
            "QLineEdit:focus{border-color:" + C_ACCENT + ";background:#FFFFFF;}"
        )
        round_row.addWidget(round_lab)
        round_row.addWidget(self.round_edit)
        round_row.addStretch(1)
        bom_l.addLayout(round_row)
        cl.addWidget(bom_card)

        # 报告文件池卡片
        rep_card = Card()
        rep_l = QVBoxLayout(rep_card)
        rep_l.setSpacing(8)
        rep_l.setContentsMargins(16, 14, 16, 14)
        rep_l.addWidget(_section_header(
            "② 基准报告文件池（拖入文件夹/多个文件，按图号匹配 BOM 行；文件名含 图号_描述_N件_日期-已用）",
            QStyle.SP_FileIcon))
        self.rep_list = DropFileList()
        self.rep_list.files_dropped.connect(self._add_reports)
        rep_l.addWidget(self.rep_list)
        rep_btn = QHBoxLayout()
        self.btn_rep = self._make_button("选择报告", primary=False, icon_pm=QStyle.SP_FileIcon)
        self.btn_rep.clicked.connect(self._select_reports)
        self.btn_rep_clear = self._make_button("清空", primary=False, icon_pm=QStyle.SP_TrashIcon)
        self.btn_rep_clear.clicked.connect(self._clear_reports)
        rep_btn.addStretch()
        rep_btn.addWidget(self.btn_rep)
        rep_btn.addWidget(self.btn_rep_clear)
        rep_l.addLayout(rep_btn)
        cl.addWidget(rep_card)

        # 输出目录卡片
        out_card = Card()
        out_l = QVBoxLayout(out_card)
        out_l.setSpacing(8)
        out_l.setContentsMargins(16, 14, 16, 14)
        out_l.addWidget(_section_header("③ 输出目录", QStyle.SP_DirIcon))
        out_row = QHBoxLayout()
        self.out_edit = QLabel("默认：源文件所在目录")
        self.out_edit.setStyleSheet(
            "QLabel{background:" + C_BG + ";border:1px solid " + C_BORDER +
            ";border-radius:6px;padding:9px 12px;font-size:12px;color:" + C_TEXT + ";}"
        )
        self.btn_out = self._make_button("选择目录", primary=False, icon_pm=QStyle.SP_DirOpenIcon)
        self.btn_out.clicked.connect(self._select_out)
        out_row.addWidget(self.out_edit, 1)
        out_row.addWidget(self.btn_out)
        out_l.addLayout(out_row)
        cl.addWidget(out_card)

        # 进度
        prog_group = QVBoxLayout()
        prog_group.setSpacing(6)
        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.progress.setTextVisible(True)
        self.progress.setFormat("%p%")
        self.progress.setMinimumHeight(22)
        self.progress.setMaximumHeight(22)
        self.progress.setStyleSheet(
            "QProgressBar{background:" + C_BG + ";border:none;border-radius:11px;font-size:11px;"
            "color:" + C_TEXT_2 + ";text-align:center;}"
            "QProgressBar::chunk{background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 " +
            C_PROG_A + ",stop:1 " + C_PROG_B + ");border-radius:11px;}"
        )
        prog_group.addWidget(self.progress)
        self._prog_anim = QPropertyAnimation(self.progress, b"value")
        self._prog_anim.setDuration(500)
        self._prog_anim.setEasingCurve(QEasingCurve.Type.OutCubic)
        self.status = QLabel("")
        self.status.setWordWrap(True)
        self.status.setStyleSheet(f"font-size:13px;font-weight:bold;color:{C_OK};padding:6px 0;")
        self.status.hide()
        prog_group.addWidget(self.status)
        cl.addLayout(prog_group)

        # 结果列表（单击任意行弹出完整详情）
        cl.addWidget(_section_header("处理结果（逐任务实时 · 单击任意行查看完整信息）",
                                     QStyle.SP_FileDialogDetailedView))
        self.result_list = QListWidget()
        self.result_list.setMinimumHeight(120)
        # 点击仅用于查看详情，不需要选中高亮（避免选中态导致行渲染异常）
        self.result_list.setSelectionMode(QAbstractItemView.NoSelection)
        self.result_list.setStyleSheet(
            "QListWidget{border:1px solid " + C_BORDER + ";border-radius:8px;background:#FFFFFF;"
            "padding:4px;font-size:12px;}"
            "QListWidget::item{padding:4px 8px;margin:2px 0;border-radius:5px;}"
        )
        self.result_list.itemClicked.connect(self._on_item_clicked)
        cl.addWidget(self.result_list)

        # 底部按钮
        btn_row = QHBoxLayout()
        btn_row.setSpacing(12)
        self.btn_start = QPushButton("开始抽取")
        self.btn_start.setIcon(QIcon(_style_icon(QStyle.SP_MediaPlay, 18)))
        self.btn_start.setCursor(Qt.PointingHandCursor)
        self.btn_start.setMinimumHeight(46)
        self.btn_start.setStyleSheet(self._btn_primary_style())
        self.btn_start.clicked.connect(self._start)
        self.btn_open_dir = QPushButton("打开输出目录")
        self.btn_open_dir.setIcon(QIcon(_style_icon(QStyle.SP_DirOpenIcon, 18)))
        self.btn_open_dir.setCursor(Qt.PointingHandCursor)
        self.btn_open_dir.setEnabled(False)
        self.btn_open_dir.setMinimumHeight(46)
        self.btn_open_dir.setStyleSheet(self._btn_secondary_style())
        self.btn_open_dir.clicked.connect(self._open_dir)
        self.btn_open = QPushButton("查看输出文件")
        self.btn_open.setIcon(QIcon(_style_icon(QStyle.SP_FileDialogDetailedView, 18)))
        self.btn_open.setCursor(Qt.PointingHandCursor)
        self.btn_open.setEnabled(False)
        self.btn_open.setMinimumHeight(46)
        self.btn_open.setStyleSheet(self._btn_secondary_style())
        self.btn_open.clicked.connect(self._open_out)
        btn_row.addWidget(self.btn_start, 3)
        btn_row.addWidget(self.btn_open_dir, 2)
        btn_row.addWidget(self.btn_open, 2)
        cl.addLayout(btn_row)

        self._refresh_bom()
        self._refresh_reports()
        self._update_start_state()

    # 按钮样式
    def _btn_primary_style(self):
        return (
            "QPushButton{background:" + C_ACCENT + ";color:#FFFFFF;font-weight:bold;font-size:14px;"
            "border:none;border-radius:6px;padding:10px 20px;}"
            "QPushButton:hover{background:" + C_ACCENT_HOVER + ";}"
            "QPushButton:pressed{background:" + C_ACCENT_PRESS + ";}"
            "QPushButton:disabled{background:#AEBECB;color:#F0F4F8;}"
        )

    def _btn_secondary_style(self):
        return (
            "QPushButton{background:#FFFFFF;color:" + C_PRIMARY + ";border:1px solid " + C_BORDER +
            ";font-weight:bold;font-size:13px;border-radius:6px;padding:10px 20px;}"
            "QPushButton:hover{background:#EAF4FD;border-color:" + C_ACCENT + ";}"
            "QPushButton:pressed{background:#D6EAFB;}"
            "QPushButton:disabled{color:#C3CDD6;border-color:#E5EAF0;}"
        )

    def _make_button(self, text, primary, icon_pm=None):
        btn = QPushButton(text)
        if icon_pm is not None:
            btn.setIcon(QIcon(_style_icon(icon_pm, 16)))
        btn.setCursor(Qt.PointingHandCursor)
        btn.setMinimumHeight(38)
        if primary:
            btn.setStyleSheet(self._btn_primary_style())
        else:
            btn.setStyleSheet(
                "QPushButton{background:" + C_BG + ";color:" + C_PRIMARY + ";border:1px solid " +
                C_BORDER + ";font-weight:500;font-size:12px;border-radius:6px;padding:8px 16px;}"
                "QPushButton:hover{background:#EAF4FD;border-color:" + C_ACCENT + ";}"
                "QPushButton:pressed{background:#D6EAFB;}"
            )
        return btn

    # 文件列表
    def _refresh_bom(self):
        self.bom_list.clear_files()
        for f in self.bom_files:
            self.bom_list.add_file_item(f)

    def _refresh_reports(self):
        self.rep_list.clear_files()
        for f in self.report_files:
            self.rep_list.add_file_item(f)

    def _add_bom(self, files):
        files = [os.path.abspath(f) for f in files if f.lower().endswith(SUPPORTED_EXTENSIONS)]
        if not files:
            return
        self.bom_files = [files[0]]   # 仅取第一个（BOM 单文件）
        self._refresh_bom()
        self._update_start_state()

    def _add_reports(self, files):
        changed = False
        for f in files:
            fp = os.path.abspath(f)
            if fp.lower().endswith(SUPPORTED_EXTENSIONS) and fp not in self.report_files:
                self.report_files.append(fp)
                changed = True
        if changed:
            self._refresh_reports()
        self._update_start_state()

    def _select_bom(self):
        files, _ = QFileDialog.getOpenFileNames(self, "选择 BOM 文件", "", FILTER_STR)
        if files:
            self._add_bom(files)

    def _select_reports(self):
        files, _ = QFileDialog.getOpenFileNames(self, "选择检验报告", "", FILTER_STR)
        if files:
            self._add_reports(files)

    def _clear_bom(self):
        self.bom_files = []
        self.bom_list.clear_files()
        self._update_start_state()

    def _clear_reports(self):
        self.report_files = []
        self.rep_list.clear_files()
        self._update_start_state()

    def _select_out(self):
        d = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if d:
            self.out_dir = d
            self.out_edit.setText(d)

    def _update_start_state(self):
        ready = bool(self.bom_files) and bool(self.report_files) and not self._running()
        self.btn_start.setEnabled(ready)

    def _running(self):
        return self.worker is not None and self.worker.isRunning()

    def _start(self):
        if not self.bom_files:
            QMessageBox.warning(self, "提示", "请先拖入 BOM 文件。")
            return
        if not self.report_files:
            QMessageBox.warning(self, "提示", "请先拖入检验报告文件（或文件夹）。")
            return
        # 当前第几套：必填（正整数）
        txt = self.round_edit.text().strip()
        if not txt or not txt.isdigit() or int(txt) <= 0:
            QMessageBox.warning(
                self, "提示",
                "请填写「当前第几套」（正整数，例如 13）。\n"
                "转换完成后会把「日期/件数」写入 BOM 的第N套列。",
            )
            self.round_edit.setFocus()
            return
        round_no = int(txt)

        self.btn_start.setEnabled(False)
        self.btn_start.setText("抽取中…")
        self.btn_start.setIcon(QIcon(_style_icon(QStyle.SP_MediaPause, 18)))
        self.btn_open.setEnabled(False)
        self.progress.setValue(0)
        self.status.hide()
        self.result_list.clear()

        self.worker = ExtractWorker(self.bom_files[0], self.report_files, self.out_dir,
                                    round_no=round_no)
        self.worker.progress.connect(self._on_progress)
        self.worker.task_done.connect(self._on_task)
        self.worker.finished.connect(self._on_finished)
        self.worker.start()

    def _on_progress(self, pct, msg):
        self._prog_anim.stop()
        self._prog_anim.setStartValue(self.progress.value())
        self._prog_anim.setEndValue(pct)
        self._prog_anim.start()
        self.status.setStyleSheet(f"font-size:12px;color:{C_TEXT_2};padding:4px 0;")
        self.status.setText(msg)
        self.status.show()

    def _on_task(self, res):
        kind = res.get("kind", "task")
        code = res.get("error_code", "")
        if res.get("ok"):
            line = (f"✅ {res.get('bh', '')} {res.get('qty', '')}件 → "
                    f"{os.path.basename(res.get('output', ''))}")
            color = C_OK
        elif kind == "already":
            line = (f"⚠️ {res.get('bh','')} 第{res.get('round_no','')}套已做过"
                    f"（BOM 已有：{res.get('existing','')}），跳过")
            color = C_WARN
        elif kind == "bom_write":
            extra = f" {res['qty']}件" if res.get("qty") else ""
            line = f"⚠️ BOM 回写失败 {res.get('bh', '')}{extra} [{code}]"
            color = C_WARN
        elif kind == "bom":
            line = f"❌ BOM 读取失败 [{code}]"
            color = C_ERR
        elif kind == "file":
            line = f"❌ {os.path.basename(res.get('file', ''))} [{code}]"
            color = C_ERR
        else:
            line = f"❌ {res.get('bh', '')} {res.get('qty', '')}件 [{code}]"
            color = C_ERR
        it = QListWidgetItem(line)
        it.setForeground(QColor(color))
        it.setData(Qt.UserRole, dict(res))       # 完整信息（单击查看）
        it.setData(Qt.UserRole + 1, line)        # 原文快照：详情关闭后用于复原行显示
        it.setToolTip(_detail_text(res))          # 悬停也能看全文
        self.result_list.addItem(it)
        self.result_list.scrollToBottom()

    def _on_item_clicked(self, item):
        """单击任意结果行 -> 弹出完整详情框（长错误信息可滚动查看/复制）。"""
        data = item.data(Qt.UserRole)
        if not isinstance(data, dict):
            return
        # 记录点击前原文，详情关闭后强制复原，保证行显示与点击前完全一致
        saved = item.data(Qt.UserRole + 1) or item.text()
        try:
            DetailDialog(data, self).exec()
        except Exception:  # noqa: BLE001
            pass
        try:
            item.setText(saved)
        except Exception:  # noqa: BLE001
            pass

    def _on_finished(self, summary):
        self.last_outputs = [
            r.get("output") for r in summary.get("tasks", [])
            if r.get("ok") and r.get("output") and os.path.isfile(r["output"])
        ]
        self.btn_start.setText("开始抽取")
        self.btn_start.setIcon(QIcon(_style_icon(QStyle.SP_MediaPlay, 18)))
        self.btn_open.setEnabled(bool(self.last_outputs))
        self.btn_open_dir.setEnabled(True)

        ok = summary.get("ok_count", 0)
        fail = summary.get("fail_count", 0)
        # 文件解析错误也计入失败提示
        fe = summary.get("file_errors", [])
        skipped = summary.get("skipped_rows", [])
        bwe = summary.get("bom_write_errors", [])

        if ok and not fail and not fe and not bwe:
            txt = f"✅ 全部完成：共 {ok} 个任务"
            color = C_OK
        else:
            parts = []
            if ok:
                parts.append(f"成功 {ok} 个")
            if fail:
                parts.append(f"任务失败 {fail} 个")
            if fe:
                parts.append(f"文件解析失败 {len(fe)} 个")
            if bwe:
                parts.append(f"BOM 回写失败 {len(bwe)} 处")
            if skipped:
                parts.append(f"跳过空数量行 {len(skipped)} 个")
            txt = "⚠️ " + " ｜ ".join(parts) + "（单击列表行查看完整信息）"
            color = C_ERR if (fail or fe) and not ok else C_WARN

        if ok and not bwe:
            txt += "\nBOM 已回写"

        self.status.setStyleSheet(f"font-size:13px;font-weight:bold;color:{color};padding:6px 0;")
        self.status.setText(txt)
        self.status.show()

        # 清空输入，便于继续
        self.report_files = []
        self.rep_list.clear_files()
        self.btn_start.setEnabled(False)

    def _open_dir(self):
        target = self.out_dir or (
            os.path.dirname(self.last_outputs[0]) if self.last_outputs else None
        )
        if target and os.path.isdir(target):
            self._open_path(target)

    def _open_out(self):
        if not self.last_outputs:
            self._open_dir()
            return
        if len(self.last_outputs) == 1:
            self._open_path(self.last_outputs[0])
            return
        menu = QMenu(self)
        for p in self.last_outputs:
            act = menu.addAction(os.path.basename(p))
            act.triggered.connect(lambda checked=False, path=p: self._open_path(path))
        btn = self.btn_open
        menu.exec(btn.mapToGlobal(btn.rect().bottomLeft()))

    def _open_path(self, target):
        if not target:
            return
        try:
            os.startfile(target)
        except Exception:
            try:
                subprocess.Popen(["explorer", os.path.dirname(target) or target])
            except Exception:
                pass


# --------------------------------------------------------------------------- #
# 弹窗可读性：强制浅色主题
#   系统处于深色模式时，QMessageBox 的背景取系统调色板（深色），而全局样式又把
#   文字设为深色 #2C3E50，造成「深底 + 深字」看不清。这里用 Fusion 样式 +
#   显式浅色调色板 + QMessageBox 显式配色，保证弹窗始终白底深字、按钮清晰。
# --------------------------------------------------------------------------- #
_DIALOG_QSS = (
    "QMessageBox{background:#FFFFFF;}"
    "QMessageBox QLabel{color:" + C_TEXT + ";font-size:13px;background:transparent;}"
    "QMessageBox QPushButton{min-width:78px;min-height:30px;padding:4px 18px;"
    "background:" + C_ACCENT + ";color:#FFFFFF;font-weight:bold;font-size:13px;"
    "border:none;border-radius:6px;}"
    "QMessageBox QPushButton:hover{background:" + C_ACCENT_HOVER + ";}"
    "QMessageBox QPushButton:pressed{background:" + C_ACCENT_PRESS + ";}"
    "QDialog{background:#FFFFFF;}"
    "QDialog QLabel{color:" + C_TEXT + ";}"
    "QDialog QPushButton{min-width:78px;min-height:30px;padding:4px 18px;"
    "background:" + C_ACCENT + ";color:#FFFFFF;font-weight:bold;font-size:13px;"
    "border:none;border-radius:6px;}"
    "QDialog QPushButton:hover{background:" + C_ACCENT_HOVER + ";}"
    "QDialog QPushButton:pressed{background:" + C_ACCENT_PRESS + ";}"
)


def _apply_light_palette(app):
    """强制浅色调色板（不受系统深色模式影响）。"""
    pal = QPalette()
    role = QPalette.ColorRole
    pal.setColor(role.Window, QColor("#FFFFFF"))
    pal.setColor(role.WindowText, QColor(C_TEXT))
    pal.setColor(role.Base, QColor("#FFFFFF"))
    pal.setColor(role.AlternateBase, QColor(C_BG))
    pal.setColor(role.Text, QColor(C_TEXT))
    pal.setColor(role.Button, QColor(C_BG))
    pal.setColor(role.ButtonText, QColor(C_TEXT))
    pal.setColor(role.ToolTipBase, QColor("#FFFFFF"))
    pal.setColor(role.ToolTipText, QColor(C_TEXT))
    pal.setColor(role.Highlight, QColor(C_ACCENT))
    pal.setColor(role.HighlightedText, QColor("#FFFFFF"))
    try:
        pal.setColor(role.PlaceholderText, QColor(C_TEXT_2))
    except Exception:  # noqa: BLE001
        pass
    app.setPalette(pal)


def main():
    for attr in ("AA_EnableHighDpiScaling", "AA_UseHighDpiPixmaps"):
        if hasattr(Qt, attr):
            QApplication.setAttribute(getattr(Qt, attr), True)
    app = QApplication(sys.argv)
    app.setApplicationName("提取文件数据")
    app.setStyle("Fusion")           # 统一各平台外观，弹窗配色可控
    _apply_light_palette(app)        # 浅色底，避免深色模式下弹窗「深底深字」
    app.setStyleSheet(
        "QMessageBox,QDialog{font-family:'Microsoft YaHei','Segoe UI';}"
        + _DIALOG_QSS
    )
    win = ExtractApp()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
