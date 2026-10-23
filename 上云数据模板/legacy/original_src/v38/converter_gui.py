#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
检验报告转换工具 —— 原生桌面 GUI (PySide6)
=========================================

功能（保持不变）:
    - 选择 / 拖拽多个文件（支持 .xls / .xlsx / .csv）
    - 选择输出目录（默认输出到源文件所在目录）
    - 点击「开始转换」后台执行，界面不冻结
    - 实时显示整体进度（多文件按权重累计）与状态文字
    - 转换完成弹出结果提示，可一键打开输出目录

转换核心由 converter 模块提供（Excel COM 保留图片/合并，自动回退 openpyxl）。

版本管理:
    - 从 VERSION 文件读取当前版本号
    - 窗口标题和右上角徽章都用 VERSION
    - 每次打包递增 VERSION
"""
import os as _os
import sys as _sys
# VERSION 文件查找：优先 exe 所在目录（dist/VERSION），回退 __file__ 目录（开发模式）
_exe_dir = _os.path.dirname(_os.path.abspath(_sys.executable)) if getattr(_sys, 'frozen', False) else _os.path.dirname(_os.path.abspath(__file__))
_VERSION_FILE = _os.path.join(_exe_dir, "VERSION")
try:
    with open(_VERSION_FILE, "r", encoding="utf-8") as _vf:
        VERSION = _vf.read().strip()
except Exception:
    VERSION = "0"
APP_TITLE = f"检验报告格式转换工具 v{VERSION}"

import os
import sys
import subprocess

from PySide6.QtCore import Qt, QThread, Signal, QPropertyAnimation, QEasingCurve
from PySide6.QtGui import (
    QDragEnterEvent, QDropEvent, QFont,
    QIcon, QColor, QIntValidator,
)
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QProgressBar, QFileDialog, QMessageBox, QListWidget,
    QListWidgetItem, QFrame, QAbstractItemView, QSizePolicy, QStyle,
    QGraphicsDropShadowEffect, QFileIconProvider, QLineEdit,
)

import datetime
import converter  # 真实转换核心（同一目录）

# --------------------------------------------------------------------------- #
# 配色方案（Fluent / Material 风的柔和蓝灰系）
# --------------------------------------------------------------------------- #
C_PRIMARY = "#2C3E50"   # 深蓝灰（主色）
C_ACCENT = "#3498DB"    # 亮蓝（强调色）
C_ACCENT_HOVER = "#2980B9"
C_ACCENT_PRESS = "#21618C"
C_BG = "#F5F7FA"        # 浅灰白背景
C_CARD = "#FFFFFF"      # 卡片白
C_TEXT = "#2C3E50"      # 主文字
C_TEXT_2 = "#7F8C8D"    # 次要文字
C_BORDER = "#DCE3EA"    # 浅边框
C_PROG_A = "#3498DB"    # 进度渐变起
C_PROG_B = "#2ECC71"    # 进度渐变止
C_OK = "#27AE60"        # 完成绿
C_ERR = "#E74C3C"       # 错误红
C_WARN = "#E67E22"      # 处理中橙蓝

# 支持的输入格式
SUPPORTED_EXTENSIONS = (".xls", ".xlsx", ".csv")
FILTER_STR = "Excel / CSV 文件 (*.xls *.xlsx *.csv)"

# 文件类型图标映射（使用系统图标提供者）
_icon_provider = QFileIconProvider()


def _style_icon(standard_pixmap, size=18):
    """取得 Qt 内置图标（SP_*）。"""
    return QApplication.style().standardIcon(standard_pixmap).pixmap(size, size)


def _file_icon(path):
    """根据扩展名返回系统文件图标。"""
    return _icon_provider.icon(QFileIconProvider.IconType.File)


# --------------------------------------------------------------------------- #
# 后台转换工作线程
# --------------------------------------------------------------------------- #
class ConvertWorker(QThread):
    """在子线程中顺序转换多个文件，避免阻塞 UI。"""

    progress = Signal(int, str)
    finished = Signal(list)

    def __init__(self, files, out_dir, default_year=None):
        super().__init__()
        self.files = files
        self.out_dir = out_dir or None
        self.default_year = default_year

    def run(self):
        results = []
        total = max(len(self.files), 1)
        for idx, f in enumerate(self.files):
            base = os.path.basename(f)
            self.progress.emit(int(idx / total * 100), f"准备: {base}")

            def on_progress(pct, msg):
                overall = int((idx + pct / 100.0) / total * 100)
                self.progress.emit(overall, f"[{idx + 1}/{total}] {base} - {msg}")

            try:
                res = converter.convert_file(
                    f, self.out_dir, on_progress=on_progress,
                    default_year=self.default_year,
                )
                # res = {"output": ..., "total_qty": ..., ...}
                out_path = res.get("output") if isinstance(res, dict) else None
                results.append({"file": f, "ok": True, "res": res, "output": out_path})
            except Exception as e:  # noqa: BLE001
                results.append({"file": f, "ok": False, "error": str(e)})

        self.progress.emit(100, "全部完成")
        self.finished.emit(results)


# --------------------------------------------------------------------------- #
# 合并的拖拽+文件列表控件
# --------------------------------------------------------------------------- #
class DropFileList(QListWidget):
    """一个 QListWidget：空时显示拖拽提示，有文件时显示文件项，始终可拖入。"""

    files_dropped = Signal(list)

    STYLE_NORMAL = (
        "QListWidget{"
        "  border:2px dashed #9AA7B5;"
        "  border-radius:12px;"
        "  background:#FFFFFF;"
        "  padding:8px;"
        "}"
        "QListWidget::item{"
        "  padding:6px 10px;margin:2px 4px;"
        "  border-radius:6px;"
        "}"
        "QListWidget::item:hover{background:#EAF4FD;}"
    )
    STYLE_HOVER = (
        "QListWidget{"
        "  border:2px dashed #3498DB;"
        "  border-radius:12px;"
        "  background:#EAF4FD;"
        "  padding:8px;"
        "}"
        "QListWidget::item{"
        "  padding:6px 10px;margin:2px 4px;"
        "  border-radius:6px;"
        "}"
        "QListWidget::item:hover{background:#D0EBFF;}"
    )

    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)
        self.setSelectionMode(QAbstractItemView.NoSelection)
        self.setMinimumHeight(140)
        self.setMaximumHeight(200)
        self.setStyleSheet(self.STYLE_NORMAL)
        self._empty_hint = None

    def _show_empty_hint(self):
        """无文件时显示占位提示。"""
        self.clear()
        item = QListWidgetItem("  拖入 .xls / .xlsx / .csv 文件，或点击下方按钮选择")
        item.setForeground(Qt.gray)
        item.setFlags(item.flags() & ~Qt.ItemIsEnabled)
        item.setTextAlignment(Qt.AlignCenter)
        self._empty_hint = item
        self.addItem(item)

    def _hide_empty_hint(self):
        """添加文件后移除占位提示。"""
        if self._empty_hint is not None:
            self.takeItem(self.row(self._empty_hint))
            self._empty_hint = None

    def add_file_item(self, filepath):
        """添加一个带图标的文件项。"""
        self._hide_empty_hint()
        base = os.path.basename(filepath)
        ext = os.path.splitext(filepath)[1].lower()
        item = QListWidgetItem(base)
        item.setToolTip(filepath)
        item.setData(Qt.UserRole, filepath)
        item.setIcon(_file_icon(filepath))
        if ext == ".csv":
            item.setForeground(Qt.darkGreen)
        elif ext == ".xlsx":
            item.setForeground(Qt.darkBlue)
        else:
            item.setForeground(Qt.black)
        self.addItem(item)

    def clear_files(self):
        """清空所有文件，恢复提示。"""
        self.clear()
        self._show_empty_hint()

    # ---- 拖拽事件 ----
    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            self.setStyleSheet(self.STYLE_HOVER)
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        # QAbstractItemView 需在处理中持续接受，否则 dropEvent 不会触发
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
        ext_lower = tuple(SUPPORTED_EXTENSIONS)
        files = [u.toLocalFile() for u in urls
                 if u.toLocalFile().lower().endswith(ext_lower)]
        if files:
            self.files_dropped.emit(files)
        event.acceptProposedAction()


# --------------------------------------------------------------------------- #
# 卡片容器（带背景与内边距，用于分组）
# --------------------------------------------------------------------------- #
class Card(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("card")
        self.setStyleSheet(
            "QFrame#card{"
            f"  background:{C_CARD};"
            f"  border:1px solid {C_BORDER};"
            "  border-radius:12px;"
            "}"
        )


def _section_header(text, icon_pm=None):
    """生成带图标的分区小标题。"""
    h = QHBoxLayout()
    h.setContentsMargins(0, 0, 0, 6)
    if icon_pm is not None:
        ic = QLabel()
        ic.setPixmap(_style_icon(icon_pm, 16))
        h.addWidget(ic)
    lab = QLabel(text)
    lab.setStyleSheet(
        f"font-weight:bold;font-size:13px;color:{C_PRIMARY};"
    )
    h.addWidget(lab)
    h.addStretch()
    wrap = QWidget()
    wrap.setLayout(h)
    return wrap


# --------------------------------------------------------------------------- #
# 主窗口
# --------------------------------------------------------------------------- #
class ConverterApp(QWidget):
    def __init__(self):
        super().__init__()
        self.input_files = []
        self.out_dir = ""
        self.worker = None
        self.last_outputs = []  # 最近一次转换的输出文件路径列表
        self._build_ui()

    # ----------------------------- UI 构建 ----------------------------- #
    def _build_ui(self):
        self.setWindowTitle(APP_TITLE)
        self.setMinimumSize(660, 500)
        self.resize(720, 560)

        # 全局字体 / 背景
        self.setStyleSheet(
            f"QWidget#MainWindow{{background:{C_BG};}}"
            "QWidget{"
            "  font-family:'Microsoft YaHei','Segoe UI','PingFang SC',sans-serif;"
            "  color:#2C3E50;"
            "}"
        )
        self.setObjectName("MainWindow")

        # 根布局：留边，使中央卡片的阴影可见
        root = QVBoxLayout(self)
        root.setSpacing(0)
        root.setContentsMargins(18, 18, 18, 14)

        # === 中央卡片（带阴影） ===
        card = Card()
        shadow = QGraphicsDropShadowEffect(card)
        shadow.setBlurRadius(24)
        shadow.setColor(QColor(0, 0, 0, 60))
        shadow.setOffset(0, 6)
        card.setGraphicsEffect(shadow)

        cl = QVBoxLayout(card)
        cl.setSpacing(16)
        cl.setContentsMargins(22, 20, 22, 20)
        root.addWidget(card)

        # ===== 标题区 =====
        title_row = QHBoxLayout()
        title = QLabel("检验报告格式转换工具")
        title.setFont(QFont("Microsoft YaHei", 19, QFont.Weight.Bold))
        title.setStyleSheet(f"color:{C_PRIMARY};")
        title_row.addWidget(title)

        title_row.addStretch()

        # 年份按钮（位于徽章左边，默认只读，双击或点击进入编辑）
        year_layout = QHBoxLayout()
        year_layout.setSpacing(2)
        year_label = QLabel("年份:")
        year_label.setStyleSheet(f"color:{C_TEXT_2};font-size:11px;")
        year_layout.addWidget(year_label)
        self.year_edit = QLineEdit(str(datetime.datetime.now().year))
        self.year_edit.setFixedWidth(54)
        self.year_edit.setMaxLength(4)
        self.year_edit.setAlignment(Qt.AlignCenter)
        self.year_edit.setValidator(QIntValidator(2000, 2099, self))
        self.year_edit.setReadOnly(True)  # 默认只读，避免焦点落在输入框
        self.year_edit.setCursor(Qt.PointingHandCursor)
        self.year_edit.setStyleSheet(
            f"QLineEdit{{border:1px solid {C_BORDER};border-radius:6px;"
            f"padding:2px 4px;font-size:11px;color:{C_PRIMARY};"
            f"background:#F5F7FA;}}"
            f"QLineEdit:hover{{border:1px solid {C_ACCENT};background:#FFFFFF;}}"
        )
        # 双击或点击进入编辑模式
        self.year_edit.mouseDoubleClickEvent = lambda e: self._edit_year()
        self.year_edit.mousePressEvent = lambda e: self._edit_year()
        year_layout.addWidget(self.year_edit)
        title_row.addLayout(year_layout)

        badge = QLabel(f"v{VERSION}")
        badge.setStyleSheet(
            f"background:{C_ACCENT};color:#FFFFFF;font-size:11px;"
            "font-weight:bold;padding:3px 12px;border-radius:10px;"
        )
        title_row.addWidget(badge)
        cl.addLayout(title_row)

        # ===== 输入文件卡片 =====
        in_card = Card()
        in_l = QVBoxLayout(in_card)
        in_l.setSpacing(10)
        in_l.setContentsMargins(16, 14, 16, 14)
        in_l.addWidget(_section_header("输入文件", QStyle.SP_FileIcon))

        # 合并的拖拽+文件列表（空时显示提示，有文件时显示项）
        self.file_list = DropFileList()
        self.file_list.files_dropped.connect(self._add_dropped_files)
        in_l.addWidget(self.file_list)

        btn_row = QHBoxLayout()
        btn_row.setSpacing(8)
        self.btn_select = self._make_button(
            "选择文件", primary=False, icon_pm=QStyle.SP_FileIcon)
        self.btn_select.clicked.connect(self._select_files)
        self.btn_clear = self._make_button(
            "清空列表", primary=False, icon_pm=QStyle.SP_TrashIcon)
        self.btn_clear.clicked.connect(self._clear_files)
        btn_row.addStretch()
        btn_row.addWidget(self.btn_select)
        btn_row.addWidget(self.btn_clear)
        in_l.addLayout(btn_row)
        cl.addWidget(in_card)

        # ===== 输出目录卡片 =====
        out_card = Card()
        out_l = QVBoxLayout(out_card)
        out_l.setSpacing(10)
        out_l.setContentsMargins(16, 14, 16, 14)
        out_l.addWidget(_section_header("输出目录", QStyle.SP_DirIcon))

        out_row = QHBoxLayout()
        out_row.setSpacing(8)
        self.out_edit = QLabel("默认：源文件所在目录")
        self.out_edit.setStyleSheet(
            "QLabel{"
            f"  background:{C_BG};border:1px solid {C_BORDER};border-radius:6px;"
            f"  padding:9px 12px;font-size:12px;color:{C_TEXT};"
            "}"
        )
        self.btn_out = self._make_button(
            "选择目录", primary=False, icon_pm=QStyle.SP_DirOpenIcon)
        self.btn_out.clicked.connect(self._select_out)
        out_row.addWidget(self.out_edit, 1)
        out_row.addWidget(self.btn_out)
        out_l.addLayout(out_row)
        cl.addWidget(out_card)

        # ===== 进度条 + 状态 =====
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
            "QProgressBar{"
            f"  background:{C_BG};border:none;border-radius:11px;"
            f"  font-size:11px;color:{C_TEXT_2};text-align:center;"
            "}"
            "QProgressBar::chunk{"
            "  background:qlineargradient(x1:0,y1:0,x2:1,y2:0,"
            f"  stop:0 {C_PROG_A},stop:1 {C_PROG_B});"
            "  border-radius:11px;"
            "}"
        )
        prog_group.addWidget(self.progress)

        # 进度条平滑动画：源数据（COM 阻塞）不连续，靠动画过渡到目标值，观感连续
        self._prog_anim = QPropertyAnimation(self.progress, b"value")
        self._prog_anim.setDuration(500)
        self._prog_anim.setEasingCurve(QEasingCurve.Type.OutCubic)

        # 完成结果提示（仅在转换完成后显示，平时隐藏）
        self.status = QLabel("")
        self.status.setWordWrap(True)
        self.status.setStyleSheet(
            f"font-size:13px;font-weight:bold;color:{C_OK};"
            f"padding:6px 0;"
        )
        self.status.hide()
        prog_group.addWidget(self.status)
        cl.addLayout(prog_group)

        # ===== 底部大按钮区 =====
        btn_row2 = QHBoxLayout()
        btn_row2.setSpacing(12)
        self.btn_start = QPushButton("开始转换")
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

        btn_row2.addWidget(self.btn_start, 3)
        btn_row2.addWidget(self.btn_open_dir, 2)
        btn_row2.addWidget(self.btn_open, 2)
        cl.addLayout(btn_row2)

        # 初始文件列表提示
        self._refresh_file_list()
        self._update_start_state()

    # ----------------------- 按钮样式工厂 ----------------------- #
    def _btn_primary_style(self):
        return (
            "QPushButton{"
            f"  background:{C_ACCENT};"
            "  color:#FFFFFF;font-weight:bold;font-size:14px;"
            "  border:none;border-radius:6px;padding:10px 20px;"
            "}"
            f"QPushButton:hover{{background:{C_ACCENT_HOVER};}}"
            f"QPushButton:pressed{{background:{C_ACCENT_PRESS};}}"
            "QPushButton:disabled{"
            f"  background:#AEBECB;color:#F0F4F8;"
            "}"
        )

    def _btn_secondary_style(self):
        return (
            "QPushButton{"
            f"  background:#FFFFFF;color:{C_PRIMARY};"
            f"  border:1px solid {C_BORDER};"
            "  font-weight:bold;font-size:13px;border-radius:6px;padding:10px 20px;"
            "}"
            f"QPushButton:hover{{background:#EAF4FD;border-color:{C_ACCENT};}}"
            f"QPushButton:pressed{{background:#D6EAFB;}}"
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
                "QPushButton{"
                f"  background:{C_BG};color:{C_PRIMARY};"
                f"  border:1px solid {C_BORDER};"
                "  font-weight:500;font-size:12px;border-radius:6px;padding:8px 16px;"
                "}"
                f"QPushButton:hover{{background:#EAF4FD;border-color:{C_ACCENT};}}"
                f"QPushButton:pressed{{background:#D6EAFB;}}"
            )
        return btn

    # ----------------------------- 文件列表管理（合并控件） ------------- #
    def _refresh_file_list(self):
        """刷新文件列表显示。"""
        self.file_list.clear_files()
        for f in self.input_files:
            self.file_list.add_file_item(f)

    # ----------------------------- 文件选择 ----------------------------- #
    def _select_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "选择检验报告文件", "", FILTER_STR
        )
        if files:
            self._add_files(files)

    def _add_dropped_files(self, files):
        self._add_files(files)

    def _add_files(self, files):
        changed = False
        for f in files:
            if f not in self.input_files:
                self.input_files.append(f)
                changed = True
        if changed:
            self._refresh_file_list()
        self._update_start_state()

    def _clear_files(self):
        self.input_files = []
        self.file_list.clear_files()
        self._update_start_state()

    def _select_out(self):
        d = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if d:
            self.out_dir = d
            self.out_edit.setText(d)

    def _update_start_state(self):
        self.btn_start.setEnabled(bool(self.input_files) and not self._running())

    def _running(self):
        return self.worker is not None and self.worker.isRunning()

    # ----------------------------- 转换流程 ----------------------------- #
    def _start(self):
        if not self.input_files:
            QMessageBox.warning(self, "提示", "请先选择要转换的文件。")
            return
        self.btn_start.setEnabled(False)
        self.btn_start.setText("转换中…")
        self.btn_start.setIcon(QIcon(_style_icon(QStyle.SP_MediaPause, 18)))
        self.btn_open.setEnabled(False)
        self.progress.setValue(0)
        self.status.hide()

        default_year = self.year_edit.text().strip() or None
        self.worker = ConvertWorker(self.input_files, self.out_dir, default_year=default_year)
        self.worker.progress.connect(self._on_progress)
        self.worker.finished.connect(self._on_finished)
        self.worker.start()

    def _on_progress(self, pct, msg):
        # 平滑过渡到目标进度，避免 30%→100% 的跳变观感
        self._prog_anim.stop()
        self._prog_anim.setStartValue(self.progress.value())
        self._prog_anim.setEndValue(pct)
        self._prog_anim.start()

    def _on_finished(self, results):
        ok = sum(1 for r in results if r["ok"])
        fail = len(results) - ok
        # 记录成功输出的文件路径，便于"打开输出文件"按钮
        self.last_outputs = [
            r.get("output") for r in results
            if r.get("ok") and r.get("output") and os.path.isfile(r["output"])
        ]
        self.btn_start.setText("开始转换")
        self.btn_start.setIcon(QIcon(_style_icon(QStyle.SP_MediaPlay, 18)))
        self.btn_open.setEnabled(bool(self.last_outputs))
        self.btn_open_dir.setEnabled(True)

        # 界面内显示结果（不弹窗）
        lines = []
        for r in results:
            bn = os.path.basename(r["file"])
            if r["ok"]:
                lines.append(f"✔ {bn}")
            else:
                lines.append(f"✘ {bn}：{r.get('error', '')}")

        detail = "  |  ".join(lines)
        if fail == 0:
            self.status.setStyleSheet(
                f"font-size:13px;font-weight:bold;color:{C_OK};padding:6px 0;"
            )
            self.status.setText(f"✅ 转换完成：成功 {ok} 个{detail}")
        else:
            self.status.setStyleSheet(
                f"font-size:13px;font-weight:bold;color:{C_ERR};padding:6px 0;"
            )
            self.status.setText(f"⚠️ 转换完成：成功 {ok} 个，失败 {fail} 个{detail}")
        self.status.show()

        # 转换完成后自动清空拖入框，便于继续拖入新文件
        self.input_files = []
        self.file_list.clear_files()
        self.btn_start.setEnabled(False)

    def _open_dir(self):
        """打开输出目录。优先用户选择的目录，其次源文件所在目录。"""
        target = self.out_dir or (
            os.path.dirname(self.last_outputs[0]) if self.last_outputs
            else (os.path.dirname(self.input_files[0]) if self.input_files else None)
        )
        if target and os.path.isdir(target):
            self._open_path(target)

    def _open_out(self):
        """打开最近一次转换的输出文件。多个时弹出菜单选择。"""
        if not self.last_outputs:
            # 没有最近输出，回退到打开输出目录
            self._open_dir()
            return

        if len(self.last_outputs) == 1:
            self._open_path(self.last_outputs[0])
            return

        # 多个文件：弹菜单选择
        from PySide6.QtWidgets import QMenu
        menu = QMenu(self)
        for p in self.last_outputs:
            act = menu.addAction(os.path.basename(p))
            act.triggered.connect(lambda checked=False, path=p: self._open_path(path))
        # 在按钮正下方弹出
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

    def _edit_year(self):
        """点击年份框进入编辑模式：临时解除只读，编辑后焦点离开恢复只读。"""
        self.year_edit.setReadOnly(False)
        self.year_edit.setCursor(Qt.IBeamCursor)
        self.year_edit.setFocus()
        self.year_edit.selectAll()
        # 失焦后恢复只读
        orig_lost = self.year_edit.focusOutEvent

        def lost_focus(event, _orig=orig_lost):
            try:
                _orig(event)
            except Exception:
                pass
            self.year_edit.setReadOnly(True)
            self.year_edit.setCursor(Qt.PointingHandCursor)

        self.year_edit.focusOutEvent = lost_focus


# --------------------------------------------------------------------------- #
# 入口
# --------------------------------------------------------------------------- #
def main():
    # 高 DPI 缩放（Qt6 默认开启，旧属性已废弃，安全设置）
    for attr in ("AA_EnableHighDpiScaling", "AA_UseHighDpiPixmaps"):
        if hasattr(Qt, attr):
            QApplication.setAttribute(getattr(Qt, attr), True)

    app = QApplication(sys.argv)
    app.setApplicationName("检验报告格式转换工具")
    app.setStyleSheet("QMessageBox{font-family:'Microsoft YaHei','Segoe UI';}")
    win = ConverterApp()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
