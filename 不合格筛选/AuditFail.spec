# -*- mode: python ; coding: utf-8 -*-
import os
from PyInstaller.utils.hooks import collect_all

# 用 spec 所在目录定位源码与 VERSION，避免在公开仓库里写死本机绝对路径。
_ROOT = globals().get("SPECPATH") or os.getcwd()


def _keep(path):
    """保留：供显式 collect_all 的 datas/binaries 按源路径过滤。

    注意：collect_all 返回 (源绝对路径, 目标相对路径) 二元组，目标路径恒为
    "PySide6"/"win32com" 等目录名，按 d[1] 过滤永远命中不了关键字；必须取
    第 0 位源绝对路径才能匹配 qt6webengine 等文件片段。
    """
    p = (path or "").replace("\\", "/").lower()
    bad = [
        "qt6webengine", "qt6qml", "qt6quick", "qt6quick3d",
        "qt63d", "qt6multimedia", "qt6multimediawidgets",
        "qt6charts", "qt6datavisualization",
        "qt6network", "qt6sql", "qt6test",
        "qt6websockets", "qt6webchannel",
        "qt6help", "qt6designer", "qt6uitools",
        "qt6sensors", "qt6serialport", "qt6bluetooth",
        "qt6nfc", "qt6positioning", "qt6location",
        "qt6webview", "qt6opengl", "qt6xml", "qt6svg",
        "qt6pdf", "qt6remoteobjects", "qt6scxml",
        "qt6spatialaudio", "qt6texttospeech", "qt6grpc",
        "qt6protobuf", "qt6coap", "qt6mqtt",
        "qt6amlogic", "qt6bcm", "qt6evdev", "qt6libinput", "qt6tslib",
    ]
    return not any(b in p for b in bad)


def _keep_mod(name):
    """hiddenimports 过滤：只保留 QtWidgets/Core/Gui/shiboken6，排除其余 PySide6 子模块。

    这是瘦身关键：collect_all('PySide6') 的 hiddenimports 若不过滤，PyInstaller
    会分析全部 Qt 子模块（QtWebEngineCore/Qt3D/QtQuick/QtCharts 等），其 hook
    又把全量 DLL 收集进包——之前 263MB 的根源。
    """
    n = (name or "").lower()
    if not n.startswith("pyside6."):
        return True   # shiboken6 / 非 Qt 模块照收
    keep = ("pyside6.qtcore", "pyside6.qtgui", "pyside6.qtwidgets")
    return n.startswith(keep)


def _keep_win32(path):
    """win32com 收集去掉 demos/test 子目录（体积无用项）。按源路径/模块名检查。"""
    p = (path or "").replace("\\", "/").lower()
    if "/demos/" in p or "/test/" in p:
        return False
    n = p.replace("/", ".")
    return not (".demos." in n or ".test." in n)


datas = [(os.path.join(_ROOT, 'VERSION'), '.')]
binaries = []
# A1.6 根因修复：win32timezone 的 import 来源只有 win32com\test\ 下的测试模块，
# _keep_win32 过滤 test/demos 后 Analysis 不再自动收集它，而 pythoncom.Invoke 的
# 运行时路径仍需要（缺失时 COM 属性读取抛 ModuleNotFoundError → 宽表整份漏规范化）。
# 显式声明为 hiddenimport，保证 win32\lib\win32timezone.py 进 PYZ。
hiddenimports = ['win32com', 'pythoncom', 'win32timezone']
# 收集 win32com / pythoncom（过滤 demos/test，按源路径第 0 位）
for _m in ('win32com', 'pythoncom'):
    try:
        _t = collect_all(_m)
        datas += [x for x in _t[0] if _keep_win32(x[0])]
        binaries += [x for x in _t[1] if _keep_win32(x[0])]
        hiddenimports += [h for h in _t[2] if _keep_win32(h)]
    except Exception:
        pass


# Analysis 层排除：双保险，即使 hook 尝试收集这些模块也不进包
QT_EXCLUDES = [
    'PySide6.Qt3DAnimation', 'PySide6.Qt3DCore', 'PySide6.Qt3DExtras',
    'PySide6.Qt3DInput', 'PySide6.Qt3DLogic', 'PySide6.Qt3DRender',
    'PySide6.QtCharts', 'PySide6.QtDataVisualization', 'PySide6.QtGraphs',
    'PySide6.QtMultimedia', 'PySide6.QtMultimediaWidgets',
    'PySide6.QtQuick', 'PySide6.QtQuick3D', 'PySide6.QtQuickControls2',
    'PySide6.QtQuickLayouts', 'PySide6.QtQuickWidgets',
    'PySide6.QtQml', 'PySide6.QtQmlModels', 'PySide6.QtQmlWorkerScript',
    'PySide6.QtWebEngineCore', 'PySide6.QtWebEngineWidgets', 'PySide6.QtWebEngineQuick',
    'PySide6.QtWebChannel', 'PySide6.QtWebSockets', 'PySide6.QtWebView',
    'PySide6.QtSql', 'PySide6.QtTest', 'PySide6.QtSensors',
    'PySide6.QtSerialBus', 'PySide6.QtSerialPort', 'PySide6.QtBluetooth',
    'PySide6.QtNfc', 'PySide6.QtPositioning', 'PySide6.QtLocation',
    'PySide6.QtHelp', 'PySide6.QtDesigner', 'PySide6.QtUiTools',
    'PySide6.QtXml', 'PySide6.QtXmlPatterns', 'PySide6.QtSvg',
    'PySide6.QtOpenGL', 'PySide6.QtOpenGLWidgets', 'PySide6.QtOpenGLFunctions',
    'PySide6.QtPdf', 'PySide6.QtPdfWidgets', 'PySide6.QtRemoteObjects',
    'PySide6.QtScxml', 'PySide6.QtStateMachine', 'PySide6.QtTextToSpeech',
    'PySide6.QtSpatialAudio', 'PySide6.QtGrpc', 'PySide6.QtProtobuf',
    'PySide6.QtCoap', 'PySide6.QtMqtt', 'PySide6.QtNetworkAuth',
    'PySide6.QtHttpServer', 'PySide6.QtLinguist', 'PySide6.QtDBus',
    'PySide6.QtConcurrent', 'PySide6.QtShaderTools', 'PySide6.QtWaylandClient',
    'PySide6.QtWaylandCompositor', 'PySide6.QtVirtualKeyboard',
    'PySide6.QtDesignerComponents',
    'PySide6.QtNetwork', 'PySide6.QtSvgWidgets',
]


a = Analysis(
    [os.path.join(_ROOT, 'audit_fail_gui.py')],
    pathex=[_ROOT],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=QT_EXCLUDES,
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='AuditFail_v19',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=True,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

