#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
审核不合格数据 —— Web 后端 API（前后端分离版）
============================================

把 audit_fail.py 核心引擎包成 REST API，供网页前端 / 公司网站 / ERP 等系统调用。

接口（三件套 = 对接合同）：
    POST /api/audit                    上传一个或多个报告，创建审核任务
    GET  /api/tasks/{task_id}          轮询任务进度与每个文件的结论
    GET  /api/tasks/{task_id}/download/{filename}   下载审核结果文件
    GET  /api/health                   健康检查
    GET  /                             Web 前端页面（web/index.html）

设计要点：
  * 核心逻辑全部来自 audit_fail.py（判定/标红/命名/双引擎），本文件只做 API 壳。
  * 任务队列：全局单 worker 线程串行处理，与桌面版共用 Excel COM 会话；
    Excel 单线程，COM 路径必须串行；openpyxl 快路径天然毫秒级，串行也足够快。
  * 文件存储：webdata/{task_id}/uploads 与 outputs，启动时自动清理超过 24h 的旧任务。
  * 部署约束：Excel COM 依赖 Windows + 已安装 Excel（有公式文件才需要）。
  * 注意：uvicorn 请以单 worker 运行（--workers 1，默认），多 worker 会导致
    任务状态与 COM 会话各自独立、互相不认。

启动（两种方式）：
    python api.py                 # 内置 uvicorn，等价下方命令
    python -m uvicorn api:app --host 127.0.0.1 --port 8000
"""
import os
import shutil
import threading
import time
import uuid
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

import audit_fail

# ---------------------------------------------------------------- 路径与常量
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "webdata"
WEB_DIR = BASE_DIR / "web"
DATA_DIR.mkdir(exist_ok=True)

SUPPORTED_EXTS = {".xls", ".xlsx", ".xlsm"}
MAX_FILE_SIZE = 100 * 1024 * 1024  # 100MB
TASK_TTL_SECONDS = 24 * 3600       # 任务产物保留 24 小时
POLL_OK = ("pending", "running")

app = FastAPI(title="审核不合格数据 API", version="1.0")

# ---------------------------------------------------------------- 任务存储与队列
_tasks = {}                    # task_id -> dict
_queue_lock = threading.Lock()
_queue = []                    # FIFO 任务队列
_queue_cond = threading.Condition(_queue_lock)
_session = audit_fail.ExcelSession()   # 进程内共享的 Excel COM 会话（懒启动）


def _cleanup_old_tasks():
    """清理超过 TTL 的旧任务目录、内存记录与预加载缓存。"""
    now = time.time()
    with _queue_lock:
        for tid, t in list(_tasks.items()):
            if now - t["created"] > TASK_TTL_SECONDS:
                _tasks.pop(tid, None)
                d = DATA_DIR / tid
                if d.exists():
                    audit_fail.clear_preloaded(lambda p, _d=str(d): p.startswith(_d))
                    shutil.rmtree(d, ignore_errors=True)


def _worker():
    """单 worker：串行处理队列中的任务，保证 COM 会话安全。"""
    while True:
        with _queue_cond:
            while not _queue:
                _queue_cond.wait()
            task_id = _queue.pop(0)
        try:
            _process_task(task_id)
        except Exception as e:  # noqa: BLE001
            t = _tasks.get(task_id)
            if t:
                t["status"] = "error"
                t["message"] = f"任务异常: {e}"
        finally:
            _cleanup_old_tasks()


def _process_task(task_id):
    t = _tasks[task_id]
    t["status"] = "running"
    t["message"] = "审核中…"
    task_dir = DATA_DIR / task_id
    uploads = task_dir / "uploads"
    outputs = task_dir / "outputs"
    total = max(len(t["files"]), 1)

    for i, f in enumerate(t["files"]):
        src = uploads / f["name"]

        def on_progress(pct, msg, _f=f, _i=i, _total=total):
            t["progress"] = int((_i + pct / 100.0) / _total * 100)
            t["message"] = f"[{_i + 1}/{_total}] {_f['name']} - {msg}"

        try:
            res = audit_fail.audit_file(str(src), out_dir=str(outputs),
                                        on_progress=on_progress, session=_session)
            f.update(
                status="done",
                conclusion="不合格" if res["is_fail"] else "合格",
                reason=res.get("reason", ""),
                output=os.path.basename(res["output"]),
                norm=res.get("norm", {}),
                log=res.get("log", []),
            )
        except Exception as e:  # noqa: BLE001
            f.update(status="error", conclusion="失败",
                     reason=str(e)[:200], output="")
        t["progress"] = int((i + 1) / total * 100)
        t["message"] = f"已处理 {i + 1}/{total}"

    t["status"] = "done"
    t["message"] = "全部完成"
    t["progress"] = 100


threading.Thread(target=_worker, daemon=True, name="audit-worker").start()


# ---------------------------------------------------------------- API 路由
@app.get("/api/health")
def api_health():
    return {"status": "ok", "engine": "v11 双引擎+规范化+异常检测", "time": time.time()}


@app.post("/api/audit")
async def api_audit(files: list[UploadFile] = File(...)):
    """上传一个或多个 Excel 报告，创建审核任务，返回 task_id。"""
    if not files:
        raise HTTPException(400, "未收到文件")

    task_id = uuid.uuid4().hex[:12]
    task_dir = DATA_DIR / task_id
    uploads = task_dir / "uploads"
    outputs = task_dir / "outputs"
    uploads.mkdir(parents=True)
    outputs.mkdir()

    flist = []
    for up in files:
        raw = up.filename or ""
        name = os.path.basename(raw.replace("\\", "/"))
        if not name:
            raise HTTPException(400, "存在空文件名的上传项")
        ext = os.path.splitext(name)[1].lower()
        if ext not in SUPPORTED_EXTS:
            shutil.rmtree(task_dir, ignore_errors=True)
            raise HTTPException(400, f"不支持的文件类型: {name}（仅支持 .xls/.xlsx/.xlsm）")
        content = await up.read()
        if len(content) > MAX_FILE_SIZE:
            shutil.rmtree(task_dir, ignore_errors=True)
            raise HTTPException(400, f"文件超过 100MB 限制: {name}")

        dst = uploads / name
        n = 1
        while dst.exists():
            base, ext = os.path.splitext(name)
            dst = uploads / f"{base}({n}){ext}"
            n += 1
        dst.write_bytes(content)
        flist.append({"name": dst.name, "status": "pending", "conclusion": "",
                      "reason": "", "output": ""})

    t = {"id": task_id, "status": "pending", "progress": 0,
         "message": "排队中", "files": flist, "created": time.time()}
    with _queue_lock:
        _tasks[task_id] = t
    with _queue_cond:
        _queue.append(task_id)
        _queue_cond.notify()
    # 上传后立即后台预打开（openpyxl 快路径），worker 处理时命中缓存省去加载
    for f in flist:
        threading.Thread(target=audit_fail.preload_openpyxl,
                         args=(str(uploads / f["name"]),),
                         daemon=True).start()
    return {"task_id": task_id, "count": len(flist)}


@app.get("/api/tasks/{task_id}")
def api_task(task_id: str):
    t = _tasks.get(task_id)
    if not t:
        raise HTTPException(404, "任务不存在（可能已超过 24 小时被清理）")
    return {k: t[k] for k in ("id", "status", "progress", "message", "files")}


@app.get("/api/tasks/{task_id}/download/{filename}")
def api_download(task_id: str, filename: str):
    t = _tasks.get(task_id)
    if not t:
        raise HTTPException(404, "任务不存在")
    safe = os.path.basename(filename)
    if safe != filename or not safe:
        raise HTTPException(400, "非法文件名")
    p = DATA_DIR / task_id / "outputs" / safe
    if not p.is_file():
        raise HTTPException(404, "输出文件不存在")
    return FileResponse(str(p), filename=safe)


# Web 前端页面（挂载在最后，API 路由优先匹配）
if WEB_DIR.is_dir():
    app.mount("/", StaticFiles(directory=str(WEB_DIR), html=True), name="web")


# ---------------------------------------------------------------- 入口
if __name__ == "__main__":
    import uvicorn

    print("审核不合格数据 Web 服务启动: http://127.0.0.1:8000")
    print("健康检查: http://127.0.0.1:8000/api/health")
    uvicorn.run(app, host="127.0.0.1", port=8000)
