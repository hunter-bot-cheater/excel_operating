# -*- coding: utf-8 -*-
"""对外统一接口（网页端 / GUI 共用）。

设计要点（前后端分离、预留网页端）：
  - 不依赖 GUI、不弹窗、纯函数。
  - 流程方向：**报告池驱动**（非 BOM 驱动）。
      · 报告文件池中的每一个文件，用其文件名前半部分的「图号」去 BOM 匹配；
      · 图号未在 BOM 中匹配到（或对应数量为空）→ 该文件单独报 NO_BOM_MATCH 失败；
      · 只有在 BOM 中匹配到的图号才会被处理（按 BOM 行拆成逐件任务）；
      · BOM 中未被任何报告文件覆盖的图号 → 静默跳过，不报错、不输出。
  - 同一图号若有多个 BOM 行（多件），则按「扣减基准文件算法」为每个任务挑选来源：
    以剩余升序逐个扣减、能用满就用满；**单个文件不够时由多份基准文件共同承担**
    （如剩 5 + 剩 10，任务需 6 → 5 用满 + 从 10 取 1）。合计仍不足才报 NO_REMAINING。
  - 逐任务回调 on_task(dict) 并实时回调 on_progress(pct, msg)；单任务失败隔离、
    事务回滚（来源文件不变、新文件不生成）。
  - 输出 JSON 可序列化 dict：{ok_count, fail_count, tasks, file_errors, skipped_rows,
      bom_write_errors, bom_backup}
      每个 task/file = {ok, kind, bh?, qty?, output?, source_after?, error_code?, reason?, file?}
      bom_backup 恒为 None（用户要求取消 .bak 备份，仅保留字段兼容旧脚本）。
  - round_no>0 时启用 BOM 回写：每个成功任务把 `日期/件数`（日期=基准文件名日期，
      件数=本次 K）写入 BOM「第N套」列 × 该图号行；列不存在则自动插入。回写失败走
      bom_write_errors 单列告警，不计入任务失败。
  - 网页端（FastAPI/Flask）只需：
        summary = extract_core.api.run_batch(bom, files, out_dir)
        # ok=True 的 task["output"] 作为下载；ok=False 的 task["reason"] 展示。
"""
import os
from collections import OrderedDict
from typing import Callable, Dict, List, Optional

from .extractor import run_one_task_multi, create_excel_app, quit_excel_app
from .errors import ExtractError
from .filename_parser import parse_report_filename
from .bom_reader import read_bom_ex, match_report, find_round_col
from .bom_writer import BomWriter
from .report_reader import detect_layout, extract_identity


def _read_round_existing(bom_path, meta, round_no):
    """预读 BOM 主表「第round_no套」列各行已有值，返回 {1-based行号: 非空文本}。

    用于跳过「本套已做过」的行：该行此列已有日期/件数 -> 不再扣除、不生成新文件，
    只在结果里黄色提醒。读不到/无此列 -> 返回空 dict（按未做过处理）。
    """
    col = find_round_col(meta, round_no)
    if not col:
        return {}
    try:
        import openpyxl
        wb = openpyxl.load_workbook(bom_path, data_only=True, read_only=True)
    except Exception:  # noqa: BLE001
        return {}
    out = {}
    try:
        if meta.sheet not in wb.sheetnames:
            return {}
        ws = wb[meta.sheet]
        for i, row_cells in enumerate(
                ws.iter_rows(min_row=1, max_row=ws.max_row,
                             min_col=col, max_col=col, values_only=True), start=1):
            v = row_cells[0]
            if v is not None and str(v).strip():
                out[i] = str(v).strip()
    except Exception:  # noqa: BLE001
        pass
    finally:
        try:
            wb.close()
        except Exception:  # noqa: BLE001
            pass
    return out


def _plan_segments(cands: List[dict], K: int) -> Optional[List[tuple]]:
    """把本任务的 K 件拆分到多个候选来源（扣减基准文件算法）。

    规则（用户需求）：
      - 以「剩余升序」逐个扣减：先用剩余最少的文件，**能用满就用满**（用完一个再下一个）；
      - 若某个文件剩余不足 K，则把它**全部扣完**，再从下一个文件补足差额；
      - 一次任务可由多份基准文件共同承担（如剩 5 件 + 剩 10 件，任务需 6 件
        → 5 件用满 + 从 10 件里取 1 件 = 10 件那份剩 9 件）。
      - 目标：尽可能把基准文件用满、让「有剩余（没用完）的文件数」尽可能少。
      - 所有候选文件剩余合计 < K → 返回 None（调用方报 NO_REMAINING）。
    返回 [(candidate, count), ...]（sum(count) == K，按扣减顺序）。
    """
    avail = [c for c in cands if (c["total"] - c["used"]) > 0]
    if sum(c["total"] - c["used"] for c in avail) < K:
        return None
    # 剩余升序（相同剩余保持拖入顺序）；用满小文件、减少「没用完」的文件数
    order = sorted(avail, key=lambda c: (c["total"] - c["used"]))
    segs = []
    left = K
    for c in order:
        if left <= 0:
            break
        rem = c["total"] - c["used"]
        take = min(rem, left)
        if take > 0:
            segs.append((c, take))
            left -= take
    if left > 0:
        return None
    return segs


def run_batch(
    bom_path: str,
    report_paths: List[str],
    out_dir: Optional[str] = None,
    round_no: Optional[int] = None,
    on_progress: Optional[Callable[[int, str], None]] = None,
    on_task: Optional[Callable[[dict], None]] = None,
) -> dict:
    """报告池驱动批量抽取。返回汇总 dict；逐任务（含文件级失败）通过 on_task 实时回调。

    round_no：当前是第几套（>0 时启用 BOM 回写：把 `日期/件数` 写入 BOM 的「第N套」列
    × 对应图号行；列不存在则自动插入）。为 None 时不回写 BOM。
    """
    def prog(p, msg):
        if on_progress:
            on_progress(p, msg)

    def emit(res):
        if on_task:
            on_task(res)

    # 1) BOM（只读；同时取表头结构供回写定位「第N套」列）
    try:
        bom_rows, bom_meta = read_bom_ex(bom_path)
    except ExtractError as e:
        e_dict = {"ok": False, "kind": "bom", "error_code": e.code,
                  "reason": e.reason, "file": getattr(e, "file", bom_path)}
        emit(e_dict)
        return {"ok_count": 0, "fail_count": 1, "tasks": [e_dict],
                "file_errors": [e_dict], "skipped_rows": [], "bom_write_errors": [],
                "bom_backup": None}

    # 2) 解析报告候选池 + 分类（报告池驱动的核心）
    #    匹配用 match_report 多信号交叉验证：
    #      内部 DWG No / 零件名称（openpyxl 轻量读取）→ 文件名 图号_描述 整体
    candidates: List[dict] = []
    file_errors: List[dict] = []
    bom_write_errors: List[dict] = []
    matched_bh: List[str] = []
    results: List[dict] = []
    bom_writer = None
    matched_rows: "OrderedDict[str, List]" = OrderedDict()
    for p in report_paths:
        try:
            pr = parse_report_filename(p)
        except ExtractError as e:
            fe = {"ok": False, "kind": "file", "file": p,
                  "error_code": e.code, "reason": e.reason}
            file_errors.append(fe)
            results.append(fe)
            emit(fe)
            continue
        # 报告内部身份（图纸真值）：图号 DWG No + 零件名称；读不到则回退纯文件名匹配
        dwg, part = extract_identity(p)
        try:
            rows = match_report(pr.bh, bom_rows, dwg=dwg, part=part)
        except ExtractError as e:
            # 匹配歧义（如分段集合命中多个不同行）也按文件级失败单报
            fe = {"ok": False, "kind": "file", "file": p, "error_code": e.code,
                  "bh": pr.bh, "reason": e.reason}
            file_errors.append(fe)
            results.append(fe)
            emit(fe)
            continue
        if not rows:
            # 只有「放入的图号未在 BOM 匹配到」才失败；其它 BOM 图号静默跳过
            fe = {"ok": False, "kind": "file", "file": p, "error_code": "NO_BOM_MATCH",
                  "bh": pr.bh,
                  "reason": f"报告「{pr.bh}」未在 BOM 中匹配到（内部图号/文件名图号+描述联合/图号前缀/分段集合/数字键均未命中，或对应数量为空）"}
            file_errors.append(fe)
            results.append(fe)
            emit(fe)
            continue
        remaining = pr.total - pr.used
        candidates.append({
            "path": os.path.abspath(p),
            "bh": pr.bh,
            "total": pr.total,
            "used": pr.used,
            "remaining": remaining,
            "exhausted": remaining <= 0,
            "parse": pr,
        })
        if pr.bh not in matched_bh:
            matched_bh.append(pr.bh)
            matched_rows[pr.bh] = rows

    # 3) 仅对「报告覆盖到的 front」建立任务组（BOM 其余行跳过，不报错）
    groups: "OrderedDict[str, List[tuple]]" = OrderedDict()
    for bh in matched_bh:
        groups[bh] = list(enumerate(matched_rows[bh]))

    total_tasks = sum(len(v) for v in groups.values())
    done = 0

    # 预读「第N套」列已有值：该列非空 = 本套已做过，任务直接跳过并黄色提醒
    round_existing = _read_round_existing(bom_path, bom_meta, round_no) if round_no else {}

    # 批次级共享资源：一个 Excel 实例跑完整批（避免每任务重启 Excel）；
    # 布局探测结果按候选文件缓存（同一文件多任务只解析一次）。
    xl_app = None
    try:
        for bh, rows in groups.items():
            # 先对该组候选做布局探测：失败即单文件隔离（报告失败、排除出候选），
            # 不再像旧版那样把异常冒泡出 run_batch 导致整批崩溃。
            bh_cands = []
            for c in candidates:
                if c["bh"] != bh or c["exhausted"]:
                    continue
                if "info" not in c:
                    try:
                        c["info"] = detect_layout(c["path"])
                    except ExtractError as e:
                        fe = {"ok": False, "kind": "file", "file": c["path"],
                              "error_code": e.code, "bh": bh, "reason": e.reason}
                        file_errors.append(fe)
                        results.append(fe)
                        emit(fe)
                        c["exhausted"] = True  # 排除，后续任务不再选用该文件
                        continue
                bh_cands.append(c)

            for gi, (idx, row) in enumerate(rows):
                K = row.qty
                done += 1
                overall_before = int((done - 1) / total_tasks * 100) if total_tasks else 0
                overall_after = int(done / total_tasks * 100) if total_tasks else 100

                # 本套已做过：BOM 该行「第N套」列已有内容 -> 不扣除、不生成、不回写，黄色提醒
                old_val = round_existing.get(row.row)
                if old_val:
                    res = {"ok": False, "kind": "already", "task_index": idx,
                           "bh": bh, "qty": K, "error_code": "ALREADY_DONE",
                           "existing": old_val, "round_no": round_no,
                           "reason": f"第{round_no}套该行已有记录「{old_val}」，本套已做过，"
                                     f"跳过扣除（未生成新文件、未改动基准文件）"}
                    results.append(res)
                    emit(res)
                    prog(overall_after, f"任务 {done}/{total_tasks}：⚠️ 本套已做过")
                    continue

                # K 无硬上限：K>5 时生成多张表（见 extractor.SHEET_MAX）。
                # 扣减计划：把本任务 K 件拆分到多个候选来源（剩余升序，逐个用满）。
                segs = _plan_segments(bh_cands, K)
                if segs is None:
                    rem_info = ", ".join(
                        f"{os.path.basename(c['path'])}剩{c['total']-c['used']}"
                        for c in bh_cands
                    ) or "无匹配文件"
                    res = {"ok": False, "kind": "task", "task_index": idx, "bh": bh,
                           "qty": K, "error_code": "NO_REMAINING",
                           "reason": f"图号 {bh} 所有基准文件余量合计不足（需要 {K} 件；{rem_info}）"}
                    results.append(res)
                    emit(res)
                    prog(overall_after, f"任务 {done}/{total_tasks}：余量不足")
                    continue

                def _task_progress(pct, msg, _ob=overall_before, _oa=overall_after):
                    overall = int(_ob + (pct / 100.0) * (_oa - _ob))
                    prog(overall, f"任务 {done}/{total_tasks}：{msg}")

                # 懒创建共享 Excel（布局已在上方预探测，每文件仅一次）
                if xl_app is None:
                    try:
                        xl_app = create_excel_app()
                    except Exception as e:  # noqa: BLE001
                        res = {"ok": False, "kind": "task", "task_index": idx, "bh": bh,
                               "qty": K, "error_code": "COM_ERROR",
                               "reason": f"Excel 启动失败：{e}", "file": segs[0][0]["path"]}
                        results.append(res)
                        emit(res)
                        prog(overall_after, f"任务 {done}/{total_tasks}：❌ COM_ERROR")
                        continue

                try:
                    seg_input = [{"path": c["path"], "info": c.get("info"),
                                  "used": c["used"], "count": cnt} for (c, cnt) in segs]
                    res = run_one_task_multi(seg_input, out_dir,
                                             on_progress=_task_progress, xl=xl_app)
                    res["kind"] = "task"
                    res["task_index"] = idx
                    # 提交成功后更新各候选状态（供后续任务复用）
                    for (c, _cnt), sout in zip(segs, res.get("segments", [])):
                        c["used"] = sout["used_to"]
                        c["remaining"] = c["total"] - c["used"]
                        if sout.get("source_after"):
                            c["path"] = sout["source_after"]
                        if c["remaining"] <= 0:
                            c["exhausted"] = True

                    # ---- BOM 回写：把「日期/件数」写入「第N套」列 × 该图号行 ----
                    # BomWriter 自带独立 Excel 实例，与任务抽取隔离；只要启用了 round_no 即回写。
                    # write_row 会定位该图号在**所有**工作表里的行并逐表写入（多表副本同步）。
                    if round_no:
                        if bom_writer is None:
                            # 复用任务抽取的 Excel 实例，不再另起第二个（避免两实例并发、
                            # 收尾 Workbooks.Close 互相拖累卡十几秒）
                            bom_writer = BomWriter(bom_path, round_no, bom_meta, xl_app=xl_app)
                        try:
                            bom_writer.write_row(row, res.get("date", ""), K)
                        except ExtractError as e:
                            wr = {"ok": False, "kind": "bom_write", "bh": bh, "qty": K,
                                  "error_code": e.code, "reason": e.reason, "file": bom_path}
                            bom_write_errors.append(wr)
                            results.append(wr)
                            emit(wr)
                        except Exception as e:  # noqa: BLE001
                            wr = {"ok": False, "kind": "bom_write", "bh": bh, "qty": K,
                                  "error_code": "BOM_WRITE_ERROR", "reason": str(e)[:200],
                                  "file": bom_path}
                            bom_write_errors.append(wr)
                            results.append(wr)
                            emit(wr)

                    results.append(res)
                    emit(res)
                    prog(overall_after, f"任务 {done}/{total_tasks}：✅ {bh} {K}件")
                except ExtractError as e:
                    res = {"ok": False, "kind": "task", "task_index": idx, "bh": bh,
                           "qty": K, "error_code": e.code, "reason": e.reason,
                           "file": getattr(e, "file", segs[0][0]["path"])}
                    results.append(res)
                    emit(res)
                    prog(overall_after, f"任务 {done}/{total_tasks}：❌ {e.code}")
                    # Excel 实例异常时重建，保证后续任务可用
                    if e.code in ("COM_ERROR", "SAVE_ERROR") and xl_app is not None:
                        quit_excel_app(xl_app)
                        xl_app = None
                except Exception as e:  # noqa: BLE001
                    res = {"ok": False, "kind": "task", "task_index": idx, "bh": bh,
                           "qty": K, "error_code": "UNKNOWN", "reason": str(e)[:200],
                           "file": segs[0][0]["path"]}
                    results.append(res)
                    emit(res)
    finally:
        # 先保存并关闭 BOM（回写），再退出 Excel
        if bom_writer is not None:
            try:
                bom_writer.close()
            except ExtractError as e:
                wr = {"ok": False, "kind": "bom_write", "error_code": e.code,
                      "reason": e.reason, "file": bom_path}
                bom_write_errors.append(wr)
                results.append(wr)
                emit(wr)
            except Exception as e:  # noqa: BLE001
                wr = {"ok": False, "kind": "bom_write", "error_code": "BOM_WRITE_ERROR",
                      "reason": str(e)[:200], "file": bom_path}
                bom_write_errors.append(wr)
                results.append(wr)
                emit(wr)
        if xl_app is not None:
            quit_excel_app(xl_app)

    # results 中已含 file_errors（同一对象），fail_count 自然计入；
    # 回写告警（kind=bom_write）不计入任务失败数
    return {
        "ok_count": sum(1 for r in results if r.get("ok")),
        "fail_count": sum(1 for r in results
                          if not r.get("ok") and r.get("kind") not in ("bom_write", "already")),
        "tasks": results,
        "file_errors": file_errors,
        "skipped_rows": [],
        "bom_write_errors": bom_write_errors,
        "bom_backup": None,
    }


def process_report(
    report_path: str,
    bom_path: str,
    out_dir: Optional[str] = None,
    round_no: Optional[int] = None,
    on_progress: Optional[Callable[[int, str], None]] = None,
) -> dict:
    """兼容接口：单份报告 = 一个任务（用文件名图号去 BOM 匹配）。

    实际复用 run_batch（报告池驱动），返回其汇总 dict。
    """
    return run_batch(bom_path, [report_path], out_dir=out_dir, round_no=round_no,
                     on_progress=on_progress)
