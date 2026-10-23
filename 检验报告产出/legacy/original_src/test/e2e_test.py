# -*- coding: utf-8 -*-
"""端到端测试（新工作流）：BOM 行驱动 + 智能调度 + 事务回滚。

覆盖：
  1. 贪心用满换文件（同图号 2/3/1 件，两文件按序消耗至 -0）
  2. 智能调度：两文件都能承接全部 -> 选剩余最少者
  3. 余量不足 -> NO_REMAINING
  4. 无后缀 = 全新（已用 0）
  5. -0 文件被跳过（余量 0）
  6. 任务失败回滚（SAMPLE_NOT_FOUND：来源不改名、不生成新文件）
  7. K > 5 -> K_TOO_LARGE
"""
import os
import shutil
import sys
import time

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
TMP = os.path.join(HERE, f"sandbox_{int(time.time() * 1000)}")
sys.path.insert(0, os.path.dirname(HERE))
from extract_core.api import run_batch

BH = "930TEST-DM"
DESC = "测试件"
DATE = "26.4.09"


def build_report(out_dir, used, total, nsheets=5, name=None):
    """生成合成多表报告。used=None 表示无后缀（全新）。"""
    if used is None:
        fname = f"{BH}_{DESC}_{total}件_{DATE}(1).xlsx"
    else:
        fname = f"{BH}_{DESC}_{total}件_{DATE}-{used}(1).xlsx"
    if name:
        fname = name
    path = os.path.join(out_dir, fname)
    wb = openpyxl.Workbook()
    for k in range(1, nsheets + 1):
        ws = wb.create_sheet(f"2026.4.9（{total}件）-{k}")
        ws["A1"] = "Inspection Report\n检验报告"
        ws["A3"] = "DWG No:\n图号"; ws["B3"] = BH
        ws["D3"] = "Part Description:\n零件名称"; ws["F3"] = DESC
        ws["A9"] = "Item No\n（编号）"; ws["B9"] = "Specified Parameter"
        ws["E9"] = "Instrument used（检具）"
        ws["F9"] = "Total Qty\n（总数）:"; ws["H9"] = total
        ws["I9"] = "Sample Qty\n（检查数）:"; ws["L9"] = 5
        ws["F10"] = "Actual Readings（测量数据）"; ws["L10"] = "Result（结果）"
        ws["B11"] = "公称尺寸\nDimension"; ws["C11"] = "上限值\nMax"; ws["D11"] = "下限值\nMin"
        seqs = [5 * (k - 1) + i for i in range(1, 6)]
        for off, g in enumerate(seqs):
            ws.cell(row=11, column=6 + off, value=g)
        ws["K11"] = "(P合格 /F不合格)"; ws["L11"] = "最大值"; ws["M11"] = "最小值"
        for ridx in range(10):
            r = 12 + ridx
            ws.cell(row=r, column=1, value=ridx + 1)
            ws.cell(row=r, column=2, value=f"spec{r}")
            ws.cell(row=r, column=3, value=10.0); ws.cell(row=r, column=4, value=9.0)
            ws.cell(row=r, column=5, value="测高仪")
            for off, g in enumerate(seqs):
                ws.cell(row=r, column=6 + off, value=float(f"{g}.{r}"))
            ws.cell(row=r, column=11, value="P")
            ws.cell(row=r, column=12, value=0)
            ws.cell(row=r, column=13, value=0)
    wb.remove(wb["Sheet"])
    wb.save(path)
    return path


def build_bom(out_dir, rows):
    """rows = [(bh, qty), ...]"""
    path = os.path.join(out_dir, f"BOM_{int(time.time()*1000)}.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "子阶物料"; ws["B1"] = "顶层物料"; ws["C1"] = "图号"
    ws["D1"] = "组件描述"; ws["E1"] = "组件数量"; ws["F1"] = "图纸+尺寸报告+材质证明"
    for i, (bh, qty) in enumerate(rows):
        r = 2 + i
        ws[f"C{r}"] = bh
        ws[f"E{r}"] = qty
        ws[f"F{r}"] = 1
    wb.save(path)
    return path


def verify_output(out_path, K, used, total):
    wb = openpyxl.load_workbook(out_path, data_only=True)
    assert wb.sheetnames[0] == DATE, f"表名应为 {DATE}，实为 {wb.sheetnames[0]}"
    ws = wb[wb.sheetnames[0]]
    seq = [ws.cell(row=11, column=6 + j).value for j in range(5)]
    assert seq[:K] == list(range(used + 1, used + K + 1)), f"序号行前{K}应为{used+1}..{used+K}，实为{seq}"
    assert all(v is None for v in seq[K:]), f"序号行多余应留空，实为{seq}"
    assert ws.cell(row=9, column=8).value == K, "总数应为K"
    assert ws.cell(row=9, column=12).value == K, "检查数应为K"
    for j, g in enumerate([used + 1, used + 2, used + 3, used + 4, used + 5][:K]):
        col = 6 + j
        v12 = ws.cell(row=12, column=col).value
        assert abs(float(v12) - float(f"{g}.12")) < 1e-6, f"列{col}行12应为{g}.12，实为{v12}"
    print(f"  [校验通过] {os.path.basename(out_path)}：表名={DATE}, 序号={seq[:K]}, 总数/检查={K}")


def verify_pieces(out_path, gis):
    """校验输出报告各列数据 = 指定全局件号序列（可跨来源）。K=len(gis)。"""
    K = len(gis)
    wb = openpyxl.load_workbook(out_path, data_only=True)
    n_out = (K + 4) // 5
    assert len(wb.sheetnames) == n_out, f"应有 {n_out} 张表，实为 {wb.sheetnames}"
    seq_all, vals_all = [], []
    for ws in wb.worksheets:
        assert ws.cell(row=9, column=8).value == K, "总数应为K"
        for j in range(5):
            c = 6 + j
            s = ws.cell(row=11, column=c).value
            if s is not None:
                seq_all.append(s)
                vals_all.append(round(float(ws.cell(row=12, column=c).value), 2))
    wb.close()
    assert seq_all == list(gis), f"序号应照抄源报告 {list(gis)}，实为 {seq_all}"
    exp = [round(float(f"{g}.12"), 2) for g in gis]
    assert vals_all == exp, f"件号 {gis} 应为 {exp}，实为 {vals_all}"
    print(f"  [校验通过] {os.path.basename(out_path)}：件号 {gis}，{n_out} 张表")


def cur_suffix(path):
    """返回文件当前后缀数字（无后缀->None, -0->0）。"""
    import re
    base = os.path.basename(path)
    base = re.sub(r"\(\d+\)", "", base)   # 去微信 (1)
    m = re.search(r"-(\d+)\.xlsx$", base)
    if not m:
        return None
    return int(m.group(1))


def safe_rmtree(p):
    try:
        if os.path.exists(p):
            shutil.rmtree(p)
    except OSError as e:
        print("  [清理跳过]", e)


def main():
    safe_rmtree(TMP)
    os.makedirs(TMP)

    # ---- 用例1：多来源拼接扣减（单图号单行 K=2；fB 剩1 用满 + fA 补1）----
    print("=== 用例1：多来源拼接 (fA -5剩5, fB -9剩1, BOM 2件) -> fB第10+fA第6 ===")
    fA = build_report(TMP, used=5, total=10)   # rem 5
    fB = build_report(TMP, used=9, total=10)   # rem 1
    bom = build_bom(TMP, [(BH, 2)])
    s = run_batch(bom, [fA, fB], out_dir=TMP)
    print("  ok/fail:", s["ok_count"], s["fail_count"])
    assert s["ok_count"] == 1 and s["fail_count"] == 0, s["tasks"]
    # K=2：fB 剩1 先用满（第10件）+ fA 补1（第6件）
    verify_pieces(s["tasks"][0]["output"], [10, 6])
    assert cur_suffix(s["tasks"][0]["segments"][0]["source_after"]) == 0, "fB 应最终 -0"
    assert cur_suffix(s["tasks"][0]["segments"][1]["source_after"]) == 6, "fA 应 -6"
    print("  [校验通过] 多来源拼接：fB第10件用满 + fA第6件；fB 归 0")

    # ---- 用例2：智能调度选剩余最少者 ----
    print("=== 用例2：智能调度 (fA -2剩10, fB -10剩2, BOM 2件) -> 选 fB ===")
    fA = build_report(TMP, used=2, total=12)   # rem 10
    fB = build_report(TMP, used=10, total=12)  # rem 2
    bom = build_bom(TMP, [(BH, 2)])
    s = run_batch(bom, [fA, fB], out_dir=TMP)
    assert s["ok_count"] == 1
    t = s["tasks"][0]
    assert os.path.basename(t["source_before"]) == os.path.basename(fB), \
        f"应选 fB，实为 {os.path.basename(t['source_before'])}"
    assert cur_suffix(t["source_after"]) == 0, "fB 用满 -0"
    assert cur_suffix(fA) == 2, "fA 未被触碰，仍 -2"
    verify_output(t["output"], 2, 10, 12)
    print("  [校验通过] 选 fB（剩余最少），fA 仍 -2")

    # ---- 用例3：余量不足 -> NO_REMAINING ----
    print("=== 用例3：余量不足 (fA -2剩1, BOM 2件) -> NO_REMAINING ===")
    fA = build_report(TMP, used=9, total=10)   # rem 1
    bom = build_bom(TMP, [(BH, 2)])
    s = run_batch(bom, [fA], out_dir=TMP)
    assert s["fail_count"] == 1 and s["tasks"][0]["error_code"] == "NO_REMAINING"
    # 来源不改名（仍 -9，不生成新文件）
    assert cur_suffix(fA) == 9, f"来源应不变 -9，实为 {cur_suffix(fA)}"
    print("  [校验通过] NO_REMAINING，来源未改")

    # ---- 用例4：无后缀 = 全新（已用 0）----
    print("=== 用例4：无后缀=全新 (BOM 2件) -> 改名 -2 ===")
    fNew = build_report(TMP, used=None, total=10)  # 无后缀
    bom = build_bom(TMP, [(BH, 2)])
    s = run_batch(bom, [fNew], out_dir=TMP)
    assert s["ok_count"] == 1
    verify_output(s["tasks"][0]["output"], 2, 0, 10)
    assert cur_suffix(s["tasks"][0]["source_after"]) == 2, \
        f"应推进到 -2，实为 {cur_suffix(s['tasks'][0]['source_after'])}"
    print("  [校验通过] 无后缀按全新(0)处理，改名 -2")

    # ---- 用例5：-0 文件被跳过 ----
    print("=== 用例5：-0 文件跳过 (BOM 2件) -> NO_REMAINING ===")
    f0 = build_report(TMP, used=0, total=10)   # -0 -> 已用完
    bom = build_bom(TMP, [(BH, 2)])
    s = run_batch(bom, [f0], out_dir=TMP)
    assert s["fail_count"] == 1 and s["tasks"][0]["error_code"] == "NO_REMAINING"
    assert cur_suffix(f0) == 0, "来源应保持 -0"
    print("  [校验通过] -0 文件跳过，来源未改")

    # ---- 用例6：任务失败回滚（SAMPLE_NOT_FOUND）----
    print("=== 用例6：SAMPLE_NOT_FOUND 回滚 (fA -9, 总21仅录到10件, BOM 2件需11) ===")
    fA = build_report(TMP, used=9, total=21, nsheets=2)  # 总21，仅记录 1..10
    bom = build_bom(TMP, [(BH, 2)])
    s = run_batch(bom, [fA], out_dir=TMP)
    assert s["fail_count"] == 1 and s["tasks"][0]["error_code"] == "SAMPLE_NOT_FOUND", \
        f"应为 SAMPLE_NOT_FOUND，实为 {s['tasks'][0].get('error_code')}"
    assert cur_suffix(fA) == 9, f"失败应回滚，来源仍 -9，实为 {cur_suffix(fA)}"
    # 失败任务不应产生 output
    assert s["tasks"][0].get("output") is None, "失败不应生成新报告"
    print("  [校验通过] 失败回滚：来源不改名、无新文件")

    # ---- 用例7：K>5 -> 自动拆多张表 ----
    print("=== 用例7：K>5 (BOM 12件) -> ceil(12/5)=3 张表 ===")
    fA = build_report(TMP, used=None, total=21, nsheets=5)   # 全新（无后缀）
    bom = build_bom(TMP, [(BH, 12)])
    s = run_batch(bom, [fA], out_dir=TMP)
    assert s["ok_count"] == 1, f"应 1 成功，实为 {s}"
    out = s["tasks"][0]["output"]
    wb = openpyxl.load_workbook(out, data_only=True)
    assert wb.sheetnames == [f"{DATE}-1", f"{DATE}-2", f"{DATE}-3"], wb.sheetnames
    seq_all = []
    for ws in wb.worksheets:
        assert ws.cell(row=9, column=8).value == 12, "每张表总数都应为 K=12"
        assert ws.cell(row=9, column=12).value == 12, "每张表检查数都应为 K=12"
        seq_all += [ws.cell(row=11, column=6 + j).value for j in range(5)
                    if ws.cell(row=11, column=6 + j).value is not None]
    assert seq_all == list(range(1, 13)), f"序号应 1..12，实为 {seq_all}"
    last = wb.worksheets[-1]
    tail = [last.cell(row=11, column=6 + j).value for j in range(5)]
    assert tail == [11, 12, None, None, None], f"尾表后 3 列应留空，实为 {tail}"
    wb.close()
    print("  [校验通过] K=12 -> 3 张表，序号 1..12，尾表后 3 列留空，总数=12")

    # ---- 用例8：报告池驱动方向（BOM 多图号，仅匹配报告的图号处理）----
    print("=== 用例8：报告池驱动 (BOM 含 30 个无关图号 + 930TEST-DM, 仅 1 个报告) ===")
    other_rows = [(f"OTHER-{i:03d}-XX", 2) for i in range(30)]
    bom = build_bom(TMP, other_rows + [(BH, 1)])
    fA = build_report(TMP, used=3, total=10)   # 仅此报告的图号 930TEST-DM 在 BOM
    s = run_batch(bom, [fA], out_dir=TMP)
    # 只有 930TEST-DM 这一组产生任务：1 成功、0 失败；30 个无关图号静默跳过
    assert s["ok_count"] == 1, f"应仅 1 成功，实为 {s['ok_count']}"
    assert s["fail_count"] == 0, f"无关图号不应失败，实为 {s['fail_count']}"
    assert s["tasks"][0]["bh"] == f"{BH}_{DESC}"  # bh 现为 front 整体
    verify_output(s["tasks"][0]["output"], 1, 3, 10)
    print("  [校验通过] 仅匹配图号处理，其余 BOM 图号静默跳过(0失败)")

    # ---- 用例9：报告图号未在 BOM -> NO_BOM_MATCH（逐文件失败，不影响其他）----
    print("=== 用例9：报告图号未匹配 (930TEST-DM 报告 + BOM 无此图号) -> NO_BOM_MATCH ===")
    # BOM 只有无关图号，没有 930TEST-DM
    bom = build_bom(TMP, other_rows)
    fA = build_report(TMP, used=3, total=10)
    s = run_batch(bom, [fA], out_dir=TMP)
    assert s["ok_count"] == 0
    assert s["fail_count"] == 1
    fe = s["tasks"][0]
    assert fe.get("kind") == "file" and fe.get("error_code") == "NO_BOM_MATCH", \
        f"应为 NO_BOM_MATCH(file)，实为 {fe}"
    # 来源文件不应被改名（未进入任务）
    assert cur_suffix(fA) == 3, f"未匹配报告不应改名，实仍为 -3，现为 {cur_suffix(fA)}"
    print("  [校验通过] 报告图号未匹配 -> NO_BOM_MATCH，来源未改动")

    print("\n✅ 全部用例通过")
    print("  (测试沙箱保留于:", TMP, ")")


if __name__ == "__main__":
    main()
