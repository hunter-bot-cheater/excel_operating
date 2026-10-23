# -*- coding: utf-8 -*-
"""图号/文件名组合匹配测试（真实场景组合枚举）。

覆盖组合（front = 报告文件名 `_N件_日期` 之前的整体）：
  1. 图号含多个 -        ：JH10-93021NDB-DM_腔体上盖      （PM-DM 拼接）
  2. 图号本身含 _        ：93021WKD_prt_Cooling            （PM-DM 拼接）
  3. 图号含全角括号，文件名用半角：93010698-(1)_Slit Valve Port Blank（归一化后拼接）
  4. WYS 型 BOM（无图号列，图号嵌在组件描述）：93126082_远端轴套（分段集合匹配）
  5. 数字前缀相同的不同图号不混：93126082 vs JH-YCL-93126082（按图加工）
  6. 同图号多行不同描述：92011601_腔体拼焊板1-试样 / 2-试样 各自命中各自行
  7. 描述为空回退「仅图号前缀」：93011182_气路法兰（4 位年日期）
  8. PM-DM 型纯数字图号：93126082_远端轴套（拼接命中）
  9. 完全不在 BOM       -> NO_BOM_MATCH
 10. WYS 分段集合歧义    -> BOM_MULTI_MATCH（两个不等价行同时命中）
"""
import os
import sys
import time

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
TMP = os.path.join(HERE, f"combo_{int(time.time() * 1000)}")
sys.path.insert(0, os.path.dirname(HERE))
from extract_core.api import run_batch

DATE = "26.4.09"


def build_report(out_dir, front, total, used, dwg, nsheets=2):
    """合成报告：B3=DWG（图号），测量记录 1..(nsheets*5) 件。used=None 表示无后缀。"""
    suffix = "" if used is None else f"-{used}"
    fname = f"{front}_{total}件_{DATE}{suffix}(1).xlsx"
    path = os.path.join(out_dir, fname)
    wb = openpyxl.Workbook()
    for k in range(1, nsheets + 1):
        ws = wb.create_sheet(f"{DATE}（{total}件）-{k}")
        ws["A1"] = "Inspection Report\n检验报告"
        ws["A3"] = "DWG No:\n图号"
        ws["B3"] = dwg
        ws["F9"] = "Total Qty\n（总数）:"
        ws["H9"] = total
        ws["I9"] = "Sample Qty\n（检查数）:"
        ws["F10"] = "Actual Readings（测量数据）"
        ws["B11"] = "公称尺寸\nDimension"
        seqs = [5 * (k - 1) + i for i in range(1, 6)]
        for off, g in enumerate(seqs):
            ws.cell(row=11, column=6 + off, value=g)
        ws["K11"] = "Result（结果）"
        ws["L11"] = "最大值"
        ws["M11"] = "最小值"
        for ridx in range(8):
            r = 12 + ridx
            ws.cell(row=r, column=1, value=ridx + 1)
            for off, g in enumerate(seqs):
                ws.cell(row=r, column=6 + off, value=float(f"{g}.{ridx}"))
    wb.remove(wb["Sheet"])
    wb.save(path)
    return path


def build_bom_pm(out_dir, rows):
    """PM-DM 型：C=图号 D=组件描述 E=组件数量。rows=(bh, desc, qty)"""
    path = os.path.join(out_dir, f"BOMPM_{int(time.time()*1000)}.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "子阶物料"; ws["B1"] = "顶层物料"; ws["C1"] = "图号"
    ws["D1"] = "组件描述"; ws["E1"] = "组件数量"; ws["F1"] = "图纸+尺寸报告+材质证明"
    for i, (bh, desc, qty) in enumerate(rows):
        r = 2 + i
        ws[f"C{r}"] = bh
        ws[f"D{r}"] = desc
        ws[f"E{r}"] = qty
    wb.save(path)
    return path


def build_bom_wys(out_dir, rows):
    """WYS 型：A=组件描述（图号嵌在其中） B=组件数量。rows=(desc, qty)"""
    path = os.path.join(out_dir, f"BOMWYS_{int(time.time()*1000)}.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "组件描述"; ws["B1"] = "组件数量"
    for i, (desc, qty) in enumerate(rows):
        r = 2 + i
        ws[f"A{r}"] = desc
        ws[f"B{r}"] = qty
    wb.save(path)
    return path


def cur_suffix(path):
    import re
    base = os.path.basename(path)
    base = re.sub(r"\(\d+\)$", "", base)
    m = re.search(r"-(\d+)\.xlsx$", base)
    return int(m.group(1)) if m else None


def main():
    os.makedirs(TMP, exist_ok=True)
    passed = []

    def check(name, cond, detail=""):
        assert cond, f"[FAIL] {name} {detail}"
        passed.append(name)
        print(f"  ✅ {name}")

    # ---- 1. 图号含多个 - ----
    print("=== 1. 图号含多个 - (JH10-93021NDB-DM) ===")
    bom = build_bom_pm(TMP, [("JH10-93021NDB-DM", "腔体上盖", 2)])
    rep = build_report(TMP, "JH10-93021NDB-DM_腔体上盖", 10, 3, "JH10-93021NDB-DM")
    s = run_batch(bom, [rep], out_dir=TMP)
    assert s["ok_count"] == 1 and s["fail_count"] == 0, s
    t = s["tasks"][0]
    check("K=BOM数量2, used=3", t["qty"] == 2 and t["used_from"] == 3)
    check("输出名=front_2件", os.path.basename(t["output"]) == "JH10-93021NDB-DM_腔体上盖_2件_26.4.09.xlsx")
    check("基准推进 -5", cur_suffix(t["source_after"]) == 5)

    # ---- 2. 图号本身含 _ ----
    print("=== 2. 图号含 _ (93021WKD_prt) ===")
    bom = build_bom_pm(TMP, [("93021WKD_prt", "Cooling", 1)])
    rep = build_report(TMP, "93021WKD_prt_Cooling", 10, None, "93021WKD_prt")
    s = run_batch(bom, [rep], out_dir=TMP)
    assert s["ok_count"] == 1, s
    t = s["tasks"][0]
    check("无后缀=全新 used=0, K=1", t["used_from"] == 0 and t["qty"] == 1)
    check("基准改名为 -1", cur_suffix(t["source_after"]) == 1)

    # ---- 3. 全角括号图号 vs 半角文件名（归一化）----
    print("=== 3. 全角→半角归一 (93010698-（1） vs 93010698-(1)) ===")
    bom = build_bom_pm(TMP, [("93010698-（1）", "Slit Valve Port Blank", 4)])
    rep = build_report(TMP, "93010698-(1)_Slit Valve Port Blank", 21, 4, "93010698-(1)")
    s = run_batch(bom, [rep], out_dir=TMP)
    assert s["ok_count"] == 1, s
    t = s["tasks"][0]
    check("拼接命中, K=4, used=4", t["qty"] == 4 and t["used_from"] == 4)

    # ---- 4. WYS 型（无图号列）----
    print("=== 4. WYS 分段集合匹配 (93126082_远端轴套) ===")
    bom = build_bom_wys(TMP, [("远端轴套_93126082_Pla,PTFE-V0", 2)])
    rep = build_report(TMP, "93126082_远端轴套", 10, None, "93126082")
    s = run_batch(bom, [rep], out_dir=TMP)
    assert s["ok_count"] == 1, s
    t = s["tasks"][0]
    check("组件数量=2 作为 K", t["qty"] == 2)

    # ---- 5. 数字前缀相同的不同图号不混 ----
    print("=== 5. 同数字前缀不同图号 (93126082 vs JH-YCL-93126082（按图加工）) ===")
    bom = build_bom_wys(TMP, [
        ("远端轴套_93126082_Pla,PTFE-V0", 2),
        ("远端轴套_JH-YCL-93126082（按图加工）_Pla,PTFE-V0", 1),
    ])
    rep = build_report(TMP, "JH-YCL-93126082（按图加工）_远端轴套", 10, 3,
                       "JH-YCL-93126082（按图加工）")
    s = run_batch(bom, [rep], out_dir=TMP)
    assert s["ok_count"] == 1, s
    t = s["tasks"][0]
    check("只命中 JH-YCL 行（K=1 而非 2）", t["qty"] == 1)

    # ---- 6. 同图号多行不同描述，各自命中 ----
    print("=== 6. 同图号多行 (92011601 腔体拼焊板1/2-试样) ===")
    bom = build_bom_pm(TMP, [
        ("92011601", "腔体拼焊板1-试样", 1),
        ("92011601", "腔体拼焊板2-试样", 2),
    ])
    rep = build_report(TMP, "92011601_腔体拼焊板2-试样", 10, None, "92011601")
    s = run_batch(bom, [rep], out_dir=TMP)
    assert s["ok_count"] == 1, s
    t = s["tasks"][0]
    check("只命中 2-试样 行（K=2 而非 1）", t["qty"] == 2)

    # ---- 7. 描述为空 -> 仅图号前缀回退（4 位年日期）----
    print("=== 7. 仅图号回退 (93011182_气路法兰, desc 空) ===")
    bom = build_bom_pm(TMP, [("93011182", "", 1)])
    # 4 位年日期需能解析
    front = "93011182_气路法兰"
    fname = f"{front}_1件_2026.6.16(1).xlsx"
    rep = os.path.join(TMP, fname)
    build_report(TMP, front, 10, None, "93011182")
    os.replace(build_report(TMP, front, 10, None, "93011182"), rep)
    s = run_batch(bom, [rep], out_dir=TMP)
    assert s["ok_count"] == 1, s
    check("前缀回退命中, K=1", s["tasks"][0]["qty"] == 1)

    # ---- 8. PM-DM 型纯数字图号（拼接）----
    print("=== 8. PM-DM 纯数字图号 (93126082_远端轴套) ===")
    bom = build_bom_pm(TMP, [("93126082", "远端轴套", 3)])
    rep = build_report(TMP, "93126082_远端轴套", 10, 2, "93126082")
    s = run_batch(bom, [rep], out_dir=TMP)
    assert s["ok_count"] == 1, s
    t = s["tasks"][0]
    check("拼接命中, K=3, used=2→5", t["qty"] == 3 and cur_suffix(t["source_after"]) == 5)

    # ---- 9. 完全不在 BOM ----
    print("=== 9. 完全不在 BOM -> NO_BOM_MATCH ===")
    bom = build_bom_pm(TMP, [("JH10-93021NDB-DM", "腔体上盖", 2)])
    rep = build_report(TMP, "99999999_不存在件", 10, None, "99999999")
    s = run_batch(bom, [rep], out_dir=TMP)
    assert s["fail_count"] == 1 and s["tasks"][0]["error_code"] == "NO_BOM_MATCH"
    check("NO_BOM_MATCH 且来源未动", cur_suffix(rep) is None)

    # ---- 10. WYS 分段集合歧义 ----
    print("=== 10. WYS 分段集合歧义 -> BOM_MULTI_MATCH ===")
    bom = build_bom_wys(TMP, [
        ("远端轴套_93126082_Pla", 2),
        ("93126082_远端轴套_Pla", 5),
    ])
    rep = build_report(TMP, "93126082_远端轴套", 10, None, "93126082")
    s = run_batch(bom, [rep], out_dir=TMP)
    assert s["fail_count"] == 1 and s["tasks"][0]["error_code"] == "BOM_MULTI_MATCH", s
    check("歧义明确报错（不瞎选）", "歧义" in s["tasks"][0]["reason"])

    # ---- 11. BOM 无「组件数量」表头 -> 明确报错（不再把标记列/其它列当数量）----
    print("=== 11. BOM 无「组件数量」表头（只有图纸包标记列）-> BOM_READ_ERROR ===")
    path = os.path.join(TMP, "BOM_noqty.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "子阶物料"; ws["B1"] = "顶层物料"; ws["C1"] = "图号"
    ws["D1"] = "组件描述"; ws["E1"] = "图纸+尺寸报告+材质证明"  # 0/1 标记列，值全 1
    ws["C2"] = "93019282-DM"; ws["D2"] = "翻转横撑"; ws["E2"] = 1
    wb.save(path)
    rep = build_report(TMP, "93019282-DM_翻转横撑", 10, 1, "93019282-DM")
    s = run_batch(path, [rep], out_dir=TMP)
    assert s["fail_count"] == 1 and not s["tasks"][0].get("ok")
    err = s["tasks"][0]
    assert err["error_code"] == "BOM_READ_ERROR", err
    check("无「组件数量」表头 -> BOM_READ_ERROR", "组件数量" in err["reason"])

    # ---- 12. 表头含「数量」但不是「组件数量」-> 也要报错（不模糊匹配）----
    print("=== 12. 表头「组件计量单位数量」-> 仍报错（证明只认精确四字） ===")
    path = os.path.join(TMP, "BOM_otherqty.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "图号"; ws["B1"] = "组件描述"; ws["C1"] = "组件计量单位数量"
    ws["A2"] = "93019282-DM"; ws["B2"] = "翻转横撑"; ws["C2"] = 4
    wb.save(path)
    s = run_batch(path, [rep], out_dir=TMP)
    assert s["fail_count"] == 1 and s["tasks"][0]["error_code"] == "BOM_READ_ERROR", s
    check("「组件计量单位数量」不被当作数量列", True)

    print(f"\n✅ 组合测试全部通过：{len(passed)}/12")
    print("  覆盖：多-图号 / 图号含_ / 全半角归一 / WYS分段集合 / 同前缀不同图号 /")
    print("        同图号多行 / 仅图号回退 / 纯数字图号 / 未匹配 / 歧义报错 /")
    print("        无「组件数量」表头报错 / 「数量」近似列名不被误认")
    print("  (沙箱保留于:", TMP, ")")


if __name__ == "__main__":
    main()
