# -*- coding: utf-8 -*-
"""BOM 回写专项测试。

覆盖：
  A. BOM 已有「第N套」列 -> 直接写入，不插列
  B. BOM 无「第N套」列 -> 自动插入（第X套列最右、图纸等无关列最左），列名「第N套」
  C. 同图号多行 -> 每行各自写；日期取基准文件名日期、件数 = K
  D. 不再生成任何 .bak 备份（用户要求：BOM 不需要备份）
  E. round_no 留空 -> 不回写、无备份
"""
import os
import re
import sys
import time
import shutil

HERE = os.path.dirname(os.path.abspath(__file__))
TMP = os.path.join(HERE, f"sandbox_bomwrite_{int(time.time()*1000)}")
sys.path.insert(0, os.path.dirname(HERE))

import openpyxl
from extract_core.api import run_batch

BH = "93019131"
DESC = "氮气弹簧安装块(左)"
DATE = "2026.6.15"

BASE_HEADERS = ["排序层级", "所在层级", "描述/title", "图号", "组件数量", "分类", "E物料码",
                "第11套", "第12套"]


def build_report(out_dir, used, total, name=None):
    """造与真实测量报告同构的报告（单表，DWG=BH）。

    注意：无后缀 = 全新未用；`-0` = 已用完。故 used=0 时不写后缀。
    """
    suffix = f"-{used}" if used else ""
    fname = name or f"{BH}_{DESC}_{total}件_{DATE}{suffix}.xlsx"
    path = os.path.join(out_dir, fname)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"2026.6.15（{total}件）"
    ws["A1"] = "Inspection Report\n检验报告"
    ws["A3"] = "DWG No:\n图号"; ws["B3"] = BH
    ws["D3"] = "Part Description:\n零件名称"; ws["F3"] = DESC
    ws["A9"] = "Item No\n（编号）"; ws["B9"] = "Specified Parameter"
    ws["E9"] = "Instrument used（检具）"
    ws["F9"] = "Total Qty\n（总数）:"; ws["H9"] = total
    ws["I9"] = "Sample Qty\n（检查数）:"; ws["L9"] = 5
    ws["F10"] = "Actual Readings（测量数据）"; ws["L10"] = "Result（结果）"
    ws["B11"] = "公称尺寸\nDimension"; ws["C11"] = "上限值\nMax"; ws["D11"] = "下限值\nMin"
    seqs = [1, 2, 3, 4, 5]
    for off, g in enumerate(seqs):
        ws.cell(row=11, column=6 + off, value=g)
    ws["K11"] = "(P合格 /F不合格)"; ws["L11"] = "最大值"; ws["M11"] = "最小值"
    for ridx in range(10):
        r = 12 + ridx
        ws.cell(row=r, column=1, value=ridx + 1)
        ws.cell(row=r, column=2, value=f"spec{ridx}")
        ws.cell(row=r, column=3, value=10.0); ws.cell(row=r, column=4, value=9.0)
        ws.cell(row=r, column=5, value="测高仪")
        for off, g in enumerate(seqs):
            ws.cell(row=r, column=6 + off, value=float(f"{g}.{r}"))
        ws.cell(row=r, column=11, value="P")
        ws.cell(row=r, column=12, value=0); ws.cell(row=r, column=13, value=0)
    wb.save(path)
    return path


def build_bom(out_dir, headers, rows, name):
    """headers=表头列表；rows=每行 dict（列名->值）。返回 (路径, {列名:列号})。"""
    path = os.path.join(out_dir, name)
    wb = openpyxl.Workbook()
    ws = wb.active
    for c, h in enumerate(headers, start=1):
        ws.cell(row=1, column=c, value=h)
    for i, row in enumerate(rows):
        r = 2 + i
        for h, v in row.items():
            ws.cell(row=r, column=headers.index(h) + 1, value=v)
    wb.save(path)
    return path


def read_sheet(path):
    """返回 (headers: {表头文本: 列号}, ws)。"""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    hdrs = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=1, column=c).value
        if v is not None and str(v).strip():
            hdrs[str(v).strip()] = c
    return hdrs, ws


def main():
    if os.path.exists(TMP):
        shutil.rmtree(TMP)
    os.makedirs(TMP)
    fails = []

    def case_dir(name):
        d = os.path.join(TMP, name)
        os.makedirs(d, exist_ok=True)
        return d

    # ---------- A. 已有「第13套」列：直接写入 ----------
    print("=== A. BOM 已有「第13套」列 -> 直接写入 ===")
    dA = case_dir("A")
    hdrsA = BASE_HEADERS[:9] + ["第13套", "图纸", "材质报告", "尺寸报告"]
    bomA = build_bom(dA, hdrsA, [{"图号": BH, "组件数量": 1, "第13套": None}], "BOM_A.xlsx")
    repA = build_report(dA, used=0, total=10)
    sA = run_batch(bomA, [repA], out_dir=dA, round_no=13)
    print("  ok/fail:", sA["ok_count"], sA["fail_count"], "| 回写错误:", sA["bom_write_errors"])
    try:
        assert sA["ok_count"] == 1 and sA["fail_count"] == 0, f"任务应成功: {sA['tasks']}"
        assert not sA["bom_write_errors"], f"不应有回写错误: {sA['bom_write_errors']}"
        hdrs, ws = read_sheet(bomA)
        assert "第13套" in hdrs, "第13套列应仍存在"
        col = hdrs["第13套"]
        val = ws.cell(row=2, column=col).value
        assert val == f"{DATE}/1", f"应为 {DATE}/1，实为 {val!r}"
        assert "第11套" in hdrs and "第12套" in hdrs and "图纸" in hdrs
        print(f"  [OK] 第13套(第{col}列) 第2行 = {val!r}")
    except Exception as e:
        fails.append(f"A: {e}")
        print("  [FAIL]", e)

    # ---------- B. 无「第13套」列：自动插入 ----------
    print("=== B. BOM 无「第13套」列 -> 自动插入（第12套右、图纸左）===")
    dB = case_dir("B")
    hdrsB = BASE_HEADERS + ["图纸", "材质报告", "尺寸报告"]
    bomB = build_bom(dB, hdrsB, [{"图号": BH, "组件数量": 1}], "BOM_B.xlsx")
    repB = build_report(dB, used=0, total=10)
    sB = run_batch(bomB, [repB], out_dir=dB, round_no=13)
    print("  ok/fail:", sB["ok_count"], sB["fail_count"], "| 回写错误:", sB["bom_write_errors"])
    try:
        assert sB["ok_count"] == 1 and not sB["bom_write_errors"], f"{sB['tasks']}"
        hdrs, ws = read_sheet(bomB)
        assert "第13套" in hdrs, "应插入第13套列"
        c13, c12, c_dwg = hdrs["第13套"], hdrs["第12套"], hdrs["图纸"]
        assert c13 == c12 + 1, f"第13套应紧邻第12套右侧（{c12}->{c13}）"
        assert c_dwg == c13 + 1, f"图纸应在第13套右侧（{c13}->{c_dwg}）"
        val = ws.cell(row=2, column=c13).value
        assert val == f"{DATE}/1", f"应为 {DATE}/1，实为 {val!r}"
        print(f"  [OK] 第12套={c12} 列, 新插第13套={c13} 列, 图纸={c_dwg} 列; 值={val!r}")
    except Exception as e:
        fails.append(f"B: {e}")
        print("  [FAIL]", e)

    # ---------- C. 同图号多行（表一图号唯一口径）-> 直接报错，不回写 ----------
    print("=== C. 同图号多行（2 行）-> 匹配直接报错，BOM 不回写 ===")
    dC = case_dir("C")
    hdrsC = BASE_HEADERS + ["图纸", "材质报告", "尺寸报告"]
    bomC = build_bom(dC, hdrsC,
                     [{"图号": BH, "组件数量": 1},
                      {"图号": BH, "组件数量": 1}],
                     "BOM_C.xlsx")
    repC = build_report(dC, used=0, total=10)
    sC = run_batch(bomC, [repC], out_dir=dC, round_no=12)
    print("  ok/fail:", sC["ok_count"], sC["fail_count"], "| 回写错误:", sC["bom_write_errors"])
    try:
        assert sC["ok_count"] == 0, f"同图号多行不应有成功任务: {sC['tasks']}"
        assert any(t.get("error_code") == "BOM_MULTI_MATCH" for t in sC["tasks"]), \
            f"应报 BOM_MULTI_MATCH: {sC['tasks']}"
        hdrs, ws = read_sheet(bomC)
        assert "第13套" not in hdrs, "报错后不应插入新套列"
        assert set(hdrs) == set(hdrsC), f"BOM 表头不应变化: {set(hdrsC) ^ set(hdrs)}"
        c12 = hdrs["第12套"]
        assert ws.cell(row=2, column=c12).value is None and \
            ws.cell(row=3, column=c12).value is None, "两行都不应回写"
        print("  [OK] 同图号多行 -> BOM_MULTI_MATCH 报错；BOM 未插列未回写")
    except Exception as e:
        fails.append(f"C: {e}")
        print("  [FAIL]", e)

    # ---------- D. 不生成备份 ----------
    print("=== D. 已取消 .bak 备份（母版就地修改）===")
    try:
        assert not sB.get("bom_backup"), f"不应生成备份，实为 {sB.get('bom_backup')}"
        leftovers = [f for f in os.listdir(dB)
                     if ".bak" in f.lower() and f.lower().endswith((".xls", ".xlsx"))]
        assert not leftovers, f"BOM 同目录不应残留 .bak 文件：{leftovers}"
        print("  [OK] 无 .bak 残留，BOM 母版已就地更新")
    except Exception as e:
        fails.append(f"D: {e}")
        print("  [FAIL]", e)

    # ---------- E. round_no 留空 -> 不回写 ----------
    print("=== E. round_no 缺省 -> 不回写、无备份 ===")
    dE = case_dir("E")
    hdrsE = BASE_HEADERS + ["图纸", "材质报告", "尺寸报告"]
    bomE = build_bom(dE, hdrsE, [{"图号": BH, "组件数量": 1}], "BOM_E.xlsx")
    repE = build_report(dE, used=0, total=10)
    sE = run_batch(bomE, [repE], out_dir=dE)
    try:
        assert sE["ok_count"] == 1, f"任务应成功: {sE['tasks']}"
        assert not sE.get("bom_backup"), f"不应有备份，实为 {sE.get('bom_backup')}"
        hdrs, _ws = read_sheet(bomE)
        assert "第13套" not in hdrs and set(hdrs) == set(hdrsE), "BOM 不应被改动"
        print("  [OK] 未回写、无备份、BOM 表头未变")
    except Exception as e:
        fails.append(f"E: {e}")
        print("  [FAIL]", e)

    # ---------- F. 多表单：只认第一张有效数据表 ----------
    print("=== F. BOM 多表单 -> 只匹配/只回写第一张有效数据表 ===")
    dF = case_dir("F")
    bomF = os.path.join(dF, "BOM_F.xlsx")
    wbf = openpyxl.Workbook()
    ws1 = wbf.active; ws1.title = "Sheet1"
    ws1.append(["图号", "组件描述", "组件数量", "第13套"])
    ws1.append([BH, DESC, 1, None])                    # Sheet1 第2行（真相源）
    ws2 = wbf.create_sheet("Sheet2")
    ws2.append(["图号", "组件描述", "组件数量"])
    ws2.append([BH, DESC, 99])                         # 不同数量 -> 旧实现会多建任务，新实现忽略
    ws2.append(["93099999", "别的件", 2])
    ws3 = wbf.create_sheet("Sheet3")
    ws3.append(["图号", "组件描述", "组件数量"])
    ws3.append([BH, DESC, 1])                          # 相同副本 -> 新实现忽略
    wbf.save(bomF)
    repF = build_report(dF, used=0, total=10)
    sF = run_batch(bomF, [repF], out_dir=dF, round_no=13)
    print("  ok/fail:", sF["ok_count"], sF["fail_count"], "| 回写错误:", sF["bom_write_errors"])
    try:
        assert sF["ok_count"] == 1 and not sF["bom_write_errors"], \
            f"应只 1 个任务成功（只认 Sheet1），实为 {sF['tasks']}"
        wbr = openpyxl.load_workbook(bomF, data_only=True)
        s1 = wbr["Sheet1"]; s2 = wbr["Sheet2"]; s3 = wbr["Sheet3"]
        assert s1.cell(row=2, column=4).value == f"{DATE}/1", \
            f"Sheet1 第2行应写入，实为 {s1.cell(row=2, column=4).value!r}"
        assert s2.cell(row=1, column=4).value is None, "Sheet2 不应被插入第13套列"
        assert s3.cell(row=1, column=4).value is None, "Sheet3 不应被插入第13套列"
        print(f"  [OK] 只回写 Sheet1（第2行={s1.cell(row=2, column=4).value!r}）；Sheet2/Sheet3 未插列未写入")
    except Exception as e:
        fails.append(f"F: {e}")
        print("  [FAIL]", e)

    print("\n==== 结果 ====")
    if fails:
        print("❌ 失败项:")
        for f in fails:
            print("  -", f)
        raise SystemExit(1)
    print("✅ BOM 回写全部验证通过（直接写/自动插列/多行/无备份/留空不回写）")
    print("  沙箱:", TMP)


if __name__ == "__main__":
    main()
