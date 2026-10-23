# -*- coding: utf-8 -*-
"""match_report 多信号交叉验证专项测试（v6 新增）。

覆盖：
  A. 真实场景1：front=93019282-DM_翻转横撑 -> 必须命中「翻转横撑」行（qty 1），
     绝不能命中同图号前缀的「弹簧压板」行（qty 3）——旧版②级前缀的 bug。
  B. 近邻图号防误撞：93019282 不命中 93019281（翻转横撑同名）。
  C. 报告内部双键（dwg+part）精确命中（L0）。
  D. 内部图号与零件名冲突时不采信图号（L1 回退）-> 仍命中翻转横撑行。
  E. 物料长描述列被识别（91010070 型 BOM）。
  F. 同数字多行 -> 描述消歧（93021561_焊接管 -> 唯一 CF 行）。
  G. 同数字多行且描述无法消歧 -> BOM_MULTI_MATCH。
  H. 回归：原 ①②③④ 行为不变。
"""
import os
import sys
import time

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
TMP = os.path.join(HERE, f"matchrep_{int(time.time() * 1000)}")
sys.path.insert(0, os.path.dirname(HERE))
os.makedirs(TMP, exist_ok=True)

from extract_core.bom_reader import read_bom_ex, match_front, match_report
from extract_core.errors import ExtractError
from extract_core.report_reader import extract_identity

fails = []


def build_bom(name, rows, desc_hdr="组件描述"):
    """PM-DM 型 BOM：C=图号 D=描述列 E=组件数量。rows=(bh, desc, qty)"""
    path = os.path.join(TMP, name)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "子阶物料"; ws["B1"] = "顶层物料"; ws["C1"] = "图号"
    ws["D1"] = desc_hdr; ws["E1"] = "组件数量"; ws["F1"] = "图纸"
    for i, (bh, desc, qty) in enumerate(rows):
        r = 2 + i
        ws[f"C{r}"] = bh
        ws[f"D{r}"] = desc
        ws[f"E{r}"] = qty
    wb.save(path)
    return path


def load(name, rows, desc_hdr="组件描述"):
    bom = build_bom(name, rows, desc_hdr)
    rows, _meta = read_bom_ex(bom)
    return rows


def check(case, got, want_qty=None, want_desc=None, want_code=None, want_empty=False):
    if want_empty:
        if got == []:
            print(f"  [OK] {case}: 未命中（符合预期）")
        else:
            fails.append(f"{case}: 期望未命中，实为 {got}")
            print(f"  [FAIL] {case}: 期望未命中，实为 {got}")
        return
    if want_code:
        if isinstance(got, tuple) and got[0] == want_code:
            print(f"  [OK] {case}: -> {want_code}")
        else:
            fails.append(f"{case}: 期望 {want_code}，实为 {got}")
            print(f"  [FAIL] {case}: 期望 {want_code}，实为 {got}")
        return
    if not isinstance(got, list) or len(got) != 1:
        fails.append(f"{case}: 期望单行命中，实为 {got}")
        print(f"  [FAIL] {case}: 实为 {got}")
        return
    r = got[0]
    ok_qty = want_qty is None or r.qty == want_qty
    ok_desc = want_desc is None or r.desc == want_desc
    if ok_qty and ok_desc:
        print(f"  [OK] {case}: -> 图号={r.bh!r} 描述={r.desc!r} 件数={r.qty} 行{r.row}")
    else:
        fails.append(f"{case}: 期望 件数={want_qty} 描述={want_desc}，实为 {r}")
        print(f"  [FAIL] {case}: 实为 {r}")


def call_match(case, front, rows, dwg="", part=""):
    try:
        return match_report(front, rows, dwg=dwg, part=part)
    except ExtractError as e:
        return (e.code, e.reason)


print("=== A. 真实场景1：93019282-DM_翻转横撑（旧版②级前缀 bug）===")
rows_a = load("BOM_A.xlsx", [
    ("93019281-DM", "翻转横撑", 2),
    ("JH-YCL-93019282-DM（按图加工）", "翻转横撑", 1),
    ("93019282-DM", "弹簧压板", 3),
])
check("A1 纯文件名匹配（L2 双键联合）",
      call_match("A1", "93019282-DM_翻转横撑", rows_a),
      want_qty=1, want_desc="翻转横撑")
check("A2 携带内部身份（L0 双键全等）",
      call_match("A2", "93019282-DM_翻转横撑", rows_a,
                 dwg="JH-YCL-93019282-DM\n（按图加工）", part="翻转横撑"),
      want_qty=1, want_desc="翻转横撑")
check("A3 内部图号+零件名冲突（L1 不采信图号）",
      call_match("A3", "93019282-DM_翻转横撑", rows_a,
                 dwg="93019282-DM", part="翻转横撑"),
      want_qty=1, want_desc="翻转横撑")
# 93019281 同名近邻不得命中
check("A4 近邻图号 93019281 不被误命中",
      call_match("A4", "93019282-DM_翻转横撑", rows_a),
      want_qty=1, want_desc="翻转横撑")

print("=== B. 同数字多行 -> 描述消歧（93021561 焊接管族）===")
rows_b = load("BOM_B.xlsx", [
    ("93021561", "CF100焊接管", 1),
    ("93021561", "ISO200焊接管", 4),
    ("93021561", "水路管接头", 2),
    ("93021561", "CF40焊接管", 1),
])
check("B1 描述唯一收敛",
      call_match("B1", "93021561_CF100焊接管", rows_b),
      want_qty=1, want_desc="CF100焊接管")
check("B2 内部零件名收敛（L1）",
      call_match("B2", "93021561_CF100焊接管", rows_b,
                 dwg="93021561", part="CF100焊接管"),
      want_qty=1, want_desc="CF100焊接管")
check("B3 描述过于宽泛（多个不同行含'焊接管'）-> 歧义报错",
      call_match("B3", "93021561_焊接管", rows_b),
      want_code="BOM_MULTI_MATCH")

print("=== C. 物料长描述列识别（91010070 型 BOM）===")
rows_c = load("BOM_C.xlsx", [
    ("93010414", "保护盖_93010414_AL6061-T6", 1),
    ("93010412-新PD", "保护盖2_93010412-新PD_AL6061-T6", 1),
    ("93010410", "保护盖1_93010410_AL6061-T6", 4),
], desc_hdr="物料长描述")
check("C1 front=93010410_保护盖1",
      call_match("C1", "93010410_保护盖1", rows_c),
      want_qty=4, want_desc="保护盖1_93010410_AL6061-T6")
check("C2 携带内部身份",
      call_match("C2", "93010410_保护盖1", rows_c,
                 dwg="JH-YCL-93010410", part="保护盖1"),
      want_qty=4, want_desc="保护盖1_93010410_AL6061-T6")
check("C3 内部图号全等直接命中（L1）",
      call_match("C3", "93010410_保护盖1", rows_c,
                 dwg="93010410", part="保护盖1"),
      want_qty=4, want_desc="保护盖1_93010410_AL6061-T6")

print("=== D. 歧义保护：同数字多行且无法消歧 ===")
rows_d = load("BOM_D.xlsx", [
    ("90000001-A", "AAA", 2),
    ("90000001-B", "BBB", 3),
])
check("D1 描述都对不上 front -> BOM_MULTI_MATCH",
      call_match("D1", "90000001_别的零件", rows_d),
      want_code="BOM_MULTI_MATCH")

print("=== F. 项目前缀混排（91010070_E.1_20260814_Pd93010410_保护盖1）===")
rows_f = load("BOM_F.xlsx", [
    ("", "Pd主腔体_91010070", 1),
    ("JH-YCL-93010410（按图加工）", "保护盖1", 4),
    ("93010412-新PD", "保护盖2", 1),
])
check("F1 图号藏在第4段（Pd93010410），无内部身份",
      call_match("F1", "91010070_E.1_20260814_Pd93010410_保护盖1", rows_f),
      want_qty=4, want_desc="保护盖1")
check("F2 同一文件名带内部身份",
      call_match("F2", "91010070_E.1_20260814_Pd93010410_保护盖1", rows_f,
                 dwg="JH-YCL-93010410\n（按图加工）", part="保护盖1"),
      want_qty=4, want_desc="保护盖1")

print("=== E. 回归：原 ①②③④ 行为不变 ===")
check("E1 ① 拼接全等", call_match("E1", "93011182_气路法兰",
      load("BOM_E1.xlsx", [("93011182", "气路法兰", 1)])), want_qty=1)
check("E2 ② 图号前缀（描述为空回退）", call_match("E2", "93011182_别的描述",
      load("BOM_E2.xlsx", [("93011182", "气路法兰", 1)])), want_qty=1)
check("E3 ③ WYS 分段集合", call_match("E3", "93126082_远端轴套",
      load("BOM_E3.xlsx", [("", "远端轴套_93126082_Pla,PTFE-V0", 2)])), want_qty=2)
check("E4 ④ 数字键（图号带 JH-YCL 前缀）", call_match("E4", "93010410_保护盖1",
      load("BOM_E4.xlsx", [("JH-YCL-93010410（按图加工）", "保护盖1", 4)])), want_qty=4)
check("E5 短数字键不参与（标准件）", call_match("E5", "1191_圆柱销",
      load("BOM_E5.xlsx", [("GB/T 119.1", "圆柱销", 5)])), want_empty=True)

print("=== G. 同图号多行 -> 直接报错（表一图号唯一口径）===")
rows_g = load("BOM_G.xlsx", [
    ("92011601", "腔体拼焊板1-试样", 1),
    ("92011601", "腔体拼焊板2-试样", 2),
    ("92011601", "腔体拼焊板3-试样", 3),
])
# G1 纯图号命中多行 -> 报错（不再返回多行任务）
check("G1 纯图号 front=92011601 -> 多行直接报错",
      call_match("G1", "92011601", rows_g),
      want_code="BOM_MULTI_MATCH")
# G2 描述段命中多行（泛化）-> 报错
check("G2 描述段命中多行（泛化）-> 报错",
      call_match("G2", "92011601_腔体拼焊板", rows_g),
      want_code="BOM_MULTI_MATCH")
# G3 描述段对不上任何行 -> 报错
check("G3 描述段对不上任何行 -> 报错",
      call_match("G3", "92011601_别的描述", rows_g),
      want_code="BOM_MULTI_MATCH")
# G3b 描述精确收敛到唯一行 -> 单行放行（唯一例外）
check("G3b 描述精确命中其中一行 -> 单行放行",
      call_match("G3b", "92011601_腔体拼焊板2-试样", rows_g),
      want_qty=2, want_desc="腔体拼焊板2-试样")
# G4 纯图号 + 单行 -> 单行
check("G4 纯图号 + 单行 -> 单行",
      call_match("G4", "93019282-DM", rows_a),
      want_qty=3, want_desc="弹簧压板")
# G5 数字键路径：纯图号多行 -> 报错
rows_g5 = load("BOM_G5.xlsx", [
    ("90000001-A", "AAA", 2),
    ("90000001-B", "BBB", 3),
])
check("G5 数字键纯图号 front=90000001 -> 多行报错",
      call_match("G5", "90000001", rows_g5),
      want_code="BOM_MULTI_MATCH")
# G6 内部 dwg 多行 + 无零件名 + 纯图号 -> 报错
check("G6 内部 dwg 多行 + 纯图号 -> 多行报错",
      call_match("G6", "93021561", rows_b, dwg="93021561", part=""),
      want_code="BOM_MULTI_MATCH")

print("\n==== 结果 ====")
if fails:
    print("❌ 失败项:")
    for f in fails:
        print("  -", f)
    raise SystemExit(1)
print("✅ 多信号匹配全部通过（双键全等 / 零件名收敛 / 双键联合 / 描述消歧 /")
print("   物料长描述识别 / 近邻图号防撞 / 歧义保护 / 原四级回归）")
print("  沙箱:", TMP)
