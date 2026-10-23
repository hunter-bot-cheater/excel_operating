# -*- coding: utf-8 -*-
"""边界场景实测：用户提出的 4 类多义/重复情形（纯匹配层，无 COM）。

A. 「必须同时满足图号+描述才能认证」？—— 澄清：联合信号优先，但单键也可命中
B. 同时满足多个【不等价】行 -> 歧义报错
C. 两行图号一致、描述是子集关系 -> 精确行优先（拼接全等先行）
D. 图号是子集、描述一致 -> 精确行优先
E. 两行完全一样（图号+描述+数量都同）-> 不是歧义，返回 2 行 = 2 个任务
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from extract_core.bom_reader import match_report
from extract_core.models import BomRow

def R(bh, desc, qty, row):
    return BomRow(bh=bh, desc=desc, qty=qty, row=row, sheet="Sheet1", locs=[])

fails = []

def show(tag, got):
    if isinstance(got, list):
        return f"{len(got)}行: " + " | ".join(
            f"{r.bh}／{r.desc}×{r.qty}(r{r.row})" for r in got)
    return str(got)

# ---------- A. 单键也可命中（不是必须双满足） ----------
rows_a = [R("93011182", "气路法兰", 1, 2)]
r = match_report("93011182", rows_a)          # 只图号（无描述段）
assert len(r) == 1 and r[0].desc == "气路法兰", f"A1 {r}"
r = match_report("93011182_气路法兰", rows_a)  # 图号+描述
assert len(r) == 1, f"A2 {r}"
r = match_report("93011182", rows_a, dwg="93011182", part="气路法兰")  # 仅内部身份
assert len(r) == 1, f"A3 {r}"
print("[OK] A 单键可命中：只写图号 / 图号+描述 / 仅内部身份 均可认证（联合信号优先，非必须）")

# ---------- B. 同时满足多个不等价行 -> 歧义 ----------
rows_b = [R("93021561", "CF100焊接管", 1, 2), R("93021561", "CF40焊接管", 2, 3),
          R("93021561", "ISO200焊接管", 3, 4)]
for front in ("93021561_焊接管", "93021561_CF"):
    try:
        match_report(front, rows_b)
        fails.append(f"B: {front} 应歧义却未报")
    except Exception as e:
        assert "BOM_MULTI_MATCH" in str(e) and "歧义" in str(e), f"B {front}: {e}"
print("[OK] B 多行同时满足（不等价）-> BOM_MULTI_MATCH 歧义报错，列候选")

# ---------- C. 图号一致、描述是子集 ----------
rows_c = [R("93021561", "焊接管", 1, 2), R("93021561", "CF100焊接管", 2, 3)]
r = match_report("93021561_焊接管", rows_c)   # 有精确行 -> 拼接全等优先
assert len(r) == 1 and r[0].desc == "焊接管" and r[0].qty == 1, f"C1 {r}"
r = match_report("93021561_CF100焊接管", rows_c)
assert len(r) == 1 and r[0].qty == 2, f"C2 {r}"
rows_c2 = [R("93021561", "CF100焊接管", 2, 3)]  # 无精确行 -> 描述包含命中
r = match_report("93021561_焊接管", rows_c2)
assert len(r) == 1 and r[0].qty == 2, f"C3 {r}"
print("[OK] C 描述子集：存在精确行 -> 精确命中（不歧义）；无精确行 -> 包含命中；多行公共子串 -> 歧义")

# ---------- D. 图号是子集、描述一致 ----------
rows_d = [R("93021561", "X", 1, 2), R("93021561-A", "X", 2, 3)]
r = match_report("93021561_X", rows_d)        # 精确拼接行优先
assert len(r) == 1 and r[0].bh == "93021561", f"D1 {r}"
r = match_report("93021561-A_X", rows_d)
assert len(r) == 1 and r[0].bh == "93021561-A", f"D2 {r}"
rows_d2 = [R("93021561", "X1", 1, 2), R("93021561-A", "X", 2, 3)]
r = match_report("93021561_X", rows_d2)       # 描述"x"是"x1"子串 -> 数字键兜底收敛到精确描述行
assert len(r) == 1 and r[0].bh == "93021561-A", f"D3 {r}"
print("[OK] D 图号子集：精确拼接行优先；无精确行时按图号数字键+描述收敛到唯一行")

# ---------- E. 两行完全一样 -> 多行同图号 -> 直接报错（v6.2 唯一行口径） ----------
rows_e = [R("93011111", "测试件", 2, 5), R("93011111", "测试件", 2, 6)]
for tag, kw in (("E1 拼接全等", dict(front="93011111_测试件")),
                ("E2 纯图号", dict(front="93011111")),
                ("E3 内部身份", dict(front="93011111_测试件", dwg="93011111", part="测试件"))):
    try:
        match_report(**kw, rows=rows_e)
        fails.append(f"{tag}: 完全相同的两行应报错，却未报")
    except Exception as e:
        assert "BOM_MULTI_MATCH" in str(e), f"{tag}: {e}"
print("[OK] E 两行完全一样（等价行）-> 同图号多行直接报错（不再拆 2 个任务）")

if fails:
    print("\nFAILED:", *fails, sep="\n  ")
    sys.exit(1)
print("\n==== 边界场景全部通过 ====")
