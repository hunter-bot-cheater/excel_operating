# -*- coding: utf-8 -*-
"""AuditFail 引擎交付前回归测试 —— 每版打包前必须全绿。

用法：
    python regression_test.py

覆盖：
  1. 真实客户样本判定断言（合格/不合格）；
  2. 布局模板变体专项（结果列已有 F 时，即使上下限/表头识别退化也不得漏判）；
  3. 老 .xls → .xlsx 转换路径；
  4. 输出文件存在性、非空、扩展名。

新增样本：把真实客户文件复制进 test_samples/ 并在 CASES 里登记预期。
"""
import os
import shutil
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import audit_fail as af

SAMPLES_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "test_samples")

# (文件名, 预期 is_fail, 说明)；预期 None = 只要求处理成功、不硬断言判定
CASES = [
    ("93014428_prt-（1）_Shutter Storage_1件_26.9.01(1).xlsx", True,
     "布局干扰（Dimension说明行）+ 结果列已有F（历史漏判样本）"),
    ("93022011_Lid到位压块_1件_2026.8.29.xlsx", False, "合格基线"),
    ("93010LKS-DM_LeftSupportArm_1件_26.6.18(1).xlsx", False, "合格（读数均在公差内）"),
    ("93021MXK_调整板_1件_26.5.13..(2).xls", None, "老 .xls → .xlsx 转换路径"),
    ("93020825_prt_bellow转接件_下阶_25件_2026.8.11-1.xlsx", None, "25 件多表拆分"),
    ("93022499-新PD_托指垫块下_84件_2026.4.26-27.xlsx", None, "84 件多表跨表累加"),
]

_SHUTTER = "93014428_prt-（1）_Shutter Storage_1件_26.9.01(1).xlsx"


def _run_one(src, outdir):
    """复制到临时文件（保留扩展名）→ audit_file → 返回 (res, copy_path)。"""
    ext = src.rsplit(".", 1)[-1]
    tmp = tempfile.mkdtemp(prefix="reg_")
    copy = os.path.join(tmp, f"in.{ext}")
    shutil.copyfile(src, copy)
    os.chmod(copy, 0o666)
    out = os.path.join(outdir, os.path.basename(copy))
    try:
        res = af.audit_file(copy, out_dir=outdir)
        return res, out
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def case_tests(outdir):
    passed = failed = 0
    for name, expect, desc in CASES:
        src = os.path.join(SAMPLES_DIR, name)
        if not os.path.exists(src):
            print(f"[FAIL] 样本缺失: {name}")
            failed += 1
            continue
        try:
            res, out = _run_one(src, outdir)
            got = res.get("is_fail")
            out_ok = bool(res.get("output")) and os.path.exists(res.get("output"))
            ext_ok = res.get("output", "").lower().endswith(".xlsx")
            judge_ok = (expect is None) or (got == expect)
            if judge_ok and out_ok and ext_ok:
                print(f"[PASS] {name} | is_fail={got} | {desc}")
                passed += 1
            else:
                print(f"[FAIL] {name} | is_fail={got}（预期 {expect}）| 输出存在={out_ok} xlsx={ext_ok} | {desc}")
                print(f"       reason: {str(res.get('reason'))[:160]}")
                failed += 1
        except Exception as e:
            print(f"[FAIL] {name} 异常: {e}")
            failed += 1
    return passed, failed


def variant_tests(outdir):
    """模板变体专项：结果列已有 F 时，上下限/表头识别退化也不得漏判。"""
    passed = failed = 0
    base = os.path.join(SAMPLES_DIR, _SHUTTER)
    if not os.path.exists(base):
        print("[FAIL] 变体基底样本缺失（Shutter Storage）")
        return 0, 1
    tmp = tempfile.mkdtemp(prefix="var_")
    try:
        # 变体A：删掉子表头行的 上限/下限/公称 文字 → 上下限识别退化，结果列 F 必须仍检出
        from openpyxl import load_workbook
        wb = load_workbook(base)
        ws = wb.worksheets[0]
        for cell in (ws.cell(row=10, column=3), ws.cell(row=10, column=4), ws.cell(row=10, column=5)):
            cell.value = None
        vA = os.path.join(tmp, "varA.xlsx")
        wb.save(vA)
        resA, _ = _run_one(vA, outdir)
        if resA.get("is_fail") is True:
            print("[PASS] 变体A 无上下限表头+结果列F → 判不合格")
            passed += 1
        else:
            print(f"[FAIL] 变体A 无上下限表头+结果列F → is_fail={resA.get('is_fail')}")
            failed += 1

        # 变体B：改掉 Item No 表头 → 走兜底标签定位，结果列 F 必须仍检出
        wb = load_workbook(base)
        ws = wb.worksheets[0]
        ws.cell(row=9, column=1).value = "编号(无ItemNo字样)"
        vB = os.path.join(tmp, "varB.xlsx")
        wb.save(vB)
        resB, _ = _run_one(vB, outdir)
        if resB.get("is_fail") is True:
            print("[PASS] 变体B 无ItemNo表头+结果列F → 判不合格")
            passed += 1
        else:
            print(f"[FAIL] 变体B 无ItemNo表头+结果列F → is_fail={resB.get('is_fail')}")
            failed += 1
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    return passed, failed


def main():
    print("=" * 70)
    print("AuditFail 引擎交付前回归测试")
    print("=" * 70)
    outdir = tempfile.mkdtemp(prefix="regout_")
    try:
        p1, f1 = case_tests(outdir)
        p2, f2 = variant_tests(outdir)
        total_p = p1 + p2
        total_f = f1 + f2
        print("=" * 70)
        print(f"结果：通过 {total_p} / 失败 {total_f}")
        if total_f:
            print("结论：FAIL —— 禁止打包交付")
            sys.exit(1)
        print("结论：PASS —— 可以打包交付")
    finally:
        shutil.rmtree(outdir, ignore_errors=True)


if __name__ == "__main__":
    main()
