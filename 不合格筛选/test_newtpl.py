# -*- coding: utf-8 -*-
"""JO-WL-PZ-010 中微A0 模板兼容专项测试。

覆盖：布局识别 / 数值越界判定 / 上下限空白填 / / 总数与检查数 / 材质补 6061-T6 /
Part No 图号 / 旧模板零回归（由 regression_test.py 另行保证）。
"""
import os
import shutil
import sys
import tempfile

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import audit_fail as af

TPL = (r"C:\Users\胡雨航\xwechat_files\wxid_9zhv2r31ydny22_b109"
       r"\msg\file\2026-11"
       r"\TM LID, C6V3_615-05411-000，615-05411-000 刻字 25-07-0021 2025-07-15.xlsx")

fails = []


def check(name, cond, detail=""):
    if cond:
        print("  [OK] %s" % name)
    else:
        print("  [FAIL] %s %s" % (name, detail))
        fails.append("%s %s" % (name, detail))


def load_grid(path):
    wb = openpyxl.load_workbook(path, data_only=False)
    ws = wb.worksheets[0]
    grid, nrows, ncols, _hf = af._openpyxl_grid(ws)
    return wb, ws, grid, nrows, ncols


def main():
    if not os.path.exists(TPL):
        print("样本缺失，跳过:", TPL)
        return
    wb, ws, grid, nrows, ncols = load_grid(TPL)

    print("=== A. 布局识别 ===")
    lay = af._detect_layout(grid, nrows, ncols)
    check("表头行 r5", lay["header_row"] == 4, str(lay["header_row"]))
    check("规格列 = C/E/F", (lay["spec_col"], lay["max_col"], lay["min_col"]) == (2, 4, 5),
          str((lay["spec_col"], lay["max_col"], lay["min_col"])))
    check("结果列 = M", lay["result_col"] == 12, str(lay["result_col"]))
    check("读数列 = H..L（5 位）", lay["reading_cols"] == [7, 8, 9, 10, 11],
          str(lay["reading_cols"]))
    check("数据起始 r8", lay["data_start"] == 7, str(lay["data_start"]))

    print("=== B. 判定 ===")
    sf, reason = af._scan_one_sheet(grid, nrows, ncols)
    rows = []
    for f in sf:
        rows.append(f["r"])
    check("结果列 4 个 F 全检出", rows == [83, 106, 107, 147], str(rows))

    # 数值越界：清空结果列 F，把 r22 的 H 改成超上限 33（上限 32.2）
    g2 = []
    for row in grid:
        g2.append(list(row))
    for r in range(7, 151):
        g2[r][12] = ""
    g2[21][7] = "33"
    sf2, reason2 = af._scan_one_sheet(g2, nrows, ncols)
    hit22 = False
    for f in sf2:
        if f["r"] == 22:
            hit22 = True
    check("H22=33 超上限 32.2 → 判 F", hit22, reason2[:120])

    print("=== C. 规范化探测 ===")
    si = af._probe_norm_sheet(grid, nrows, ncols, 1)
    check("n_meas = 1（只有 H 列有读数）", si["n_meas"] == 1, str(si["n_meas"]))
    check("斜杠 44 格", len(si["slash_cells"]) == 44, str(len(si["slash_cells"])))
    only_ef = True
    for _s, _r, c in si["slash_cells"]:
        if c not in (4, 5):
            only_ef = False
            break
    check("斜杠只落在 E/F 列", only_ef)
    check("页脚不补结果列", len(si["result_fill"]) == 0, str(len(si["result_fill"])))
    check("材质标签已定位", si["material"] is not None and
          si["material"]["r"] == 3 and si["material"]["c"] == 4, str(si["material"]))
    check("图号 = 615-05411-000",
          si["drawing_no"] is not None and si["drawing_no"]["value"] == "615-05411-000",
          str(si["drawing_no"]))
    wb.close()

    print("=== D. 总数 / 检查数不被误改 ===")
    wb, ws, grid, nrows, ncols = load_grid(TPL)
    norm = af._collect_norm([(grid, nrows, ncols, set())])
    check("Total Qty 原值 1 保持不变", len(norm["total_qty"]) == 0, str(norm["total_qty"]))
    check("Sample Qty 原值 1 保持不变", len(norm["sample_qty"]) == 0, str(norm["sample_qty"]))
    check("图号核对 ok", af._dwg_match("615-05411-000", os.path.basename(TPL))[0] == "ok")
    wb.close()

    print("=== E. 端到端：材质为空补 6061-T6 ===")
    tmp = tempfile.mkdtemp(prefix="newtpl_")
    try:
        src = os.path.join(tmp, "in.xlsx")
        shutil.copyfile(TPL, src)
        outdir = os.path.join(tmp, "out")
        os.makedirs(outdir, exist_ok=True)
        res = af.audit_file(src, out_dir=outdir)
        out = res.get("output")
        check("输出文件生成", bool(out) and os.path.exists(out), str(out))
        if out and os.path.exists(out):
            wbo = openpyxl.load_workbook(out, data_only=False)
            wso = wbo.worksheets[0]
            check("H4 = 6061-T6", str(wso["H4"].value).strip() == af.MATERIAL_DEFAULT,
                  repr(wso["H4"].value))
            check("E10/F10 空白已填 /", str(wso["E10"].value).strip() == "/" and
                  str(wso["F10"].value).strip() == "/",
                  "%r/%r" % (wso["E10"].value, wso["F10"].value))
            check("数值行 E22 未被填 /", str(wso["E22"].value).strip() == "32.200",
                  repr(wso["E22"].value))
            check("结果列 F 仍为 4 处", res.get("is_fail") is True)
            check("材质操作写入日志", any(x.get("op") == "材质" for x in res.get("log", [])))
            wbo.close()

        # 材质已有值 → 不改
        src2 = os.path.join(tmp, "in2.xlsx")
        shutil.copyfile(TPL, src2)
        wb2 = openpyxl.load_workbook(src2)
        wb2.worksheets[0]["H4"] = "SUS304"
        wb2.save(src2)
        wb2.close()
        outdir2 = os.path.join(tmp, "out2")
        os.makedirs(outdir2, exist_ok=True)
        res2 = af.audit_file(src2, out_dir=outdir2)
        out2 = res2.get("output")
        if out2 and os.path.exists(out2):
            wb3 = openpyxl.load_workbook(out2)
            check("已有材质 SUS304 不被覆盖", str(wb3.worksheets[0]["H4"].value).strip() == "SUS304",
                  repr(wb3.worksheets[0]["H4"].value))
            wb3.close()
        else:
            check("已有材质副本输出生成", False)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    # === F. 不信任结果列：数据超差但结果列未标 F，程序必须独立扫描抓出 ===
    print("=== F. 不信任结果列（独立扫描） ===")
    for mode, mval in (("结果列写合格P", "P"), ("结果列空白", None)):
        tmpf = tempfile.mkdtemp(prefix="trust_")
        try:
            srcf = os.path.join(tmpf, "in.xlsx")
            shutil.copyfile(TPL, srcf)
            wbf = openpyxl.load_workbook(srcf)
            wsf = wbf.worksheets[0]
            # 行8：下限0 上限0 实测0；改成实测5（超差），结果列保持 P 或空白（不标F）
            wsf.cell(row=8, column=5).value = "0.000"
            wsf.cell(row=8, column=6).value = "0.000"
            wsf.cell(row=8, column=8).value = 5
            wsf.cell(row=8, column=13).value = mval
            wbf.save(srcf)
            wbf.close()
            outdirf = os.path.join(tmpf, "out")
            os.makedirs(outdirf, exist_ok=True)
            resf = af.audit_file(srcf, out_dir=outdirf)
            basef = os.path.basename(resf.get("output") or "")
            check("%s -> 判不合格" % mode, resf.get("is_fail") is True,
                  "is_fail=%s" % resf.get("is_fail"))
            check("%s -> 前缀含 数据不合格_" % mode, "数据不合格_" in basef, basef)
            check("%s -> 原因含 超出公差" % mode, "超出公差" in resf.get("reason", ""),
                  resf.get("reason", ""))
        finally:
            shutil.rmtree(tmpf, ignore_errors=True)

    print()
    if fails:
        print("FAILED %d:" % len(fails))
        for f in fails:
            print("  -", f)
        sys.exit(1)
    print("新模板兼容专项测试全部通过")


if __name__ == "__main__":
    main()
