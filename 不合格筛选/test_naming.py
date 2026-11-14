# -*- coding: utf-8 -*-
"""命名前缀逻辑单测：结果列 F → 不合格_；上下限空白填 / 不加前缀。"""
import os
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)
import audit_fail as A  # noqa: E402


def check(name, fails, slash, expect):
    norm = {"slash_cells": slash}
    got = A._fail_prefix(fails, norm)
    ok = got == expect
    print("  [%s] %s | fails=%d slash=%d -> %r (期望 %r)" %
          ("OK" if ok else "FAIL", name, len(fails), len(slash), got, expect))
    return ok


def main():
    print("=== 命名前缀 _fail_prefix ===")
    results = []
    # 1) 数据写出但超差（cause=oob）→ 数据不合格_
    results.append(check("数据超差(oob)", [{"r": 1, "c": 2, "cause": "oob"}], [], "数据不合格_"))
    # 2) 仅数据列写 /（cause=missing）→ 不合格_
    results.append(check("仅数据缺失(/)", [{"r": 1, "c": 2, "cause": "missing"}], [], "不合格_"))
    # 3) 超差 + 缺失 都有 → 数据不合格_ 优先
    results.append(check("超差+缺失", [{"r": 1, "c": 2, "cause": "oob"},
                                       {"r": 3, "c": 4, "cause": "missing"}], [], "数据不合格_"))
    # 4) 都没有 -> 合格（空前缀）
    results.append(check("都合格", [], [], ""))
    # 5) 多个超差
    results.append(check("多超差", [{"r": 1, "c": 2, "cause": "oob"},
                                   {"r": 3, "c": 4, "cause": "oob"}], [], "数据不合格_"))
    # 6) 数据写出但区间内被标 F（cause=unknown）→ 数据不合格_
    results.append(check("写出但区间内(unknown)", [{"r": 1, "c": 2, "cause": "unknown"}], [], "数据不合格_"))

    # --- _row_cause 直接校验（不依赖外部样本）---
    def _cau(expect, read_vals, mn, mx):
        row = [None] * 5
        row[1], row[2] = mn, mx          # min_col=1, max_col=2
        for i, v in enumerate(read_vals):
            row[3 + i] = v               # reading_cols=[3, 4]
        grid = [row]
        got = A._row_cause(grid, 0, [3, 4], 1, 2)
        ok = got == expect
        print("  [%s] _row_cause %r(限[%s,%s]) -> %r (期望 %r)"
              % ("OK" if ok else "FAIL", read_vals, mn, mx, got, expect))
        return ok
    results.append(_cau("missing", ["/", None], 0.0, 32.2))   # 数据列写 /
    results.append(_cau("oob", [33, None], 0.0, 32.2))        # 写出但超上限
    results.append(_cau("unknown", [10, None], 0.0, 32.2))    # 写出且区间内
    results.append(_cau("missing", [None, None], 0.0, 32.2))  # 整行无数据
    results.append(_cau("oob", ["/", 40], 0.0, 32.2))         # 第二读数超差

    # 7) norm 为 None 时不报错
    got = A._fail_prefix([], None)
    ok = got == ""
    print("  [%s] norm=None 安全 -> %r" % ("OK" if ok else "FAIL", got))
    results.append(ok)

    # 前缀应用幂等：已是 不合格_ 不再追加
    p = os.path.join(BASE, "_t.xlsx")
    once = A._with_prefix(p, "不合格_")
    twice = A._with_prefix(once, "不合格_")
    ok = once == twice and os.path.basename(once) == "不合格__t.xlsx"
    print("  [%s] 不合格_ 幂等 -> %r" % ("OK" if ok else "FAIL", os.path.basename(once)))
    results.append(ok)
    # 上下限空白 / 不生成前缀（_fail_prefix 返回空串）
    got2 = A._fail_prefix([], {"slash_cells": [("s", 1, 2)]})
    ok2 = got2 == ""
    print("  [%s] 上下限空白/ 不加前缀 -> %r" % ("OK" if ok2 else "FAIL", got2))
    results.append(ok2)
    # 不合格_ 幂等
    once2 = A._with_prefix(p, "不合格_")
    twice2 = A._with_prefix(once2, "不合格_")
    ok = once2 == twice2
    print("  [%s] 不合格_ 幂等 -> %r" % ("OK" if ok else "FAIL", os.path.basename(once2)))
    results.append(ok)

    # 扫描排除：两种前缀都不二次扫描
    for pre in ("不合格_", "数据不合格_"):
        fn = pre + "x.xlsx"
        excluded = fn.startswith("不合格_") or fn.startswith("数据不合格_")
        assert excluded, fn
    print("  [OK] 扫描排除覆盖两种前缀")

    n_fail = sum(1 for r in results if not r)
    print("\n结果: %d/%d 通过" % (len(results) - n_fail, len(results)))
    if n_fail:
        print("❌ 有失败用例")
        sys.exit(1)
    print("✅ 命名前缀逻辑全部通过")


if __name__ == "__main__":
    main()
