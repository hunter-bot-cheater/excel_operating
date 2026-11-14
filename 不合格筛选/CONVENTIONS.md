# 代码规范约束（AuditFail）

这些不是风格偏好，是踩过坑之后定下来的硬约束。改代码前先读完，违反任意一条都会导致客户拿到错的结果或丢掉原文件。

## 1. 判定口径只允许有一处实现

`audit_fail._is_fail_token` 与 `D:\文件转换\converter.py::_is_fail_token` 必须逐字符等价。

改判定时：

1. 先改 `audit_fail.py` 的 `_FAIL_TOKENS` / `_is_fail_token`
2. 同步改 `converter.py` 的同名函数
3. 跑 `python regression_test.py`（8 个真实样本）
4. 跑 `python analysis/a18/verify_token_a18.py`（35 项逐条断言）
5. 两边用同一组词跑一次对比，输出必须完全一致

**为什么**：两个工具对同一份报告必须给出同一个结论，否则客户会拿两个互相矛盾的结果来问。

**反面教材**：曾经把 `startswith("f")` 改成整格匹配，修好了 `FQC` 误判，却漏掉了 `F(略)` 这种「F + 中文备注」。现在靠 `_only_cjk` 兜住这一类。

## 2. 输出绝不覆盖原文件

所有落盘必须走 `_write_output(write_fn, base_path)`，它内部调 `_unique_target` 做 `(1)(2)` 递增避让。

- 不许直接 `wb.save(原路径)`
- 不许自己写 `while os.path.exists(...)` 循环
- 图号错误的改名路径也必须走 `_unique_target`（`exclude=` 排除自身，否则会把自己判成占用）

**为什么**：客户的报告是唯一证据，覆盖一次就没了。这个工具的名字叫「不覆盖」，覆盖就是 P0 事故。

## 3. Excel COM 的生命周期

| 规则 | 原因 |
|---|---|
| 每个 worker 进程一个独立 Excel，不许跨进程共享 | COM 在 Windows 上跨进程共享会 RPC 超时 |
| 退出必须 `Quit()` + 按 PID `taskkill`，PID 要在 `Quit` **之前**取 | `Quit` 之后窗口销毁，`hwnd` 失效就取不到 PID，僵尸进程会一直挂着并锁住文件 |
| 只杀自己 `DispatchEx` 出来那个 PID | 用户可能自己开着 Excel，误杀会丢他的数据 |
| 实例死掉后不要复用坏句柄 | 坏句柄每次调用都会等 RPC 超时，比崩溃更难排查 |

## 4. 异常处理

- 不许裸 `except: pass` 吞掉。要么 `_dbg(...)` 留痕，要么明确降级并把降级原因返回给调用方。
- 单文件失败必须隔离，不能冒泡出批处理把整批带崩。
- 写盘失败要删掉半成品，否则残留文件会污染后续 `(n)` 编号。

`_dbg()` 只在调试开关打开时输出，生产环境静默，但代码里保留调用点，排查时能立刻打开。

## 5. 并发与进度

- 并发上限 2，份数 < 6 或核数 < 4 时退化为串行。实测 2 路在目标机器上最快。
- 进度 = (已完成份数 + 当前份内百分比) / 总份数，**只许前进不许回退**。
  单份文件必须发中间进度（读取 → 逐表扫描 → 汇总 → 写出），否则只有一份文件时
  进度条会从 0 直接跳到 100，看着像没跑。
- UI 线程不许做重活，也不许碰 COM。
- 判断「在跑还是卡死」看有没有事件，不看总耗时：worker 每 5s 发一次心跳，
  父进程只按「多久没有任何事件」判卡死，大文件慢不会被误杀。

## 6. 界面行为

- 结果列表逐条实时插入（`file_done` 信号），不等整批结束。
- 结束时只更新汇总文案，不重建列表，避免重复行。
- 颜色语义固定：合格绿 `#27AE60`、不合格与失败红 `#E74C3C`、
  提醒（图号不完全一致等）橙黄 `#E6A23C`。提醒不算失败，汇总里单独列。

## 7. 提交前检查清单

```bash
python -m py_compile audit_fail.py audit_fail_gui.py
python regression_test.py                      # 8/8
python test_samples/_verify_a_unify.py
python analysis/a18/verify_token_a18.py        # 改了判定才跑
```

顺手确认：跑完不留无窗口 EXCEL（`quit()` 里已带强杀，若仍有残留说明 PID 取晚了）。

## 8. 新增一个判定词 / 新增一种文件格式

- 加判定词：改 `_FAIL_TOKENS`，同步 `converter.py`，跑第 1 条的全套检查。
- 新增格式：在 `_read_source` 加分支，保证 `_sheet_grid` 返回值结构不变（`(grid, nrows, ncols)`，0-based），否则下游全崩。
- 新增模板分支必须是**兼容回退**：只在旧逻辑识别失败（返回 None / 空）时启用，不得改写旧模板已能识别的路径。改完必须跑 `regression_test.py`（8/8）＋ `test_newtpl.py`。
- 尾部签名/说明区不得计入数据区。判据：最后一个结果值之后、A 列全为非编号文本时截断；尾部若有真实编号行（结果列待补 P/F）保持原行为。

## 关窗口 / 取消

- **GUI 必须在 `closeEvent` 里先取消后台任务**：`worker.stop()` → `pool.cancel()`。
  QThread 还在跑就让进程退出，子进程和 Excel 都成了孤儿，且没有 finally 收尾。
- **取消要杀干净**：`cancel()` 里除了杀 worker，还要调 `sweep_orphan_excel(批次前快照)`。
  Excel 是 DCOM 拉起的，不在 worker 进程树里；`ready` 未到时连 PID 都不知道，
  唯一可靠判据是「批次前不存在 + 无可见窗口」。
- **批次收尾也扫一遍**：`run()` 的 finally 里同样调 `sweep_orphan_excel`，
  这是「跑十批不留一个 EXCEL」的最后一道保险；用户自己开的有窗口实例永远跳过。
- **不加无限的等待**：任何 `while`、`join`、等 `ready`、等进程退出，都必须有 deadline
  （本项目统一 3-5s），并在超时后清理而不是继续等。
