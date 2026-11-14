# 变更记录

## 0.1.4 - 2026-11-14（v19 补丁，未升版本）

- 命名前缀按失败原因细分：**数据写出但超出上下限 → `数据不合格_`**；**仅因数据列未写数据（填 `/`）被判不合格 → `不合格_`**；图号不符 → `图号错误_`；上下限空白补 `/`（规格留空）不加任何前缀。
- 取消"信任检验员结果列"：即便结果列未标 F，也**逐行独立扫描实测值 × 上下限**，把超差行抓出来判 `数据不合格_`。
- 新增 `test_naming.py`（15 项前缀单测）、`test_newtpl.py`（含"不信任结果列"小节）；`regression_test.py` 8/8、`test_newtpl.py` 全过。
- 发布物 `dist/AuditFail_v19.exe` 已重打包覆盖（不升版本）；`dist/` 加入 `.gitignore`，二进制不入库。

## 0.1.3 - 2026-11-14（v19）

- 同步来源目录的 `audit_fail.py`、`audit_fail_gui.py`、`VERSION`、`AuditFail.spec`、`CONVENTIONS.md`、`regression_test.py`。
- `AuditFail.spec` 改用 spec 所在目录的相对路径，去掉本机绝对路径，便于仓库内复现打包。
- 新增中微A0 模板兼容：表头行的「基准值 / 最大值 / 最小值」列定位、r7 连续序号行读数位列定位、尾部签名区截断、Part No 图号回退。
- 新增材质空值补全 `6061-T6`，只在值格为空时写，已有材质不覆盖。
- 判定口径、命名规则、输出不覆盖源文件、旧模板路径均未改写。`regression_test.py` 8/8 通过，专项 20 项通过。
- 同步前副本保留在 `legacy/snapshots/2026-11-14-pre-sync/`。发布产物 `dist/AuditFail_v19.exe` 不在仓库内。

## 0.1.2 - 2026-10-23

- 同步来源的审核核心、GUI、打包配置、版本和 `regression_test.py`。
- 同步前副本保留在 `legacy/snapshots/2026-10-23-pre-sync/`。

## 0.1.1 - 2026-10-05

- 同步原始来源的 `audit_fail.py`、GUI、旧 API、前端页面、打包配置和验证脚本；新增 `verify_v11.py`。
- 同步前版本已保留于 `legacy/snapshots/2026-10-05-pre-sync/`。

## 0.1.0 - 2026-09-20

- 从只读备份 `D:\文件转换不改格式` 复制原始源码至 `legacy/original_src`。
- 新增不改变核心逻辑的 CLI/API/服务包装、配置、文档和回归测试骨架。
