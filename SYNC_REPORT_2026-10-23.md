# 2026-10-23 增量同步报告

## 同步范围与保护措施

本次仅从四个原始来源读取代码：`D:\文件转换不改格式`、`D:\提取文件数据`、`D:\文件转换`、`D:\特殊文件格式转换\检验报告模板转换`。写入仅发生在 `D:\文件操作程序`；没有移动、重命名、删除或写入原始来源。

来源文件覆盖前，旧 legacy 复制件已保存至各项目 `legacy/snapshots/2026-10-23-pre-sync/`。来源中已不再出现的历史样例、fixture 或辅助文件继续保留，未删除。

## 同步明细

| 项目 | 新增 | 覆盖同步 | 结论 |
| --- | --- | --- | --- |
| 不合格筛选 | `regression_test.py` | `audit_fail.py`、`audit_fail_gui.py`、`AuditFail.spec`、`VERSION` | 已同步 5 文件 |
| 检验报告产出 | `_run_tests.py`、`_verify.py`、`test/_repro_multi_sheet.py`、`test/_selftest_run.py`、`test/fake_xlsx_test.py`、`test/gauge_row_test.py` | `extract_core/bom_writer.py`、`extractor.py`、`report_reader.py`、`ExtractTool.spec`、`test/k5_test.py` | 已同步 11 文件 |
| 上云数据模板 | 无 | 无 | 当前 legacy 核心、GUI、模板与来源一致 |
| 特殊文件格式转换 | 无 | `converter.py`、`converter_gui.py`、`VERSION` | 已同步 3 文件 |

后端服务/API/前端占位没有被来源代码覆盖。它们继续通过各项目的 legacy loader 调用同步后的核心实现，因此旧 GUI、兼容入口、新 CLI/API 都使用当前同一份业务逻辑。

## 验证方法

1. 每个已同步文件均应与来源执行 SHA-256 比对。
2. 运行全部兼容入口测试：`python -m pytest -q 不合格筛选/backend/tests 检验报告产出/backend/tests 上云数据模板/backend/tests 特殊文件格式转换/backend/tests`。
3. 对实际业务金样，在来源程序与目标项目根目录兼容入口中使用输入副本执行，再比对文件名、业务结果、Excel 格式、公式、图片/形状和可打开性。
4. Excel COM 路径必须在已安装 Microsoft Excel 的 Windows 环境进行人工回归；不能只用 openpyxl 或 ZIP 二进制哈希判断格式一致性。
