# 2026-10-05 增量同步报告

## 原则

本次只读取以下原始来源：`D:\文件转换不改格式`、`D:\提取文件数据`、`D:\文件转换`、`D:\特殊文件格式转换`。所有写入均位于 `D:\文件操作程序`；未删除任何原始文件，也没有从原始目录移动任何文件。

## 三个既有项目同步结果

| 项目 | 新增 | 覆盖同步（已先快照） | 未删除的历史副本 |
| --- | --- | --- | --- |
| 不合格筛选 | `verify_v11.py` | `api.py`、`audit_fail.py`、`audit_fail_gui.py`、`AuditFail.spec`、`verify_v8.py`、`web/index.html` | 输入样例继续保留 |
| 检验报告产出 | 无 | `extractor.py`、`models.py`、`report_reader.py`、`e2e_test.py`、`k5_test.py`、`multi_source_test.py`、`xls_support_test.py` | `_fix.py` 和历史 fixtures/realinput 继续保留 |
| 上云数据模板 | 无 | 无，源码副本已与来源一致 | `template_nda.xlsx` 继续保留 |

被覆盖的旧版本已复制到各自 `legacy/snapshots/2026-10-05-pre-sync/`；早期同步目录中的同内容副本也保留，未做删除。因此同步不会丢失此前版本。业务包装层没有改写，继续按导入的 legacy 模块执行。

## 新增项目

来源 `D:\特殊文件格式转换\检验报告模板转换` 映射至 `D:\文件操作程序\特殊文件格式转换`。

- 已复制：`converter.py`、`converter_gui.py`、`InspectionConverter.spec`、`template.xlsx`、`VERSION`、`test_in/` 样例。
- 未复制：`build/`、`dist/`、`test_out/`，因为它们是打包二进制或运行输出；原目录完整保留。
- 新接口：`POST /api/v1/special-format-conversions/convert`，任务状态和下载接口见该项目 README。

## 验证要求

同步后应对四个项目在输入副本上分别运行原入口和目标兼容入口，核对文件名、业务结果、Excel 样式和可打开性。特殊格式项目重点核对模板布局、上下限 `/`、读数范围和 F 红色标记；涉及 COM 的 `.xls` 或复杂模板需要在 Windows+Excel 环境人工验证。
