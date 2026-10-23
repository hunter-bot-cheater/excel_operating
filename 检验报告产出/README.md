# 检验报告产出

## 项目用途与来源

本项目从只读备份 `D:\提取文件数据` 整理而来：依据 BOM 和检验大报告抽取指定件数的测量数据，生成独立检验报告，并按旧规则回写 BOM 套数。原 `extract_core`、GUI、规格和测试源码完整保留于 `legacy/original_src`；新层不重写匹配、拆分、写回或命名算法。

## 功能清单

- 输入：一个 `.xls/.xlsx/.xlsm` BOM 与一个或多个同格式检验大报告（旧 GUI 还支持文件夹递归收集）。
- 输出：按 BOM 数量抽取的独立报告，报告文件名已用量后缀和 BOM 的 `第N套` 列遵从原逻辑。
- 处理：报告文件名解析，DWG/零件名/文件名/数字键多信号匹配，按可用件数分段，生成报告并可选择 `round_no` 回写。
- 边界：无匹配 BOM 行静默跳过；解析、匹配歧义和写入失败以旧错误码返回。API 始终在上传副本上执行，避免修改用户原文件。

## 目录结构

`legacy/original_src/extract_core` 是原核心包；根目录 `extract_core` 保留旧导入路径；`backend/app/services` 只适配 `run_batch`；`domain`、`adapters`、`api/v1`、`frontend/src/api`、`config`、`docs`、`tests`、`output`、`logs` 以职责分层。

## 环境、安装与配置

- Windows 10/11、Python 3.10+、`openpyxl`；原项目对于 Excel 写入/`.xls` 支持依赖桌面 Excel 与 `pywin32`。GUI 需要 PySide6。
- 安装：`python -m pip install -r requirements.txt`。
- 从 `config/config.example.yaml` 或 `.env.example` 复制环境配置；默认 API 临时副本位于 `backend/.runtime`，不写死用户路径，也不保存敏感信息。

## 运行方式

- 旧 GUI：`python extract_gui.py`。
- 旧核心调用：`from extract_core.api import run_batch`，调用 `run_batch(bom_path, report_paths, out_dir, round_no)`。
- API：`python scripts/run_api.py`，或从 `backend` 执行 `python -m uvicorn app.main:app --host 127.0.0.1 --port 8001 --workers 1`。
- 自动检查：`python -m pytest`；端到端测试须用复制的 BOM/报告样例。

## 核心函数/类

| 旧函数/类 | 参数/返回 | 异常/说明 |
| --- | --- | --- |
| `extract_core.api.run_batch` | `bom_path, report_paths, out_dir, round_no`；返回任务汇总 | 保留 `BOM_READ_ERROR`、`NO_BOM_MATCH` 等原错误结果 |
| `process_report` | 单报告兼容处理入口 | 旧 API 的公共函数 |
| `extractor.run_one_task[_multi]` | 任务段、输出和进度回调 | 负责 Excel 抽取与报告生成 |
| `BomWriter` | BOM 回写适配 | 仅 `round_no` 有效时按旧规则写入 |

## API、数据格式与错误

- `POST /api/v1/inspection-reports/generate`：multipart `bom`、`reports`，可选表单 `round_no`；返回 `task_id`。
- `GET /api/v1/inspection-reports/tasks/{task_id}`：任务状态、进度和旧结果。
- `GET /api/v1/inspection-reports/tasks/{task_id}/downloads/{filename}`：输出文件流。
- 统一成功包络：`{"code":0,"message":"ok|accepted","data":{},"trace_id":"..."}`。422 为扩展名/参数错误，404 为未知任务或输出；业务失败保留在 `status=failed` 或旧汇总的文件错误中。
- BOM 表头、报告文件名和 Excel 字段严格遵从 [数据结构](docs/数据结构.md) 与原 [需求文档](legacy/original_src/需求文档.md)。

## 测试与验收

运行 pytest 验证兼容包。对每个金样，在原项目与新项目的**复制输入**中运行，比较输出文件集合、工作表值/公式/格式、来源报告已用量后缀、BOM `第N套` 写回值和所有错误码。原逻辑存在写入副作用，绝不能对三份原始备份测试。详见 [测试与验收](docs/测试与验收.md)。

## 扩展兼容性、安全与日志

已预留 `/api/v1`、CORS、`auth_enabled`、异步任务、前端 baseURL。后续可外置任务状态到数据库/队列、将文件输出到对象存储、增加 Docker 和权限，但服务层必须继续调用原 `run_batch`。上传会隔离到运行时副本；日志不记录表格内容/凭据，响应携带 `trace_id`。

## 常见问题

- **BOM 表头不识别**：原规则只接受精确的 `组件数量` 列，请按原需求文档调整。
- **多行匹配**：这是原 `BOM_MULTI_MATCH` 保护，不能静默任选一行。
- **Excel 被占用**：关闭 BOM/报告/输出文件后重试；COM 任务应单 worker 执行。

## 变更记录与旧新对应表

详见 [CHANGELOG](CHANGELOG.md) 与 [扩展兼容性](docs/扩展兼容性.md)。

| 旧文件/函数 | 新位置/接口 |
| --- | --- |
| `extract_core/api.py` / `run_batch` | `legacy/original_src/extract_core/api.py`；`backend/app/services/report_generation_service.py` |
| `extract_gui.py` | 根目录同名 GUI 兼容入口 |
| `extract_core` 包 | 根目录同名兼容包，保持旧 import 方式 |
