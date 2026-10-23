# 不合格筛选

## 项目用途与来源

本项目从只读备份 `D:\文件转换不改格式` 整理而来，用于审核 `.xls/.xlsx/.xlsm` 检验报告，筛选不合格项并输出规范化后的工作簿。项目根目录的 `audit_fail.py`、`api.py`、`audit_fail_gui.py` 是可直接维护和运行的原始程序；`legacy/original_src` 保留迁移副本。**业务判定、命名、编码、默认参数和 Excel 格式规则均未改写。**

## 功能清单

- 输入：单文件、文件夹或 API 上传的 `.xls/.xlsx/.xlsm`；输出：原名或 `不合格_` 前缀的同格式工作簿，同名自动避让。
- 处理：全工作表扫描结果/测量/上下限；命中时只标红结果单元格并按旧规则写 `F`；同时保留旧版客户代码 `0998`、最终判定和上下限 `/` 规范化。
- 边界：无公式的普通 xlsx/xlsm 用 openpyxl；公式、xls 或复杂对象降级 Windows Excel COM。源文件不会被覆盖。

## 目录结构

根目录原始程序为维护源；`legacy/original_src` 是迁移时副本；`backend/app/services` 是用例包装，`adapters` 从项目根加载原始代码；`domain` 是请求契约；`api/v1` 是 FastAPI 路由；`frontend/src/api` 是前端请求占位；`config`、`docs`、`tests`、`output`、`logs` 分别存放示例配置、文档、验证、运行产物和日志。

## 环境、安装与配置

- Windows 10/11、Python 3.10+；要完整处理公式、xls 或复杂 Excel 文件，请安装桌面 Excel 和 `pywin32`。
- `python -m pip install -r requirements.txt`
- 复制 `config/config.example.yaml` 或 `.env.example` 后按环境修改；不得把账号、密码、密钥写入仓库。默认输出路径在项目 `output/`，API 临时文件在 `backend/.runtime/`。

## 运行方式

- 旧 GUI：`python audit_fail_gui.py`
- 旧核心调用：`python -c "import audit_fail; print(audit_fail.audit_file(r'输入.xlsx', r'output'))"`
- 新 API：`python scripts/run_api.py`，或在 `backend` 下运行 `python -m uvicorn app.main:app --host 127.0.0.1 --port 8000 --workers 1`。
- 测试：`python -m pytest`。COM 场景必须在安装 Excel 的 Windows 机器上人工复核。

## 核心函数/类

| 旧函数/类 | 参数/返回 | 异常与说明 |
| --- | --- | --- |
| `audit_file(input_path, out_dir=None, on_progress=None, session=None)` | 返回 `input/output/is_fail/reason/norm` 字典 | 文件不存在、Excel/COM 或解析错误向调用者抛出 |
| `audit_paths(paths, out_dir=None, on_progress=None)` | 返回 `total/fail` | 批量共享一个 Excel COM 会话 |
| `scan_fails(path)` / `detect_fail(path)` | 不改变旧调用签名 | 仅扫描/判定；实现仍在 legacy |
| `ExcelSession` | COM 会话复用 | 只能串行使用 |

## API、数据格式与错误

- `POST /api/v1/nonconforming/screen`：multipart 字段 `files`，提交后返回 `task_id`。
- `GET /api/v1/nonconforming/tasks/{task_id}`：查询进度、旧函数返回值和可下载文件。
- `GET /api/v1/nonconforming/tasks/{task_id}/downloads/{filename}`：下载输出。
- 所有成功响应为 `{"code":0,"message":"ok|accepted","data":{},"trace_id":"..."}`；422 是输入扩展名错误，404 是任务/文件不存在，任务 `status=failed` 保留原异常文本。
- 文件必须是原工具支持的 Excel 格式；保留原工作簿编码、模板字段、工作表和命名规则。字段细节见 [数据结构](docs/数据结构.md)。

## 测试与验收

运行 `python -m pytest` 验证兼容入口和源码副本。再以原样本在“原目录”和新项目的**副本**各运行一次，比对文件名、每张表的值/公式/合并区域/行列尺寸、红色结果格以及 Excel 可打开性；不要要求 xlsx ZIP 字节完全一致。详见 [测试与验收](docs/测试与验收.md)。

## 扩展兼容性、安全与日志

接口已预留 `/api/v1`、CORS 配置、`auth_enabled`、Bearer JWT/API Key 接入点、异步任务 ID 和前端 baseURL。Docker、数据库、队列、对象存储、多用户与权限只能在不改变 legacy 业务函数的适配层增加。上传文件会先复制到运行时目录；日志带时间，响应带 `trace_id`；不要记录文件内容、账号或密钥。

## 常见问题

- **公式结果不同**：确认 Windows Excel 可用，保持单 worker，并以 Excel COM 路径复测。
- **输出被占用**：关闭 Excel 中打开的目标文件后重试。
- **接口任务重启后消失**：当前任务状态为进程内实现；接入数据库/队列前不承诺跨进程持久化。

## 变更记录与旧新对应表

见 [CHANGELOG](CHANGELOG.md)、[扩展兼容性](docs/扩展兼容性.md)。

| 旧文件/函数 | 新位置/接口 |
| --- | --- |
| `audit_fail.py` / `audit_file` | 根目录原程序；`backend/app/services/screening_service.py` 从根目录加载 |
| `audit_fail_gui.py` | 根目录原 GUI 入口 |
| `api.py` | 根目录原 API 入口；新 REST API 为 `/api/v1/nonconforming/*` |
