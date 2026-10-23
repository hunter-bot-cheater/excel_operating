# 上云数据模板

## 项目用途与来源

本项目从只读备份 `D:\文件转换` 整理而来，将 `.xls/.xlsx/.csv` 检验报告转换为“上云数据”标准 `.xlsx` 模板。原转换实现、GUI、模板、历史版本和验证源码均复制在 `legacy/original_src`；新模块只是稳定的服务/API/兼容层，未改写原格式转换算法。

## 功能清单

- 输入：`.xls`、`.xlsx` 或 UTF-8-SIG CSV；输出：标准 `.xlsx`，默认由原 `compute_output_path` 命名为 `【上云数据】【尺寸】<图号>.xlsx` 并自动避让重名。
- 处理：解析原报告布局、图号、日期、规格/上下限、读数和结果，重建上云模板；按原规则清理图号、补全短日期、统一合格结果、处理不合格红色和检具代码。
- 边界：Excel 文件优先 Windows COM 原地处理以保留列宽、图片、文本框等；不可用时回退 openpyxl 重建。CSV 无形状复制路径。

## 目录结构

`legacy/original_src/converter.py` 是唯一核心算法；根目录 `converter.py` 保留旧 import；`backend/app/services` 仅委派原 `convert_file`；`domain`、`adapters`、`api/v1`、`frontend/src/api`、`config`、`docs`、`tests`、`output`、`logs` 是标准化层。

## 环境、安装与配置

- Windows 10/11、Python 3.10+、`openpyxl`、`xlrd`；若要保留 xls/xlsx 的内嵌对象，安装 Excel 和 `pywin32`。GUI 需要 PySide6。
- 安装：`python -m pip install -r requirements.txt`。
- 使用 `config/config.example.yaml` 或 `.env.example` 配置运行路径、CORS、鉴权开关；默认没有密钥、账号或硬编码用户路径。

## 运行方式

- 旧 GUI：`python converter_gui.py`。
- 旧核心调用：`from converter import convert_file`，调用 `convert_file(input_path, out_dir=None, on_progress=None, default_year=None)`。
- API：`python scripts/run_api.py`，或在 `backend` 下执行 `python -m uvicorn app.main:app --host 127.0.0.1 --port 8002 --workers 1`。
- 测试：`python -m pytest`；完整格式回归按 [测试与验收](docs/测试与验收.md) 人工完成。

## 核心函数/类

| 旧函数/类 | 参数/返回 | 异常/说明 |
| --- | --- | --- |
| `convert_file` | `input_path, out_dir, on_progress, default_year, xl`；返回 `input/output/total_qty/tech/dim` | 扩展名、布局、依赖、COM/输出锁定错误会抛出 |
| `compute_output_path` | 输入、输出目录、总件数、年份 | 保留旧文件名规则 |
| `ensure_deps` | 无 | 旧入口可能有安装副作用；服务层不调用它 |

## API、数据格式与错误

- `POST /api/v1/cloud-templates/convert`：multipart `file`，可选 `default_year`（四位数），返回 `task_id`。
- `GET /api/v1/cloud-templates/tasks/{task_id}`：查询进度和原函数返回。
- `GET /api/v1/cloud-templates/tasks/{task_id}/downloads/{filename}`：下载生成 xlsx。
- 成功返回 `{"code":0,"message":"ok|accepted","data":{},"trace_id":"..."}`；422 表示扩展名或年份参数无效，404 为任务/输出不存在，`failed` 状态保留原异常。
- 模板字段和格式详见 [数据结构](docs/数据结构.md)，原模板在 `legacy/original_src/template_nda.xlsx`。

## 测试与验收

运行 pytest 先验证兼容入口。将金样复制后，分别运行原 `converter.py` 和根目录兼容入口；对比输出命名、工作表、单元格值/公式、合并区域、列宽、行高、字体、红色标记、图片/形状和 Excel 可打开性。由于重建/COM 保存会改变 ZIP 元数据，不以二进制相同为标准。

## 扩展兼容性、安全与日志

已提供 `/api/v1`、异步 `task_id`、CORS/`auth_enabled` 配置和前端 baseURL；后续可添加 JWT/API Key、Docker、数据库/队列、对象存储、多用户权限，但适配层必须继续调用 legacy 的 `convert_file`。上传文件先隔离复制，日志不写入敏感内容，响应含 `trace_id`。

## 常见问题

- **图片/文本框没有保留**：确认 Windows Excel/pywin32 可用；COM 失败时原实现会回退。
- **短日期年份不正确**：在 API 传入 `default_year` 或沿用旧 GUI 输入。
- **输出被占用**：关闭目标 Excel 文件后重试。

## 变更记录与旧新对应表

详见 [CHANGELOG](CHANGELOG.md) 与 [扩展兼容性](docs/扩展兼容性.md)。

| 旧文件/函数 | 新位置/接口 |
| --- | --- |
| `converter.py` / `convert_file` | `legacy/original_src/converter.py`；`backend/app/services/cloud_template_service.py` |
| `converter_gui.py` | 根目录同名 GUI 兼容入口 |
| `template_nda.xlsx` | `legacy/original_src/template_nda.xlsx` |
