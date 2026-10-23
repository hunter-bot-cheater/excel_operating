# 特殊文件格式转换

## 项目用途与来源

本项目从只读来源 `D:\特殊文件格式转换\检验报告模板转换` 整理而来。它把特殊布局的原始检验数据转换为内置 `template.xlsx`（LJG-FH56-1 格式）检验报告：数据取自源工作簿，模板负责布局和样式。项目根目录的 `converter.py`、`converter_gui.py`、`template.xlsx`、`VERSION` 是可直接维护和运行的原始程序及资源；`legacy/original_src/检验报告模板转换` 保留迁移副本。

## 功能清单

- 输入：`.xls/.xlsx/.xlsm` 特殊格式检验报告；输出：源图号命名的 `.xlsx`，不合格时前缀为 `不合格_`，同名由原逻辑自动避让。
- 处理：识别表头和数据区，抽取图号、报告编号、温度、检验人、项目、读数与上下限；按内容识别行类型，合并读数范围，空上下限填 `/`。
- 判定：源结果为 F 或测量值超上下限时，输出结果格标红并写 F。
- 格式：优先 Excel COM 写入模板以保留样式；不可用时由 openpyxl 回退。`.xls` 需要本机 Excel COM 转换。

## 目录结构

根目录的 `converter.py`、`converter_gui.py`、`template.xlsx` 是原始维护文件；`legacy/original_src/检验报告模板转换` 为迁移副本并保留原输入样例；`backend/app` 提供服务、领域契约、适配器、配置和 API，适配器从项目根加载原始代码；`frontend/src/api` 是 Web 前端请求占位；`docs`、`config`、`tests`、`output`、`logs` 用于交付和运行。

## 环境、安装与配置

- Windows 10/11、Python 3.10+、`openpyxl`；完整 `.xls` 和模板 COM 写入需要桌面 Excel、`pywin32`。GUI 依赖 PySide6。
- 安装：`python -m pip install -r requirements.txt`。
- 参考 `config/config.example.yaml` 与 `.env.example` 配置端口、路径、CORS 和鉴权开关。默认值使用项目相对目录；不保存账号、密码或密钥。

## 运行方式

- 旧 GUI：`python converter_gui.py`。
- 旧 CLI：`python converter.py 输入文件.xlsx -o output`；可加 `--no-com` 强制旧 openpyxl 路径。
- 旧代码调用：`from converter import convert_file`，签名为 `convert_file(input_path, out_dir=None, on_progress=None, xl=None, template_path=None, force_openpyxl=False)`。
- 新 API：`python scripts/run_api.py`，或在 `backend` 下运行 `python -m uvicorn app.main:app --host 127.0.0.1 --port 8003 --workers 1`。
- 验证：`python -m pytest`；格式和 Excel COM 行为再按测试文档人工验收。

## 核心函数/类

| 旧函数 | 参数/返回 | 说明 |
| --- | --- | --- |
| `convert_file` | 返回 `input/output/is_fail/rows/tech/slash_filled/log` | 唯一业务转换入口；COM 失败自动回退 openpyxl |
| `find_template_path` | 可选模板路径 | 默认发现复制的 `template.xlsx` |
| `prepare_rows` / `row_is_fail` | 解析项目行 | 保留原不合格判定规则 |
| `compute_output_path` | 解析结果、输入、输出目录 | 保留旧命名及不合格前缀规则 |

## API、数据格式与错误

- `POST /api/v1/special-format-conversions/convert`：multipart `file`，可选 `force_openpyxl`；返回 `task_id`。
- `GET /api/v1/special-format-conversions/tasks/{task_id}`：返回状态、进度和旧函数结果。
- `GET /api/v1/special-format-conversions/tasks/{task_id}/downloads/{filename}`：下载生成工作簿。
- 成功响应：`{"code":0,"message":"ok|accepted","data":{},"trace_id":"..."}`；422 是扩展名错误，404 是任务或文件不存在；原转换异常在任务 `failed` 状态中保留。
- 支持格式、字段和模板要求见 [数据结构](docs/数据结构.md)。

## 测试与验收

先运行 `python -m pytest` 验证兼容入口。再将原 `test_in` 的样例复制到临时目录，以原程序和本项目兼容入口分别运行；比较输出文件名、图号/报告编号/温度/检验人、项目行、上下限 `/`、读数范围、F 标红、模板布局、图片/对象和 Excel 可打开性。不要用 ZIP 字节完全一致作为唯一标准。详见 [测试与验收](docs/测试与验收.md)。

## 扩展兼容性、安全与日志

已预留 `/api/v1`、异步任务、CORS、`auth_enabled` 和前端 `VITE_API_BASE_URL`。后续可在不修改 legacy 核心的前提下接入 JWT/API Key、数据库任务表、Redis/Celery、对象存储、Docker、多用户和权限。上传文件先复制至任务隔离目录；日志避免记录表格内容或凭据，响应带 `trace_id`。

## 常见问题

- **`.xls` 失败**：旧逻辑要求 Windows Excel COM；安装 Excel/pywin32 后重试。
- **模板找不到**：确认 `legacy/original_src/检验报告模板转换/template.xlsx` 未移动。
- **样式与原模板不完全一致**：检查是否走了 COM；`force_openpyxl` 仅应作为兼容降级选项。

## 变更记录与旧新对应表

见 [CHANGELOG](CHANGELOG.md) 和 [扩展兼容性](docs/扩展兼容性.md)。

| 旧文件/函数 | 新位置/接口 |
| --- | --- |
| `converter.py` / `convert_file` | 根目录原程序；`backend/app/services/special_format_service.py` 从根目录加载 |
| `converter_gui.py` | 根目录原 GUI 入口 |
| `template.xlsx` | 根目录原模板，legacy 中保留副本 |
