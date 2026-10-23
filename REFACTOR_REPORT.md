# 标准化工程整理报告

生成日期：2026-09-20

## 完成结论

已在 `D:\文件操作程序` 创建四个相互独立的标准化子项目，并将可运行源码、模板、说明和必要验证资产复制到各自的 `legacy/original_src/`。本次没有对以下原始备份目录执行写入、移动、重命名或删除：

- `D:\文件转换不改格式`
- `D:\提取文件数据`
- `D:\文件转换`
- `D:\特殊文件格式转换`

实际映射经源码确认，与建议映射一致：

| 原始目录 | 目标项目 | 核心旧实现 |
| --- | --- | --- |
| `D:\文件转换不改格式` | `不合格筛选` | `audit_fail.py` / `audit_fail_gui.py` / `api.py` |
| `D:\提取文件数据` | `检验报告产出` | `extract_core` / `extract_gui.py` |
| `D:\文件转换` | `上云数据模板` | `converter.py` / `converter_gui.py` |
| `D:\特殊文件格式转换\检验报告模板转换` | `特殊文件格式转换` | `converter.py` / `converter_gui.py` |

详细扫描依据见 [MAPPING_REPORT.md](MAPPING_REPORT.md)。

## 迁移内容与复制完整性

已复制 178 个源码、说明、打包配置、测试脚本和必要模板/样例文件；排除了 `build/`、`dist/`、`__pycache__/`、打包二进制、构建日志和大量带时间戳的临时运行输出。它们不是业务源码，仍完整保留在原始备份目录。

对每一个复制件执行 SHA-256 与相对路径原文件比对：**178 已检查，0 不一致**。复制范围的具体理由见映射报告“复制范围说明”。

每个新项目均包含：

- `legacy/original_src/`：未重写的原始逻辑复制件；
- 根目录旧调用兼容入口（`audit_fail.py`、`api.py`、`audit_fail_gui.py`；`extract_core`、`extract_gui.py`；`converter.py`、`converter_gui.py`）；
- `backend/app/{services,domain,adapters,schemas,core,api/v1}`：只做调用适配、配置、任务和响应，不改业务算法；
- `frontend/src/api/client.js`：环境变量 baseURL 的前端分离占位；
- `config/`、`.env.example`、`requirements.txt`、`pyproject.toml`、`.gitignore`、`docs/`、`tests/`、`output/`、`logs/`。

## 新增模块与接口

| 项目 | 服务包装 | 提交接口 | 查询接口 |
| --- | --- | --- | --- |
| 不合格筛选 | `screening_service.screen` → 旧 `audit_paths` | `POST /api/v1/nonconforming/screen` | `GET /api/v1/nonconforming/tasks/{task_id}` |
| 检验报告产出 | `report_generation_service.generate` → 旧 `extract_core.api.run_batch` | `POST /api/v1/inspection-reports/generate` | `GET /api/v1/inspection-reports/tasks/{task_id}` |
| 上云数据模板 | `cloud_template_service.convert` → 旧 `converter.convert_file` | `POST /api/v1/cloud-templates/convert` | `GET /api/v1/cloud-templates/tasks/{task_id}` |

每个项目还提供对应任务下载接口和 `/api/v1/health`。成功 JSON 均为 `code/message/data/trace_id` 包络，文件提交使用 multipart。任务当前为进程内、单 worker 的异步队列，专门保持 Excel COM 的单线程约束；`auth_enabled`、CORS、JWT/API Key 接入点和前端 baseURL 已配置化预留。

## 旧功能和新功能的运行方式

1. 进入对应项目，执行 `python -m pip install -r requirements.txt`。
2. 旧桌面方式：分别运行 `python audit_fail_gui.py`、`python extract_gui.py`、`python converter_gui.py`。
3. 旧代码调用：使用根目录兼容入口的 `audit_fail.audit_file/audit_paths`、`extract_core.api.run_batch`、`converter.convert_file`；签名和默认值由 legacy 原实现保持。
4. 新 API：分别运行 `python scripts/run_api.py`，端口为 8000、8001、8002；或使用各项目 README 中的 Uvicorn 命令。

旧不合格筛选 API 也保留根目录 `api.py` 与 `start_api.bat` 兼容入口，真实旧 API 文件保留在 `legacy/original_src/api.py`。

## 验证结果

- 新增 Python 包装层已执行 `py_compile`：通过。
- 三项目兼容入口测试已执行：`6 passed`。
- 不合格筛选服务已用复制的 `不合格_test_93022608_injected.xlsx` 执行冒烟：返回 `total=1`、`fail=1`，并在目标项目 `output/smoke/` 生成 `不合格_test_93022608_injected.xlsx`。输入复制件未被修改。
- FastAPI 运行时导入尚未执行：当前 Python 环境缺少 `fastapi`，并非源码错误；依赖已写入三个 `requirements.txt`。完成安装后应按 README 运行健康检查和上传接口。

## 功能未丢失的验证方法

完整业务回归不能只用自动测试确认，原因是 Excel COM、公式重算、图片/形状、工作簿保护和业务金样在当前环境中不可全部无损模拟。必须按每项目 `docs/测试与验收.md` 在**输入副本**上执行以下对比：

- 不合格筛选：比对输出数量/命名、失败结论、标红结果格、规范化字段、公式、隐藏行、合并区域和 Excel 打开结果。
- 检验报告产出：比对小报告集合、BOM `第N套` 回写、报告已用量后缀、匹配错误码、每张表数据与格式。
- 上云数据模板：比对输出命名、模板字段/公式、值、合并区域、行列尺寸、字体/红色、不规则对象和 Excel 打开结果。

Excel/xlsx 保存后 ZIP 内部元数据可能不同，因此对格式转换和 COM 保存路径不应以二进制完全相同为唯一验收条件；应同时记录输入哈希、Excel 版本、执行参数和人工结论。

## 未决问题与后续建议

1. 确认业务方认可的金样输入/输出、允许的 Excel 版本，以及“上云”平台字段契约后，补充到 `tests/fixtures` 并固化回归断言。
2. 当前任务状态存于进程内；若部署为网站，应在服务层外接数据库/Redis 队列、对象存储和认证网关，但不得改写 `legacy/original_src`。
3. 检验报告产出的旧逻辑会回写 BOM 并推进来源报告的已用量后缀。API 已通过上传副本隔离；直接使用旧函数前仍需由用户确认输入文件可被改写。
4. Windows Excel COM 场景应保持单 worker，并由实际安装 Excel 的 Windows worker 执行。生产发布前先安装 requirements，再跑 `/api/v1/health`、上传冒烟和金样人工验收。

## 2026-10-05 增量同步

三个既有项目中已变更的 legacy 源码已按来源覆盖同步，覆盖前副本保存于各项目 `legacy/snapshots/2026-10-05-pre-sync/`；上云数据模板没有源码差异。新增“特殊文件格式转换”项目，默认 API 端口为 8003。完整清单见 [SYNC_REPORT_2026-10-05.md](SYNC_REPORT_2026-10-05.md)。
