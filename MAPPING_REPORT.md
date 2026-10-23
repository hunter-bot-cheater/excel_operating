# 原始目录功能映射分析报告

> 扫描时间：2026-09-20。扫描方式：只读目录清单、Python 源码入口/导入/公共函数和现有需求文档检查。本报告生成时尚未对三个原始目录执行写入、移动、重命名或删除操作。

## 结论

默认映射与源码实际功能一致，采用下表的映射。

| 原始只读备份目录 | 实际功能 | 目标子项目 | 判断依据 |
| --- | --- | --- | --- |
| `D:\文件转换不改格式` | 识别检验报告不合格项，规范化客户代码/上下限，并在保持工作簿格式的前提下输出结果 | `不合格筛选` | 核心文件名为 `audit_fail.py`，其文档和 API 均为“审核不合格数据” |
| `D:\提取文件数据` | 根据 BOM 和检验大报告抽取测量数据，生成小报告并回写 BOM 套数列 | `检验报告产出` | `需求文档.md`、`extract_core` 和 GUI 明确描述“提取文件数据/生成独立小报告” |
| `D:\文件转换` | 将 `.xls/.xlsx/.csv` 检验报告转换为“上云数据”标准 `.xlsx` 模板 | `上云数据模板` | `converter.py` 输出以 `【上云数据】【尺寸】` 为前缀的标准化工作簿 |

## 目录一：`D:\文件转换不改格式`

### 结构与文件清单（按用途归类）

- 核心业务：`audit_fail.py`（55,589 bytes）；历史兼容版本：`audit_fail_orig_v5.py`。
- 桌面入口：`audit_fail_gui.py`；历史 GUI：`audit_fail_gui_orig_v5.py`。
- HTTP 入口：`api.py`；启动脚本：`start_api.bat`；静态前端：`web/index.html`。
- 验证/测试脚本：`test_api.py`、`test_api_v7.py`、`validate_real.py`、`verify_v6.py` 至 `verify_v10.py`、`verify_v7_com.py`。
- 打包配置和版本：`AuditFail.spec`、`VERSION`；样例输入：`不合格_test_93022608_injected.xlsx`。
- 临时/构建产物：`build/`、`dist/`、`__pycache__/`（包含多份 PyInstaller `.exe/.pkg/.pyz`，不是迁移核心源码）。

### 入口、依赖和运行方式

- 旧 GUI：`python audit_fail_gui.py`，依赖 PySide6。
- 旧 API：`python api.py` 或 `python -m uvicorn api:app --host 127.0.0.1 --port 8000`，现有 API 为 `/api/audit`、`/api/tasks/{task_id}`、下载接口和健康检查。
- 核心公共兼容函数：`scan_fails`、`detect_fail`、`audit_file`、`audit_paths`、`preload_openpyxl`、`clear_preloaded`。
- 运行依赖：Python、`openpyxl`；可选/慢路径依赖 Windows Excel 与 `pywin32` (`win32com`)；GUI/API 还需要 PySide6/FastAPI/Uvicorn。

### 当前功能、输入输出与异常

- 输入 `.xls/.xlsx/.xlsm` 检验报告，扫描所有工作表的结果、测量值和上下限。
- 无公式且不含复杂元素时使用 openpyxl 快路径；公式、`.xls` 或复杂工作簿使用 Excel COM 重算慢路径。
- 任一工作表命中不合格时，输出文件名前增加 `不合格_`，且仅将对应结果单元格标红/按规则写入 `F`；不覆盖输入文件，同名输出自动增加 `(1)` 序号。
- 同时执行报告规范化：客户代码统一为 `0998`、缺少最终判定时报 `final_missing`、有效数据区空的上下限写入 `/`。
- 典型异常：不支持扩展名、Excel/COM 不可用、公式重算失败、工作簿无法解析、输出被锁定。核心会避开输出目录和已有 `不合格_` 文件以防重复扫描。

### 迁移风险与处理

- COM 依赖 Windows 和已安装 Excel，必须单线程；新接口将保留该约束。
- Excel 图表、图片、透视表、隐藏行和公式是格式保真风险点；必须用原样本进行字节外结构与人工 Excel 打开对比。

## 目录二：`D:\提取文件数据`

### 结构与文件清单（按用途归类）

- 核心包：`extract_core/__init__.py`、`api.py`、`extractor.py`、`bom_reader.py`、`bom_writer.py`、`report_reader.py`、`filename_parser.py`、`models.py`、`errors.py`。
- 桌面入口：`extract_gui.py`；辅助脚本：`_dl.py`、`_fix.py`。
- 规格、依赖、打包：`需求文档.md`、`requirements.txt`、`ExtractTool.spec`、`VERSION`。
- 测试源：`test/*.py`（BOM 写回、端到端、多源、重名、xls、真实样例等）；`test/fixtures` 和 `test/realinput` 为验证样例。
- 临时/构建产物：`build/`、`dist/`、`__pycache__/`、大量带时间戳的 `test/combo_*`、`test/sandbox_*`、`test/matchrep_*`、`test/numkey_*` 等运行输出，不作为核心源码迁移。

### 入口、依赖和运行方式

- 旧 GUI：`python extract_gui.py`。
- 可供集成的旧服务函数：`extract_core.api.run_batch`、`extract_core.api.process_report`；抽取执行函数为 `extract_core.extractor.run_one_task[_multi]`。
- 依赖清单原文：`PySide6>=6.0`、`openpyxl>=3.1`、`pywin32>=306`、`pyinstaller>=6.0`。

### 当前功能、输入输出与异常

- 输入一个 BOM 和一组 `.xls/.xlsx/.xlsm` 大报告（文件、目录可混合）。
- 解析 `图号_组件描述_总件数件_日期-已用` 文件名；用报告内部 DWG/零件名、文件名前缀和数字键与 BOM 逐级匹配。
- 依据 BOM 数量、已用量、报告可用量切分抽取段，生成独立小报告；文件名的已用量后缀会随抽取推进；BOM 回写/插入 `第N套` 列。
- `.xls`/复杂写入经 Excel COM，其他路径使用 openpyxl；生成文件需保留模板布局。
- 典型异常码/情况：`BOM_READ_ERROR`、`BOM_MULTI_MATCH`、`FILE_PARSE_ERROR`、`NO_BOM_MATCH`、Excel 被占用或写入失败。未被报告覆盖的 BOM 行按原逻辑静默跳过。

### 迁移风险与处理

- 原业务可能就地重命名/推进输入报告的已用量后缀、回写 BOM，属于有状态操作；新包装默认仍调用原逻辑，API 仅在显式允许时执行，测试一律在副本上进行。
- 匹配规则复杂且已有大量边界测试；不重写算法，只通过适配器调用复制的 `extract_core`。

## 目录三：`D:\文件转换`

### 结构与文件清单（按用途归类）

- 核心业务：`converter.py`（当前版本）；历史版本：`v38/converter.py`。
- 桌面入口：`converter_gui.py`；历史 GUI：`v38/converter_gui.py`。
- 验证/诊断：`test_convert.py`、`verify_dump.py`、`verify_pvd_nda.py`、`pvd_compare.py`、`f_red_test.py`、`dump_source.py`。
- 模板/打包：`template_nda.xlsx`、`InspectionConverter_v12.spec`、`v38/InspectionConverter_v38.spec`。
- 性能试验和历史备份：`bench_tmp/*.py`、`bench_tmp/backup/*.py`，以及其中的临时输入/输出 Excel。
- 构建/输出：`build/`、`dist/`、`test_out/`、`__pycache__/`、`build_log*.txt`，不是迁移核心源码。

### 入口、依赖和运行方式

- 旧 GUI：`python converter_gui.py`。
- 核心兼容函数：`convert_file(input_path, out_dir=None, on_progress=None, default_year=None, xl=None)`；依赖检查函数 `ensure_deps`。
- 支持输入：`.xls`、`.xlsx`、`.csv`；运行时需 `openpyxl`、`xlrd`。特定 `.xls` 或形状复制场景依赖 Windows Excel/pywin32。

### 当前功能、输入输出与异常

- 读取不同来源的检验报告，抽取图号、日期、测量项目、规格上下限和结果。
- 重建为标准 `.xlsx` 上云模板，统一合格判定、数值格式、图号/文件名清理、日期补全、检具代码、上下限不合格标红等规则。
- 输出名由 `compute_output_path` 计算，现有样例显示为 `【上云数据】【尺寸】<图号>.xlsx`，同名自动避让。
- 典型异常：不支持输入类型、必需布局/表头无法识别、依赖缺失、Excel COM/形状复制失败、输出文件被占用。

### 迁移风险与处理

- 该工具是“格式转换/模板重建”，不能以字节相同作为验收标准；要按单元格值、公式、样式、合并区域、图片/形状以及文件名对比。
- `ensure_deps` 会在旧脚本中自动安装依赖；新服务不调用该副作用函数，改由声明式依赖安装，但兼容入口仍保留原行为。

## 共性技术识别与待确认事项

- 已识别：三个项目均为 Python；前两个有 PySide6 GUI；第一个已有 FastAPI + HTML 页面；三个项目都处理 Excel；不存在 VBA 源文件、数据库实现、真实云接口或消息队列实现。Windows Excel COM 是可选但重要的兼容适配器。
- 未发现明文账号、密码或密钥；新配置只提供示例，不写入敏感信息。
- 待人工确认：每个项目最终用户认可的金样输入/输出对、允许的 Excel 版本、是否允许 API 任务修改原始 BOM/报告文件、以及“上云”平台的实际字段契约。未确认前，新 API 仅提供骨架和显式路径参数，旧兼容入口保持原逻辑。

## 复制范围说明

下一阶段会将可执行 Python、GUI、API、批处理、规格、模板、打包配置和必要测试/样例复制到各项目 `legacy/original_src/`。`build/`、`dist/`、`__pycache__/`、运行日志、打包二进制和时间戳运行输出不作为“原始源码”复制；其路径和用途已在本报告记录，原始备份仍完整保留。

## 2026-10-05 新增目录：`D:\特殊文件格式转换\检验报告模板转换`

### 映射结论

| 原始只读备份目录 | 实际功能 | 目标子项目 | 判断依据 |
| --- | --- | --- | --- |
| `D:\特殊文件格式转换\检验报告模板转换` | 将特殊布局检验报告转换为 LJG-FH56-1 模板，识别不合格并保留模板格式 | `特殊文件格式转换` | `converter.py` 文档、`template.xlsx`、GUI 和打包配置均明确该功能 |

### 文件、入口与依赖

- 核心：`converter.py`；桌面入口：`converter_gui.py`；模板：`template.xlsx`；版本：`VERSION`；打包配置：`InspectionConverter.spec`；样例输入：`test_in/*.xlsx`。
- 支持 `.xls/.xlsx/.xlsm`，依赖 `openpyxl`；`.xls` 与高保真模板写入依赖 Windows Excel COM 和 `pywin32`；GUI 依赖 PySide6。
- `convert_file(input_path, out_dir=None, on_progress=None, xl=None, template_path=None, force_openpyxl=False)` 是旧核心兼容入口。
- 处理：按内容检测布局，读取表头与项目，空上下限填 `/`，读数合并，不合格结果写 F/红底；优先 COM 写模板，失败回退 openpyxl。
- `build/`、`dist/` 和 `test_out/` 是二进制/运行产物，没有迁移；原始目录保持不变。
