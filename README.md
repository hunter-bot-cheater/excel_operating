# 文件操作程序工程总览

本目录将四个原始只读备份工程标准化为彼此独立的子项目。原始目录未被移动、重命名、删除或修改；可运行源码的复制件均在各项目 `legacy/original_src`。

| 原始备份 | 新项目 | 核心旧入口 | 新 API |
| --- | --- | --- | --- |
| `D:\文件转换不改格式` | [不合格筛选](不合格筛选/README.md) | `audit_fail.py` | `POST /api/v1/nonconforming/screen` |
| `D:\提取文件数据` | [检验报告产出](检验报告产出/README.md) | `extract_core.api.run_batch` | `POST /api/v1/inspection-reports/generate` |
| `D:\文件转换` | [上云数据模板](上云数据模板/README.md) | `converter.convert_file` | `POST /api/v1/cloud-templates/convert` |
| `D:\特殊文件格式转换\检验报告模板转换` | [特殊文件格式转换](特殊文件格式转换/README.md) | `converter.convert_file` | `POST /api/v1/special-format-conversions/convert` |

## 统一规范

- 后端均为 FastAPI 结构，统一 `/api/v1` 前缀、JSON 包络 `code/message/data/trace_id`、multipart 上传、任务提交/查询/下载三段式接口。
- 三个服务默认单 worker，以兼容 Windows Excel COM；前端目录只提供 baseURL API 封装，不锁定具体框架。
- 配置位于每项目 `config/config.example.yaml` 和 `.env.example`；不含账号、密码或密钥。所有新文本均为 UTF-8。

## 启动

分别进入子项目执行 `python -m pip install -r requirements.txt`，然后运行其 `python scripts/run_api.py`（默认端口依次为 8000、8001、8002、8003），或运行同名 GUI 兼容入口。详细的输入、输出、错误和验收要求以各项目 README 为准。

## 回滚

本次未改原始目录；回滚标准化层只需停止新 API/GUI 并继续从原始目录运行。若需要撤回目标目录内容，请先保留 `MAPPING/_REPORT.md` 与 `REFACTOR_REPORT.md`，再由人工确认后处理；本次不会自动删除任何文件。

映射依据详见 [MAPPING_REPORT.md](MAPPING_REPORT.md)。增量同步记录见 [2026-10-05](SYNC_REPORT_2026-10-05.md) 与 [2026-10-23](SYNC_REPORT_2026-10-23.md)。

## GitHub 版本控制

本目录是 Git 仓库根目录，远程为 `https://github.com/hunter-bot-cheater/excel_operating.git`，四个项目分别位于本目录的四个子文件夹。日常更新可执行：

```powershell
git pull --rebase origin main
git add .
git commit -m "描述本次修改"
git push origin main
```

根 `.gitignore` 会排除运行输出、真实测试工作簿、本机路径诊断脚本和本地快照，避免把用户数据误提交到公开仓库。需要共享测试时请使用脱敏/合成样例。
