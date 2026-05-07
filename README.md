# PreSubmissionAIQC

中文名称：投稿前AI数据真实性与合理性质控系统

PreSubmissionAIQC 是一个 Windows 桌面端投稿前质控助手。用户上传包含补充数据、原始表格、图片、PDF、分析脚本等文件的 zip 包后，软件会执行本地文件完整性检查、表格解析、确定性数值规则检查、图片待检包生成、外部 AI 工具状态整合和报告输出。

## 软件用途

- 投稿前发现数据文件缺失、解析失败、异常数值模式和需要人工复核的风险信号。
- 生成结构化输出：`raw_file_manifest.csv`、`checksum_manifest.csv`、`sheet_inventory.xlsx`、`numeric_qc_results.xlsx`、`QC_issue_log.xlsx`、`image_check_package.zip`、`external_ai_status.xlsx`、`final_QC_report.pdf` 和 `final_QC_report.html`。
- 帮助实验室把 Red / Orange 问题集中到人工复核清单中。

## 适用数据类型

- Excel / CSV / TSV 表格。
- RNA-seq DEG、代谢组、GO/KEGG enrichment、gene list、figure source、通用数值表。
- PDF、PNG、JPG、TIFF、CZI、ND2 等图片或稿件文件。
- Python、R、ipynb 等脚本文件会进入 manifest。

## 不适用范围

- 本软件不能直接定性研究不端。
- 本软件不内置 Proofig / Imagetwin / DataSeer 的商业检测模型。
- 本地模块不做复杂图像篡改判断，只生成图片待检包和导入外部报告。
- DataSeer 类检查不等于数值真实性检查。

## 安装开发依赖

```powershell
python -m pip install -r requirements.txt
```

需要 Python 3.11+。如果当前机器的 `python` 指向旧版本，请先修正 PATH 或使用 Python Launcher。

## 开发运行

```powershell
python main.py
```

## 打包 exe

```powershell
build_exe.bat
```

打包后运行：

```text
dist/PreSubmissionAIQC/PreSubmissionAIQC.exe
```

## 给未安装环境用户的发布方式

最终用户不需要安装 Python、PySide6、pandas、reportlab、keyring 或 PyInstaller。推荐发布方式是生成 one-folder 便携版或 zip 包：

```powershell
build_release.bat
```

该脚本会生成：

```text
release/PreSubmissionAIQC-portable.zip
```

把这个 zip 发给用户，用户解压后双击：

```text
PreSubmissionAIQC.exe
```


## 用户使用流程

1. 启动 GUI。
2. 输入 Project Name。
3. 选择待检查 zip 压缩包。
4. 可选选择 `data_dictionary.xlsx`、`sample_info.xlsx` 和外部 AI 报告。
5. 如有官方 API 权限，在“设置 API Key / 外部AI工具”中填写 endpoint 和 API key。
6. 点击“开始检查”。
7. 检查完成后在结果概览、表格预览和输出按钮中查看结果。

## API key 安全说明

- API key 只通过 `keyring` 保存到系统安全凭据存储，例如 Windows Credential Manager。
- Endpoint、启用状态和 LLM model name 可保存到 `config/external_ai_settings.json`。
- API key 不写入配置文件、日志、PDF、HTML 或 `runs` 输出目录。
- 如果 keyring 不可用，软件会提示 API key 仅本次会话使用，不会明文降级保存。

## 外部AI工具说明

- Proofig AI：科学图片完整性筛查，适用于重复、复用、旋转、翻转、裁剪、局部复制、blot/gel 风险和 AI 生成图风险筛查。
- Imagetwin：图片重复、复用、篡改、抄袭和 AI 生成图筛查。
- DataSeer：数据共享声明、代码共享、协议共享和投稿政策合规检查，不用于直接判断数值真实性。
- LLM：仅用于总结 QC 结果和生成返查建议，不直接定性研究不端。

Proofig / Imagetwin / DataSeer 是外部商业或机构服务，本软件不内置其商业检测模型。

## 无 API 时的手动检查包

没有官方 API 权限时，本软件仍会生成：

```text
runs/<run>/outputs/image_check/image_check_package.zip
```

用户可手动上传到对应平台，获得 PDF / CSV / XLSX 报告后再通过 GUI 导入。

## 输出文件说明

- `outputs/tables/raw_file_manifest.csv`：文件清单。
- `outputs/tables/checksum_manifest.csv`：MD5 / SHA256 校验清单。
- `outputs/tables/sheet_inventory.xlsx`：表格和 sheet 解析结果。
- `outputs/tables/numeric_qc_results.xlsx`：数值规则检查结果。
- `outputs/tables/QC_issue_log.xlsx`：人工复核问题清单。
- `outputs/image_check/image_check_package.zip`：图片和 PDF 待检包。
- `outputs/external_ai/external_ai_status.xlsx`：外部工具状态。
- `outputs/reports/final_QC_report.pdf`：最终 PDF 报告。
- `outputs/reports/final_QC_report.html`：HTML 摘要报告。

## 风险等级

- Red：投稿前必须解决。
- Orange：需要回查原始记录，未复核前不建议投稿。
- Yellow：可解释异常或轻微问题，建议记录。
- Green：未见明显风险信号。

Red / Orange 问题必须由人工结合原始记录复核。

## 已知限制

- 外部 AI adapter 是通用 skeleton，需要用户按官方 API 文档配置真实 endpoint。
- PDF 报告为 MVP 结构化报告，版式后续可继续增强。
- 本地图片模块只打包和登记，不做商业级图片完整性判断。
- Excel 解析依赖 pandas/openpyxl；损坏、加密或特殊格式文件会登记为解析失败并继续流程。

## 免责声明

本软件的本地数值检查用于发现风险信号，不能直接定性研究不端。本报告用于投稿前数据质量和研究诚信风险筛查。自动化结果不能直接定性研究不端。所有 Red 和 Orange 问题均需结合原始记录、实验记录本、仪器导出文件、未裁剪原始图片和人工复核确认。本报告中的外部AI图片检查结果依赖用户提供的 API 服务或用户导入的外部工具报告。
