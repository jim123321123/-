# PreSubmissionAIQC

中文名称：投稿前 AI 数据质量控制系统

PreSubmissionAIQC 是一个 Windows 桌面端工具，用于在投稿前对原始数据压缩包进行自动化质量筛查。用户双击 exe 打开窗口，选择待检查的 zip 压缩包后，软件会解析其中的表格、图片、PDF 和脚本文件，执行本地数值规则、附表区块审计、图片重复检查、外部 AI 工具状态整合，并生成中文报告。

软件的目标是帮助科研人员提前发现“需要人工复核的数据质量信号”。它不是研究诚信定性工具，也不会直接判定研究不端。

## 核心特点

- 图形化界面：用户通过窗口完成上传、API key 设置、运行检查和结果查看。
- 不需要用户安装 Python：可通过打包好的便携版运行，解压后双击 `PreSubmissionAIQC.exe`。
- 无 API 也可运行：本地表格规则、附表区块审计、图片重复检查、报告生成均可离线完成。
- 中文报告：面向没有计算机基础的科研人员，尽量用简单中文说明问题、依据和复核建议。
- 单一风险系统：只使用 `Red / Orange / Yellow` 颜色风险等级。
- 可视化问题定位：右侧按规则分组展示结果，点击规则后只显示有问题的 sheet，并在表格中高亮对应单元格、行或列。
- PDF 表格自动换行：报告中的长文本和宽表格会控制在页面宽度内。

## 风险颜色

| 颜色 | 含义 | 建议处理 |
|---|---|---|
| Red | 必须优先处理的问题 | 投稿前逐项核对并解决，或形成清楚的书面解释。 |
| Orange | 需要回查原始记录的问题 | 建议回到原始数据、分析脚本、仪器导出文件中确认原因。 |
| Yellow | 建议记录说明的问题 | 通常风险较低，但建议保留解释，便于审稿或内部复核。 |

默认界面只显示 `Red` 重点问题，用户可以切换查看 `Orange`、`Yellow` 或全部问题。

## 适用数据

- Excel / CSV / TSV 表格。
- 补充表、figure source data、RNA-seq DEG、代谢组、GO/KEGG 富集分析、gene list、通用数值表。
- PNG / JPG / TIFF 等图片文件。
- PDF、Python、R、ipynb 等文件会进入文件清单；表格和图片会进入对应规则检查。

## 主要检查规则

当前版本包含以下几类检查。

### 数值规则 N001-N011

- N001：两列长期保持固定差值。
- N002：两列可由简单直线关系解释。
- N003：多列小数尾部异常相同。
- N004：末位数字分布异常。
- N005：百分比和计数不匹配。
- N006：重复的连续数值序列。
- N007：数值列呈等差序列。
- N008：某列大量重复同一个数值。
- N009：均值、样本量和小数精度之间不匹配。
- N010：不同文件中出现高度相似的数据块。
- N011：同一结果区的小数位数不一致。

### 基础表格质量规则

- 完全重复的数据行。
- 高度相似的数据行。
- 完全相同的数据列。
- 两列高度相关。
- 固定倍数关系。
- 连续固定间隔变化。
- p 值或 q 值超出 0 到 1。
- p 值或 q 值为 0。
- fold change 小于或等于 0。
- 丰度、计数或强度出现负数。
- 百分比超出 0 到 100。
- 极大值或无穷大替代值。
- 高缺失率的列或行。
- 富集分析 Count 与 Genes 数量不一致。
- 富集分析条目为空或重复。

### 附表区块审计

该部分整合了 `data_audit.py` 的检查思路，主要检查补充表或 figure source data 中的连续数值区块：

- 两个数值区块完全相同。
- 两个数值区块存在固定倍数关系。
- 两个数值区块存在固定差值关系。
- 同一区块内有完全重复的行。
- 同一区块内有完全重复的列。
- 大表或区块过多时跳过部分两两审计，并在报告中提示。

### 图片规则 I001-I002

- I001：发现完全相同的图片文件。
- I002：发现视觉上高度相似的图片。

本地图片规则用于初筛，不能替代 Proofig、Imagetwin 等专业图像平台，也不能替代人工查看未裁剪原图。

更完整的规则说明见：

- `检测规则说明.md`
- `检测规则说明.docx`

## 用户使用流程

1. 解压便携版压缩包。
2. 双击 `PreSubmissionAIQC.exe`。
3. 点击上传按钮，选择待检查的 zip 压缩包。
4. 如有外部工具 API，可在界面中填写 API key；没有 API 也可以直接运行本地检查。
5. 点击开始检查。
6. 检查完成后，在右侧结果区按规则查看问题。
7. 点击某条规则后，下方会显示涉及的问题 sheet；点击 sheet 可查看表格内容和高亮位置。
8. 打开输出目录中的 PDF、HTML、Excel、CSV 或 JSON 报告进行保存和复核。

## 高亮逻辑

软件会根据规则性质决定高亮范围：

- 单元格级问题：只高亮具体有问题的单元格，例如百分比/计数不匹配、小数位异常、重复值。
- 行级问题：高亮相关行，例如完全重复行、高度相似行。
- 列关系问题：高亮相关列，例如两列高度相关、固定倍数、固定差值、简单直线关系。
- 区块问题：高亮相关数据块或区块范围。
- 图片或外部报告问题：可能无法映射到 sheet 单元格，会在文字证据中说明。

## 输出文件

每次运行会在 `runs/<运行时间_项目名>/outputs/` 下生成结果。

主要输出包括：

- `outputs/tables/raw_file_manifest.csv`：压缩包内文件清单。
- `outputs/tables/checksum_manifest.csv`：MD5 / SHA256 校验清单。
- `outputs/tables/sheet_inventory.xlsx`：表格和 sheet 解析结果。
- `outputs/tables/numeric_qc_results.xlsx`：数值规则检查结果。
- `outputs/tables/block_audit_results.xlsx`：附表区块审计结果。
- `outputs/tables/image_qc_results.xlsx`：本地图片规则检查结果。
- `outputs/tables/QC_issue_log.xlsx`：人工复核问题清单。
- `outputs/reports/findings.csv`：问题明细 CSV。
- `outputs/reports/report.json`：结构化 JSON 报告。
- `outputs/reports/final_QC_report.html`：HTML 报告。
- `outputs/reports/final_QC_report.pdf`：PDF 报告。
- `outputs/image_check/image_check_package.zip`：可手动上传到外部图片平台的图片检查包。
- `outputs/external_ai/external_ai_status.xlsx`：外部工具状态。

导出的用户报告只显示颜色风险等级。

## 无 API 时如何使用

没有外部 API key 时，软件仍会完成：

- 文件清单和校验。
- Excel / CSV / TSV 解析。
- 本地数值规则检查。
- `data_audit.py` 风格的附表区块审计。
- 本地图片重复和相似图片初筛。
- 图片检查包生成。
- 中文 PDF / HTML / Excel / CSV / JSON 报告生成。

如果需要 Proofig、Imagetwin、DataSeer 等外部平台结果，可手动上传 `image_check_package.zip`，拿到外部报告后再通过软件导入。

## API key 安全说明

- API key 通过 `keyring` 保存到系统安全凭据存储，例如 Windows Credential Manager。
- API key 不写入报告、日志、配置文件或 `runs` 输出目录。
- 如果系统 keyring 不可用，软件会提示 API key 仅本次会话使用，不会明文保存。
- Endpoint、启用状态和 LLM model name 可保存到 `config/external_ai_settings.json`。

## 开发运行

开发环境需要 Python 3.11+。

```powershell
python -m pip install -r requirements.txt
python main.py
```

## 打包 exe

生成本机可运行目录：

```powershell
python -m PyInstaller --noconfirm --clean pre_submission_ai_qc.spec
```

或使用脚本：

```powershell
build_exe.bat
```

生成便携 zip：

```powershell
build_release.bat
```

生成结果：

```text
dist/PreSubmissionAIQC/PreSubmissionAIQC.exe
release/PreSubmissionAIQC-portable.zip
```

最终用户只需要解压 `PreSubmissionAIQC-portable.zip`，然后双击 `PreSubmissionAIQC.exe`。

## Git 与数据安全

仓库默认忽略以下目录和产物：

- `build/`
- `dist/`
- `release/`
- `runs/`
- 本地日志和临时输出

请不要把原始数据 zip、运行输出、图片检查包或包含未公开数据的报告提交到 GitHub。

## 已知限制

- 自动化结果不能直接定性研究不端，只能提示需要复核的数据模式。
- 本地图片规则只做重复和相似图片初筛，不做商业级图像篡改判定。
- 外部 AI 工具需要用户自己的 API 或手动导入外部报告。
- 损坏、加密或特殊格式表格可能无法解析，会在报告中记录为解析失败。
- 规则命中不等于一定有错误，合理的单位换算、归一化、仪器精度、设计梯度、检测下限和阴性/阳性对照都可能解释部分信号。

## 免责声明

本软件用于投稿前数据质量和研究诚信风险筛查。自动化结果不能直接定性研究不端。所有 Red 和 Orange 项都需要结合原始记录、实验记录本、仪器导出文件、未裁剪原始图片和人工复核确认；Yellow 项建议记录解释。
