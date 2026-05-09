from __future__ import annotations

import re

import pandas as pd


RISK_LABELS = {
    "Red": "Red（必须处理）",
    "Orange": "Orange（需要回查）",
    "Yellow": "Yellow（建议记录说明）",
    "Green": "Green（未见明显问题）",
}

ISSUE_LABELS = {
    "exact_duplicate_rows": "发现完全重复的数据行",
    "near_duplicate_rows": "发现高度相似的数据行",
    "near_duplicate_scan_skipped_large_sheet": "大表已跳过近似重复行两两扫描",
    "duplicate_numeric_columns": "发现内容完全相同的数据列",
    "duplicate_constant_control_columns": "发现完全相同的常数对照列",
    "high_column_correlation": "两个数值列高度相关",
    "fixed_ratio_columns": "两个数值列存在固定倍数关系",
    "constant_delta_detector": "相邻数值长期保持固定差值",
    "affine_relation_detector": "两列数值可用简单直线关系解释",
    "duplicate_decimal_detector": "多列数值的小数尾部重复",
    "equal_difference_run": "连续数值呈固定间隔变化",
    "terminal_digit_anomaly": "数值末位数字分布异常",
    "invalid_p_or_q_value": "p 值或 q 值超出 0 到 1 的合理范围",
    "zero_p_or_q_value": "p 值或 q 值为 0",
    "invalid_fold_change": "fold change 小于或等于 0",
    "negative_abundance_value": "丰度、计数或强度出现负数",
    "invalid_percentage": "百分比超出 0 到 100 的范围",
    "extreme_or_infinite_value": "出现极大值或无穷大替代值",
    "high_column_missingness": "某一列缺失值过多",
    "high_row_missingness": "某些行缺失值过多",
    "enrichment_count_gene_mismatch": "富集分析中的 Count 与基因列表数量不一致",
    "empty_enrichment_term": "富集分析条目名称为空",
    "duplicate_enrichment_term": "富集分析条目重复",
    "percent_count_consistency_detector": "百分比和计数不匹配",
    "duplicate_series_detector": "发现重复的连续数值序列",
    "arithmetic_sequence_detector": "数值列呈等差序列",
    "repeated_value_detector": "某列大量重复同一个值",
    "grim_like_detector": "均值、样本量和小数精度之间不匹配",
    "cross_file_reuse_detector": "不同文件中出现高度相似的数据块",
    "decimal_precision_detector": "同一结果区的小数位数不一致",
    "composite_numeric_pattern": "同一区域出现多种数值异常信号",
    "composite_percent_precision_pattern": "同一区域同时存在百分比/计数和小数精度问题",
    "exact_duplicate_image_detector": "发现完全相同的图片文件",
    "perceptual_duplicate_detector": "发现视觉上高度相似的图片",
    "table_parse_failed": "表格文件解析失败",
    "external_report_finding": "外部检查报告提示风险",
    "external_ai_manual_required": "外部 AI 工具需要手动检查",
    "external_ai_failed": "外部 AI 工具调用失败",
    "block_identical": "两个数值区块完全相同",
    "block_fixed_ratio": "两个数值区块存在固定倍数关系",
    "block_fixed_difference": "两个数值区块存在固定差值关系",
    "block_internal_duplicate_rows": "同一区块内有完全重复的行",
    "block_internal_duplicate_columns": "同一区块内有完全重复的列",
    "block_duplicate_constant_control_columns": "同一区块内有完全相同的常数对照列",
    "block_audit_skipped_large_sheet": "大表跳过了附表区块两两审计",
    "block_audit_skipped_many_blocks": "区块数量过多，跳过了附表区块两两审计",
}

ACTION_BY_ISSUE = {
    "near_duplicate_scan_skipped_large_sheet": "该表行数很大，已保留其他本地规则检查；如需近似重复行筛查，建议按实验分组拆分后单独复核。",
    "constant_delta_detector": "请核对该列是否来自排序、公式拖拽、批量填充或人工录入；如确属实验设计或仪器导出规律，请在复核记录中说明。",
    "affine_relation_detector": "请确认两列是否有合法换算、归一化或公式关系；没有明确解释时，应回查原始记录和分析脚本。",
    "duplicate_decimal_detector": "请回到原始导出文件，确认这些列的小数尾部重复是否来自四舍五入、格式设置或复制粘贴。",
    "terminal_digit_anomaly": "请核对仪器精度、四舍五入规则和人工录入记录，确认末位数字偏倚是否有合理来源。",
    "percent_count_consistency_detector": "请核对百分比、分子、分母和总数来源，优先检查报告中列出的不匹配行。",
    "duplicate_series_detector": "请确认重复序列是否代表同一批样本、同一模板或重复导出；必要时回查样本编号和处理脚本。",
    "arithmetic_sequence_detector": "请确认该列是否为编号、梯度、时间点或设计变量；若不是设计变量，请回查录入和公式。",
    "repeated_value_detector": "请确认大量重复值是否是合法的阴性/阳性对照、检测下限或缺省值；如果不是，需要核对数据来源。",
    "grim_like_detector": "请核对均值、样本量和小数位设置，确认摘要统计是否能由原始整数计数得到。",
    "cross_file_reuse_detector": "请确认不同文件中相似数据块是否是同一数据的合理复用；如不是，请检查是否重复导出或误放文件。",
    "decimal_precision_detector": "请确认同一结果区的小数位数是否应当统一；若混用不同精度，请说明来源。",
    "composite_numeric_pattern": "同一文件/表格同时触发多条规则，建议把该区域列为优先复核对象，逐项核对原始记录。",
    "composite_percent_precision_pattern": "请优先复核该区域的百分比、计数和小数位设置，确认计算过程和导出格式。",
    "exact_duplicate_image_detector": "请打开原始未裁剪图片，确认两张图片是否本应相同；如果代表不同实验条件，需要重点回查图片来源。",
    "perceptual_duplicate_detector": "请人工比对原始图片和实验条件，确认是否为重复使用、轻微裁剪、压缩或同一视野的合理复用。",
    "exact_duplicate_rows": "请优先核对这些行对应的样本或实验条件，确认是否为重复录入、重复导出或样本编号错误。",
    "near_duplicate_rows": "请比较这些行对应的样本或实验条件，确认是否存在误复制、模板填充或样本混淆。",
    "fixed_ratio_columns": "请确认两个指标是否存在合法单位换算或归一化关系；如果没有明确解释，应回查原始数据。",
    "equal_difference_run": "请检查该列是否由人工填充、公式拖拽、排序或批量录入造成。",
    "invalid_p_or_q_value": "请回到统计软件原始输出，确认 p/q 值导出是否错误。",
    "extreme_or_infinite_value": "请确认极大值是否为软件把 Inf 或计算失败结果替换成边界值。",
    "high_column_missingness": "请确认该列是否本应为空；如果不是，请补充数据来源或说明缺失原因。",
    "high_row_missingness": "请确认这些行对应的样本是否缺少关键测量值。",
    "block_identical": "请核对两个区块是否本应重复展示；如果不是，请回查复制粘贴或重复导出问题。",
    "block_fixed_ratio": "请确认两个区块是否存在合法单位换算或归一化关系，并在复核记录中说明。",
    "block_fixed_difference": "请确认固定差值是否来自合法基线校正、平移处理或公式计算。",
    "table_parse_failed": "请检查表格文件是否损坏、加密或格式不受支持。",
    "external_ai_manual_required": "如需外部图片平台筛查，请配置官方接口，或手动上传软件生成的图片检查包后导入报告。",
}


def explain_final_status(final_status: str, red_count: int, orange_count: int, yellow_count: int) -> str:
    if final_status == "Fail":
        return (
            f"结论：暂不建议直接投稿。本次检查发现 {red_count} 个必须优先处理的问题，"
            f"另有 {orange_count} 个需要回查原始记录的问题和 {yellow_count} 个建议记录说明的问题。"
        )
    if final_status == "Conditional Fail":
        return f"结论：需要完成重点复核后再考虑投稿。未发现 Red 问题，但有 {orange_count} 个需要回查原始记录的问题。"
    if final_status == "Conditional Pass":
        return f"结论：总体可以继续推进，但建议记录 {yellow_count} 个轻微信号或解释性问题。"
    return "结论：本地规则未发现明显风险信号。仍建议保留原始记录、分析脚本和未裁剪图片，以备审稿或内部复核。"


def _value(row: pd.Series, key: str) -> str:
    value = row.get(key, "")
    if pd.isna(value):
        return ""
    return str(value).strip()


def _issue_label(issue_type: str) -> str:
    return ISSUE_LABELS.get(issue_type, issue_type or "未分类问题")


def _location(row: pd.Series) -> str:
    sample_or_variable = _value(row, "sample_or_variable")
    issue_type = _value(row, "issue_type")
    if sample_or_variable and sample_or_variable.lower() != "nan":
        return sample_or_variable
    if "row" in issue_type:
        return "见问题说明中的行号"
    if "column" in issue_type or "missingness" in issue_type:
        return "见问题说明中的列名"
    return "见问题说明"


def _join_rows(rows: str) -> str:
    parts = re.split(r"\s*(?:,|and|与|、)\s*", rows.strip())
    return "、".join(part for part in parts if part)


def _translate_evidence(evidence: str) -> str:
    if not evidence:
        return ""
    text = evidence.strip()
    translations = [
        (
            r"Rows (.+?) are exact duplicates across (\d+) comparable(?: numeric)? columns\.",
            lambda m: f"第 {_join_rows(m.group(1))} 行在 {m.group(2)} 个可比较字段中完全相同。",
        ),
        (
            r"Rows (\d+) and (\d+) share ([\d.]+%) near-identical numeric values across (\d+) columns\.",
            lambda m: f"第 {m.group(1)} 行和第 {m.group(2)} 行在 {m.group(4)} 个数值字段中有 {m.group(3)} 的数值高度相似。",
        ),
        (
            r"Columns (.+?) and (.+?) show Pearson r = ([\d.\-]+) across (\d+) paired values\.",
            lambda m: f"列 {m.group(1)} 与列 {m.group(2)} 在 {m.group(4)} 个配对数值中的相关系数为 {m.group(3)}。",
        ),
        (
            r"Columns (.+?) and (.+?) are exactly identical across (\d+) rows\.",
            lambda m: f"列 {m.group(1)} 与列 {m.group(2)} 在 {m.group(3)} 行中数值完全相同。",
        ),
        (
            r"Rows (\d+)-(\d+) in column (.+?) show constant adjacent difference (.+?)\.",
            lambda m: f"列 {m.group(3)} 的第 {m.group(1)} 到 {m.group(2)} 行，相邻数值差值持续为 {m.group(4)}。",
        ),
        (
            r"Column (.+?) missingness is ([\d.]+%)\.",
            lambda m: f"列 {m.group(1)} 的缺失比例为 {m.group(2)}。",
        ),
        (
            r"(\d+) rows exceed high missingness threshold\.",
            lambda m: f"有 {m.group(1)} 行的缺失值比例超过阈值。",
        ),
        (
            r"Column (.+?) contains extreme or infinite values such as (.+?)\.",
            lambda m: f"列 {m.group(1)} 包含极大值或无穷大替代值，例如 {m.group(2)}。",
        ),
        (
            r"Column (.+?) contains values outside \[0, 1\]\.",
            lambda m: f"列 {m.group(1)} 中存在小于 0 或大于 1 的值。",
        ),
        (
            r"Column (.+?) contains fold change values <= 0\.",
            lambda m: f"列 {m.group(1)} 中存在小于或等于 0 的 fold change。",
        ),
        (
            r"Column (.+?) contains negative abundance/count/intensity values\.",
            lambda m: f"列 {m.group(1)} 中存在负数形式的丰度、计数或强度。",
        ),
        (
            r"Row (\d+) Count=(\d+) but Genes contains (\d+) entries\.",
            lambda m: f"第 {m.group(1)} 行 Count 为 {m.group(2)}，但 Genes 字段实际列出 {m.group(3)} 个基因。",
        ),
    ]
    for pattern, builder in translations:
        match = re.match(pattern, text)
        if match:
            return builder(match)
    return text


def _problem_text(row: pd.Series) -> str:
    issue_type = _value(row, "issue_type")
    evidence = _translate_evidence(_value(row, "evidence"))
    rule_id = _value(row, "rule_id")
    prefix = f"{rule_id} " if rule_id else ""
    label = f"{prefix}{_issue_label(issue_type)}"
    return f"{label}。证据：{evidence}" if evidence else label


def _action_text(row: pd.Series) -> str:
    issue_type = _value(row, "issue_type")
    recommended = _value(row, "recommended_action")
    return ACTION_BY_ISSUE.get(issue_type) or recommended or "请结合原始记录、实验记录本和数据处理脚本进行人工复核。"


def build_plain_issue_table(issue_log: pd.DataFrame, limit: int | None = None) -> pd.DataFrame:
    columns = ["风险等级", "文件", "表格/页面", "具体位置", "发现的问题", "建议怎么做"]
    if issue_log is None or issue_log.empty:
        return pd.DataFrame(columns=columns)
    source = issue_log.copy()
    if limit is not None:
        source = source.head(limit)
    rows = []
    for _, row in source.iterrows():
        risk = _value(row, "risk_level")
        rows.append(
            {
                "风险等级": RISK_LABELS.get(risk, risk or "未分级"),
                "文件": _value(row, "file_name") or "未指定文件",
                "表格/页面": _value(row, "sheet_or_panel") or "未指定",
                "具体位置": _location(row),
                "发现的问题": _problem_text(row),
                "建议怎么做": _action_text(row),
            }
        )
    return pd.DataFrame(rows, columns=columns)


def build_priority_review_text(issue_log: pd.DataFrame) -> str:
    if issue_log is None or issue_log.empty:
        return "本次未形成需要人工复核的问题清单。"
    priority = issue_log[issue_log["risk_level"].isin(["Red", "Orange"])] if "risk_level" in issue_log else issue_log
    if priority.empty:
        return "本次没有 Red 或 Orange 问题；建议保存完整原始数据和报告备查。"
    grouped = priority.groupby(["file_name", "sheet_or_panel"], dropna=False).size().sort_values(ascending=False).head(5)
    parts = []
    for (file_name, sheet_name), count in grouped.items():
        file_text = "未指定文件" if pd.isna(file_name) or not str(file_name).strip() else str(file_name)
        sheet_text = "未指定表格/页面" if pd.isna(sheet_name) or not str(sheet_name).strip() else str(sheet_name)
        parts.append(f"{file_text} / {sheet_text}：{int(count)} 项")
    return "建议优先复核以下位置：" + "；".join(parts) + "。"
