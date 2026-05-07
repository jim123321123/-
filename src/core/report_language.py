from __future__ import annotations

import re

import pandas as pd


RISK_LABELS = {
    "Red": "Red（必须处理）",
    "Orange": "Orange（需要回查）",
    "Yellow": "Yellow（建议记录）",
    "Green": "Green（未见明显问题）",
}

ISSUE_LABELS = {
    "exact_duplicate_rows": "发现完全重复的数据行",
    "near_duplicate_rows": "发现高度相似的数据行",
    "duplicate_numeric_columns": "发现内容完全相同的数据列",
    "high_column_correlation": "发现两个数据列异常高度相关",
    "fixed_ratio_columns": "发现两个数据列存在固定倍数关系",
    "equal_difference_run": "发现连续数据呈固定间隔变化",
    "terminal_digit_anomaly": "发现数值末位数字分布异常",
    "invalid_p_or_q_value": "发现 p 值或 q 值超出 0 到 1 的合理范围",
    "zero_p_or_q_value": "发现 p 值或 q 值为 0",
    "invalid_fold_change": "发现 fold change 小于或等于 0",
    "negative_abundance_value": "发现丰度、计数或强度为负数",
    "extreme_or_infinite_value": "发现极大值或无穷大替代值",
    "high_column_missingness": "发现某一列大量缺失",
    "high_row_missingness": "发现某些行大量缺失",
    "enrichment_count_gene_mismatch": "发现富集分析中的 Count 与 Genes 数量不一致",
    "empty_enrichment_term": "发现富集分析条目名称为空",
    "duplicate_enrichment_term": "发现富集分析条目重复",
    "table_parse_failed": "表格文件解析失败",
    "external_report_finding": "外部检查报告提示风险",
}

ACTION_BY_ISSUE = {
    "exact_duplicate_rows": "请优先核对这些行对应的原始记录，确认是否为复制粘贴、重复导出或样本编号错配。",
    "near_duplicate_rows": "请比较这些行对应的样本或实验条件，确认是否存在误复制、模板填充或样本混淆。",
    "fixed_ratio_columns": "请确认这两个指标是否确实存在合法单位换算或公式关系；如果没有明确解释，应回查原始数据。",
    "equal_difference_run": "请检查该列是否由人工填充、公式拖拽、排序或批量录入造成。",
    "terminal_digit_anomaly": "请核对仪器精度、四舍五入规则和人工录入记录。",
    "invalid_p_or_q_value": "请回到统计软件原始输出，确认 p/q 值导出是否错误。",
    "extreme_or_infinite_value": "请确认该极大值是否是软件把 Inf 或计算失败结果替换成边界值。",
    "high_column_missingness": "请确认该列是否本应为空；如果不是，请补充数据来源或说明缺失原因。",
    "high_row_missingness": "请确认这些行对应的样本是否缺少关键测量值。",
}


def explain_final_status(final_status: str, red_count: int, orange_count: int, yellow_count: int) -> str:
    if final_status == "Fail":
        return (
            f"结论：暂不建议直接投稿。本次检查发现 {red_count} 个必须优先处理的问题，"
            f"另有 {orange_count} 个需要回查原始记录的问题和 {yellow_count} 个建议记录说明的问题。"
        )
    if final_status == "Conditional Fail":
        return (
            f"结论：需要完成重点复核后再考虑投稿。未发现 Red 问题，但有 {orange_count} 个需要回查原始记录的问题。"
        )
    if final_status == "Conditional Pass":
        return f"结论：总体可继续推进，但建议记录 {yellow_count} 个轻微异常或解释性问题。"
    return "结论：本地规则未发现明显风险信号。仍建议保留原始记录、分析脚本和未裁剪图片以备审稿或内部复核。"


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
    if sample_or_variable:
        if sample_or_variable.lower() == "nan":
            return "见问题说明"
        return sample_or_variable
    if "row" in issue_type:
        return "见问题说明中的行号"
    if "column" in issue_type or "missingness" in issue_type:
        return "见问题说明中的列名"
    return "见问题说明"


def _problem_text(row: pd.Series) -> str:
    issue_type = _value(row, "issue_type")
    evidence = _value(row, "evidence")
    label = _issue_label(issue_type)
    translated = _translate_evidence(evidence)
    return f"{label}。证据：{translated}" if translated else label


def _join_rows(rows: str) -> str:
    parts = re.split(r"\s*(?:,|and|，|、)\s*", rows.strip())
    parts = [part for part in parts if part]
    return "、".join(parts)


def _translate_evidence(evidence: str) -> str:
    if not evidence:
        return ""
    text = evidence.strip()
    match = re.match(r"Rows (.+?) are exact duplicates across (\d+) comparable columns\.", text)
    if match:
        return f"第 {_join_rows(match.group(1))} 行在 {match.group(2)} 个可比较字段中完全相同。"
    match = re.match(r"Rows (\d+) and (\d+) share ([\d.]+%) near-identical numeric values across (\d+) columns\.", text)
    if match:
        return f"第 {match.group(1)} 行和第 {match.group(2)} 行在 {match.group(4)} 个数值字段中有 {match.group(3)} 的数值高度相似。"
    match = re.match(r"Columns (.+?) and (.+?) show Pearson r = ([\d.\-]+) across (\d+) paired values\.", text)
    if match:
        return f"列 {match.group(1)} 与列 {match.group(2)} 在 {match.group(4)} 个配对数值中的相关系数为 {match.group(3)}。"
    match = re.match(r"Columns (.+?) and (.+?) are exactly identical across (\d+) rows\.", text)
    if match:
        return f"列 {match.group(1)} 与列 {match.group(2)} 在 {match.group(3)} 行中数值完全相同。"
    match = re.match(r"Rows (\d+)-(\d+) in column (.+?) show constant adjacent difference (.+?)\.", text)
    if match:
        return f"列 {match.group(3)} 的第 {match.group(1)} 到 {match.group(2)} 行，相邻数值差值持续为 {match.group(4)}。"
    match = re.match(r"Column (.+?) missingness is ([\d.]+%)\.", text)
    if match:
        return f"列 {match.group(1)} 的缺失比例为 {match.group(2)}。"
    match = re.match(r"(\d+) rows exceed high missingness threshold\.", text)
    if match:
        return f"有 {match.group(1)} 行的缺失值比例超过阈值。"
    match = re.match(r"Column (.+?) contains extreme or infinite values such as (.+?)\.", text)
    if match:
        return f"列 {match.group(1)} 包含极大值或无穷大替代值，例如 {match.group(2)}。"
    match = re.match(r"Column (.+?) has ([\d.]+%) values ending in 0 or 5 across (\d+) values\.", text)
    if match:
        return f"列 {match.group(1)} 在 {match.group(3)} 个数值中，有 {match.group(2)} 的末位数字为 0 或 5。"
    match = re.match(r"(.+?) ~= (.+?) x (.+?), matched (\d+)/(\d+) rows, ratio_cv=(.+?)\.", text)
    if match:
        return f"列 {match.group(1)} 与列 {match.group(2)} 大致呈固定倍数 {match.group(3)}，符合 {match.group(4)}/{match.group(5)} 行。"
    match = re.match(r"Column (.+?) contains values outside \[0, 1\]\.", text)
    if match:
        return f"列 {match.group(1)} 中存在小于 0 或大于 1 的值。"
    match = re.match(r"Column (.+?) contains fold change values <= 0\.", text)
    if match:
        return f"列 {match.group(1)} 中存在小于或等于 0 的 fold change。"
    match = re.match(r"Column (.+?) contains negative abundance/count/intensity values\.", text)
    if match:
        return f"列 {match.group(1)} 中存在负数形式的丰度、计数或强度。"
    match = re.match(r"Row (\d+) Count=(\d+) but Genes contains (\d+) entries\.", text)
    if match:
        return f"第 {match.group(1)} 行 Count 为 {match.group(2)}，但 Genes 字段实际列出 {match.group(3)} 个基因。"
    return text


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
    high = issue_log[issue_log["risk_level"].isin(["Red", "Orange"])] if "risk_level" in issue_log else issue_log
    if high.empty:
        return "本次没有 Red 或 Orange 问题；建议保存完整原始数据和报告备查。"
    grouped = high.groupby(["file_name", "sheet_or_panel"], dropna=False).size().sort_values(ascending=False).head(5)
    parts = []
    for (file_name, sheet_name), count in grouped.items():
        file_text = "未指定文件" if pd.isna(file_name) or not str(file_name).strip() else str(file_name)
        sheet_text = "未指定表格/页面" if pd.isna(sheet_name) or not str(sheet_name).strip() else str(sheet_name)
        parts.append(f"{file_text} / {sheet_text}：{int(count)} 项")
    return "建议优先复核以下位置：" + "；".join(parts) + "。"
