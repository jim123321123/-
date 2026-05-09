from __future__ import annotations

import ast
import json
import re
from typing import Any

import pandas as pd


ISSUE_TITLES = {
    "exact_duplicate_rows": "发现数字内容完全重复的数据行",
    "near_duplicate_rows": "发现数字内容高度相似的数据行",
    "duplicate_numeric_columns": "发现两列数字内容完全相同",
    "duplicate_constant_control_columns": "发现完全相同的常数对照列",
    "high_column_correlation": "两个数值列高度相关",
    "fixed_ratio_columns": "两个数值列存在固定倍数关系",
    "constant_delta_detector": "两个数值列存在固定差值关系",
    "affine_relation_detector": "两个数值列可由简单直线关系解释",
    "duplicate_decimal_detector": "两个数值列的小数尾部异常相同",
    "equal_difference_run": "连续数值呈固定间隔变化",
    "terminal_digit_anomaly": "数值末位数字分布异常",
    "percent_count_consistency_detector": "百分比和计数不匹配",
    "duplicate_series_detector": "发现重复的连续数值序列",
    "arithmetic_sequence_detector": "数值列呈等差序列",
    "repeated_value_detector": "某列大量重复同一个数值",
    "grim_like_detector": "均值、样本量和小数精度之间不匹配",
    "cross_file_reuse_detector": "不同文件中出现高度相似的数据块",
    "decimal_precision_detector": "同一结果区的小数位数不一致",
    "invalid_p_or_q_value": "p 值或 q 值超出合理范围",
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
}


MECHANISMS = {
    "exact_duplicate_rows": "可能原理：如果两行的多个数字完全一致，常见原因包括重复粘贴、重复导出、样本编号错配，或把同一组测量值误放到两个样本名下。",
    "near_duplicate_rows": "可能原理：如果两行大部分数字几乎相同，可能是复制上一行后只改了少数数值，也可能是模板填充或样本顺序混淆造成。",
    "duplicate_numeric_columns": "可能原理：两列数字逐行完全相同，可能来自复制整列、公式引用错列，或同一指标被重复导出到两个变量名下。",
    "duplicate_constant_control_columns": "可能原理：常数对照列完全相同有时是合理的归一化结果，但如果变量名代表不同实验条件，则可能是复制或导出模板造成。",
    "high_column_correlation": "可能原理：两个独立指标如果几乎同步变化，可能是合法的生物学/单位换算关系；若没有设计依据，也可能提示一列由另一列复制、缩放或公式生成。",
    "fixed_ratio_columns": "可能原理：固定倍数常见于单位换算或归一化；如果两列本应是独立测量，固定倍数可能说明其中一列是由另一列批量乘以常数得到。",
    "constant_delta_detector": "可能原理：固定差值常见于基线校正；如果没有处理说明，也可能是复制一列后统一加减常数得到新列。",
    "affine_relation_detector": "可能原理：简单直线关系说明一列几乎可由另一列乘以固定系数再加固定值算出，可能是合法换算，也可能是公式派生或批量修改。",
    "duplicate_decimal_detector": "可能原理：不同列整数部分不同但小数尾部频繁相同，可能是复制小数部分、手工改整数部分，或格式化造成的异常一致。",
    "equal_difference_run": "可能原理：连续固定间隔可能是实验设计中的梯度或时间点；若该列应为真实测量值，则可能来自拖拽填充、排序编号或人工构造序列。",
    "terminal_digit_anomaly": "可能原理：真实测量的末位数字通常不会长期集中在少数数字；过度集中可能来自人工四舍五入、手工录入习惯或仪器精度设置。",
    "percent_count_consistency_detector": "可能原理：百分比应能由分子和分母计算得到；不一致可能来自复制错单元格、四舍五入设置错误或分母引用错误。",
    "duplicate_series_detector": "可能原理：一段连续数字序列重复出现，可能是同一批数据被复制到另一处，或模板区域未被真实数据覆盖。",
    "arithmetic_sequence_detector": "可能原理：等差序列可能是设计变量；如果应为实验测量值，则可能来自自动填充、排序编号或公式生成。",
    "repeated_value_detector": "可能原理：大量重复同一数值可能是检测下限、阴性对照或缺省值；如果无合理解释，也可能是批量填充或缺失值替代。",
    "grim_like_detector": "可能原理：给定样本量和小数位时，均值只能由有限的原始整数计数产生；不匹配可能说明均值、样本量或小数位记录有误。",
    "cross_file_reuse_detector": "可能原理：不同文件出现相同数据块，可能是合理复用同一原始数据；若文件代表不同实验，则可能是重复导出或误放文件。",
    "decimal_precision_detector": "可能原理：同一结果区的小数位数通常来自同一仪器或同一导出格式；混用精度可能提示手工修改、拼接不同来源数据或格式设置不一致。",
    "invalid_p_or_q_value": "可能原理：p/q 值理论范围是 0 到 1；超出范围通常说明导出列错误、格式转换错误或把其他统计量误放入该列。",
    "zero_p_or_q_value": "可能原理：统计软件通常不会输出真正的 0，而是极小值；0 可能表示被截断、格式显示过短或导出时精度丢失。",
    "invalid_fold_change": "可能原理：fold change 通常应为正数；小于等于 0 可能是列含义不是 fold change，或计算/导出过程有误。",
    "negative_abundance_value": "可能原理：丰度、计数和强度通常不应为负；负数可能来自背景扣除、标准化或错误替换，需要确认处理方法。",
    "invalid_percentage": "可能原理：百分比通常应在 0 到 100 之间；超出范围可能是比例和百分比单位混用，或公式引用错误。",
    "extreme_or_infinite_value": "可能原理：极大值常见于除以 0、无穷大被替换为边界值，或软件计算失败后被导出为异常数字。",
    "high_column_missingness": "可能原理：整列大量缺失可能是该指标未检测、导出不完整，或合并表格时列错位。",
    "high_row_missingness": "可能原理：某些行大量缺失可能对应样本检测失败、文件合并错位，或样本信息没有填完整。",
}


def _details(value: Any) -> dict[str, Any]:
    if isinstance(value, dict):
        return value
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return {}
    if isinstance(value, str):
        text = value.strip()
        if not text:
            return {}
        for loader in (json.loads, ast.literal_eval):
            try:
                parsed = loader(text)
            except Exception:
                continue
            if isinstance(parsed, dict):
                return parsed
    return {}


def _value(row: pd.Series, key: str) -> str:
    value = row.get(key, "")
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return str(value).strip()


def _columns(row: pd.Series, details: dict[str, Any]) -> list[str]:
    columns = []
    related = _value(row, "related_columns")
    if related:
        columns.extend(part.strip() for part in re.split(r"[;,；，]", related) if part.strip())
    for key in ("column", "left_column", "right_column", "mean_column", "n_column", "numerator", "denominator", "percent"):
        value = str(details.get(key, "") or "").strip()
        if value:
            columns.append(value)
    evidence = _value(row, "evidence")
    for pattern in (
        r"Columns\s+(.+?)\s+and\s+(.+?)\s+show",
        r"Columns\s+(.+?)\s+and\s+(.+?)\s+are",
        r"(.+?)\s*~=\s*(.+?)\s*x",
    ):
        match = re.search(pattern, evidence, flags=re.I)
        if match:
            columns.extend([match.group(1).strip(), match.group(2).strip()])
    return list(dict.fromkeys(columns))


def issue_title(row: pd.Series) -> str:
    issue_type = _value(row, "issue_type")
    rule_id = _value(row, "rule_id")
    title = ISSUE_TITLES.get(issue_type, issue_type or "未分类问题")
    return f"{rule_id}：{title}" if rule_id else title


def evidence_text(row: pd.Series) -> str:
    issue_type = _value(row, "issue_type")
    evidence = _value(row, "evidence")
    details = _details(row.get("details", ""))
    cols = _columns(row, details)

    if issue_type == "exact_duplicate_rows":
        rows = _value(row, "row_index") or ", ".join(re.findall(r"\d+", evidence))
        n = re.search(r"across\s+(\d+)", evidence)
        return f"软件只比较数字列后发现，第 {rows} 行的数字内容完全相同。比较字段数：{n.group(1) if n else '见问题表'}。文字字段相同不会作为本规则的判断依据。"
    if issue_type == "near_duplicate_rows":
        match = re.search(r"Rows\s+(\d+)\s+and\s+(\d+).*?([\d.]+%).*?(\d+)\s+columns", evidence)
        if match:
            return f"第 {match.group(1)} 行和第 {match.group(2)} 行在 {match.group(4)} 个数字字段中有 {match.group(3)} 的数值高度相似。这里比较的是数字接近程度，不比较文字是否相同。"
    if issue_type in {"duplicate_numeric_columns", "duplicate_constant_control_columns"} and len(cols) >= 2:
        return f"列“{cols[0]}”和列“{cols[1]}”逐行比较后，数字内容完全一致。该规则判断的是两列整体关系，因此界面高亮相关列。"
    if issue_type == "high_column_correlation" and len(cols) >= 2:
        match = re.search(r"r\s*=\s*([\d.\-]+).*?across\s+(\d+)", evidence)
        return f"列“{cols[0]}”和列“{cols[1]}”的数字变化趋势几乎同步。相关系数为 {match.group(1) if match else '见问题表'}，参与比较的配对数值约 {match.group(2) if match else '见问题表'} 个。相关系数越接近 1 或 -1，两列越像是由同一趋势或同一公式产生。"
    if issue_type == "fixed_ratio_columns" and len(cols) >= 2:
        ratio = re.search(r"x\s*([\d.eE+\-]+)", evidence)
        return f"列“{cols[0]}”和列“{cols[1]}”之间接近固定倍数关系。固定倍数约为 {ratio.group(1) if ratio else '见问题表'}。这说明一列可能可以由另一列乘以固定数得到。"
    if issue_type == "constant_delta_detector" and len(cols) >= 2:
        delta = details.get("delta", "")
        hit = details.get("hit_rows", "")
        n = details.get("n", "")
        return f"列“{cols[0]}”和列“{cols[1]}”之间长期保持相同差值。固定差值约为 {delta}，命中 {hit}/{n} 行。"
    if issue_type == "affine_relation_detector" and len(cols) >= 2:
        return f"列“{cols[0]}”和列“{cols[1]}”几乎满足一条直线关系：后一列约等于前一列乘以 {details.get('slope', '某个系数')} 再加 {details.get('intercept', '某个常数')}。拟合程度 R²={details.get('r2', '见问题表')}。"
    if issue_type == "duplicate_decimal_detector" and len(cols) >= 2:
        return f"列“{cols[0]}”和列“{cols[1]}”的小数尾部在多行中异常一致。命中比例约为 {details.get('match_rate', '见问题表')}，示例行已在表格中标出。"
    if issue_type == "equal_difference_run":
        match = re.search(r"Rows\s+(\d+)-(\d+)\s+in column\s+(.+?)\s+show.*?difference\s+(.+?)\.", evidence)
        if match:
            return f"列“{match.group(3)}”从第 {match.group(1)} 行到第 {match.group(2)} 行，相邻两个数字之间的差值持续为 {match.group(4)}。"
    if issue_type == "terminal_digit_anomaly":
        return f"列“{details.get('column') or _value(row, 'sample_or_variable')}”的末位有效数字分布不均匀。出现最多的末位数字是 {details.get('dominant_digit', '见问题表')}，占比约 {details.get('dominant_ratio', '见问题表')}，统计检验 p 值为 {details.get('p_value', '见问题表')}。"
    if issue_type == "percent_count_consistency_detector":
        return f"软件用分子列、分母列重新计算百分比后，发现报告中的百分比与重新计算结果不一致。相关列包括：{', '.join(cols) if cols else '见问题表'}。"
    if issue_type == "arithmetic_sequence_detector":
        return f"列“{details.get('column') or _value(row, 'sample_or_variable')}”中相邻数字经常保持固定步长 {details.get('diff', '见问题表')}，命中比例约为 {details.get('match_rate', '见问题表')}。"
    if issue_type == "repeated_value_detector":
        return f"列“{details.get('column') or _value(row, 'sample_or_variable')}”中同一个数值反复出现，占比达到 {details.get('dominant_ratio', '见问题表')}。重复值为 {details.get('dominant_value', '见问题表')}。"
    if issue_type == "decimal_precision_detector":
        return f"列“{details.get('column') or _value(row, 'sample_or_variable')}”中多数数字使用 {details.get('dominant_precision', '见问题表')} 位小数，但有部分数字的小数位数不同。"
    if issue_type == "cross_file_reuse_detector":
        return f"软件在不同文件或不同 sheet 中找到了高度相似的数据窗口。涉及文件包括：{details.get('first_file', '')} 和 {details.get('second_file', '')}。"

    if evidence and re.search(r"[\u4e00-\u9fff]", evidence):
        return evidence
    if cols:
        return f"该规则主要依据这些列或对象进行判断：{', '.join(cols)}。详细数值请结合下方高亮区域和问题列表查看。"
    return "该规则根据表格中的数字模式触发。请结合高亮区域、问题列表和原始记录进行复核。"


def action_text(row: pd.Series) -> str:
    issue_type = _value(row, "issue_type")
    recommended = _value(row, "recommended_action")
    fallback = {
        "high_column_correlation": "请确认两列是否本应具有生物学相关性、单位换算关系或由同一计算流程得到；如果没有明确解释，应回查原始数据和分析脚本。",
        "duplicate_numeric_columns": "请确认两列是否是同一指标的合理重复展示；如果变量名代表不同指标，需要核对是否复制错列。",
        "near_duplicate_rows": "请核对两行对应的样本编号、实验条件和原始记录，确认是否存在样本混淆、复制粘贴或重复导出。",
        "exact_duplicate_rows": "请优先核对这些行对应的样本或实验条件，确认是否为重复录入、重复导出或样本编号错误。",
    }
    return fallback.get(issue_type) or recommended or "请结合原始记录、实验记录本、仪器导出文件和分析脚本进行人工复核。"


def mechanism_text(row: pd.Series) -> str:
    return MECHANISMS.get(_value(row, "issue_type"), "可能原理：该规则提示的是数据模式异常，不等同于结论；需要结合实验设计、仪器导出方式和分析脚本判断是否合理。")


def highlight_text(targets: list[Any], issue_type: str) -> str:
    rows = sorted({row for target in targets for row in getattr(target, "rows", set())})
    columns = sorted({col for target in targets for col in getattr(target, "columns", set())})
    cells = sorted({cell for target in targets for cell in getattr(target, "cells", set())})
    if issue_type in {"exact_duplicate_rows", "near_duplicate_rows", "high_row_missingness"} and rows:
        return "高亮范围：第 " + "、".join(str(row) for row in rows[:20]) + " 行（整行高亮，因为规则判断对象是行）。"
    if columns and issue_type in {
        "duplicate_numeric_columns",
        "duplicate_constant_control_columns",
        "high_column_correlation",
        "fixed_ratio_columns",
        "constant_delta_detector",
        "affine_relation_detector",
        "high_column_missingness",
    }:
        return "高亮范围：列 " + "、".join(columns[:12]) + "（整列高亮，因为规则判断对象是列之间的整体关系）。"
    if cells:
        text = "、".join(f"{col}{row}" for row, col in cells[:16])
        if len(cells) > 16:
            text += f" 等 {len(cells)} 个单元格"
        return f"高亮范围：{text}。"
    if columns:
        return "高亮范围：列 " + "、".join(columns[:12]) + "。"
    if rows:
        return "高亮范围：第 " + "、".join(str(row) for row in rows[:20]) + " 行。"
    return "高亮范围：该问题没有精确到具体单元格；请阅读判断依据并查看问题列表。"
