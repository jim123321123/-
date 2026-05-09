from __future__ import annotations

import math
import re
from itertools import combinations
from typing import Any

import numpy as np
import pandas as pd
from scipy.stats import chisquare

from .table_profiler import normalize_column_name


DEFAULT_THRESHOLDS = {
    "duplicate": {
        "near_duplicate_similarity_threshold": 0.95,
        "red_duplicate_similarity_threshold": 0.99,
        "red_column_correlation_threshold": 0.999,
        "max_near_duplicate_rows": 2000,
    },
    "ratio_pattern": {"min_matched_rows": 8, "ratio_cv_threshold": 0.01},
    "equal_difference": {
        "min_run_length_red": 8,
        "min_run_length_orange": 6,
        "absolute_tolerance": 1.0e-8,
        "relative_tolerance": 0.001,
    },
    "digit_pattern": {
        "max_zero_or_five_fraction_orange": 0.45,
        "max_zero_or_five_fraction_red": 0.60,
        "min_n_for_digit_test": 20,
    },
    "correlation": {"high_corr_threshold": 0.995, "red_corr_threshold": 0.999, "min_non_missing_pairs": 10},
    "missingness": {
        "high_column_missing_fraction_yellow": 0.5,
        "high_column_missing_fraction_orange": 0.8,
        "high_row_missing_fraction_yellow": 0.5,
        "high_row_missing_fraction_orange": 0.8,
    },
    "extreme_values": {"float_max_warning_threshold": 1.0e300},
    "outlier": {"robust_z_threshold": 3.5},
    "numeric": {
        "constant_delta_min_n": 6,
        "constant_delta_match_rate_high": 0.90,
        "constant_delta_match_rate_medium": 0.75,
        "affine_min_n": 8,
        "affine_r2_high": 0.999999,
        "affine_r2_medium": 0.999,
        "duplicate_decimal_min_n": 10,
        "decimal_match_rate_high": 0.80,
        "decimal_match_rate_medium": 0.60,
        "terminal_digit_min_n": 30,
        "terminal_digit_p_medium": 0.001,
        "terminal_digit_p_high": 0.000001,
        "terminal_digit_dominant_ratio_medium": 0.35,
        "terminal_digit_dominant_ratio_high": 0.50,
        "arithmetic_sequence_min_n": 10,
        "arithmetic_sequence_match_rate": 0.85,
        "repeated_value_min_n": 20,
        "repeated_value_unique_ratio": 0.20,
        "repeated_value_dominant_ratio": 0.50,
        "cross_file_window_rows": 5,
        "cross_file_window_cols": 2,
    },
}


def merged_thresholds(thresholds: dict[str, Any] | None) -> dict[str, Any]:
    result = {key: value.copy() for key, value in DEFAULT_THRESHOLDS.items()}
    for section, values in (thresholds or {}).items():
        if isinstance(values, dict):
            result.setdefault(section, {}).update(values)
    return result


def _issue(
    issues: list[dict[str, Any]],
    issue_type: str,
    risk_level: str,
    file_name: str,
    sheet_name: str,
    evidence: str,
    column_name: str = "",
    row_index: str = "",
    related_columns: str = "",
    action: str = "回查原始记录、仪器导出文件和数据处理脚本，确认该风险信号是否有合理解释。",
    rule_id: str = "",
    severity: str = "",
    details: dict[str, Any] | None = None,
) -> None:
    severity = severity or {"Red": "HIGH", "Orange": "MEDIUM", "Yellow": "LOW"}.get(risk_level, risk_level)
    issues.append(
        {
            "issue_id": f"NUM{len(issues) + 1:03d}",
            "module": "Numeric Forensics",
            "rule_id": rule_id,
            "severity": severity,
            "risk_level": risk_level,
            "issue_type": issue_type,
            "file_name": file_name,
            "sheet_name": sheet_name,
            "row_index": row_index,
            "column_name": column_name,
            "related_columns": related_columns,
            "evidence": evidence,
            "recommended_action": action,
            "details": details or {},
            "need_human_review": "Yes" if risk_level in {"Red", "Orange"} else "Recommended",
            "affects_submission": "Yes" if risk_level in {"Red", "Orange"} else "Review",
        }
    )


def _risk_from_severity(severity: str) -> str:
    return {"CRITICAL": "Red", "HIGH": "Red", "MEDIUM": "Orange", "LOW": "Yellow"}.get(severity, "Yellow")


def _downgrade_severity(severity: str) -> str:
    order = ["LOW", "MEDIUM", "HIGH", "CRITICAL"]
    if severity not in order:
        return severity
    return order[max(order.index(severity) - 1, 0)]


def _numeric_columns(df: pd.DataFrame) -> list[str]:
    columns = []
    for column in df.columns:
        if _is_identifier_or_layout_column(column):
            continue
        numeric = pd.to_numeric(df[column], errors="coerce")
        if numeric.notna().sum() >= 2:
            columns.append(str(column))
    return columns


def _is_identifier_or_layout_column(column: Any) -> bool:
    norm = normalize_column_name(column)
    if re.fullmatch(r"column\d+", norm):
        return False
    tokens = (
        "id",
        "identifier",
        "pubchem",
        "cas",
        "kegg",
        "hmdb",
        "compound",
        "compid",
        "chemicalid",
        "sort",
        "order",
        "rank",
        "index",
    )
    return any(token in norm for token in tokens)


def _is_coordinate_or_enrichment_sheet(sheet_name: str, df: pd.DataFrame) -> bool:
    name = normalize_column_name(sheet_name)
    if any(token in name for token in ("location", "peak", "lad", "coordinate", "genomic", "bed")):
        return True
    header_text = " ".join(str(col).lower() for col in df.columns)
    sample_text = " ".join(str(value).lower() for value in df.head(5).to_numpy().ravel() if not pd.isna(value))
    combined = f"{name} {header_text} {sample_text}"
    return any(token in combined for token in ("go:", "fold enrichment", "benjamini", "bonferroni", "pop total", "pop hits"))


def _is_figure_sheet(sheet_name: str) -> bool:
    name = normalize_column_name(sheet_name)
    return "figure" in name or name.startswith("fig")


def _is_derived_column(column: Any) -> bool:
    norm = normalize_column_name(column)
    tokens = (
        "normalized",
        "ratio",
        "percent",
        "percentage",
        "rate",
        "foldchange",
        "fold",
        "log2fc",
        "mean",
        "average",
        "sum",
        "calculated",
        "standardized",
        "标准化",
        "比例",
        "均值",
        "总和",
    )
    return any(token in norm for token in tokens)


def _is_category_column(column: Any) -> bool:
    norm = normalize_column_name(column)
    return any(token in norm for token in ("group", "condition", "category", "sex", "status", "type", "class", "label"))


def _is_design_sequence_column(column: Any) -> bool:
    norm = normalize_column_name(column)
    return any(token in norm for token in ("time", "dose", "concentration", "standardcurve", "gradient", "浓度", "时间", "梯度"))


def _looks_like_bounded_percent_or_heatmap(sheet_name: str, column: Any, values: pd.Series) -> bool:
    if values.empty:
        return False
    norm_col = normalize_column_name(column)
    norm_sheet = normalize_column_name(sheet_name)
    finite = pd.to_numeric(values, errors="coerce").replace([np.inf, -np.inf], np.nan).dropna()
    if finite.empty:
        return False
    if finite.between(0, 100).mean() < 0.98:
        return False
    rounded = finite.round(9)
    counts = rounded.value_counts()
    dominant_value = float(counts.index[0])
    dominant_ratio = int(counts.iloc[0]) / len(rounded)
    column_hint = any(token in norm_col for token in ("percent", "percentage", "rate", "score", "scaled", "normalized"))
    sheet_hint = any(token in norm_sheet for token in ("heatmap", "heat", "pathway", "score", "percent", "percentage"))
    boundary_saturation = dominant_value in {0.0, 1.0, 100.0} and dominant_ratio >= 0.35
    return (column_hint or sheet_hint) and boundary_saturation


def _simple_number(value: float) -> bool:
    if not np.isfinite(value):
        return False
    simple = {0.1, 0.2, 0.25, 0.5, 1.0, 2.0, 5.0, 10.0, 100.0, -1.0, -0.1}
    return any(math.isclose(value, item, rel_tol=1e-8, abs_tol=1e-8) for item in simple)


def _decimal_tail(value: Any, digits: int = 2) -> str | None:
    if pd.isna(value):
        return None
    text = str(value).strip()
    if "e" in text.lower():
        text = f"{float(value):.{digits + 4}f}"
    match = re.search(r"[-+]?\d+\.([0-9]+)", text)
    if not match:
        return None
    decimals = match.group(1)
    if len(decimals) < digits:
        decimals = decimals.ljust(digits, "0")
    return decimals[:digits]


def _last_significant_digit(value: Any) -> int | None:
    if pd.isna(value):
        return None
    text = str(value).strip()
    if not text:
        return None
    try:
        numeric = float(text)
    except (TypeError, ValueError):
        return None
    if not np.isfinite(numeric):
        return None
    if "e" in text.lower():
        text = f"{numeric:.12f}"
    text = text.rstrip("0").rstrip(".")
    digits = re.sub(r"\D", "", text)
    if not digits:
        return None
    return int(digits[-1])


def _is_constant_control(values: pd.Series) -> bool:
    clean = pd.to_numeric(values, errors="coerce").replace([np.inf, -np.inf], np.nan).dropna()
    if len(clean) < 3:
        return False
    unique = np.unique(np.round(clean.to_numpy(dtype=float), 9))
    if len(unique) != 1:
        return False
    return any(math.isclose(float(unique[0]), expected, abs_tol=1e-9) for expected in (0.0, 1.0, 100.0))


def _numeric_frame(df: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    return pd.DataFrame({column: pd.to_numeric(df[column], errors="coerce") for column in columns})


def _check_exact_duplicates(issues, file_name, sheet_name, df):
    cols = _numeric_columns(df)
    if len(cols) < 2:
        return
    comparable = _numeric_frame(df, cols).dropna(axis=1, how="all")
    if comparable.empty:
        return
    minimum_non_missing = max(2, int(comparable.shape[1] * 0.3))
    comparable = comparable[comparable.notna().sum(axis=1) >= minimum_non_missing]
    if comparable.empty:
        return
    duplicated = comparable.duplicated(keep=False)
    if duplicated.any():
        rows = [str(i + 2) for i in comparable.index[duplicated].tolist()]
        risk = "Red" if comparable.shape[1] >= 4 else "Yellow"
        _issue(
            issues,
            "exact_duplicate_rows",
            risk,
            file_name,
            sheet_name,
            f"Rows {', '.join(rows)} are exact duplicates across {comparable.shape[1]} comparable numeric columns.",
            row_index=", ".join(rows),
            action="回查原始仪器导出文件、实验记录本和数据录入记录，确认是否存在复制粘贴或样本ID错配。",
        )


def _check_near_duplicate_rows(issues, file_name, sheet_name, df, thresholds):
    if _is_coordinate_or_enrichment_sheet(sheet_name, df):
        return
    cols = _numeric_columns(df)
    if len(cols) < 5:
        return
    max_rows = thresholds["duplicate"].get("max_near_duplicate_rows", 2000)
    if len(df) > max_rows:
        _issue(
            issues,
            "near_duplicate_scan_skipped_large_sheet",
            "Yellow",
            file_name,
            sheet_name,
            f"Sheet has {len(df)} rows; pairwise near-duplicate row scan was skipped because it would require more than {max_rows}x{max_rows} comparisons.",
            action="该表行数很大，软件已保留完整重复、列关系、范围和缺失等检查；如需近似重复行筛查，建议按样本分组或抽样后单独复核。",
        )
        return
    ndf = _numeric_frame(df, cols)
    orange = thresholds["duplicate"]["near_duplicate_similarity_threshold"]
    red = thresholds["duplicate"]["red_duplicate_similarity_threshold"]
    for left, right in combinations(range(len(ndf)), 2):
        a = ndf.iloc[left]
        b = ndf.iloc[right]
        mask = a.notna() & b.notna()
        if mask.sum() < 5:
            continue
        denom = np.maximum(np.maximum(np.abs(a[mask].to_numpy()), np.abs(b[mask].to_numpy())), 1.0)
        close = np.isclose(a[mask].to_numpy(), b[mask].to_numpy(), rtol=1e-3, atol=1e-8) | (
            np.abs(a[mask].to_numpy() - b[mask].to_numpy()) / denom < 0.01
        )
        similarity = float(close.mean())
        if similarity >= orange:
            _issue(
                issues,
                "near_duplicate_rows",
                "Red" if similarity >= red else "Orange",
                file_name,
                sheet_name,
                f"Rows {left + 2} and {right + 2} share {similarity:.1%} near-identical numeric values across {int(mask.sum())} columns.",
                row_index=f"{left + 2}, {right + 2}",
            )
            return


def _check_duplicate_and_related_columns(issues, file_name, sheet_name, df, thresholds):
    cols = _numeric_columns(df)
    if len(cols) < 2:
        return
    skip_structural_relationships = _is_coordinate_or_enrichment_sheet(sheet_name, df)
    ndf = _numeric_frame(df, cols)
    min_pairs = thresholds["correlation"]["min_non_missing_pairs"]
    for a, b in combinations(cols, 2):
        pair = ndf[[a, b]].replace([np.inf, -np.inf], np.nan).dropna()
        pair = pair[(pair[a].abs() <= thresholds["extreme_values"]["float_max_warning_threshold"]) & (pair[b].abs() <= thresholds["extreme_values"]["float_max_warning_threshold"])]
        if len(pair) < 2:
            continue
        if pair[a].equals(pair[b]):
            if _is_constant_control(pair[a]) and _is_constant_control(pair[b]):
                _issue(
                    issues,
                    "duplicate_constant_control_columns",
                    "Yellow",
                    file_name,
                    sheet_name,
                    f"Columns {a} and {b} are identical constant control-style values across {len(pair)} rows.",
                    related_columns=f"{a}; {b}",
                    action="该信号常见于归一化对照列（例如全为 1）或填充值；请确认是否为实验设计导致，通常不应单独作为严重错误。",
                )
                continue
            risk = "Red" if len(pair) >= 8 else "Orange" if len(pair) >= 4 else "Yellow"
            _issue(
                issues,
                "duplicate_numeric_columns",
                risk,
                file_name,
                sheet_name,
                f"Columns {a} and {b} are exactly identical across {len(pair)} rows.",
                related_columns=f"{a}; {b}",
            )
            continue
        if not skip_structural_relationships and len(pair) >= min_pairs and pair[a].std() > 0 and pair[b].std() > 0:
            corr = float(pair[a].corr(pair[b]))
            if abs(corr) > thresholds["correlation"]["high_corr_threshold"]:
                risk = "Red" if abs(corr) > thresholds["correlation"]["red_corr_threshold"] else "Orange"
                _issue(
                    issues,
                    "high_column_correlation",
                    risk,
                    file_name,
                    sheet_name,
                    f"Columns {a} and {b} show Pearson r = {corr:.4f} across {len(pair)} paired values.",
                    related_columns=f"{a}; {b}",
                )
        if not skip_structural_relationships:
            ratio_pair = pair[(pair[a] != 0) & np.isfinite(pair[a]) & np.isfinite(pair[b])]
            if len(ratio_pair) >= thresholds["ratio_pattern"]["min_matched_rows"]:
                ratios = ratio_pair[b] / ratio_pair[a]
                if ratios.mean() != 0:
                    cv = float(ratios.std(ddof=0) / abs(ratios.mean()))
                    if cv < thresholds["ratio_pattern"]["ratio_cv_threshold"]:
                        _issue(
                            issues,
                            "fixed_ratio_columns",
                            "Red" if len(ratio_pair) == len(pair) else "Orange",
                            file_name,
                            sheet_name,
                            f"{b} ~= {a} x {ratios.mean():.4g}, matched {len(ratio_pair)}/{len(pair)} rows, ratio_cv={cv:.4g}.",
                            related_columns=f"{a}; {b}",
                        )


def _check_constant_delta(issues, file_name, sheet_name, df, thresholds):
    cfg = thresholds["numeric"]
    cols = _numeric_columns(df)
    if len(cols) < 2 or _is_coordinate_or_enrichment_sheet(sheet_name, df):
        return
    ndf = _numeric_frame(df, cols)
    for a, b in combinations(cols, 2):
        pair = ndf[[a, b]].replace([np.inf, -np.inf], np.nan).dropna()
        if len(pair) < cfg["constant_delta_min_n"]:
            continue
        delta = (pair[b] - pair[a]).round(6)
        if delta.abs().max() <= 1e-12:
            continue
        counts = delta.value_counts()
        fixed_delta = float(counts.index[0])
        hit_rows = int(counts.iloc[0])
        match_rate = hit_rows / len(pair)
        if float(delta.std(ddof=0)) < 1e-9:
            severity = "HIGH"
        elif match_rate >= cfg["constant_delta_match_rate_high"]:
            severity = "HIGH"
        elif match_rate >= cfg["constant_delta_match_rate_medium"]:
            severity = "MEDIUM"
        else:
            continue
        _issue(
            issues,
            "constant_delta_detector",
            _risk_from_severity(severity),
            file_name,
            sheet_name,
            f"列 {b} 与列 {a} 在 {hit_rows}/{len(pair)} 行中存在相同差值 {fixed_delta:+.6f}，疑似复制后统一加减常数。",
            related_columns=f"{a}; {b}",
            action="建议核对原始记录、仪器导出文件和分析脚本，确认该固定差值是否来自合法校正或单位换算。",
            rule_id="N001",
            severity=severity,
            details={"left_column": a, "right_column": b, "n": len(pair), "hit_rows": hit_rows, "delta": fixed_delta, "match_rate": match_rate},
        )


def _check_affine_relation(issues, file_name, sheet_name, df, thresholds):
    cfg = thresholds["numeric"]
    cols = _numeric_columns(df)
    if len(cols) < 2 or _is_coordinate_or_enrichment_sheet(sheet_name, df):
        return
    ndf = _numeric_frame(df, cols)
    for a, b in combinations(cols, 2):
        pair = ndf[[a, b]].replace([np.inf, -np.inf], np.nan).dropna()
        if len(pair) < cfg["affine_min_n"] or pair[a].std() == 0 or pair[b].std() == 0:
            continue
        x = pair[a].to_numpy(dtype=float)
        y = pair[b].to_numpy(dtype=float)
        slope, intercept = np.polyfit(x, y, 1)
        predicted = slope * x + intercept
        residual = y - predicted
        ss_res = float(np.sum(residual**2))
        ss_tot = float(np.sum((y - y.mean()) ** 2))
        if ss_tot <= 0:
            continue
        r2 = 1 - ss_res / ss_tot
        max_abs_residual = float(np.max(np.abs(residual)))
        simple_relation = _simple_number(float(slope)) or _simple_number(float(intercept))
        if r2 >= cfg["affine_r2_high"] and max_abs_residual < 1e-8:
            severity = "HIGH"
        elif r2 >= cfg["affine_r2_medium"] and simple_relation:
            severity = "HIGH" if _simple_number(float(slope)) and _simple_number(float(intercept)) else "MEDIUM"
        else:
            continue
        note = ""
        if _is_derived_column(a) or _is_derived_column(b):
            severity = _downgrade_severity(severity)
            note = " 其中一列可能是合法派生列，需结合分析脚本复核。"
        _issue(
            issues,
            "affine_relation_detector",
            _risk_from_severity(severity),
            file_name,
            sheet_name,
            f"列 {b} 与列 {a} 近似满足 B = {slope:.6g} * A + {intercept:.6g}，R²={r2:.8f}，最大残差={max_abs_residual:.3g}。{note}",
            related_columns=f"{a}; {b}",
            action="建议核对两列是否为合法派生、标准化、单位换算或公式计算；若无明确解释，请回查原始记录。",
            rule_id="N002",
            severity=severity,
            details={"left_column": a, "right_column": b, "n": len(pair), "slope": float(slope), "intercept": float(intercept), "r2": float(r2), "max_abs_residual": max_abs_residual},
        )


def _check_duplicate_decimal(issues, file_name, sheet_name, df, thresholds):
    cfg = thresholds["numeric"]
    cols = _numeric_columns(df)
    if len(cols) < 2:
        return
    for a, b in combinations(cols, 2):
        left = df[a]
        right = df[b]
        rows = []
        matches = 0
        large_integer_diffs = 0
        terminal_05 = 0
        for idx in df.index:
            av = pd.to_numeric(left.loc[idx], errors="coerce")
            bv = pd.to_numeric(right.loc[idx], errors="coerce")
            if pd.isna(av) or pd.isna(bv):
                continue
            atail = _decimal_tail(left.loc[idx], 2)
            btail = _decimal_tail(right.loc[idx], 2)
            if atail is None or btail is None:
                continue
            rows.append((idx, av, bv, atail, btail))
            if atail == btail:
                matches += 1
                if abs(int(float(av)) - int(float(bv))) >= 5:
                    large_integer_diffs += 1
            if atail[-1:] in {"0", "5"} and btail[-1:] in {"0", "5"}:
                terminal_05 += 1
        if len(rows) < cfg["duplicate_decimal_min_n"]:
            continue
        if terminal_05 / len(rows) > 0.7:
            continue
        match_rate = matches / len(rows)
        if match_rate >= cfg["decimal_match_rate_high"]:
            severity = "HIGH"
        elif match_rate >= cfg["decimal_match_rate_medium"]:
            severity = "MEDIUM"
        else:
            continue
        if large_integer_diffs / max(matches, 1) > 0.5 and severity == "MEDIUM":
            severity = "HIGH"
        examples = [f"row {idx + 2}: {av} vs {bv}" for idx, av, bv, _, _ in rows[:3]]
        _issue(
            issues,
            "duplicate_decimal_detector",
            _risk_from_severity(severity),
            file_name,
            sheet_name,
            f"列 {a} 与列 {b} 的小数后两位在 {matches}/{len(rows)} 行中相同，比例 {match_rate:.1%}。示例：{'; '.join(examples)}。",
            related_columns=f"{a}; {b}",
            action="建议核对是否存在复制小数部分、手工改整数部分或格式化造成的异常一致。",
            rule_id="N003",
            severity=severity,
            details={"left_column": a, "right_column": b, "n": len(rows), "matched_rows": matches, "match_rate": match_rate, "examples": examples},
        )


def _longest_equal_diff_run(values: pd.Series, atol: float, rtol: float) -> tuple[int, int, float]:
    clean = pd.to_numeric(values, errors="coerce").dropna().to_numpy(dtype=float)
    if len(clean) < 5:
        return 0, 0, 0.0
    diffs = np.diff(clean)
    best_len = current_len = 1
    best_start = current_start = 0
    for i in range(1, len(diffs)):
        if math.isclose(diffs[i], diffs[i - 1], abs_tol=atol, rel_tol=rtol):
            current_len += 1
        else:
            if current_len > best_len:
                best_len, best_start = current_len, current_start
            current_len = 1
            current_start = i
    if current_len > best_len:
        best_len, best_start = current_len, current_start
    return best_len + 1, best_start, diffs[best_start]


def _check_equal_difference(issues, file_name, sheet_name, df, thresholds):
    if _is_coordinate_or_enrichment_sheet(sheet_name, df):
        return
    cfg = thresholds["equal_difference"]
    for col in _numeric_columns(df):
        run_len, start, diff = _longest_equal_diff_run(df[col], cfg["absolute_tolerance"], cfg["relative_tolerance"])
        if run_len >= cfg["min_run_length_orange"]:
            if math.isclose(diff, 0.0, abs_tol=cfg["absolute_tolerance"]):
                continue
            _issue(
                issues,
                "equal_difference_run",
                "Red" if run_len >= cfg["min_run_length_red"] else "Orange",
                file_name,
                sheet_name,
                f"Rows {start + 2}-{start + run_len + 1} in column {col} show constant adjacent difference {diff:.6g}.",
                column_name=col,
            )


def _terminal_digit(value: float) -> int | None:
    if not np.isfinite(value):
        return None
    text = f"{abs(value):.12g}".rstrip("0").rstrip(".")
    digits = re.sub(r"\D", "", text)
    if not digits:
        return None
    return int(digits[-1])


def _check_terminal_digits(issues, file_name, sheet_name, df, thresholds):
    if _is_coordinate_or_enrichment_sheet(sheet_name, df):
        return
    cfg = thresholds["numeric"]
    for col in _numeric_columns(df):
        values = df[col].dropna()
        if _looks_like_bounded_percent_or_heatmap(sheet_name, col, values):
            continue
        digits = [_last_significant_digit(v) for v in values]
        digits = [d for d in digits if d is not None]
        if len(digits) < cfg["terminal_digit_min_n"]:
            continue
        counts = {digit: digits.count(digit) for digit in range(10)}
        observed = np.array([counts[digit] for digit in range(10)], dtype=float)
        p_value = float(chisquare(observed, np.ones(10) * (len(digits) / 10)).pvalue)
        dominant_digit = max(counts, key=counts.get)
        dominant_ratio = counts[dominant_digit] / len(digits)
        if p_value < cfg["terminal_digit_p_high"] and dominant_ratio > cfg["terminal_digit_dominant_ratio_high"]:
            severity = "HIGH"
        elif p_value < cfg["terminal_digit_p_medium"] and dominant_ratio > cfg["terminal_digit_dominant_ratio_medium"]:
            severity = "MEDIUM"
        else:
            continue
        if any(token in normalize_column_name(col) for token in ("score", "grade", "level", "rank", "rounded", "评分", "等级")):
            severity = _downgrade_severity(severity)
        _issue(
            issues,
            "terminal_digit_anomaly",
            _risk_from_severity(severity),
            file_name,
            sheet_name,
            f"列 {col} 的末位有效数字分布不均匀，主导数字 {dominant_digit} 占 {dominant_ratio:.1%}，chisquare p={p_value:.2g}。",
            column_name=col,
            action="建议核对仪器精度、四舍五入规则、人工录入习惯和原始记录；该结果仅提示末位数字模式异常。",
            rule_id="N004",
            severity=severity,
            details={"digit_counts": counts, "p_value": p_value, "dominant_digit": dominant_digit, "dominant_ratio": dominant_ratio, "n": len(digits)},
        )


def _check_ranges(issues, file_name, sheet_name, df, thresholds):
    extreme = thresholds["extreme_values"]["float_max_warning_threshold"]
    for column in df.columns:
        norm = normalize_column_name(column)
        if _is_identifier_or_layout_column(column):
            continue
        values = pd.to_numeric(df[column], errors="coerce")
        if values.notna().sum() == 0:
            continue
        bad_extreme = values[np.isinf(values) | (values.abs() > extreme)]
        if len(bad_extreme):
            _issue(
                issues,
                "extreme_or_infinite_value",
                "Orange",
                file_name,
                sheet_name,
                f"Column {column} contains extreme or infinite values such as {bad_extreme.iloc[0]}.",
                column_name=str(column),
            )
        if norm in {"p", "pvalue", "qvalue", "padj", "fdr"}:
            invalid = values[(values < 0) | (values > 1)]
            if len(invalid):
                _issue(
                    issues,
                    "invalid_p_or_q_value",
                    "Red",
                    file_name,
                    sheet_name,
                    f"Column {column} contains values outside [0, 1].",
                    column_name=str(column),
                )
            zeros = values[values == 0]
            if len(zeros):
                _issue(
                    issues,
                    "zero_p_or_q_value",
                    "Yellow" if len(zeros) < max(3, len(values) * 0.1) else "Orange",
                    file_name,
                    sheet_name,
                    f"Column {column} contains {len(zeros)} zero p/q values.",
                    column_name=str(column),
                    action="建议报告为 < threshold 或回查原始统计软件输出。",
                )
        if norm in {"foldchange", "fc"}:
            invalid = values[values <= 0]
            if len(invalid):
                _issue(
                    issues,
                    "invalid_fold_change",
                    "Red",
                    file_name,
                    sheet_name,
                    f"Column {column} contains fold change values <= 0.",
                    column_name=str(column),
                )
        if any(token in norm for token in ("fpkm", "counts", "intensity", "abundance", "concentration")):
            invalid = values[values < 0]
            if len(invalid):
                _issue(
                    issues,
                    "negative_abundance_value",
                    "Red",
                    file_name,
                    sheet_name,
                    f"Column {column} contains negative abundance/count/intensity values.",
                    column_name=str(column),
                )
        if "percent" in norm:
            invalid = values[(values < 0) | (values > 100)]
            if len(invalid):
                _issue(issues, "invalid_percentage", "Orange", file_name, sheet_name, f"Column {column} has values outside 0-100.", column_name=str(column))


def _check_missingness(issues, file_name, sheet_name, df, thresholds):
    cfg = thresholds["missingness"]
    if df.empty:
        return
    for column, frac in df.isna().mean(axis=0).items():
        if frac > cfg["high_column_missing_fraction_yellow"]:
            non_missing = int(df[column].notna().sum())
            numeric_non_missing = int(pd.to_numeric(df[column], errors="coerce").notna().sum())
            minimum_content = max(3, min(10, int(len(df) * 0.05)))
            if non_missing < minimum_content and numeric_non_missing < 3:
                continue
            risk = "Orange" if frac > cfg["high_column_missing_fraction_orange"] else "Yellow"
            if _is_figure_sheet(sheet_name):
                risk = "Yellow"
            _issue(
                issues,
                "high_column_missingness",
                risk,
                file_name,
                sheet_name,
                f"Column {column} missingness is {frac:.1%}.",
                column_name=str(column),
            )
    row_fracs = df.isna().mean(axis=1)
    high_rows = row_fracs[row_fracs > cfg["high_row_missing_fraction_yellow"]]
    if df.shape[1] >= 6 and len(high_rows) / max(len(df), 1) > 0.25:
        return
    if len(high_rows):
        risk = "Orange" if high_rows.max() > cfg["high_row_missing_fraction_orange"] and len(high_rows) >= 3 else "Yellow"
        _issue(
            issues,
            "high_row_missingness",
            risk,
            file_name,
            sheet_name,
            f"{len(high_rows)} rows exceed high missingness threshold.",
        )


def _split_genes(value: Any) -> list[str]:
    if pd.isna(value):
        return []
    return [part for part in re.split(r"[,;|/\s]+", str(value).strip()) if part]


def _check_enrichment(issues, file_name, sheet_name, df):
    columns = {normalize_column_name(col): col for col in df.columns}
    count_col = columns.get("count")
    genes_col = columns.get("genes") or columns.get("geneid")
    term_col = columns.get("term") or columns.get("pathway")
    if term_col:
        missing = df[df[term_col].isna() | (df[term_col].astype(str).str.strip() == "")]
        if len(missing):
            _issue(issues, "empty_enrichment_term", "Orange", file_name, sheet_name, f"{len(missing)} enrichment rows have empty Term/pathway.")
        duplicated = df[term_col].duplicated(keep=False)
        if duplicated.any():
            _issue(issues, "duplicate_enrichment_term", "Yellow", file_name, sheet_name, f"{int(duplicated.sum())} rows contain duplicate Term/pathway values.")
    if count_col and genes_col:
        counts = pd.to_numeric(df[count_col], errors="coerce")
        for idx, count in counts.dropna().items():
            genes = _split_genes(df.at[idx, genes_col])
            if int(count) != len(genes):
                _issue(
                    issues,
                    "enrichment_count_gene_mismatch",
                    "Orange",
                    file_name,
                    sheet_name,
                    f"Row {idx + 2} Count={int(count)} but Genes contains {len(genes)} entries.",
                    row_index=str(idx + 2),
                    related_columns=f"{count_col}; {genes_col}",
                )
                break


def _column_candidates(df: pd.DataFrame, tokens: tuple[str, ...]) -> list[str]:
    result = []
    for column in df.columns:
        norm = normalize_column_name(column)
        if any(token in norm for token in tokens):
            result.append(str(column))
    return result


def _check_percent_count_consistency(issues, file_name, sheet_name, df, thresholds):
    numerators = _column_candidates(df, ("positive", "pos", "count", "events", "yes", "hits", "阳性", "计数", "事件数"))
    denominators = _column_candidates(df, ("total", "cells", "denominator", "总数", "细胞数", "样本数"))
    percents = _column_candidates(df, ("percent", "percentage", "rate", "ratio", "%", "阳性率", "比例", "率"))
    for pct_col in percents:
        pct = pd.to_numeric(df[pct_col], errors="coerce")
        for num_col in numerators:
            for den_col in denominators:
                if num_col == den_col or pct_col in {num_col, den_col}:
                    continue
                num = pd.to_numeric(df[num_col], errors="coerce")
                den = pd.to_numeric(df[den_col], errors="coerce")
                expected = num / den.replace(0, np.nan) * 100
                diff = (pct - expected).abs()
                bad = diff > 0.05
                comparable = diff.notna()
                if comparable.sum() < 3:
                    continue
                if bad.sum() / comparable.sum() >= 0.25:
                    examples = [
                        f"row {idx + 2}: reported={pct.loc[idx]}, expected={expected.loc[idx]:.4g}, diff={diff.loc[idx]:.4g}"
                        for idx in diff[bad].head(3).index
                    ]
                    _issue(
                        issues,
                        "percent_count_consistency_detector",
                        "Red",
                        file_name,
                        sheet_name,
                        f"列 {pct_col} 与 {num_col}/{den_col} 计算结果不一致，{int(bad.sum())}/{int(comparable.sum())} 行超出容差。示例：{'; '.join(examples)}。",
                        related_columns=f"{num_col}; {den_col}; {pct_col}",
                        action="建议核对阳性数、总数和百分比计算流程，确认是否存在录入、四舍五入或引用单元格错误。",
                        rule_id="N005",
                        severity="HIGH",
                        details={"numerator": num_col, "denominator": den_col, "percent": pct_col, "bad_rows": int(bad.sum()), "n": int(comparable.sum()), "examples": examples},
                    )
                    return


def _check_duplicate_series(issues, file_name, sheet_name, df, thresholds):
    cols = _numeric_columns(df)
    ndf = _numeric_frame(df, cols)
    for col in cols:
        if _is_category_column(col):
            continue
        values = ndf[col].replace([np.inf, -np.inf], np.nan).dropna().round(9).tolist()
        if len(values) < 10:
            continue
        seen: dict[tuple[float, ...], int] = {}
        window = 5
        for index in range(0, len(values) - window + 1):
            key = tuple(values[index : index + window])
            if key in seen and abs(index - seen[key]) >= window:
                _issue(
                    issues,
                    "duplicate_series_detector",
                    "Orange",
                    file_name,
                    sheet_name,
                    f"列 {col} 中发现长度为 {window} 的重复数值片段，第一次约在第 {seen[key] + 2} 行，第二次约在第 {index + 2} 行。",
                    column_name=col,
                    action="建议核对是否存在局部复制粘贴，或确认该片段是否由实验设计导致。",
                    rule_id="N006",
                    severity="MEDIUM",
                    details={"column": col, "window": window, "first_row": seen[key] + 2, "second_row": index + 2},
                )
                break
            seen[key] = index


def _check_arithmetic_sequence(issues, file_name, sheet_name, df, thresholds):
    if _is_coordinate_or_enrichment_sheet(sheet_name, df):
        return
    cfg = thresholds["numeric"]
    for col in _numeric_columns(df):
        clean = pd.to_numeric(df[col], errors="coerce").dropna()
        if len(clean) < cfg["arithmetic_sequence_min_n"]:
            continue
        diffs = np.round(np.diff(clean.to_numpy(dtype=float)), 9)
        if len(diffs) == 0:
            continue
        values, counts = np.unique(diffs, return_counts=True)
        index = int(np.argmax(counts))
        match_rate = counts[index] / len(diffs)
        longest = _longest_equal_diff_run(clean, 1e-8, 0.001)[0]
        if match_rate >= cfg["arithmetic_sequence_match_rate"] or longest >= 8:
            severity = "MEDIUM"
            if _is_design_sequence_column(col):
                severity = "LOW"
            _issue(
                issues,
                "arithmetic_sequence_detector",
                _risk_from_severity(severity),
                file_name,
                sheet_name,
                f"列 {col} 呈现固定步长模式，常见差值 {values[index]:.6g}，占比 {match_rate:.1%}，最长等差段 {longest} 个数值。",
                column_name=col,
                action="建议核对该列是否为实验设计梯度、时间、剂量或标准曲线；如不是，请回查录入和公式填充。",
                rule_id="N007",
                severity=severity,
                details={"column": col, "diff": float(values[index]), "match_rate": float(match_rate), "longest_run": int(longest)},
            )


def _check_repeated_values(issues, file_name, sheet_name, df, thresholds):
    if _is_coordinate_or_enrichment_sheet(sheet_name, df):
        return
    cfg = thresholds["numeric"]
    for col in _numeric_columns(df):
        if _is_category_column(col):
            continue
        values = pd.to_numeric(df[col], errors="coerce").replace([np.inf, -np.inf], np.nan).dropna()
        if len(values) < cfg["repeated_value_min_n"]:
            continue
        counts = values.round(9).value_counts()
        unique_ratio = len(counts) / len(values)
        dominant_value = float(counts.index[0])
        dominant_ratio = int(counts.iloc[0]) / len(values)
        if abs(dominant_value) > 1e-12 and dominant_ratio > cfg["repeated_value_dominant_ratio"]:
            severity = "HIGH"
        elif unique_ratio < cfg["repeated_value_unique_ratio"]:
            severity = "MEDIUM"
        else:
            continue
        if _looks_like_bounded_percent_or_heatmap(sheet_name, col, values):
            severity = "LOW"
        _issue(
            issues,
            "repeated_value_detector",
            _risk_from_severity(severity),
            file_name,
            sheet_name,
            f"列 {col} 的重复值比例较高，unique_ratio={unique_ratio:.3f}，最常见非空值 {dominant_value} 占 {dominant_ratio:.1%}。",
            column_name=col,
            action="建议核对是否存在批量填充、模板复制、缺省值替代或真实离散分组变量。",
            rule_id="N008",
            severity=severity,
            details={"column": col, "n": len(values), "unique_ratio": unique_ratio, "dominant_value": dominant_value, "dominant_ratio": dominant_ratio},
        )


def _check_grim_like(issues, file_name, sheet_name, df, thresholds):
    mean_cols = _column_candidates(df, ("mean", "average", "均值"))
    n_cols = _column_candidates(df, ("sample_size", "samplesize", "n", "total", "样本量"))
    for mean_col in mean_cols:
        for n_col in n_cols:
            if mean_col == n_col:
                continue
            mean_values = pd.to_numeric(df[mean_col], errors="coerce")
            n_values = pd.to_numeric(df[n_col], errors="coerce")
            product = mean_values * n_values
            comparable = product.dropna()
            if len(comparable) < 3:
                continue
            bad = (comparable - comparable.round()).abs() > 1e-6
            if bad.sum() and (n_values.dropna() <= 200).all():
                examples = [f"row {idx + 2}: mean*n={product.loc[idx]:.6g}" for idx in comparable[bad].head(3).index]
                _issue(
                    issues,
                    "grim_like_detector",
                    "Orange",
                    file_name,
                    sheet_name,
                    f"列 {mean_col} 与样本量列 {n_col} 存在 {int(bad.sum())} 行 mean*n 非整数。该规则仅适用于整数/离散评分均值。",
                    related_columns=f"{mean_col}; {n_col}",
                    action="请先确认该均值是否来自整数计数或离散评分；若是，请核对样本量、均值和四舍五入流程。",
                    rule_id="N009",
                    severity="MEDIUM",
                    details={"mean_column": mean_col, "n_column": n_col, "bad_rows": int(bad.sum()), "examples": examples},
                )
                return


def _check_decimal_precision(issues, file_name, sheet_name, df, thresholds):
    for col in _numeric_columns(df):
        values = df[col].dropna()
        if len(values) < thresholds["numeric"]["terminal_digit_min_n"]:
            continue
        lengths = []
        for value in values:
            text = str(value)
            if "e" in text.lower():
                continue
            match = re.search(r"\.([0-9]+)", text)
            if match:
                lengths.append(len(match.group(1).rstrip("0")))
            else:
                lengths.append(0)
        if len(lengths) < thresholds["numeric"]["terminal_digit_min_n"]:
            continue
        counts = pd.Series(lengths).value_counts()
        dominant_len = int(counts.index[0])
        ratio = int(counts.iloc[0]) / len(lengths)
        if ratio >= 0.95:
            _issue(
                issues,
                "decimal_precision_detector",
                "Yellow",
                file_name,
                sheet_name,
                f"列 {col} 中 {ratio:.1%} 的数值具有相同小数位长度 {dominant_len}。这可能只是正常格式化。",
                column_name=col,
                action="如多列独立实验数据都具有完全相同的小数格式，建议核对仪器导出和格式化流程。",
                rule_id="N011",
                severity="LOW",
                details={"column": col, "n": len(lengths), "dominant_precision": dominant_len, "ratio": ratio},
            )


def _check_cross_file_reuse(issues, parsed_sheets, thresholds):
    rows = thresholds["numeric"].get("cross_file_window_rows", 5)
    cols_required = thresholds["numeric"].get("cross_file_window_cols", 2)
    seen: dict[tuple[float, ...], tuple[str, str, int, str]] = {}
    for file_name, sheet_name, _profile, df in parsed_sheets:
        cols = _numeric_columns(df)
        if len(cols) < cols_required:
            continue
        ndf = _numeric_frame(df, cols).replace([np.inf, -np.inf], np.nan)
        for start in range(0, max(0, len(ndf) - rows + 1)):
            window = ndf.iloc[start : start + rows, :cols_required]
            if window.isna().any().any():
                continue
            key = tuple(np.round(window.to_numpy(dtype=float).ravel(), 9))
            location = (file_name, sheet_name, start + 2, f"{cols[0]}-{cols[cols_required - 1]}")
            if key in seen:
                first = seen[key]
                if first[0] != file_name or first[1] != sheet_name:
                    severity = "HIGH" if any(token in f"{first[0]} {first[1]} {file_name} {sheet_name}".lower() for token in ("control", "treat", "wt", "ko", "day", "drug", "vehicle")) else "MEDIUM"
                    _issue(
                        issues,
                        "cross_file_reuse_detector",
                        _risk_from_severity(severity),
                        file_name,
                        sheet_name,
                        f"发现跨文件/跨 sheet 重复数据块：{first[0]} / {first[1]} 第 {first[2]} 行附近 与 {file_name} / {sheet_name} 第 {start + 2} 行附近相同。",
                        row_index=str(start + 2),
                        related_columns=location[3],
                        action="建议核对两个位置是否应共享同一数据块；若属于不同实验条件，请重点回查原始数据和复制粘贴记录。",
                        rule_id="N010",
                        severity=severity,
                        details={"first_file": first[0], "first_sheet": first[1], "first_row": first[2], "second_file": file_name, "second_sheet": sheet_name, "second_row": start + 2, "window_rows": rows, "window_cols": cols_required},
                    )
                    return
            else:
                seen[key] = location


def _add_composite_numeric_risk(issues: list[dict[str, Any]]) -> None:
    by_area: dict[tuple[str, str], set[str]] = {}
    for issue in issues:
        key = (issue.get("file_name", ""), issue.get("sheet_name", ""))
        by_area.setdefault(key, set()).add(issue.get("rule_id", ""))
    for (file_name, sheet_name), rules in by_area.items():
        if {"N001", "N003", "N004"}.issubset(rules):
            _issue(
                issues,
                "composite_numeric_pattern",
                "Red",
                file_name,
                sheet_name,
                "同一数据区域同时命中固定差值、小数部分重复和末位数字异常，统计模式不自然，需优先复核。",
                action="建议优先核对原始记录、仪器导出、样本编号和分析脚本。",
                rule_id="COMPOSITE_N001_N003_N004",
                severity="CRITICAL",
                details={"rules": sorted(rules)},
            )
        if {"N005", "N011"}.issubset(rules):
            _issue(
                issues,
                "composite_percent_precision_pattern",
                "Red",
                file_name,
                sheet_name,
                "同一数据区域同时命中百分比/计数不一致与小数位整齐度异常，需优先复核计算流程。",
                action="建议核对原始计数、百分比公式、四舍五入和表格格式化流程。",
                rule_id="COMPOSITE_N005_N011",
                severity="HIGH",
                details={"rules": sorted(rules)},
            )


def run_numeric_forensics(parsed_sheets: list[tuple[str, str, str, pd.DataFrame]], thresholds: dict[str, Any] | None = None) -> list[dict[str, Any]]:
    cfg = merged_thresholds(thresholds)
    issues: list[dict[str, Any]] = []
    for file_name, sheet_name, profile, df in parsed_sheets:
        if df.empty:
            continue
        _check_exact_duplicates(issues, file_name, sheet_name, df)
        _check_near_duplicate_rows(issues, file_name, sheet_name, df, cfg)
        _check_duplicate_and_related_columns(issues, file_name, sheet_name, df, cfg)
        _check_constant_delta(issues, file_name, sheet_name, df, cfg)
        _check_affine_relation(issues, file_name, sheet_name, df, cfg)
        _check_duplicate_decimal(issues, file_name, sheet_name, df, cfg)
        _check_equal_difference(issues, file_name, sheet_name, df, cfg)
        _check_terminal_digits(issues, file_name, sheet_name, df, cfg)
        _check_ranges(issues, file_name, sheet_name, df, cfg)
        _check_missingness(issues, file_name, sheet_name, df, cfg)
        _check_percent_count_consistency(issues, file_name, sheet_name, df, cfg)
        _check_duplicate_series(issues, file_name, sheet_name, df, cfg)
        _check_arithmetic_sequence(issues, file_name, sheet_name, df, cfg)
        _check_repeated_values(issues, file_name, sheet_name, df, cfg)
        _check_grim_like(issues, file_name, sheet_name, df, cfg)
        _check_decimal_precision(issues, file_name, sheet_name, df, cfg)
        if profile == "enrichment":
            _check_enrichment(issues, file_name, sheet_name, df)
    _check_cross_file_reuse(issues, parsed_sheets, cfg)
    _add_composite_numeric_risk(issues)
    return issues


def write_numeric_results(issues: list[dict[str, Any]], output_path) -> pd.DataFrame:
    df = pd.DataFrame(issues)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(output_path, index=False)
    return df
