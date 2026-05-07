from __future__ import annotations

import math
import re
from itertools import combinations
from typing import Any

import numpy as np
import pandas as pd

from .table_profiler import normalize_column_name


DEFAULT_THRESHOLDS = {
    "duplicate": {
        "near_duplicate_similarity_threshold": 0.95,
        "red_duplicate_similarity_threshold": 0.99,
        "red_column_correlation_threshold": 0.999,
    },
    "ratio_pattern": {"min_matched_rows": 8, "ratio_cv_threshold": 0.01},
    "equal_difference": {
        "min_run_length_red": 5,
        "min_run_length_orange": 4,
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
) -> None:
    issues.append(
        {
            "issue_id": f"NUM{len(issues) + 1:03d}",
            "module": "Numeric Forensics",
            "risk_level": risk_level,
            "issue_type": issue_type,
            "file_name": file_name,
            "sheet_name": sheet_name,
            "row_index": row_index,
            "column_name": column_name,
            "related_columns": related_columns,
            "evidence": evidence,
            "recommended_action": action,
            "need_human_review": "Yes" if risk_level in {"Red", "Orange"} else "Recommended",
            "affects_submission": "Yes" if risk_level in {"Red", "Orange"} else "Review",
        }
    )


def _numeric_columns(df: pd.DataFrame) -> list[str]:
    columns = []
    for column in df.columns:
        numeric = pd.to_numeric(df[column], errors="coerce")
        if numeric.notna().sum() >= 2:
            columns.append(str(column))
    return columns


def _numeric_frame(df: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    return pd.DataFrame({column: pd.to_numeric(df[column], errors="coerce") for column in columns})


def _check_exact_duplicates(issues, file_name, sheet_name, df):
    comparable = df.dropna(axis=1, how="all")
    if comparable.empty:
        return
    duplicated = comparable.duplicated(keep=False)
    if duplicated.any():
        rows = [str(i + 2) for i in comparable.index[duplicated].tolist()]
        _issue(
            issues,
            "exact_duplicate_rows",
            "Red",
            file_name,
            sheet_name,
            f"Rows {', '.join(rows)} are exact duplicates across {comparable.shape[1]} comparable columns.",
            row_index=", ".join(rows),
            action="回查原始仪器导出文件、实验记录本和数据录入记录，确认是否存在复制粘贴或样本ID错配。",
        )


def _check_near_duplicate_rows(issues, file_name, sheet_name, df, thresholds):
    cols = _numeric_columns(df)
    if len(cols) < 5:
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
    ndf = _numeric_frame(df, cols)
    min_pairs = thresholds["correlation"]["min_non_missing_pairs"]
    for a, b in combinations(cols, 2):
        pair = ndf[[a, b]].replace([np.inf, -np.inf], np.nan).dropna()
        pair = pair[(pair[a].abs() <= thresholds["extreme_values"]["float_max_warning_threshold"]) & (pair[b].abs() <= thresholds["extreme_values"]["float_max_warning_threshold"])]
        if len(pair) < 2:
            continue
        if pair[a].equals(pair[b]):
            risk = "Red" if len(pair) >= 8 else "Orange"
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
        if len(pair) >= min_pairs and pair[a].std() > 0 and pair[b].std() > 0:
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
    cfg = thresholds["equal_difference"]
    for col in _numeric_columns(df):
        run_len, start, diff = _longest_equal_diff_run(df[col], cfg["absolute_tolerance"], cfg["relative_tolerance"])
        if run_len >= cfg["min_run_length_orange"]:
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
    cfg = thresholds["digit_pattern"]
    for col in _numeric_columns(df):
        values = pd.to_numeric(df[col], errors="coerce").dropna().to_numpy(dtype=float)
        digits = [_terminal_digit(v) for v in values]
        digits = [d for d in digits if d is not None]
        if len(digits) < 10:
            continue
        frac = sum(1 for d in digits if d in {0, 5}) / len(digits)
        if frac > cfg["max_zero_or_five_fraction_orange"]:
            risk = "Red" if frac > cfg["max_zero_or_five_fraction_red"] and len(digits) >= cfg["min_n_for_digit_test"] else "Orange"
            if len(digits) < cfg["min_n_for_digit_test"]:
                risk = "Yellow"
            _issue(
                issues,
                "terminal_digit_anomaly",
                risk,
                file_name,
                sheet_name,
                f"Column {col} has {frac:.1%} values ending in 0 or 5 across {len(digits)} values.",
                column_name=col,
                action="检查仪器精度、人工录入习惯、四舍五入规则和原始记录。",
            )


def _check_ranges(issues, file_name, sheet_name, df, thresholds):
    extreme = thresholds["extreme_values"]["float_max_warning_threshold"]
    for column in df.columns:
        norm = normalize_column_name(column)
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
            _issue(
                issues,
                "high_column_missingness",
                "Orange" if frac > cfg["high_column_missing_fraction_orange"] else "Yellow",
                file_name,
                sheet_name,
                f"Column {column} missingness is {frac:.1%}.",
                column_name=str(column),
            )
    row_fracs = df.isna().mean(axis=1)
    high_rows = row_fracs[row_fracs > cfg["high_row_missing_fraction_yellow"]]
    if len(high_rows):
        _issue(
            issues,
            "high_row_missingness",
            "Orange" if high_rows.max() > cfg["high_row_missing_fraction_orange"] else "Yellow",
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


def run_numeric_forensics(parsed_sheets: list[tuple[str, str, str, pd.DataFrame]], thresholds: dict[str, Any] | None = None) -> list[dict[str, Any]]:
    cfg = merged_thresholds(thresholds)
    issues: list[dict[str, Any]] = []
    for file_name, sheet_name, profile, df in parsed_sheets:
        if df.empty:
            continue
        _check_exact_duplicates(issues, file_name, sheet_name, df)
        _check_near_duplicate_rows(issues, file_name, sheet_name, df, cfg)
        _check_duplicate_and_related_columns(issues, file_name, sheet_name, df, cfg)
        _check_equal_difference(issues, file_name, sheet_name, df, cfg)
        _check_terminal_digits(issues, file_name, sheet_name, df, cfg)
        _check_ranges(issues, file_name, sheet_name, df, cfg)
        _check_missingness(issues, file_name, sheet_name, df, cfg)
        if profile == "enrichment":
            _check_enrichment(issues, file_name, sheet_name, df)
    return issues


def write_numeric_results(issues: list[dict[str, Any]], output_path) -> pd.DataFrame:
    df = pd.DataFrame(issues)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(output_path, index=False)
    return df
