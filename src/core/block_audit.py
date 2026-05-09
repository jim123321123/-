from __future__ import annotations

from dataclasses import dataclass
from itertools import combinations
from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd


@dataclass
class NumericBlock:
    matrix: np.ndarray
    row_numbers: list[int]
    col_names: list[str]
    address: str


def _issue(
    issues: list[dict[str, Any]],
    issue_type: str,
    risk_level: str,
    file_name: str,
    sheet_name: str,
    evidence: str,
    location: str = "",
    action: str = "请回到原始 Excel 附表、实验记录和数据处理脚本，确认该区块关系是否有合理来源。",
) -> None:
    issues.append(
        {
            "issue_id": f"BLK{len(issues) + 1:03d}",
            "module": "Supplementary Table Block Audit",
            "risk_level": risk_level,
            "issue_type": issue_type,
            "file_name": file_name,
            "sheet_name": sheet_name,
            "row_index": location,
            "column_name": "",
            "related_columns": "",
            "evidence": evidence,
            "recommended_action": action,
            "need_human_review": "Yes" if risk_level in {"Red", "Orange"} else "Recommended",
            "affects_submission": "Yes" if risk_level in {"Red", "Orange"} else "Review",
        }
    )


def _numeric_mask(df: pd.DataFrame) -> pd.DataFrame:
    return df.apply(lambda col: pd.to_numeric(col, errors="coerce").notna())


def _contiguous_groups(indices: list[int], min_len: int) -> list[list[int]]:
    groups: list[list[int]] = []
    current: list[int] = []
    previous = None
    for index in indices:
        if previous is None or index == previous + 1:
            current.append(index)
        else:
            if len(current) >= min_len:
                groups.append(current)
            current = [index]
        previous = index
    if len(current) >= min_len:
        groups.append(current)
    return groups


def extract_numeric_blocks(df: pd.DataFrame, min_rows: int = 4, min_cols: int = 2) -> list[NumericBlock]:
    if df.empty:
        return []
    mask = _numeric_mask(df)
    dense_cols = [idx for idx, col in enumerate(mask.columns) if int(mask[col].sum()) >= min_rows]
    col_groups = _contiguous_groups(dense_cols, min_cols)
    blocks: list[NumericBlock] = []
    for col_group in col_groups:
        submask = mask.iloc[:, col_group]
        row_hits = submask.sum(axis=1)
        min_row_hits = max(2, int(np.ceil(len(col_group) * 0.5)))
        dense_rows = [idx for idx, value in row_hits.items() if int(value) >= min_row_hits]
        row_groups = _contiguous_groups(dense_rows, min_rows)
        for row_group in row_groups:
            raw = df.iloc[row_group, col_group]
            numeric = raw.apply(pd.to_numeric, errors="coerce").to_numpy(dtype=float)
            if np.isnan(numeric).mean() > 0.25:
                continue
            row_numbers = [idx + 2 for idx in row_group]
            col_names = [str(df.columns[idx]) for idx in col_group]
            address = f"rows {row_numbers[0]}-{row_numbers[-1]}, columns {col_names[0]}-{col_names[-1]}"
            blocks.append(NumericBlock(numeric, row_numbers, col_names, address))
    return blocks


def _same_shape(a: NumericBlock, b: NumericBlock) -> bool:
    return a.matrix.shape == b.matrix.shape


def _valid_pair(a: np.ndarray, b: np.ndarray) -> np.ndarray:
    return ~(np.isnan(a) | np.isnan(b))


def _constant_control_vector(values: np.ndarray) -> bool:
    clean = values[~np.isnan(values)]
    if len(clean) < 3:
        return False
    unique = np.unique(np.round(clean.astype(float), 9))
    if len(unique) != 1:
        return False
    return any(np.isclose(unique[0], expected, atol=1e-9) for expected in (0.0, 1.0, 100.0))


def _duplicate_risk_for_width(block: NumericBlock) -> str:
    return "Red" if block.matrix.shape[1] >= 4 else "Yellow"


def _check_identical(a: NumericBlock, b: NumericBlock) -> str | None:
    if not _same_shape(a, b):
        return None
    mask = _valid_pair(a.matrix, b.matrix)
    if int(mask.sum()) < 4:
        return None
    if np.allclose(a.matrix[mask], b.matrix[mask], rtol=0, atol=1e-9):
        rows, cols = a.matrix.shape
        return f"区块 {a.address} 与区块 {b.address} 的 {rows} 行 x {cols} 列数值完全相同。"
    return None


def _check_ratio(a: NumericBlock, b: NumericBlock) -> str | None:
    if not _same_shape(a, b):
        return None
    mask = _valid_pair(a.matrix, b.matrix) & (np.abs(a.matrix) > 1e-12)
    if int(mask.sum()) < 6:
        return None
    ratios = b.matrix[mask] / a.matrix[mask]
    mean_ratio = float(np.mean(ratios))
    if abs(mean_ratio) < 1e-12 or abs(mean_ratio - 1.0) < 1e-9:
        return None
    cv = float(np.std(ratios) / abs(mean_ratio)) if mean_ratio else np.inf
    if cv < 1e-4:
        return f"区块 {b.address} 约等于区块 {a.address} 乘以固定倍数 {mean_ratio:.6g}，倍数变异系数 {cv:.2e}。"
    return None


def _check_difference(a: NumericBlock, b: NumericBlock) -> str | None:
    if not _same_shape(a, b):
        return None
    mask = _valid_pair(a.matrix, b.matrix)
    if int(mask.sum()) < 6:
        return None
    diffs = b.matrix[mask] - a.matrix[mask]
    mean_diff = float(np.mean(diffs))
    if abs(mean_diff) < 1e-12:
        return None
    combined = np.concatenate([a.matrix[mask], b.matrix[mask]])
    range_value = float(np.ptp(combined))
    if range_value <= 0:
        return None
    relative_std = float(np.std(diffs) / range_value)
    if relative_std < 1e-4:
        return f"区块 {b.address} 约等于区块 {a.address} 加上固定差值 {mean_diff:.6g}，相对误差 {relative_std:.2e}。"
    return None


def _check_internal_duplicates(block: NumericBlock) -> list[tuple[str, str, str]]:
    findings: list[tuple[str, str, str]] = []
    matrix = block.matrix
    seen_rows: dict[tuple[Any, ...], int] = {}
    for index, row in enumerate(matrix):
        if np.isnan(row).all():
            continue
        key = tuple(np.where(np.isnan(row), None, np.round(row, 9)))
        if key in seen_rows:
            first = seen_rows[key]
            risk = _duplicate_risk_for_width(block)
            findings.append(
                (
                    "block_internal_duplicate_rows",
                    risk,
                    f"区块 {block.address} 内第 {block.row_numbers[first]} 行与第 {block.row_numbers[index]} 行数值完全相同。",
                )
            )
            break
        seen_rows[key] = index
    seen_cols: dict[tuple[Any, ...], int] = {}
    for index in range(matrix.shape[1]):
        col = matrix[:, index]
        if np.isnan(col).all():
            continue
        key = tuple(np.where(np.isnan(col), None, np.round(col, 9)))
        if key in seen_cols:
            first = seen_cols[key]
            if _constant_control_vector(col) and _constant_control_vector(matrix[:, first]):
                findings.append(
                    (
                        "block_duplicate_constant_control_columns",
                        "Yellow",
                        f"区块 {block.address} 内列 {block.col_names[first]} 与列 {block.col_names[index]} 为完全相同的常数对照值，可能是归一化对照列。",
                    )
                )
                break
            findings.append(
                (
                    "block_internal_duplicate_columns",
                    "Red",
                    f"区块 {block.address} 内列 {block.col_names[first]} 与列 {block.col_names[index]} 数值完全相同。",
                )
            )
            break
        seen_cols[key] = index
    return findings


def run_block_audit(
    parsed_sheets: list[tuple[str, str, str, pd.DataFrame]],
    max_sheet_cells: int = 200_000,
    max_blocks_per_sheet: int = 80,
) -> list[dict[str, Any]]:
    issues: list[dict[str, Any]] = []
    for file_name, sheet_name, _profile, df in parsed_sheets:
        cell_count = int(df.shape[0] * df.shape[1])
        if cell_count > max_sheet_cells:
            _issue(
                issues,
                "block_audit_skipped_large_sheet",
                "Yellow",
                file_name,
                sheet_name,
                f"该 sheet 共有 {df.shape[0]} 行 x {df.shape[1]} 列，超过区块审计上限 {max_sheet_cells} 个单元格，已跳过区块两两比较。",
                action="该表规模很大，软件已保留常规数值质控；如需区块审计，请按实验模块拆分后单独检查。",
            )
            continue
        blocks = extract_numeric_blocks(df)
        if len(blocks) > max_blocks_per_sheet:
            _issue(
                issues,
                "block_audit_skipped_many_blocks",
                "Yellow",
                file_name,
                sheet_name,
                f"该 sheet 识别出 {len(blocks)} 个数值区块，超过上限 {max_blocks_per_sheet}，已跳过区块两两比较。",
                action="请优先查看常规 QC 结果；如需区块审计，请将该 sheet 按图表或实验模块拆分。",
            )
            continue
        for block in blocks:
            for issue_type, risk, evidence in _check_internal_duplicates(block):
                _issue(issues, issue_type, risk, file_name, sheet_name, evidence, location=block.address)
        for left, right in combinations(blocks, 2):
            identical = _check_identical(left, right)
            if identical:
                _issue(issues, "block_identical", "Red", file_name, sheet_name, identical, location=f"{left.address}; {right.address}")
                continue
            ratio = _check_ratio(left, right)
            if ratio:
                _issue(issues, "block_fixed_ratio", "Orange", file_name, sheet_name, ratio, location=f"{left.address}; {right.address}")
                continue
            diff = _check_difference(left, right)
            if diff:
                _issue(issues, "block_fixed_difference", "Orange", file_name, sheet_name, diff, location=f"{left.address}; {right.address}")
    return issues


def write_block_audit_results(issues: list[dict[str, Any]], output_path: Path) -> pd.DataFrame:
    df = pd.DataFrame(issues)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(output_path, index=False)
    return df
