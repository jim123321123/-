from __future__ import annotations

import ast
import json
import re
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from typing import Any

import pandas as pd

from src.core.report_language import ISSUE_LABELS


RISK_RANK = {"Red": 3, "Orange": 2, "Yellow": 1, "": 0}


@dataclass
class HighlightTarget:
    file_name: str
    sheet_name: str
    rows: set[int] = field(default_factory=set)
    columns: set[str] = field(default_factory=set)
    cells: set[tuple[int, str]] = field(default_factory=set)
    note: str = ""
    scope: str = "cell"


def parse_details(value: Any) -> dict[str, Any]:
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


def rule_key(row: pd.Series) -> str:
    rule_id = str(row.get("rule_id", "") or "").strip()
    issue_type = str(row.get("issue_type", "") or "").strip()
    return rule_id or issue_type or "unclassified"


def rule_title(row: pd.Series) -> str:
    key = rule_key(row)
    issue_type = str(row.get("issue_type", "") or "")
    label = ISSUE_LABELS.get(issue_type, issue_type or "未分类问题")
    return f"{key} {label}" if key != issue_type else label


def summarize_rules(issue_log: pd.DataFrame) -> list[dict[str, Any]]:
    if issue_log is None or issue_log.empty:
        return []
    rows = []
    for key, group in issue_log.groupby(issue_log.apply(rule_key, axis=1), sort=False):
        first = group.iloc[0]
        risk = max((str(v) for v in group.get("risk_level", pd.Series(dtype=str)).fillna("")), key=lambda v: RISK_RANK.get(v, 0), default="")
        files = int(group["file_name"].replace("", pd.NA).dropna().nunique()) if "file_name" in group else 0
        rows.append(
            {
                "key": key,
                "title": rule_title(first),
                "count": int(len(group)),
                "risk": risk or "未分级",
                "files": files,
                "issue_type": str(first.get("issue_type", "")),
            }
        )
    rows.sort(key=lambda item: (RISK_RANK.get(item["risk"], 0), item["count"]), reverse=True)
    return rows


def table_index(parsed_sheets: list[tuple[str, str, str, pd.DataFrame]] | None) -> dict[tuple[str, str], pd.DataFrame]:
    return {(file_name, sheet_name): df for file_name, sheet_name, _, df in (parsed_sheets or [])}


def _ints(text: Any) -> list[int]:
    return [int(value) for value in re.findall(r"\d+", str(text or ""))]


def _columns_from_text(text: Any, df: pd.DataFrame) -> set[str]:
    source = str(text or "")
    columns = {str(col) for col in df.columns}
    found = {col for col in columns if col and re.search(rf"(?<!\w){re.escape(col)}(?!\w)", source)}
    for token in re.split(r"[;,，、\s]+", source):
        if token in columns:
            found.add(token)
    return found


def _numeric_columns(df: pd.DataFrame) -> set[str]:
    return {str(col) for col in df.columns if pd.to_numeric(df[col], errors="coerce").notna().sum() >= 2}


def _actual_column(df: pd.DataFrame, column: str) -> Any | None:
    for actual in df.columns:
        if str(actual) == str(column):
            return actual
    return None


def _non_empty_cells(df: pd.DataFrame, rows: set[int], columns: set[str]) -> set[tuple[int, str]]:
    cells: set[tuple[int, str]] = set()
    for row_no in rows:
        idx = row_no - 2
        if idx < 0 or idx >= len(df):
            continue
        for col in columns:
            actual_col = _actual_column(df, col)
            if actual_col is None:
                continue
            value = df.iloc[idx][actual_col]
            if not pd.isna(value):
                cells.add((row_no, col))
    return cells


def _column_data_cells(df: pd.DataFrame, columns: set[str]) -> set[tuple[int, str]]:
    rows = set(range(2, len(df) + 2))
    return _non_empty_cells(df, rows, columns)


def _last_digit(value: Any) -> str | None:
    if value is None or pd.isna(value):
        return None
    text = str(value).strip()
    if not text:
        return None
    try:
        text = format(Decimal(text), "f")
    except (InvalidOperation, ValueError):
        pass
    digits = re.sub(r"\D", "", text.rstrip("0"))
    return digits[-1] if digits else None


def _decimal_precision(value: Any) -> int | None:
    if value is None or pd.isna(value):
        return None
    text = str(value).strip()
    if not text:
        return None
    if "." in text and "e" not in text.lower():
        return len(text.split(".", 1)[1].rstrip())
    try:
        number = Decimal(text)
    except (InvalidOperation, ValueError):
        return None
    return max(-number.as_tuple().exponent, 0)


def _dominant_digit_cells(df: pd.DataFrame, column: str, digit: Any) -> set[tuple[int, str]]:
    actual_col = _actual_column(df, column)
    if actual_col is None or digit in (None, ""):
        return set()
    expected = str(digit)
    cells = set()
    for idx, value in df[actual_col].items():
        if _last_digit(value) == expected:
            cells.add((int(idx) + 2, str(column)))
    return cells


def _precision_outlier_cells(df: pd.DataFrame, column: str, dominant_precision: Any) -> set[tuple[int, str]]:
    actual_col = _actual_column(df, column)
    if actual_col is None or dominant_precision in (None, ""):
        return set()
    try:
        expected = int(dominant_precision)
    except (TypeError, ValueError):
        return set()
    cells = set()
    for idx, value in df[actual_col].items():
        precision = _decimal_precision(value)
        if precision is not None and precision != expected:
            cells.add((int(idx) + 2, str(column)))
    return cells


def _arithmetic_cells(df: pd.DataFrame, column: str, diff: Any) -> set[tuple[int, str]]:
    actual_col = _actual_column(df, column)
    if actual_col is None or diff in (None, ""):
        return set()
    try:
        expected = float(diff)
    except (TypeError, ValueError):
        return set()
    values = pd.to_numeric(df[actual_col], errors="coerce")
    cells = set()
    previous_idx: int | None = None
    previous_value: float | None = None
    for idx, value in values.items():
        if pd.isna(value):
            previous_idx = None
            previous_value = None
            continue
        current = float(value)
        if previous_idx is not None and previous_value is not None and abs((current - previous_value) - expected) < 1e-9:
            cells.add((int(previous_idx) + 2, str(column)))
            cells.add((int(idx) + 2, str(column)))
        previous_idx = int(idx)
        previous_value = current
    return cells


def _rows_from_examples(examples: Any) -> set[int]:
    rows: set[int] = set()
    if isinstance(examples, list):
        for item in examples:
            match = re.search(r"row\s+(\d+)", str(item), flags=re.I)
            if match:
                rows.add(int(match.group(1)))
    return rows


def _block_target(text: str, df: pd.DataFrame) -> tuple[set[int], set[str]]:
    rows: set[int] = set()
    columns: set[str] = set()
    row_match = re.search(r"rows\s+(\d+)\s*-\s*(\d+)", text, flags=re.I)
    if row_match:
        rows.update(range(int(row_match.group(1)), int(row_match.group(2)) + 1))
    col_match = re.search(r"columns\s+(.+?)\s*-\s*(.+)$", text, flags=re.I)
    if col_match:
        names = [str(col) for col in df.columns]
        start, end = col_match.group(1).strip(), col_match.group(2).strip()
        if start in names and end in names:
            left, right = names.index(start), names.index(end)
            if left <= right:
                columns.update(names[left : right + 1])
    return rows, columns


def _single_target(row: pd.Series, df: pd.DataFrame) -> HighlightTarget:
    issue_type = str(row.get("issue_type", "") or "")
    details = parse_details(row.get("details", ""))
    file_name = str(row.get("file_name", "") or "")
    sheet_name = str(row.get("sheet_or_panel", "") or "")
    target = HighlightTarget(file_name=file_name, sheet_name=sheet_name)
    sample = str(row.get("sample_or_variable", "") or "")
    evidence = str(row.get("evidence", "") or "")

    if issue_type in {"exact_duplicate_rows", "near_duplicate_rows"}:
        target.rows.update(_ints(sample) or _ints(evidence))
        if issue_type == "near_duplicate_rows":
            target.cells.update((row_no, col) for row_no in target.rows for col in _numeric_columns(df))
        else:
            target.cells.update(_non_empty_cells(df, target.rows, _numeric_columns(df)))
        target.note = "行级重复问题：只比较并高亮重复的数字单元格，文字相同不作为重复依据。"
        target.scope = "row"
        return target

    if issue_type.startswith("block_"):
        rows, columns = _block_target(sample or evidence, df)
        target.rows.update(rows)
        target.columns.update(columns)
        target.cells.update((row_no, col) for row_no in rows for col in columns)
        target.note = "区块级问题：高亮参与匹配的连续数值区块。"
        return target

    if issue_type in {"percent_count_consistency_detector", "grim_like_detector"}:
        cols = {
            str(details.get("numerator", "")),
            str(details.get("denominator", "")),
            str(details.get("percent", "")),
            str(details.get("mean_column", "")),
            str(details.get("n_column", "")),
        } & {str(col) for col in df.columns}
        rows = _rows_from_examples(details.get("examples", []))
        target.columns.update(cols)
        target.cells.update((row_no, col) for row_no in rows for col in cols)
        target.note = "单元格级问题：高亮示例中不一致的行与相关计算列。"
        return target

    if issue_type == "duplicate_series_detector":
        col = str(details.get("column", ""))
        window = int(details.get("window", 0) or 0)
        rows = set()
        for key in ("first_row", "second_row"):
            start = int(details.get(key, 0) or 0)
            if start and window:
                rows.update(range(start, start + window))
        if _actual_column(df, col) is not None:
            target.cells.update((row_no, col) for row_no in rows)
        target.note = "序列级问题：高亮重复序列所在的单元格。"
        return target

    if issue_type == "repeated_value_detector":
        col = str(details.get("column", sample))
        actual_col = _actual_column(df, col)
        if actual_col is not None:
            dominant = details.get("dominant_value", None)
            values = pd.to_numeric(df[actual_col], errors="coerce")
            if dominant is not None:
                rows = {int(idx) + 2 for idx, value in values.items() if pd.notna(value) and abs(float(value) - float(dominant)) < 1e-9}
                target.cells.update((row_no, col) for row_no in rows)
        target.note = "列内重复值问题：只高亮重复值所在单元格。"
        return target

    if issue_type == "duplicate_decimal_detector":
        columns = {
            str(details.get("left_column", "")),
            str(details.get("right_column", "")),
        } & {str(col) for col in df.columns}
        rows = _rows_from_examples(details.get("examples", []))
        target.columns.update(columns)
        target.cells.update(_non_empty_cells(df, rows, columns))
        target.note = "小数尾部重复问题：只高亮示例行中参与比较的两个单元格。"
        return target

    if issue_type == "equal_difference_run":
        match = re.search(r"Rows\s+(\d+)\s*-\s*(\d+)\s+in column\s+(.+?)\s+show", evidence, flags=re.I)
        if match:
            start, end, col = int(match.group(1)), int(match.group(2)), match.group(3).strip()
            if _actual_column(df, col) is not None:
                rows = set(range(start, end + 1))
                target.columns.add(col)
                target.cells.update(_non_empty_cells(df, rows, {col}))
        target.note = "固定间隔问题：只高亮证据中给出的连续单元格。"
        return target

    if issue_type == "terminal_digit_anomaly":
        col = str(details.get("column", sample))
        if _actual_column(df, col) is not None:
            target.columns.add(col)
            target.cells.update(_dominant_digit_cells(df, col, details.get("dominant_digit")))
        target.note = "末位数字问题：只高亮末位数字符合异常集中模式的单元格。"
        return target

    if issue_type == "decimal_precision_detector":
        col = str(details.get("column", sample))
        if _actual_column(df, col) is not None:
            target.columns.add(col)
            target.cells.update(_precision_outlier_cells(df, col, details.get("dominant_precision")))
        target.note = "小数位数问题：只高亮小数位数不同于该列主要格式的单元格。"
        return target

    if issue_type == "arithmetic_sequence_detector":
        col = str(details.get("column", sample))
        if _actual_column(df, col) is not None:
            target.columns.add(col)
            target.cells.update(_arithmetic_cells(df, col, details.get("diff")))
        target.note = "等差序列问题：只高亮参与连续等差关系的单元格。"
        return target

    column_fields = {
        str(details.get("column", "")),
        str(details.get("left_column", "")),
        str(details.get("right_column", "")),
        str(details.get("mean_column", "")),
        str(details.get("n_column", "")),
        sample,
    }
    target.columns.update(col for col in column_fields if _actual_column(df, col) is not None)
    target.columns.update(_columns_from_text(evidence, df))
    if target.columns:
        matched_rows = set(int(row) for row in details.get("matched_rows", []) if str(row).isdigit()) if isinstance(details.get("matched_rows", []), list) else set()
        hit_rows = set(int(row) for row in details.get("hit_rows", []) if str(row).isdigit()) if isinstance(details.get("hit_rows", []), list) else set()
        rows = matched_rows or hit_rows
        if rows:
            target.cells.update(_non_empty_cells(df, rows, target.columns))
            target.note = "列级/多列关系问题：只高亮规则明确定位到的问题单元格。"
        else:
            target.note = "列级/多列统计问题：该规则判断对象是整列或多列之间的关系，因此高亮相关列，请结合上方文字说明复核。"
            target.scope = "column"
    return target


def highlight_targets(row: pd.Series, tables: dict[tuple[str, str], pd.DataFrame]) -> list[HighlightTarget]:
    details = parse_details(row.get("details", ""))
    issue_type = str(row.get("issue_type", "") or "")
    if issue_type == "cross_file_reuse_detector":
        targets = []
        for file_key, sheet_key, row_key in [
            ("first_file", "first_sheet", "first_row"),
            ("second_file", "second_sheet", "second_row"),
        ]:
            file_name = str(details.get(file_key, "") or "")
            sheet_name = str(details.get(sheet_key, "") or "")
            df = tables.get((file_name, sheet_name))
            if df is None:
                continue
            rows = set(range(int(details.get(row_key, 0) or 0), int(details.get(row_key, 0) or 0) + int(details.get("window_rows", 0) or 0)))
            numeric_columns = _numeric_columns(df)
            columns = [str(col) for col in df.columns if str(col) in numeric_columns][: int(details.get("window_cols", 0) or 0)]
            targets.append(
                HighlightTarget(
                    file_name=file_name,
                    sheet_name=sheet_name,
                    rows=rows,
                    columns=set(columns),
                    cells={(row_no, col) for row_no in rows for col in columns},
                    note="跨文件问题：两个源表分别显示，只高亮参与复用匹配的窗口。",
                )
            )
        return targets

    file_name = str(row.get("file_name", "") or "")
    sheet_name = str(row.get("sheet_or_panel", "") or "")
    df = tables.get((file_name, sheet_name))
    if df is None:
        return []
    return [_single_target(row, df)]
