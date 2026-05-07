from __future__ import annotations

from pathlib import Path
from typing import Any

import pandas as pd

from .file_classifier import classify_file
from .table_profiler import detect_profile, make_unique_columns


SHEET_INVENTORY_COLUMNS = [
    "file_name",
    "sheet_name",
    "n_rows",
    "n_cols",
    "columns_detected",
    "detected_table_type",
    "qc_profile",
    "parse_status",
    "parse_warning",
]


ParsedSheet = tuple[str, str, str, pd.DataFrame]


def _finalize_frame(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = make_unique_columns(df.columns)
    return df


def parse_table_file(path: Path) -> tuple[list[ParsedSheet], list[dict[str, Any]]]:
    parsed: list[ParsedSheet] = []
    inventory: list[dict[str, Any]] = []
    try:
        file_type = classify_file(path)
        if file_type == "excel":
            sheets = pd.read_excel(path, sheet_name=None)
        elif path.suffix.lower() == ".tsv":
            sheets = {"TSV": pd.read_csv(path, sep="\t")}
        else:
            sheets = {"CSV": pd.read_csv(path)}
        for sheet_name, df in sheets.items():
            df = _finalize_frame(df)
            detected, profile = detect_profile(df.columns)
            parsed.append((path.name, str(sheet_name), profile, df))
            inventory.append(
                {
                    "file_name": path.name,
                    "sheet_name": str(sheet_name),
                    "n_rows": int(df.shape[0]),
                    "n_cols": int(df.shape[1]),
                    "columns_detected": ", ".join(map(str, df.columns)),
                    "detected_table_type": detected,
                    "qc_profile": profile,
                    "parse_status": "ok",
                    "parse_warning": "" if not df.empty else "empty sheet",
                }
            )
    except Exception as exc:
        inventory.append(
            {
                "file_name": path.name,
                "sheet_name": "",
                "n_rows": 0,
                "n_cols": 0,
                "columns_detected": "",
                "detected_table_type": "",
                "qc_profile": "",
                "parse_status": "failed",
                "parse_warning": str(exc),
            }
        )
    return parsed, inventory


def parse_tables(root: Path, output_path: Path | None = None) -> tuple[list[ParsedSheet], pd.DataFrame]:
    parsed: list[ParsedSheet] = []
    inventory_rows: list[dict[str, Any]] = []
    for path in sorted(p for p in root.rglob("*") if p.is_file()):
        if classify_file(path) in {"excel", "csv"}:
            sheets, rows = parse_table_file(path)
            parsed.extend(sheets)
            inventory_rows.extend(rows)
    inventory = pd.DataFrame(inventory_rows, columns=SHEET_INVENTORY_COLUMNS)
    if output_path is not None:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        inventory.to_excel(output_path, index=False)
    return parsed, inventory
