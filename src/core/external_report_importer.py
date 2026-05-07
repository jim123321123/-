from __future__ import annotations

import shutil
from pathlib import Path
from typing import Any

import pandas as pd


def guess_tool_name(path: Path) -> str:
    name = path.name.lower()
    if "proofig" in name:
        return "Proofig AI"
    if "imagetwin" in name:
        return "Imagetwin"
    if "dataseer" in name:
        return "DataSeer"
    return "External AI"


def _map_risk(value: Any) -> str:
    text = str(value).lower()
    if "red" in text or "high" in text or "fail" in text:
        return "Red"
    if "orange" in text or "medium" in text or "warning" in text:
        return "Orange"
    if "yellow" in text or "low" in text:
        return "Yellow"
    return "Yellow"


def import_external_reports(report_paths: list[Path], output_dir: Path) -> tuple[pd.DataFrame, list[dict[str, Any]]]:
    imported_dir = output_dir / "imported_reports"
    imported_dir.mkdir(parents=True, exist_ok=True)
    rows = []
    issues: list[dict[str, Any]] = []
    for index, source in enumerate([p for p in report_paths if p], start=1):
        target = imported_dir / source.name
        shutil.copy2(source, target)
        extracted_count = 0
        status = "registered"
        note = "PDF registered without OCR."
        try:
            if source.suffix.lower() == ".csv":
                df = pd.read_csv(source)
            elif source.suffix.lower() in {".xlsx", ".xls"}:
                df = pd.read_excel(source)
            else:
                df = None
            if df is not None:
                status = "parsed"
                note = f"Parsed {len(df)} rows."
                risk_col = next((c for c in df.columns if str(c).lower() in {"risk", "risk_level", "severity", "status"}), None)
                evidence_col = next((c for c in df.columns if str(c).lower() in {"evidence", "finding", "message", "issue"}), None)
                for _, row in df.iterrows():
                    if risk_col or evidence_col:
                        extracted_count += 1
                        issues.append(
                            {
                                "issue_id": "",
                                "module": "External AI Image Check",
                                "risk_level": _map_risk(row[risk_col]) if risk_col else "Yellow",
                                "issue_type": "external_report_finding",
                                "file_name": source.name,
                                "sheet_or_panel": "",
                                "sample_or_variable": "",
                                "triggered_rule": "Imported external report",
                                "evidence": str(row[evidence_col]) if evidence_col else "External report row imported.",
                                "recommended_action": "结合外部平台原始报告和未裁剪原图进行人工复核。",
                                "need_human_review": "Yes",
                                "affects_submission": "Review",
                                "review_status": "Pending",
                                "reviewer": "",
                                "review_comment": "",
                            }
                        )
        except Exception as exc:
            status = "failed"
            note = str(exc)
        rows.append(
            {
                "report_id": f"EXTREP{index:03d}",
                "tool_name_guess": guess_tool_name(source),
                "file_name": source.name,
                "file_type": source.suffix.lower().lstrip("."),
                "imported_path": str(target),
                "parse_status": status,
                "extracted_issue_count": extracted_count,
                "note": note,
            }
        )
    imported = pd.DataFrame(rows)
    imported.to_excel(output_dir / "imported_external_reports.xlsx", index=False)
    return imported, issues
