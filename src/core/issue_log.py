from __future__ import annotations

from pathlib import Path
from typing import Any

import pandas as pd

from .report_exports import serialize_issue_log


ISSUE_LOG_COLUMNS = [
    "issue_id",
    "module",
    "rule_id",
    "severity",
    "risk_level",
    "issue_type",
    "file_name",
    "sheet_or_panel",
    "sample_or_variable",
    "triggered_rule",
    "evidence",
    "recommended_action",
    "details",
    "need_human_review",
    "affects_submission",
    "review_status",
    "reviewer",
    "review_comment",
]


def numeric_issues_to_log(issues: list[dict[str, Any]]) -> list[dict[str, Any]]:
    rows = []
    for issue in issues:
        rows.append(
            {
                "issue_id": "",
                "module": issue.get("module", "Numeric Forensics"),
                "rule_id": issue.get("rule_id", ""),
                "severity": issue.get("severity", ""),
                "risk_level": issue.get("risk_level", "Yellow"),
                "issue_type": issue.get("issue_type", ""),
                "file_name": issue.get("file_name", ""),
                "sheet_or_panel": issue.get("sheet_name", ""),
                "sample_or_variable": issue.get("column_name") or issue.get("row_index", ""),
                "triggered_rule": issue.get("issue_type", ""),
                "evidence": issue.get("evidence", ""),
                "recommended_action": issue.get("recommended_action", ""),
                "details": issue.get("details", ""),
                "need_human_review": issue.get("need_human_review", "Recommended"),
                "affects_submission": issue.get("affects_submission", "Review"),
                "review_status": "Pending",
                "reviewer": "",
                "review_comment": "",
            }
        )
    return rows


def build_issue_log(
    numeric_issues: list[dict[str, Any]],
    table_inventory: pd.DataFrame | None = None,
    external_issues: list[dict[str, Any]] | None = None,
    external_status: pd.DataFrame | None = None,
) -> pd.DataFrame:
    rows = numeric_issues_to_log(numeric_issues)
    if table_inventory is not None and not table_inventory.empty:
        failed = table_inventory[table_inventory["parse_status"] == "failed"]
        for _, row in failed.iterrows():
            rows.append(
                {
                    "issue_id": "",
                    "module": "Table Parser",
                    "rule_id": "TAB001",
                    "severity": "MEDIUM",
                    "risk_level": "Orange",
                    "issue_type": "table_parse_failed",
                    "file_name": row.get("file_name", ""),
                    "sheet_or_panel": row.get("sheet_name", ""),
                    "sample_or_variable": "",
                    "triggered_rule": "table_parser",
                    "evidence": row.get("parse_warning", ""),
                    "recommended_action": "检查表格文件是否损坏、加密或格式不受支持。",
                    "details": "",
                    "need_human_review": "Yes",
                    "affects_submission": "Review",
                    "review_status": "Pending",
                    "reviewer": "",
                    "review_comment": "",
                }
            )
    rows.extend(external_issues or [])
    if external_status is not None and not external_status.empty:
        for _, row in external_status.iterrows():
            if row.get("status") in {"manual_required", "failed"}:
                rows.append(
                    {
                        "issue_id": "",
                        "module": "External AI",
                        "rule_id": "EXT001",
                        "severity": "LOW" if row.get("status") == "manual_required" else "MEDIUM",
                        "risk_level": "Yellow" if row.get("status") == "manual_required" else "Orange",
                        "issue_type": f"external_ai_{row.get('status')}",
                        "file_name": "",
                        "sheet_or_panel": "",
                        "sample_or_variable": row.get("tool", ""),
                        "triggered_rule": "external_ai_status",
                        "evidence": row.get("message", ""),
                        "recommended_action": "如需外部AI筛查，请配置官方 endpoint 或手动上传检查包并导入报告。",
                        "details": "",
                        "need_human_review": "Recommended",
                        "affects_submission": "Review",
                        "review_status": "Pending",
                        "reviewer": "",
                        "review_comment": "",
                    }
                )
    for index, row in enumerate(rows, start=1):
        prefix = {
            "Numeric Forensics": "NUM",
            "Supplementary Table Block Audit": "BLK",
            "Table Parser": "TAB",
            "External AI": "EXT",
            "External AI Image Check": "EXT",
            "Image Forensics": "IMG",
        }.get(row.get("module", ""), "QC")
        row["issue_id"] = row.get("issue_id") or f"{prefix}{index:03d}"
    return pd.DataFrame(rows, columns=ISSUE_LOG_COLUMNS)


def final_status(issue_log: pd.DataFrame) -> str:
    if issue_log.empty:
        return "Pass"
    counts = issue_log["risk_level"].value_counts()
    if counts.get("Red", 0) > 0:
        return "Fail"
    if counts.get("Orange", 0) > 0:
        return "Conditional Fail"
    if counts.get("Yellow", 0) > 0:
        return "Conditional Pass"
    return "Pass"


def write_issue_log(issue_log: pd.DataFrame, output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    serialize_issue_log(issue_log).to_excel(output_path, index=False)
