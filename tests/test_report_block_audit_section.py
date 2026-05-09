from pathlib import Path

import pandas as pd
import pytest

from src.core.report_html import generate_html_report
from src.core.report_pdf import generate_pdf_report


def _block_issue_log() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "issue_id": "BLK001",
                "module": "Supplementary Table Block Audit",
                "risk_level": "Red",
                "issue_type": "block_internal_duplicate_columns",
                "file_name": "supplement.xlsx",
                "sheet_or_panel": "Figure 1a",
                "sample_or_variable": "column_3",
                "triggered_rule": "block_internal_duplicate_columns",
                "evidence": "区块 rows 5-8 内列 column_3 与列 column_4 数值完全相同。",
                "recommended_action": "请核对重复列。",
            }
        ]
    )


def test_html_report_contains_block_audit_section(tmp_path: Path):
    issue_log = _block_issue_log()

    output = generate_html_report(
        tmp_path / "report.html",
        "Demo",
        {
            "final_status": "Fail",
            "red_count": 1,
            "orange_count": 0,
            "yellow_count": 0,
            "run_dir": str(tmp_path),
            "block_audit_issue_count": 1,
        },
        pd.DataFrame([{"file_name": "supplement.xlsx", "file_type": "excel"}]),
        pd.DataFrame([{"file_name": "supplement.xlsx", "sheet_name": "Figure 1a"}]),
        issue_log,
        pd.DataFrame(columns=["tool", "status", "message"]),
    )

    html = output.read_text(encoding="utf-8")
    assert "附表区块审计" in html
    assert "data_audit.py" in html
    assert "block_audit_results.xlsx" in html
    assert "supplement.xlsx" in html
    assert "column_3" in html


def test_pdf_report_contains_block_audit_section(tmp_path: Path):
    fitz = pytest.importorskip("fitz")
    issue_log = _block_issue_log()

    output = generate_pdf_report(
        tmp_path / "report.pdf",
        tmp_path / "figures",
        "Demo",
        "upload.zip",
        {
            "final_status": "Fail",
            "red_count": 1,
            "orange_count": 0,
            "yellow_count": 0,
            "app_version": "test",
            "block_audit_issue_count": 1,
        },
        pd.DataFrame([{"file_name": "supplement.xlsx", "file_type": "excel"}]),
        pd.DataFrame(
            [
                {
                    "file_name": "supplement.xlsx",
                    "sheet_name": "Figure 1a",
                    "n_rows": 10,
                    "n_cols": 4,
                    "qc_profile": "generic_numeric",
                    "parse_status": "parsed",
                }
            ]
        ),
        issue_log,
        pd.DataFrame(columns=["tool", "status", "message"]),
    )

    doc = fitz.open(output)
    text = "\n".join(page.get_text() for page in doc)
    assert "附表区块审计" in text
    assert "data_audit.py" in text
    assert "block_audit_results.xlsx" in text
