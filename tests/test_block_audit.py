import pandas as pd

from src.core.block_audit import run_block_audit


def issue_types(issues):
    return {issue["issue_type"] for issue in issues}


def test_block_audit_detects_duplicate_ratio_diff_and_internal_repeats():
    df = pd.DataFrame(
        {
            "label": ["r1", "r2", "r3", "r4", "r5", "r6"],
            "A1": [1, 2, 3, 4, 5, 5],
            "A2": [10, 20, 30, 40, 50, 50],
            "gap": ["", "", "", "", "", ""],
            "B1": [1, 2, 3, 4, 5, 5],
            "B2": [10, 20, 30, 40, 50, 50],
            "gap2": ["", "", "", "", "", ""],
            "C1": [2, 4, 6, 8, 10, 10],
            "C2": [20, 40, 60, 80, 100, 100],
            "gap3": ["", "", "", "", "", ""],
            "D1": [6, 7, 8, 9, 10, 10],
            "D2": [15, 25, 35, 45, 55, 55],
            "gap4": ["", "", "", "", "", ""],
            "E1": [1, 2, 3, 4, 5, 6],
            "E2": [1, 2, 3, 4, 5, 6],
        }
    )

    issues = run_block_audit([("file.xlsx", "Sheet1", "generic_numeric", df)], max_sheet_cells=10_000)
    types = issue_types(issues)

    assert "block_identical" in types
    assert "block_fixed_ratio" in types
    assert "block_fixed_difference" in types
    assert "block_internal_duplicate_rows" in types
    assert "block_internal_duplicate_columns" in types


def test_block_audit_skips_very_large_sheet_with_clear_issue():
    df = pd.DataFrame({"A": range(1001), "B": range(1001), "C": range(1001)})

    issues = run_block_audit([("large.xlsx", "BigSheet", "generic_numeric", df)], max_sheet_cells=1000)

    assert any(issue["issue_type"] == "block_audit_skipped_large_sheet" for issue in issues)


def test_block_audit_downgrades_repeated_constant_control_columns():
    df = pd.DataFrame(
        {
            "group": ["P16", "IL6", "IL8", "Lamin B1"],
            "control_1": [1, 1, 1, 1],
            "control_2": [1, 1, 1, 1],
            "treatment_1": [0.2, 0.4, 0.6, 0.8],
            "treatment_2": [0.3, 0.5, 0.7, 0.9],
        }
    )

    issues = run_block_audit([("figure.xlsx", "Figure", "generic_numeric", df)])

    constant_controls = [issue for issue in issues if issue["issue_type"] == "block_duplicate_constant_control_columns"]
    assert constant_controls
    assert constant_controls[0]["risk_level"] == "Yellow"
    assert not any(issue["issue_type"] == "block_internal_duplicate_columns" for issue in issues)
