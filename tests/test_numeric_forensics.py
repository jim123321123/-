import pandas as pd

from src.core.numeric_forensics import run_numeric_forensics


def issue_types(issues):
    return {issue["issue_type"] for issue in issues}


def test_numeric_forensics_detects_required_risk_patterns():
    df = pd.DataFrame(
        {
            "A": [1, 2, 3, 4, 5, 6, 7, 8, 9, 10],
            "B": [2, 4, 6, 8, 10, 12, 14, 16, 18, 20],
            "C": [1, 1, 1, 1, 1, 1, 1, 1, 1, 1],
            "pvalue": [0.01, 1.2, 0.03, 0.04, 0.05, 0, 0.07, 0.08, 0.09, 0.10],
            "qvalue": [0.01, 0.02, -0.1, 0.04, 0.05, 0.06, 0.07, 0.08, 0.09, 0.10],
            "fold_change": [1, 2, 3, 0, 5, 6, 7, 8, 9, 10],
            "FPKM": [1, 2, 3, 4, -1, 6, 7, 8, 9, 10],
            "extreme": [1, 2, 3, 4, 5, 1.79769e308, 7, 8, 9, 10],
            "terminal": [10, 15, 20, 25, 30, 35, 40, 45, 50, 55],
            "Count": [2] * 10,
            "Genes": ["A/B/C"] * 10,
        }
    )
    df.loc[9] = df.loc[8]

    issues = run_numeric_forensics(
        [("file.xlsx", "Sheet1", "enrichment", df)],
        thresholds={},
    )
    types = issue_types(issues)

    assert "exact_duplicate_rows" in types
    assert "near_duplicate_rows" in types
    assert "fixed_ratio_columns" in types
    assert "equal_difference_run" in types
    assert "invalid_p_or_q_value" in types
    assert "invalid_fold_change" in types
    assert "negative_abundance_value" in types
    assert "extreme_or_infinite_value" in types
    assert "terminal_digit_anomaly" in types
    assert "enrichment_count_gene_mismatch" in types
