import pandas as pd

from src.core.numeric_forensics import run_numeric_forensics


def issue_types(issues):
    return {issue["issue_type"] for issue in issues}


def test_numeric_forensics_does_not_flag_sparse_layout_columns_as_missingness():
    df = pd.DataFrame(
        {
            "value_a": [1.1, 1.2, 1.3, 1.4, 1.5],
            "value_b": [2.1, 2.2, 2.3, 2.4, 2.5],
            "column_3": [None, None, None, None, None],
            "P value": [0.01, None, None, None, None],
            "column_5": [None, None, None, None, None],
            "column_6": [None, None, None, None, None],
        }
    )

    issues = run_numeric_forensics([("figure.xlsx", "Figure 1", "generic_numeric", df)])
    types = issue_types(issues)
    assert "high_column_missingness" not in types
    assert "high_row_missingness" not in types


def test_numeric_forensics_downgrades_identical_constant_control_columns():
    df = pd.DataFrame(
        {
            "control_1": [1, 1, 1, 1, 1, 1],
            "control_2": [1, 1, 1, 1, 1, 1],
            "treatment": [1.2, 1.4, 1.1, 1.6, 1.3, 1.5],
        }
    )

    issues = run_numeric_forensics([("figure.xlsx", "Figure 2", "generic_numeric", df)])
    control_issues = [issue for issue in issues if issue["issue_type"] == "duplicate_constant_control_columns"]
    assert control_issues
    assert control_issues[0]["risk_level"] == "Yellow"
    assert "duplicate_numeric_columns" not in issue_types(issues)


def test_numeric_forensics_ignores_zero_difference_constant_runs():
    df = pd.DataFrame({"constant_control": [1, 1, 1, 1, 1, 1, 1], "signal": [1, 3, 2, 4, 3, 5, 4]})

    issues = run_numeric_forensics([("figure.xlsx", "Figure 3", "generic_numeric", df)])
    assert "equal_difference_run" not in issue_types(issues)


def test_exact_duplicate_rows_ignores_text_only_duplicates():
    df = pd.DataFrame(
        {
            "sample": ["A", "A", "B", "C"],
            "group": ["control", "control", "treated", "treated"],
            "value": [1.0, 2.0, 3.0, 4.0],
            "score": [10.0, 20.0, 30.0, 40.0],
        }
    )

    issues = run_numeric_forensics([("text_duplicates.xlsx", "Sheet1", "generic_numeric", df)])
    assert "exact_duplicate_rows" not in issue_types(issues)


def test_numeric_forensics_ignores_expected_enrichment_count_percent_relationships():
    df = pd.DataFrame(
        {
            "Term": [f"GO:{i:07d}~term" for i in range(1, 12)],
            "Count": list(range(1, 12)),
            "%": [value / 250 * 100 for value in range(1, 12)],
            "PValue": [0.001 * value for value in range(1, 12)],
            "Genes": ["A,B"] * 11,
            "List Total": [250] * 11,
            "Pop Hits": list(range(20, 31)),
            "Pop Total": [13528] * 11,
            "Fold Enrichment": [2.0 + value / 100 for value in range(1, 12)],
            "Bonferroni": [0.01 * value for value in range(1, 12)],
            "Benjamini": [0.005 * value for value in range(1, 12)],
            "FDR": [0.002 * value for value in range(1, 12)],
        }
    )

    issues = run_numeric_forensics([("enrichment.xlsx", "GO enrichment analysis", "generic_numeric", df)])
    types = issue_types(issues)
    assert "fixed_ratio_columns" not in types
    assert "high_column_correlation" not in types
    assert "terminal_digit_anomaly" not in types


def test_numeric_forensics_ignores_expected_genomic_coordinate_relationships():
    df = pd.DataFrame(
        {
            "chrom": ["chr1"] * 20,
            "start": list(range(1000, 3000, 100)),
            "end": list(range(1050, 3050, 100)),
        }
    )

    issues = run_numeric_forensics([("coords.xlsx", "ATAC_peak_location", "generic_numeric", df)])
    types = issue_types(issues)
    assert "fixed_ratio_columns" not in types
    assert "high_column_correlation" not in types
    assert "terminal_digit_anomaly" not in types


def test_v2_numeric_rules_have_positive_cases():
    fixed_delta = pd.DataFrame({"A": range(1, 13), "B": [value + 0.3 for value in range(1, 13)]})
    affine = pd.DataFrame({"A": range(1, 13), "B": [2 * value + 1 for value in range(1, 13)]})
    duplicate_decimal = pd.DataFrame({"A": [10 + i + 0.12 for i in range(12)], "B": [100 + i * 3 + 0.12 for i in range(12)]})
    terminal_digit = pd.DataFrame({"biased": [float(f"{i}.15") for i in range(1, 101)]})
    percent_count = pd.DataFrame({"positive": [5, 8, 2, 9], "total": [10, 10, 10, 10], "percent": [50, 70, 20, 95]})
    duplicate_series = pd.DataFrame({"A": [1, 2, 3, 4, 5, 9, 8, 1, 2, 3, 4, 5]})
    arithmetic = pd.DataFrame({"signal": list(range(0, 30, 2))})
    repeated = pd.DataFrame({"signal": [7] * 16 + list(range(4))})
    grim = pd.DataFrame({"mean": [1.23, 2.5, 3.1], "n": [10, 12, 14]})
    precision = pd.DataFrame({"value": [float(f"{i}.123") for i in range(1, 40)]})

    parsed = [
        ("fixed_delta.xlsx", "Sheet1", "generic_numeric", fixed_delta),
        ("affine_relation.xlsx", "Sheet1", "generic_numeric", affine),
        ("duplicate_decimal.xlsx", "Sheet1", "generic_numeric", duplicate_decimal),
        ("terminal_digit_bias.xlsx", "Sheet1", "generic_numeric", terminal_digit),
        ("percent_count_inconsistent.xlsx", "Sheet1", "generic_numeric", percent_count),
        ("duplicate_series.xlsx", "Sheet1", "generic_numeric", duplicate_series),
        ("arithmetic_sequence.xlsx", "Sheet1", "generic_numeric", arithmetic),
        ("repeated_value.xlsx", "Sheet1", "generic_numeric", repeated),
        ("grim_like.xlsx", "Sheet1", "generic_numeric", grim),
        ("decimal_precision.xlsx", "Sheet1", "generic_numeric", precision),
    ]

    issues = run_numeric_forensics(parsed, thresholds={"numeric": {"terminal_digit_min_n": 30}})
    rule_ids = {issue.get("rule_id") for issue in issues}

    assert "N001" in rule_ids
    assert "N002" in rule_ids
    assert "N003" in rule_ids
    assert "N004" in rule_ids
    assert "N005" in rule_ids
    assert "N006" in rule_ids
    assert "N007" in rule_ids
    assert "N008" in rule_ids
    assert "N009" in rule_ids
    assert "N011" in rule_ids


def test_cross_file_reuse_detector_positive_case():
    left = pd.DataFrame({"A": [1, 2, 3, 4, 5, 9], "B": [2, 3, 4, 5, 6, 10]})
    right = pd.DataFrame({"A": [1, 2, 3, 4, 5, 11], "B": [2, 3, 4, 5, 6, 12]})

    issues = run_numeric_forensics(
        [
            ("control.xlsx", "WT", "generic_numeric", left),
            ("treat.xlsx", "KO", "generic_numeric", right),
        ]
    )

    assert any(issue.get("rule_id") == "N010" for issue in issues)


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
        thresholds={"numeric": {"terminal_digit_min_n": 10}},
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


def test_large_sheet_skips_quadratic_near_duplicate_scan():
    df = pd.DataFrame(
        {
            "A": range(2501),
            "B": range(2501),
            "C": range(2501),
            "D": range(2501),
            "E": range(2501),
            "F": range(2501),
        }
    )

    issues = run_numeric_forensics(
        [("large.xlsx", "BigSheet", "generic_numeric", df)],
        thresholds={"duplicate": {"max_near_duplicate_rows": 2000}},
    )

    assert "near_duplicate_scan_skipped_large_sheet" in issue_types(issues)
