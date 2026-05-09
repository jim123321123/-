import pandas as pd

from src.gui.rule_detail import highlight_targets, table_index


def _target_for(issue, df):
    tables = table_index([("file.xlsx", "Sheet1", "generic_numeric", df)])
    return highlight_targets(pd.Series({"file_name": "file.xlsx", "sheet_or_panel": "Sheet1", **issue}), tables)[0]


def test_column_level_rule_without_rows_does_not_expand_to_full_columns():
    df = pd.DataFrame({"A": [1, 2, 3], "B": [2, 4, 6]})

    target = _target_for(
        {
            "issue_type": "fixed_ratio_columns",
            "details": {"left_column": "A", "right_column": "B"},
            "evidence": "Columns A and B have a fixed ratio.",
        },
        df,
    )

    assert target.columns == {"A", "B"}
    assert target.cells == set()
    assert target.scope == "column"


def test_exact_duplicate_rows_highlight_only_numeric_cells():
    df = pd.DataFrame({"sample": ["A", "A"], "value": [1.0, 1.0], "score": [2.0, 2.0]})

    target = _target_for(
        {
            "issue_type": "exact_duplicate_rows",
            "sample_or_variable": "2, 3",
            "evidence": "Rows 2, 3 are exact duplicates across 2 comparable numeric columns.",
        },
        df,
    )

    assert target.cells == {(2, "value"), (2, "score"), (3, "value"), (3, "score")}
    assert (2, "sample") not in target.cells
    assert target.scope == "row"


def test_terminal_digit_rule_highlights_only_matching_cells():
    df = pd.DataFrame({"value": [1.15, 2.16, 3.15, 4.17]})

    target = _target_for(
        {"issue_type": "terminal_digit_anomaly", "details": {"column": "value", "dominant_digit": 5}},
        df,
    )

    assert target.cells == {(2, "value"), (4, "value")}


def test_decimal_precision_rule_highlights_only_precision_outliers():
    df = pd.DataFrame({"value": ["1.1", "2.22", "3.3"]})

    target = _target_for(
        {"issue_type": "decimal_precision_detector", "details": {"column": "value", "dominant_precision": 1}},
        df,
    )

    assert target.cells == {(3, "value")}


def test_arithmetic_sequence_rule_highlights_only_matching_sequence_cells():
    df = pd.DataFrame({"signal": [1, 3, 5, 9]})

    target = _target_for(
        {"issue_type": "arithmetic_sequence_detector", "details": {"column": "signal", "diff": 2}},
        df,
    )

    assert target.cells == {(2, "signal"), (3, "signal"), (4, "signal")}
