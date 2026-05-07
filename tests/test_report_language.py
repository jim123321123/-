import pandas as pd

from src.core.report_language import build_plain_issue_table, explain_final_status


def test_plain_issue_table_points_to_file_sheet_and_column():
    issue_log = pd.DataFrame(
        [
            {
                "risk_level": "Red",
                "issue_type": "exact_duplicate_rows",
                "file_name": "Supplementary Figure 1.xlsx",
                "sheet_or_panel": "Pathway Heatmap",
                "sample_or_variable": "2, 3, 4",
                "evidence": "Rows 2, 3, 4 are exact duplicates across 38 comparable columns.",
                "recommended_action": "回查原始记录。",
            }
        ]
    )

    table = build_plain_issue_table(issue_log)

    assert list(table.columns) == ["风险等级", "文件", "表格/页面", "具体位置", "发现的问题", "建议怎么做"]
    assert table.iloc[0]["风险等级"] == "Red（必须处理）"
    assert table.iloc[0]["文件"] == "Supplementary Figure 1.xlsx"
    assert table.iloc[0]["表格/页面"] == "Pathway Heatmap"
    assert "发现完全重复的数据行" in table.iloc[0]["发现的问题"]
    assert "第 2、3、4 行" in table.iloc[0]["发现的问题"]


def test_final_status_explanation_is_plain_chinese():
    text = explain_final_status("Fail", 3, 2, 1)

    assert "结论：暂不建议直接投稿" in text
    assert "3 个必须优先处理的问题" in text
