import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.platypus import Paragraph

from src.core.report_pdf import _wrapped_table


def _contains_paragraph(value):
    if isinstance(value, Paragraph):
        return True
    if isinstance(value, (tuple, list)):
        return any(isinstance(item, Paragraph) for item in value)
    return False


def test_wrapped_table_fits_available_page_width():
    df = pd.DataFrame(
        [
            {
                "风险等级": "Red（必须处理）",
                "文件": "Supplementary_Figure_with_a_very_very_long_file_name_that_would_overflow.xlsx",
                "表格/页面": "A very long sheet name with many words and identifiers",
                "具体位置": "column_123456789",
                "发现的问题": "这是一段很长的问题说明，用于模拟科研人员报告中可能出现的详细证据和解释，必须在单元格内自动换行。",
                "建议怎么做": "请回查原始记录、实验记录本、仪器导出文件和数据处理脚本，并记录复核结论。",
            }
        ]
    )
    available_width = A4[0] - 2.4 * cm

    table = _wrapped_table(
        df,
        ["风险等级", "文件", "表格/页面", "具体位置", "发现的问题", "建议怎么做"],
        available_width,
        max_rows=1,
    )
    width, _ = table.wrap(available_width, A4[1])

    assert width <= available_width
    assert _contains_paragraph(table._cellvalues[1][1])
    assert _contains_paragraph(table._cellvalues[1][4])
