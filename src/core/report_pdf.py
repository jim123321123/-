from __future__ import annotations

from pathlib import Path
from typing import Any

import matplotlib.pyplot as plt
import pandas as pd
from matplotlib.font_manager import FontProperties
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Image, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from .report_exports import summarize_v2_findings
from .report_language import build_plain_issue_table, build_priority_review_text, explain_final_status
from .report_summary import generate_raw_data_overview


DISCLAIMER = (
    "本报告用于投稿前数据质量和研究诚信风险筛查。自动化结果不能直接定性研究不端。"
    "所有 Red 和 Orange 项都需要结合原始记录、实验记录本、仪器导出文件、未裁剪原始图片和人工复核确认；Yellow 项建议记录解释。"
)


def _chinese_font_path() -> Path | None:
    for font in [
        Path("C:/Windows/Fonts/msyh.ttc"),
        Path("C:/Windows/Fonts/simsun.ttc"),
        Path("C:/Windows/Fonts/simhei.ttf"),
    ]:
        if font.exists():
            return font
    return None


def _register_font() -> str:
    font = _chinese_font_path()
    if font is not None:
        pdfmetrics.registerFont(TTFont("CNFont", str(font)))
        return "CNFont"
    return "Helvetica"


def _plot_counts(counts: pd.Series, title: str, output_path: Path) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    fig, ax = plt.subplots(figsize=(5, 3))
    if counts.empty:
        counts = pd.Series({"无": 0})
    font_path = _chinese_font_path()
    font_prop = FontProperties(fname=str(font_path)) if font_path is not None else None
    counts.plot(kind="bar", ax=ax, color="#4f7cac")
    ax.set_title(title, fontproperties=font_prop)
    ax.set_ylabel("数量", fontproperties=font_prop)
    for label in ax.get_xticklabels() + ax.get_yticklabels():
        label.set_fontproperties(font_prop)
    fig.tight_layout()
    fig.savefig(output_path, dpi=160)
    plt.close(fig)
    return output_path


def _cell_style(font: str, font_size: int = 6) -> ParagraphStyle:
    return ParagraphStyle(
        name=f"Cell-{font_size}",
        fontName=font,
        fontSize=font_size,
        leading=font_size + 2,
        wordWrap="CJK",
        splitLongWords=True,
        spaceAfter=0,
        spaceBefore=0,
    )


def _table_col_widths(columns: list[str], available_width: float) -> list[float]:
    presets = {
        ("风险等级", "文件", "表格/页面", "具体位置", "发现的问题", "建议怎么做"): [0.10, 0.15, 0.13, 0.11, 0.26, 0.25],
        ("规则编号", "规则名称", "数量", "风险等级"): [0.16, 0.52, 0.12, 0.20],
        ("规则", "风险", "文件", "表格/页面", "位置", "问题"): [0.20, 0.11, 0.17, 0.14, 0.13, 0.25],
        ("文件", "表格/页面", "变量", "末位数字计数", "p值", "占比最高数字"): [0.17, 0.15, 0.15, 0.30, 0.12, 0.11],
        ("文件", "表格/页面", "位置", "不匹配行", "说明"): [0.18, 0.15, 0.15, 0.18, 0.34],
        ("规则", "风险", "图片1", "图片2", "距离", "说明"): [0.10, 0.11, 0.24, 0.24, 0.09, 0.22],
        ("file_name", "sheet_name", "n_rows", "n_cols", "qc_profile", "parse_status"): [0.25, 0.24, 0.08, 0.08, 0.22, 0.13],
        ("tool", "status", "message"): [0.22, 0.18, 0.60],
    }
    ratios = presets.get(tuple(columns), [1 / len(columns)] * len(columns))
    return [available_width * 0.999 * ratio for ratio in ratios]


def _wrapped_table(
    df: pd.DataFrame,
    columns: list[str],
    available_width: float,
    max_rows: int = 12,
    font: str | None = None,
    font_size: int = 6,
) -> Table:
    font_name = font or _register_font()
    style = _cell_style(font_name, font_size)
    header_style = _cell_style(font_name, font_size)
    if df is None or df.empty:
        data = [[Paragraph("无记录", style)]]
        widths = [available_width]
    else:
        source = df.head(max_rows)
        data = [[Paragraph(str(column), header_style) for column in columns]]
        for _, row in source.iterrows():
            data.append(
                [
                    Paragraph("" if pd.isna(row.get(column, "")) else str(row.get(column, ""))[:800], style)
                    for column in columns
                ]
            )
        widths = _table_col_widths(columns, available_width)
    table = Table(data, colWidths=widths, repeatRows=1, hAlign="LEFT", splitByRow=True)
    table.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, -1), font_name),
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e5edf5")),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#cbd5e1")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 3),
                ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]
        )
    )
    return table


def _paragraph_block(text: str, style: ParagraphStyle) -> list[Paragraph]:
    return [Paragraph(part, style) for part in text.splitlines() if part.strip()]


def _block_audit_issues(issue_log: pd.DataFrame) -> pd.DataFrame:
    if issue_log is None or issue_log.empty or "module" not in issue_log.columns:
        return pd.DataFrame()
    return issue_log[issue_log["module"] == "Supplementary Table Block Audit"].copy()


def _block_audit_summary(issue_log: pd.DataFrame, block_count: int) -> str:
    block_issues = _block_audit_issues(issue_log)
    if block_issues.empty:
        return "本次已运行从 data_audit.py 整合而来的附表区块审计，未发现需要单独列出的区块级问题。"
    red = int((block_issues["risk_level"] == "Red").sum()) if "risk_level" in block_issues else 0
    orange = int((block_issues["risk_level"] == "Orange").sum()) if "risk_level" in block_issues else 0
    yellow = int((block_issues["risk_level"] == "Yellow").sum()) if "risk_level" in block_issues else 0
    return (
        f"附表区块审计共发现 {len(block_issues)} 条信号，其中 Red {red} 条、Orange {orange} 条、Yellow {yellow} 条；"
        f"主流程记录的问题数为 {block_count}。完整明细见 block_audit_results.xlsx。"
    )


def generate_pdf_report(
    output_path: Path,
    figures_dir: Path,
    project_name: str,
    zip_name: str,
    summary: dict[str, Any],
    manifest: pd.DataFrame,
    sheet_inventory: pd.DataFrame,
    issue_log: pd.DataFrame,
    external_status: pd.DataFrame,
) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    font = _register_font()
    available_width = A4[0] - 2.4 * cm
    styles = getSampleStyleSheet()
    for style in styles.byName.values():
        style.fontName = font
        style.wordWrap = "CJK"

    file_type_fig = _plot_counts(
        manifest["file_type"].value_counts() if not manifest.empty else pd.Series(dtype=int),
        "文件类型统计",
        figures_dir / "file_type_counts.png",
    )
    risk_fig = _plot_counts(
        issue_log["risk_level"].value_counts() if not issue_log.empty else pd.Series(dtype=int),
        "风险等级统计",
        figures_dir / "risk_level_counts.png",
    )
    module_fig = _plot_counts(
        issue_log["module"].value_counts() if not issue_log.empty else pd.Series(dtype=int),
        "按模块统计风险",
        figures_dir / "risk_by_module.png",
    )

    overview = generate_raw_data_overview(manifest, sheet_inventory, issue_log, external_status)
    plain_issues = build_plain_issue_table(issue_log, limit=50)
    block_issues = _block_audit_issues(issue_log)
    plain_block_issues = build_plain_issue_table(block_issues, limit=30)
    block_count = int(summary.get("block_audit_issue_count", len(block_issues)))
    status_text = explain_final_status(
        summary.get("final_status", "Pass"),
        int(summary.get("red_count", 0)),
        int(summary.get("orange_count", 0)),
        int(summary.get("yellow_count", 0)),
    )
    priority_text = build_priority_review_text(issue_log)
    v2 = summarize_v2_findings(issue_log)

    story = [
        Paragraph("投稿前 AI 数据质量控制报告", styles["Title"]),
        Spacer(1, 0.35 * cm),
        Paragraph(f"项目名称：{project_name}", styles["Normal"]),
        Paragraph(f"上传压缩包：{zip_name}", styles["Normal"]),
        Paragraph(f"软件版本：{summary.get('app_version', '')}", styles["Normal"]),
        Paragraph(f"检查状态：{summary.get('final_status', 'Pass')}", styles["Heading2"]),
        Spacer(1, 0.35 * cm),
        Paragraph("先看结论", styles["Heading1"]),
        Paragraph(status_text, styles["Normal"]),
        Paragraph(priority_text, styles["Normal"]),
        Paragraph("报告中的风险项是复核提示，不是对研究行为的定性判断。", styles["Normal"]),
        Spacer(1, 0.25 * cm),
        Paragraph("压缩包原始数据整体情况", styles["Heading1"]),
        *_paragraph_block(overview, styles["Normal"]),
        Spacer(1, 0.25 * cm),
        Paragraph("文件组成", styles["Heading1"]),
        Image(str(file_type_fig), width=14 * cm, height=8 * cm),
        Paragraph("表格数据结构识别", styles["Heading1"]),
        _wrapped_table(
            sheet_inventory,
            ["file_name", "sheet_name", "n_rows", "n_cols", "qc_profile", "parse_status"],
            available_width,
            max_rows=12,
            font=font,
        ),
        Paragraph("附表区块审计（data_audit.py 整合结果）", styles["Heading1"]),
        Paragraph(_block_audit_summary(issue_log, block_count), styles["Normal"]),
        _wrapped_table(
            plain_block_issues,
            ["风险等级", "文件", "表格/页面", "具体位置", "发现的问题", "建议怎么做"],
            available_width,
            max_rows=30,
            font=font,
            font_size=5,
        ),
        Paragraph("新版规则汇总（N001-N011、I001-I002）", styles["Heading1"]),
        _wrapped_table(v2["rule_counts"], ["规则编号", "规则名称", "数量", "风险等级"], available_width, max_rows=80, font=font),
        Paragraph("优先复核清单", styles["Heading1"]),
        _wrapped_table(v2["top_findings"], ["规则", "风险", "文件", "表格/页面", "位置", "问题"], available_width, max_rows=20, font=font, font_size=5),
        Paragraph("N004 末位数字分布", styles["Heading1"]),
        _wrapped_table(v2["n004_digits"], ["文件", "表格/页面", "变量", "末位数字计数", "p值", "占比最高数字"], available_width, max_rows=20, font=font, font_size=5),
        Paragraph("N005 百分比/计数不一致", styles["Heading1"]),
        _wrapped_table(v2["n005_mismatches"], ["文件", "表格/页面", "位置", "不匹配行", "说明"], available_width, max_rows=20, font=font, font_size=5),
        Paragraph("I001/I002 图片重复或相似图片", styles["Heading1"]),
        _wrapped_table(v2["image_pairs"], ["规则", "风险", "图片1", "图片2", "距离", "说明"], available_width, max_rows=20, font=font, font_size=5),
        Paragraph("外部 AI 工具状态", styles["Heading1"]),
        _wrapped_table(external_status, ["tool", "status", "message"], available_width, max_rows=12, font=font),
        Paragraph("风险分级汇总", styles["Heading1"]),
        Image(str(risk_fig), width=14 * cm, height=8 * cm),
        Image(str(module_fig), width=14 * cm, height=8 * cm),
        Paragraph("需要人工复核的问题清单（前50条）", styles["Heading1"]),
        _wrapped_table(
            plain_issues,
            ["风险等级", "文件", "表格/页面", "具体位置", "发现的问题", "建议怎么做"],
            available_width,
            max_rows=50,
            font=font,
            font_size=5,
        ),
        PageBreak(),
        Paragraph("人工复核建议", styles["Heading1"]),
        Paragraph("Red 问题投稿前必须解决；Orange 问题需要回查原始记录；Yellow 问题建议记录解释。图片类信号需要结合未裁剪原图和外部平台原始报告复核。", styles["Normal"]),
        Paragraph("附录与免责声明", styles["Heading1"]),
        Paragraph(DISCLAIMER, styles["Normal"]),
    ]

    doc = SimpleDocTemplate(
        str(output_path),
        pagesize=A4,
        rightMargin=1.2 * cm,
        leftMargin=1.2 * cm,
        topMargin=1.2 * cm,
        bottomMargin=1.2 * cm,
    )
    doc.build(story)
    return output_path
