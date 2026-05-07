from __future__ import annotations

from pathlib import Path
from typing import Any

import matplotlib.pyplot as plt
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Image, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


DISCLAIMER = (
    "本报告用于投稿前数据质量和研究诚信风险筛查。自动化结果不能直接定性研究不端。"
    "所有 Red 和 Orange 问题均需结合原始记录、实验记录本、仪器导出文件、未裁剪原始图片和人工复核确认。"
    "本报告中的外部AI图片检查结果依赖用户提供的 API 服务或用户导入的外部工具报告。"
)


def _register_font() -> str:
    candidates = [
        Path("C:/Windows/Fonts/msyh.ttc"),
        Path("C:/Windows/Fonts/simsun.ttc"),
        Path("C:/Windows/Fonts/simhei.ttf"),
    ]
    for font in candidates:
        if font.exists():
            pdfmetrics.registerFont(TTFont("CNFont", str(font)))
            return "CNFont"
    return "Helvetica"


def _plot_counts(counts: pd.Series, title: str, output_path: Path) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    fig, ax = plt.subplots(figsize=(5, 3))
    if counts.empty:
        counts = pd.Series({"None": 0})
    counts.plot(kind="bar", ax=ax, color="#4f7cac")
    ax.set_title(title)
    ax.set_ylabel("Count")
    fig.tight_layout()
    fig.savefig(output_path, dpi=160)
    plt.close(fig)
    return output_path


def _small_table(df: pd.DataFrame, columns: list[str], max_rows: int = 12):
    if df is None or df.empty:
        return [["No records"]]
    data = [columns]
    for _, row in df.head(max_rows).iterrows():
        data.append([str(row.get(col, ""))[:80] for col in columns])
    return data


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
    styles = getSampleStyleSheet()
    for style in styles.byName.values():
        style.fontName = font

    file_type_fig = _plot_counts(manifest["file_type"].value_counts() if not manifest.empty else pd.Series(dtype=int), "File type counts", figures_dir / "file_type_counts.png")
    risk_fig = _plot_counts(issue_log["risk_level"].value_counts() if not issue_log.empty else pd.Series(dtype=int), "Risk level counts", figures_dir / "risk_level_counts.png")
    module_fig = _plot_counts(issue_log["module"].value_counts() if not issue_log.empty else pd.Series(dtype=int), "Risk by module", figures_dir / "risk_by_module.png")

    story = [
        Paragraph("投稿前AI数据真实性与合理性质控报告", styles["Title"]),
        Spacer(1, 0.4 * cm),
        Paragraph(f"项目名称：{project_name}", styles["Normal"]),
        Paragraph(f"上传压缩包：{zip_name}", styles["Normal"]),
        Paragraph(f"软件版本：{summary.get('app_version', '')}", styles["Normal"]),
        Paragraph(f"检查状态：{summary.get('final_status', 'Pass')}", styles["Heading2"]),
        Spacer(1, 0.5 * cm),
        Paragraph("检查范围与方法", styles["Heading1"]),
        Paragraph("本软件执行本地确定性规则检查，并整合外部AI工具状态或导入报告。外部工具不内置在本软件中。", styles["Normal"]),
        Paragraph("文件完整性检查", styles["Heading1"]),
        Image(str(file_type_fig), width=14 * cm, height=8 * cm),
        Paragraph("表格数据结构识别", styles["Heading1"]),
    ]
    story.append(Table(_small_table(sheet_inventory, ["file_name", "sheet_name", "n_rows", "n_cols", "qc_profile", "parse_status"]), repeatRows=1))
    story.extend(
        [
            Paragraph("数值型原始数据真实性与合理性检查", styles["Heading1"]),
            Paragraph("已检查重复行、近似重复、重复列、高相关、固定倍数、等差、尾数、p/q值范围、极端值、缺失率和表类型专项规则。", styles["Normal"]),
            Paragraph("图片完整性AI检查", styles["Heading1"]),
            Paragraph("本地模块生成 image_check_package.zip 和 image_inventory.csv；Proofig / Imagetwin 需通过官方 API 或手动上传检查包完成。", styles["Normal"]),
            Paragraph("外部AI工具调用状态", styles["Heading1"]),
            Table(_small_table(external_status, ["tool", "status", "message"]), repeatRows=1),
            Paragraph("风险分级汇总", styles["Heading1"]),
            Image(str(risk_fig), width=14 * cm, height=8 * cm),
            Image(str(module_fig), width=14 * cm, height=8 * cm),
            Paragraph("问题清单（前50条）", styles["Heading1"]),
            Table(_small_table(issue_log, ["issue_id", "risk_level", "module", "issue_type", "evidence"], 50), repeatRows=1),
            PageBreak(),
            Paragraph("人工复核建议", styles["Heading1"]),
            Paragraph("Red 问题投稿前必须解决；Orange 问题需回查原始记录；Yellow 问题建议记录解释。图片AI标记需结合未裁剪原图和外部平台原始报告复核。", styles["Normal"]),
            Paragraph("附录与免责声明", styles["Heading1"]),
            Paragraph(DISCLAIMER, styles["Normal"]),
        ]
    )

    for item in story:
        if isinstance(item, Table):
            item.setStyle(
                TableStyle(
                    [
                        ("FONTNAME", (0, 0), (-1, -1), font),
                        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                        ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ]
                )
            )
    doc = SimpleDocTemplate(str(output_path), pagesize=A4, rightMargin=1.2 * cm, leftMargin=1.2 * cm, topMargin=1.2 * cm, bottomMargin=1.2 * cm)
    doc.build(story)
    return output_path
