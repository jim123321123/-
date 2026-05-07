from __future__ import annotations

from html import escape
from pathlib import Path
from typing import Any

import pandas as pd

from .report_language import build_plain_issue_table, build_priority_review_text, explain_final_status
from .report_summary import generate_raw_data_overview


DISCLAIMER = (
    "本报告用于投稿前数据质量和研究诚信风险筛查。自动化结果不能直接定性研究不端。"
    "所有 Red 和 Orange 问题均需结合原始记录、实验记录本、仪器导出文件、未裁剪原始图片和人工复核确认。"
    "本报告中的外部AI图片检查结果依赖用户提供的 API 服务或用户导入的外部工具报告。"
)


def _table(df: pd.DataFrame, limit: int = 100) -> str:
    if df is None or df.empty:
        return "<p>No records.</p>"
    return df.head(limit).to_html(index=False, escape=True)


def _paragraphs(text: str) -> str:
    return "\n".join(f"  <p>{escape(part)}</p>" for part in text.splitlines() if part.strip())


def generate_html_report(
    output_path: Path,
    project_name: str,
    summary: dict[str, Any],
    manifest: pd.DataFrame,
    sheet_inventory: pd.DataFrame,
    issue_log: pd.DataFrame,
    external_status: pd.DataFrame,
) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    overview = generate_raw_data_overview(manifest, sheet_inventory, issue_log, external_status)
    plain_issues = build_plain_issue_table(issue_log, limit=100)
    status_text = explain_final_status(
        summary.get("final_status", "Pass"),
        int(summary.get("red_count", 0)),
        int(summary.get("orange_count", 0)),
        int(summary.get("yellow_count", 0)),
    )
    priority_text = build_priority_review_text(issue_log)
    html = f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <title>投稿前AI数据真实性与合理性质控报告</title>
  <style>
    body {{ font-family: "Microsoft YaHei", Arial, sans-serif; margin: 32px; color: #222; }}
    h1, h2 {{ color: #17324d; }}
    table {{ border-collapse: collapse; width: 100%; margin: 12px 0 24px; font-size: 12px; }}
    th, td {{ border: 1px solid #ddd; padding: 6px; text-align: left; vertical-align: top; }}
    th {{ background: #f2f5f7; }}
    .summary {{ display: grid; grid-template-columns: repeat(4, minmax(120px, 1fr)); gap: 10px; }}
    .item {{ border: 1px solid #ddd; padding: 10px; border-radius: 6px; }}
    .overview {{ background: #f7fafc; border-left: 4px solid #4f7cac; padding: 12px 16px; margin: 16px 0 28px; }}
    .conclusion {{ background: #fff8e5; border-left: 4px solid #d9822b; padding: 12px 16px; margin: 16px 0 28px; }}
  </style>
</head>
<body>
  <h1>投稿前AI数据真实性与合理性质控报告</h1>
  <div class="summary">
    <div class="item"><strong>项目</strong><br>{escape(project_name)}</div>
    <div class="item"><strong>状态</strong><br>{escape(summary.get("final_status", "Pass"))}</div>
    <div class="item"><strong>Red</strong><br>{summary.get("red_count", 0)}</div>
    <div class="item"><strong>Orange</strong><br>{summary.get("orange_count", 0)}</div>
    <div class="item"><strong>Yellow</strong><br>{summary.get("yellow_count", 0)}</div>
    <div class="item"><strong>文件数</strong><br>{len(manifest)}</div>
    <div class="item"><strong>Sheet 数</strong><br>{len(sheet_inventory)}</div>
    <div class="item"><strong>输出目录</strong><br>{escape(str(summary.get("run_dir", "")))}</div>
  </div>
  <h2>先看结论</h2>
  <div class="conclusion">
    <p>{escape(status_text)}</p>
    <p>{escape(priority_text)}</p>
    <p>报告中的 Red 和 Orange 不是“定罪结论”，而是提醒课题组必须回到原始记录逐项确认。</p>
  </div>
  <h2>压缩包原始数据整体情况</h2>
  <div class="overview">
{_paragraphs(overview)}
  </div>
  <h2>外部AI状态</h2>
  {_table(external_status)}
  <h2>需要人工复核的问题清单（前100条）</h2>
  <p>下面的表格已经把技术规则翻译成中文。请优先看“文件”“表格/页面”“具体位置”和“建议怎么做”。</p>
  {_table(plain_issues, 100)}
  <h2>文件清单</h2>
  {_table(manifest, 100)}
  <h2>Sheet Inventory</h2>
  {_table(sheet_inventory, 100)}
  <h2>免责声明</h2>
  <p>{escape(DISCLAIMER)}</p>
</body>
</html>
"""
    output_path.write_text(html, encoding="utf-8")
    return output_path
