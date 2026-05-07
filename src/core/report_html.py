from __future__ import annotations

from html import escape
from pathlib import Path
from typing import Any

import pandas as pd


def _table(df: pd.DataFrame, limit: int = 100) -> str:
    if df is None or df.empty:
        return "<p>No records.</p>"
    return df.head(limit).to_html(index=False, escape=True)


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
  <h2>外部AI状态</h2>
  {_table(external_status)}
  <h2>问题清单（前100条）</h2>
  {_table(issue_log, 100)}
  <h2>文件清单</h2>
  {_table(manifest, 100)}
  <h2>Sheet Inventory</h2>
  {_table(sheet_inventory, 100)}
  <h2>免责声明</h2>
  <p>本报告用于投稿前数据质量和研究诚信风险筛查。自动化结果不能直接定性研究不端。所有 Red 和 Orange 问题均需结合原始记录、实验记录本、仪器导出文件、未裁剪原始图片和人工复核确认。本报告中的外部AI图片检查结果依赖用户提供的 API 服务或用户导入的外部工具报告。</p>
</body>
</html>
"""
    output_path.write_text(html, encoding="utf-8")
    return output_path
