from __future__ import annotations

from html import escape
from pathlib import Path
from typing import Any

import pandas as pd

from .report_exports import summarize_v2_findings
from .report_language import build_plain_issue_table, build_priority_review_text, explain_final_status
from .report_summary import generate_raw_data_overview


DISCLAIMER = (
    "本报告用于投稿前数据质量和研究诚信风险筛查。自动化结果不能直接定性研究不端。"
    "所有 Red 和 Orange 项都需要结合原始记录、实验记录本、仪器导出文件、未裁剪原始图片和人工复核确认；Yellow 项建议记录解释。"
)


def _table(df: pd.DataFrame, limit: int = 100) -> str:
    if df is None or df.empty:
        return "<p>无记录。</p>"
    return df.head(limit).to_html(index=False, escape=True)


def _paragraphs(text: str) -> str:
    return "\n".join(f"  <p>{escape(part)}</p>" for part in text.splitlines() if part.strip())


def _block_audit_issues(issue_log: pd.DataFrame) -> pd.DataFrame:
    if issue_log is None or issue_log.empty or "module" not in issue_log.columns:
        return pd.DataFrame()
    return issue_log[issue_log["module"] == "Supplementary Table Block Audit"].copy()


def _block_audit_section(issue_log: pd.DataFrame, block_count: int) -> str:
    block_issues = _block_audit_issues(issue_log)
    plain_block_issues = build_plain_issue_table(block_issues, limit=50)
    if block_issues.empty:
        summary = "本次已运行从 data_audit.py 整合而来的附表区块审计，未在数值区块中发现需要单独列出的重复、固定倍数或固定差值问题。"
    else:
        red = int((block_issues["risk_level"] == "Red").sum()) if "risk_level" in block_issues else 0
        orange = int((block_issues["risk_level"] == "Orange").sum()) if "risk_level" in block_issues else 0
        yellow = int((block_issues["risk_level"] == "Yellow").sum()) if "risk_level" in block_issues else 0
        summary = (
            f"本次已运行从 data_audit.py 整合而来的附表区块审计，共发现 {len(block_issues)} 条区块级风险信号，"
            f"其中 Red {red} 条、Orange {orange} 条、Yellow {yellow} 条。完整明细见 block_audit_results.xlsx。"
        )
    return f"""
  <h2>附表区块审计（data_audit.py 整合结果）</h2>
  <div class="overview">
    <p>{escape(summary)}</p>
    <p>这部分主要检查补充表或 figure source data 中的数值区块：两个区块是否完全相同，是否存在固定倍数或固定差值关系，以及同一区块内部是否有重复行或重复列。</p>
    <p>本次记录到主报告的附表区块审计问题数：{block_count}</p>
  </div>
  {_table(plain_block_issues, 50)}
"""


def _v2_sections(issue_log: pd.DataFrame) -> str:
    v2 = summarize_v2_findings(issue_log)
    return f"""
  <h2>新版规则汇总（N001-N011、I001-I002）</h2>
  <p>下表按规则编号汇总触发次数。规则只提示“需要复核的信号”，不直接给出定性结论。</p>
  {_table(v2["rule_counts"], 100)}

  <h2>优先复核清单</h2>
  <p>这里列出 Red/Orange 项。建议先核对这些位置的原始记录、处理脚本和未裁剪原图。</p>
  {_table(v2["top_findings"], 20)}

  <h2>N004 末位数字分布</h2>
  <p>该表展示触发 N004 的变量及末位数字计数。若仪器精度、四舍五入或录入方式可以解释，应在复核记录中说明。</p>
  {_table(v2["n004_digits"], 30)}

  <h2>N005 百分比/计数不一致</h2>
  <p>该表列出百分比、分子、分母或总数之间不一致的行，便于直接回查计算过程。</p>
  {_table(v2["n005_mismatches"], 30)}

  <h2>I001/I002 图片重复或相似图片</h2>
  <p>该表列出完全重复或视觉上高度相似的图片对。若两张图片代表不同实验条件，应优先打开原始未裁剪图片人工比对。</p>
  {_table(v2["image_pairs"], 30)}
"""


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
    block_count = int(summary.get("block_audit_issue_count", len(_block_audit_issues(issue_log))))
    html = f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <title>投稿前 AI 数据质量控制报告</title>
  <style>
    body {{ font-family: "Microsoft YaHei", Arial, sans-serif; margin: 32px; color: #1f2937; line-height: 1.55; }}
    h1, h2 {{ color: #17324d; }}
    table {{ border-collapse: collapse; width: 100%; margin: 12px 0 24px; font-size: 12px; table-layout: fixed; }}
    th, td {{ border: 1px solid #ddd; padding: 6px; text-align: left; vertical-align: top; overflow-wrap: anywhere; word-break: break-word; }}
    th {{ background: #f2f5f7; }}
    .summary {{ display: grid; grid-template-columns: repeat(4, minmax(120px, 1fr)); gap: 10px; }}
    .item {{ border: 1px solid #ddd; padding: 10px; border-radius: 6px; background: #fff; }}
    .overview {{ background: #f7fafc; border-left: 4px solid #4f7cac; padding: 12px 16px; margin: 16px 0 28px; }}
    .conclusion {{ background: #fff8e5; border-left: 4px solid #d9822b; padding: 12px 16px; margin: 16px 0 28px; }}
  </style>
</head>
<body>
  <h1>投稿前 AI 数据质量控制报告</h1>
  <div class="summary">
    <div class="item"><strong>项目</strong><br>{escape(project_name)}</div>
    <div class="item"><strong>状态</strong><br>{escape(summary.get("final_status", "Pass"))}</div>
    <div class="item"><strong>Red</strong><br>{summary.get("red_count", 0)}</div>
    <div class="item"><strong>Orange</strong><br>{summary.get("orange_count", 0)}</div>
    <div class="item"><strong>Yellow</strong><br>{summary.get("yellow_count", 0)}</div>
    <div class="item"><strong>输出目录</strong><br>{escape(str(summary.get("run_dir", "")))}</div>
  </div>
  <h2>先看结论</h2>
  <div class="conclusion">
    <p>{escape(status_text)}</p>
    <p>{escape(priority_text)}</p>
    <p>报告中的风险项是复核提示，不是对研究行为的定性判断。</p>
  </div>
  <h2>压缩包原始数据整体情况</h2>
  <div class="overview">
{_paragraphs(overview)}
  </div>
  <h2>外部 AI 状态</h2>
  {_table(external_status)}
  {_block_audit_section(issue_log, block_count)}
  {_v2_sections(issue_log)}
  <h2>需要人工复核的问题清单（前100条）</h2>
  <p>下面的表格已把技术规则翻译成中文。请优先看“文件”“表格/页面”“具体位置”和“建议怎么做”。</p>
  {_table(plain_issues, 100)}
  <h2>文件清单</h2>
  {_table(manifest, 100)}
  <h2>Sheet 清单</h2>
  {_table(sheet_inventory, 100)}
  <h2>免责声明</h2>
  <p>{escape(DISCLAIMER)}</p>
</body>
</html>
"""
    output_path.write_text(html, encoding="utf-8")
    return output_path
