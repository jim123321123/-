from __future__ import annotations

import pandas as pd


def _count_text(counts: pd.Series) -> str:
    if counts.empty:
        return "未识别到文件类型"
    return "、".join(f"{key} {int(value)} 个" for key, value in counts.items())


def _risk_count(issue_log: pd.DataFrame, level: str) -> int:
    if issue_log is None or issue_log.empty or "risk_level" not in issue_log:
        return 0
    return int((issue_log["risk_level"] == level).sum())


def _external_api_text(external_status: pd.DataFrame) -> str:
    if external_status is None or external_status.empty or "status" not in external_status:
        return "未提供外部 AI 工具状态；本报告仍可基于本地确定性规则完成原始数据质量概览。"
    statuses = set(external_status["status"].dropna().astype(str))
    if statuses and statuses.issubset({"skipped"}):
        return "未配置可用的外部 AI API；本次整体描述和风险汇总均由本地文件清单、表格解析和数值规则检查生成。"
    if "manual_required" in statuses:
        return "部分外部 AI 工具需要手动上传检查包或导入外部报告；本地 QC 结果不依赖外部 API。"
    return "外部 AI 工具状态已记录；本段整体描述仍以本地原始数据清单和确定性规则检查为依据。"


def generate_raw_data_overview(
    manifest: pd.DataFrame,
    sheet_inventory: pd.DataFrame,
    issue_log: pd.DataFrame,
    external_status: pd.DataFrame,
) -> str:
    file_count = 0 if manifest is None else len(manifest)
    file_type_text = "未识别到文件类型"
    table_count = 0
    pdf_count = 0
    image_count = 0
    if manifest is not None and not manifest.empty and "file_type" in manifest:
        counts = manifest["file_type"].fillna("unknown").astype(str).value_counts()
        file_type_text = _count_text(counts)
        table_count = int(counts.get("excel", 0) + counts.get("csv", 0))
        pdf_count = int(counts.get("pdf", 0))
        image_count = int(counts.get("image", 0))

    sheet_count = 0 if sheet_inventory is None else len(sheet_inventory)
    parsed_ok = 0
    failed = 0
    profile_text = "未形成可统计的表格类型"
    if sheet_inventory is not None and not sheet_inventory.empty:
        if "parse_status" in sheet_inventory:
            parsed_ok = int((sheet_inventory["parse_status"] == "ok").sum())
            failed = int((sheet_inventory["parse_status"] == "failed").sum())
        if "qc_profile" in sheet_inventory:
            profiles = sheet_inventory["qc_profile"].fillna("unknown").astype(str).value_counts()
            profile_text = "、".join(f"{key} {int(value)} 个" for key, value in profiles.items())

    red = _risk_count(issue_log, "Red")
    orange = _risk_count(issue_log, "Orange")
    yellow = _risk_count(issue_log, "Yellow")
    risk_text = f"Red {red} 项、Orange {orange} 项、Yellow {yellow} 项"
    if red > 0:
        status_text = "整体状态为 Fail，说明存在投稿前必须人工复核并处理的高风险信号。"
    elif orange > 0:
        status_text = "整体状态为 Conditional Fail，说明存在需要回查原始记录的中高风险信号。"
    elif yellow > 0:
        status_text = "整体状态为 Conditional Pass，说明存在轻微信号或需要记录解释的问题。"
    else:
        status_text = "整体状态为 Pass，未见本地规则触发的明显风险信号。"

    parts = [
        f"本次上传的压缩包共识别 {file_count} 个文件，其中 {file_type_text}。",
        (
            f"数据结构上，软件识别到 {table_count} 个表格文件、{pdf_count} 个 PDF 文件和 "
            f"{image_count} 个图片/原始图像文件；共解析 {sheet_count} 个 sheet，其中 {parsed_ok} 个解析成功、{failed} 个解析失败。"
        ),
        f"表格类型初步分布为：{profile_text}。",
        f"本地确定性规则检查共形成风险计数：{risk_text}。{status_text}",
        _external_api_text(external_status),
        (
            "这段描述只反映压缩包内可被本软件读取和规则化检查的原始数据概况，不等同于对研究真实性或研究不端的定性判断；"
            "Red 和 Orange 项仍需结合原始记录、实验记录本、仪器导出文件和人工复核确认。"
        ),
    ]
    return "\n".join(parts)
