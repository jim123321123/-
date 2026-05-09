from __future__ import annotations

import ast
import json
from pathlib import Path
from typing import Any

import pandas as pd

from .report_language import ISSUE_LABELS


def _details(value: Any) -> dict[str, Any]:
    if isinstance(value, dict):
        return value
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return {}
    if isinstance(value, str):
        text = value.strip()
        if not text:
            return {}
        for loader in (json.loads, ast.literal_eval):
            try:
                parsed = loader(text)
            except Exception:
                continue
            if isinstance(parsed, dict):
                return parsed
    return {}


def _json_safe(value: Any) -> Any:
    if isinstance(value, dict):
        return {str(key): _json_safe(val) for key, val in value.items()}
    if isinstance(value, (list, tuple, set)):
        return [_json_safe(item) for item in value]
    if isinstance(value, pd.Timestamp):
        return value.isoformat()
    if pd.isna(value) if not isinstance(value, (dict, list, tuple, set)) else False:
        return ""
    return value


def serialize_issue_log(issue_log: pd.DataFrame) -> pd.DataFrame:
    if issue_log is None or issue_log.empty:
        return pd.DataFrame() if issue_log is None else issue_log.copy()
    out = issue_log.copy()
    if "severity" in out.columns:
        out = out.drop(columns=["severity"])
    if "details" in out.columns:
        out["details"] = out["details"].apply(
            lambda value: json.dumps(_json_safe(value), ensure_ascii=False) if isinstance(value, dict) else value
        )
    return out


def write_findings_csv(issue_log: pd.DataFrame, output_path: Path) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    serialize_issue_log(issue_log).to_csv(output_path, index=False, encoding="utf-8-sig")
    return output_path

def _max_risk(values: pd.Series) -> str:
    order = {"Red": 3, "Orange": 2, "Yellow": 1, "": 0}
    normalized = [str(value) for value in values.fillna("")]
    return max(normalized, key=lambda item: order.get(item, 0), default="")


def summarize_v2_findings(issue_log: pd.DataFrame) -> dict[str, pd.DataFrame]:
    empty = {
        "rule_counts": pd.DataFrame(columns=["规则编号", "规则名称", "数量", "风险等级"]),
        "top_findings": pd.DataFrame(columns=["规则", "风险", "文件", "表格/页面", "位置", "问题"]),
        "n004_digits": pd.DataFrame(columns=["文件", "表格/页面", "变量", "末位数字计数", "p值", "占比最高数字"]),
        "n005_mismatches": pd.DataFrame(columns=["文件", "表格/页面", "位置", "不匹配行", "说明"]),
        "image_pairs": pd.DataFrame(columns=["规则", "风险", "图片1", "图片2", "距离", "说明"]),
    }
    if issue_log is None or issue_log.empty:
        return empty

    source = issue_log.copy()
    if "rule_id" not in source.columns:
        source["rule_id"] = ""
    if "issue_type" not in source.columns:
        source["issue_type"] = ""
    if "risk_level" not in source.columns:
        source["risk_level"] = ""
    if "issue_id" not in source.columns:
        source["issue_id"] = source.index.astype(str)

    grouped = (
        source.groupby(["rule_id", "issue_type"], dropna=False)
        .agg(count=("issue_id", "count"), risk=("risk_level", _max_risk))
        .reset_index()
        .sort_values(["rule_id", "count"], ascending=[True, False])
    )
    empty["rule_counts"] = pd.DataFrame(
        [
            {
                "规则编号": row["rule_id"] or "未编号",
                "规则名称": ISSUE_LABELS.get(row["issue_type"], row["issue_type"]),
                "数量": int(row["count"]),
                "风险等级": row["risk"] or "未分级",
            }
            for _, row in grouped.iterrows()
        ],
        columns=empty["rule_counts"].columns,
    )

    risk_rank = {"Red": 3, "Orange": 2, "Yellow": 1}
    ranked = source.assign(_risk_rank=source["risk_level"].astype(str).map(risk_rank).fillna(0)).sort_values(
        "_risk_rank", ascending=False
    )
    top = ranked[ranked["_risk_rank"] >= 2].head(20)
    empty["top_findings"] = pd.DataFrame(
        [
            {
                "规则": f"{row.get('rule_id', '')} {ISSUE_LABELS.get(row.get('issue_type', ''), row.get('issue_type', ''))}".strip(),
                "风险": row.get("risk_level", ""),
                "文件": row.get("file_name", ""),
                "表格/页面": row.get("sheet_or_panel", ""),
                "位置": row.get("sample_or_variable", ""),
                "问题": row.get("evidence", ""),
            }
            for _, row in top.iterrows()
        ],
        columns=empty["top_findings"].columns,
    )

    n004_rows = []
    n005_rows = []
    image_rows = []
    for _, row in source.iterrows():
        details = _details(row.get("details", ""))
        rule_id = str(row.get("rule_id", ""))
        issue_type = str(row.get("issue_type", ""))
        if rule_id == "N004":
            digit_counts = details.get("digit_counts", {})
            counts_text = "；".join(f"{key}:{value}" for key, value in digit_counts.items())
            n004_rows.append(
                {
                    "文件": row.get("file_name", ""),
                    "表格/页面": row.get("sheet_or_panel", ""),
                    "变量": row.get("sample_or_variable", ""),
                    "末位数字计数": counts_text,
                    "p值": details.get("p_value", ""),
                    "占比最高数字": details.get("dominant_digit", ""),
                }
            )
        if rule_id == "N005":
            n005_rows.append(
                {
                    "文件": row.get("file_name", ""),
                    "表格/页面": row.get("sheet_or_panel", ""),
                    "位置": row.get("sample_or_variable", ""),
                    "不匹配行": details.get("mismatched_rows", ""),
                    "说明": row.get("evidence", ""),
                }
            )
        if rule_id in {"I001", "I002"} or issue_type in {"exact_duplicate_image_detector", "perceptual_duplicate_detector"}:
            image_rows.append(
                {
                    "规则": rule_id,
                    "风险": row.get("risk_level", ""),
                    "图片1": details.get("image_1", row.get("file_name", "")),
                    "图片2": details.get("image_2", ""),
                    "距离": details.get("distance", ""),
                    "说明": row.get("evidence", ""),
                }
            )

    empty["n004_digits"] = pd.DataFrame(n004_rows, columns=empty["n004_digits"].columns)
    empty["n005_mismatches"] = pd.DataFrame(n005_rows, columns=empty["n005_mismatches"].columns)
    empty["image_pairs"] = pd.DataFrame(image_rows, columns=empty["image_pairs"].columns)
    return empty


def write_report_json(
    output_path: Path,
    summary: dict[str, Any],
    manifest: pd.DataFrame,
    sheet_inventory: pd.DataFrame,
    issue_log: pd.DataFrame,
    external_status: pd.DataFrame,
) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    v2 = summarize_v2_findings(issue_log)
    payload = {
        "summary": _json_safe(summary),
        "file_count": int(len(manifest)) if manifest is not None else 0,
        "sheet_count": int(len(sheet_inventory)) if sheet_inventory is not None else 0,
        "external_status": serialize_issue_log(external_status).to_dict(orient="records")
        if external_status is not None
        else [],
        "rule_counts": v2["rule_counts"].to_dict(orient="records"),
        "top_findings": v2["top_findings"].to_dict(orient="records"),
        "findings": serialize_issue_log(issue_log).to_dict(orient="records") if issue_log is not None else [],
    }
    output_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return output_path
