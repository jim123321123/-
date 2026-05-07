from __future__ import annotations

from pathlib import Path
from typing import Any

import pandas as pd
import requests


def _status(tool: str, status: str, message: str, needs_manual_upload: bool = False, report_path: str = "", raw_response_path: str = "") -> dict[str, Any]:
    return {
        "tool": tool,
        "status": status,
        "message": message,
        "report_path": report_path,
        "raw_response_path": raw_response_path,
        "needs_manual_upload": needs_manual_upload,
    }


def _post_file(tool: str, file_path: Path | None, api_key: str | None, endpoint: str | None) -> dict[str, Any]:
    if not api_key:
        return _status(tool, "skipped", "API key not provided.")
    if not endpoint:
        return _status(
            tool,
            "manual_required",
            "API key exists but endpoint is not configured. Please upload the generated package to the external platform manually.",
            True,
        )
    try:
        files = {}
        handle = None
        if file_path and file_path.exists():
            handle = file_path.open("rb")
            files["file"] = (file_path.name, handle)
        response = requests.post(endpoint, headers={"Authorization": f"Bearer {api_key}"}, files=files, timeout=60)
        response.raise_for_status()
        return _status(tool, "submitted", f"Request submitted. HTTP {response.status_code}.")
    except Exception as exc:
        return _status(tool, "failed", f"External API request failed: {exc}")
    finally:
        if "handle" in locals() and handle is not None:
            handle.close()


def run_proofig_check(image_package_path: Path, api_key: str | None, endpoint: str | None) -> dict[str, Any]:
    return _post_file("Proofig AI", image_package_path, api_key, endpoint)


def run_imagetwin_check(image_package_path: Path, api_key: str | None, endpoint: str | None) -> dict[str, Any]:
    return _post_file("Imagetwin", image_package_path, api_key, endpoint)


def run_dataseer_check(manuscript_path: Path | None, api_key: str | None, endpoint: str | None) -> dict[str, Any]:
    return _post_file("DataSeer", manuscript_path, api_key, endpoint)


def run_llm_summary(qc_summary_json: dict[str, Any], api_key: str | None, endpoint: str | None, model_name: str | None) -> dict[str, Any]:
    if not api_key:
        return _status("LLM", "skipped", "API key not provided.")
    if not endpoint:
        return _status("LLM", "manual_required", "API key exists but endpoint is not configured.", True)
    try:
        response = requests.post(
            endpoint,
            headers={"Authorization": f"Bearer {api_key}"},
            json={"model": model_name, "qc_summary": qc_summary_json},
            timeout=60,
        )
        response.raise_for_status()
        return _status("LLM", "submitted", f"Request submitted. HTTP {response.status_code}.")
    except Exception as exc:
        return _status("LLM", "failed", f"External API request failed: {exc}")


def write_external_ai_status(statuses: list[dict[str, Any]], output_path: Path) -> pd.DataFrame:
    df = pd.DataFrame(statuses)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(output_path, index=False)
    return df
