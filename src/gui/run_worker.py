from __future__ import annotations

import logging
from pathlib import Path
from typing import Any

import pandas as pd
from PySide6.QtCore import QThread, Signal

from src import APP_VERSION
from src.core import credential_store
from src.core.external_ai_adapters import (
    run_dataseer_check,
    run_imagetwin_check,
    run_llm_summary,
    run_proofig_check,
    write_external_ai_status,
)
from src.core.external_report_importer import import_external_reports
from src.core.image_package import create_image_check_package
from src.core.issue_log import build_issue_log, final_status, write_issue_log
from src.core.manifest import write_manifests
from src.core.numeric_forensics import run_numeric_forensics, write_numeric_results
from src.core.report_html import generate_html_report
from src.core.report_pdf import generate_pdf_report
from src.core.table_parser import parse_tables
from src.core.utils import copy_input_file, create_run_dir, extract_zip, load_yaml


SERVICES = {
    "proofig": "PreSubmissionAIQC_Proofig",
    "imagetwin": "PreSubmissionAIQC_Imagetwin",
    "dataseer": "PreSubmissionAIQC_DataSeer",
    "llm": "PreSubmissionAIQC_LLM",
}


def _counts(issue_log: pd.DataFrame) -> dict[str, int]:
    if issue_log.empty:
        return {"red_count": 0, "orange_count": 0, "yellow_count": 0}
    values = issue_log["risk_level"].value_counts()
    return {
        "red_count": int(values.get("Red", 0)),
        "orange_count": int(values.get("Orange", 0)),
        "yellow_count": int(values.get("Yellow", 0)),
    }


def run_qc_pipeline(options: dict[str, Any], emit=None, should_stop=lambda: False) -> dict[str, Any]:
    base_dir = Path(options.get("base_dir", Path.cwd()))
    project_name = options.get("project_name") or "project"
    zip_path = Path(options["zip_path"])
    external_settings = options.get("external_settings", {})

    def step(progress: int, text: str) -> None:
        if emit:
            emit(progress, text)
        if should_stop():
            raise InterruptedError("用户已请求停止检查。")

    step(3, "正在创建运行目录")
    run_dir = create_run_dir(base_dir, project_name)
    log_path = run_dir / "logs" / "run.log"
    logging.basicConfig(filename=log_path, level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s", force=True)
    logging.info("Run started")

    try:
        step(8, "正在复制输入文件")
        input_zip = copy_input_file(zip_path, run_dir / "input", "upload.zip")
        for key in ("data_dictionary", "sample_info", "external_image_report", "external_dataseer_report"):
            if options.get(key):
                copy_input_file(Path(options[key]), run_dir / "input")

        step(14, "正在解压 zip")
        extract_zip(input_zip, run_dir / "extracted")

        step(22, "正在生成 manifest")
        manifest, _, _ = write_manifests(run_dir / "extracted", run_dir / "outputs" / "tables")

        step(33, "正在解析 Excel / CSV")
        parsed_sheets, sheet_inventory = parse_tables(run_dir / "extracted", run_dir / "outputs" / "tables" / "sheet_inventory.xlsx")

        step(45, "正在执行数值取证检查")
        thresholds = load_yaml(base_dir / "config" / "qc_thresholds.yaml")
        numeric_issues = run_numeric_forensics(parsed_sheets, thresholds)
        numeric_df = write_numeric_results(numeric_issues, run_dir / "outputs" / "tables" / "numeric_qc_results.xlsx")

        step(58, "正在生成图片检查包")
        image_inventory, image_zip = create_image_check_package(run_dir / "extracted", run_dir / "outputs" / "image_check")

        step(68, "正在检查外部AI工具状态")
        session_keys = options.get("session_keys", {})

        def secret(tool: str) -> str | None:
            return session_keys.get(tool) or credential_store.get_secret(SERVICES[tool], "api_key")

        statuses = [
            run_proofig_check(image_zip, secret("proofig"), external_settings.get("proofig", {}).get("endpoint")),
            run_imagetwin_check(image_zip, secret("imagetwin"), external_settings.get("imagetwin", {}).get("endpoint")),
            run_dataseer_check(None, secret("dataseer"), external_settings.get("dataseer", {}).get("endpoint")),
            run_llm_summary({}, secret("llm"), external_settings.get("llm", {}).get("endpoint"), external_settings.get("llm", {}).get("model")),
        ]
        external_status = write_external_ai_status(statuses, run_dir / "outputs" / "external_ai" / "external_ai_status.xlsx")

        step(76, "正在导入外部AI检查报告")
        reports = [Path(options[key]) for key in ("external_image_report", "external_dataseer_report") if options.get(key)]
        imported_reports, external_issues = import_external_reports(reports, run_dir / "outputs" / "external_ai")

        step(84, "正在生成 issue log")
        issue_log = build_issue_log(numeric_issues, sheet_inventory, external_issues, external_status)
        write_issue_log(issue_log, run_dir / "outputs" / "tables" / "QC_issue_log.xlsx")
        status = final_status(issue_log)
        summary = {
            "run_dir": str(run_dir),
            "final_status": status,
            "app_version": APP_VERSION,
            "file_count": int(len(manifest)),
            "table_count": int(len(manifest[manifest["file_type"].isin(["excel", "csv"])])) if not manifest.empty else 0,
            "sheet_count": int(len(sheet_inventory)),
            "image_count": int(len(image_inventory[image_inventory["source_type"].str.contains("image", na=False)])) if not image_inventory.empty else 0,
            "pdf_count": int(len(manifest[manifest["file_type"] == "pdf"])) if not manifest.empty else 0,
            "numeric_issue_count": int(len(numeric_df)),
            "external_ai_status": "; ".join(f"{row.tool}:{row.status}" for row in external_status.itertuples()),
            **_counts(issue_log),
        }

        step(92, "正在生成 PDF 报告")
        generate_html_report(run_dir / "outputs" / "reports" / "final_QC_report.html", project_name, summary, manifest, sheet_inventory, issue_log, external_status)
        generate_pdf_report(
            run_dir / "outputs" / "reports" / "final_QC_report.pdf",
            run_dir / "outputs" / "figures",
            project_name,
            zip_path.name,
            summary,
            manifest,
            sheet_inventory,
            issue_log,
            external_status,
        )

        step(100, "检查完成")
        return {
            "summary": summary,
            "manifest": manifest,
            "sheet_inventory": sheet_inventory,
            "issue_log": issue_log,
            "run_dir": str(run_dir),
        }
    except Exception:
        logging.exception("Run failed")
        raise


class RunWorker(QThread):
    progress = Signal(int)
    current_step = Signal(str)
    log_message = Signal(str)
    result_summary = Signal(dict)
    error_message = Signal(str)
    finished_ok = Signal(dict)

    def __init__(self, options: dict[str, Any]):
        super().__init__()
        self.options = options
        self._stop_requested = False

    def stop(self) -> None:
        self._stop_requested = True

    def run(self) -> None:
        try:
            def emit(progress: int, text: str) -> None:
                self.progress.emit(progress)
                self.current_step.emit(text)
                self.log_message.emit(text)

            result = run_qc_pipeline(self.options, emit=emit, should_stop=lambda: self._stop_requested)
            self.result_summary.emit(result["summary"])
            self.finished_ok.emit(result)
        except Exception as exc:
            self.error_message.emit(str(exc))
