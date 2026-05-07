from __future__ import annotations

from pathlib import Path

import pandas as pd
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QFileDialog,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QProgressBar,
    QTabWidget,
    QTableView,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from src import APP_VERSION
from src.core.utils import load_json, open_path
from src.gui.run_worker import RunWorker
from src.gui.settings_dialog import SettingsDialog
from src.gui.widgets import DataFrameModel


class MainWindow(QMainWindow):
    def __init__(self, base_dir: Path):
        super().__init__()
        self.base_dir = base_dir
        self.worker: RunWorker | None = None
        self.session_keys: dict[str, str] = {}
        self.paths: dict[str, Path | None] = {
            "zip_path": None,
            "data_dictionary": None,
            "sample_info": None,
            "external_image_report": None,
            "external_dataseer_report": None,
        }
        self.result: dict | None = None
        self.models = {
            "manifest": DataFrameModel(),
            "sheets": DataFrameModel(),
            "issues": DataFrameModel(),
        }
        self.setWindowTitle("投稿前AI数据真实性与合理性质控系统")
        self.resize(1220, 840)
        self._build_ui()

    def _build_ui(self) -> None:
        root = QWidget()
        layout = QVBoxLayout(root)
        layout.addWidget(self._project_group())
        layout.addWidget(self._upload_group())
        layout.addWidget(self._config_group())
        layout.addWidget(self._run_group())
        layout.addWidget(self._result_group())
        layout.addWidget(self._tables_group(), stretch=1)
        layout.addWidget(self._output_group())
        self.setCentralWidget(root)

    def _project_group(self) -> QGroupBox:
        group = QGroupBox("项目信息")
        grid = QGridLayout(group)
        self.project_name = QLineEdit("Demo_Project")
        grid.addWidget(QLabel("Project Name"), 0, 0)
        grid.addWidget(self.project_name, 0, 1)
        grid.addWidget(QLabel(f"软件版本：{APP_VERSION}"), 0, 2)
        grid.addWidget(QLabel(f"当前运行目录：{self.base_dir}"), 1, 0, 1, 3)
        settings = QPushButton("设置 API Key / 外部AI工具")
        history = QPushButton("打开历史输出目录")
        settings.clicked.connect(self._open_settings)
        history.clicked.connect(lambda: open_path(self.base_dir / "runs"))
        grid.addWidget(settings, 0, 3)
        grid.addWidget(history, 1, 3)
        return group

    def _upload_group(self) -> QGroupBox:
        group = QGroupBox("数据上传")
        grid = QGridLayout(group)
        specs = [
            ("zip_path", "选择 zip 压缩包", "Zip files (*.zip)"),
            ("data_dictionary", "可选 data_dictionary.xlsx", "Excel files (*.xlsx *.xls)"),
            ("sample_info", "可选 sample_info.xlsx", "Excel files (*.xlsx *.xls)"),
            ("external_image_report", "外部AI图片检查报告", "Reports (*.pdf *.csv *.xlsx *.xls)"),
            ("external_dataseer_report", "DataSeer/其他报告", "Reports (*.pdf *.csv *.xlsx *.xls)"),
        ]
        self.path_labels: dict[str, QLabel] = {}
        for row, (key, text, pattern) in enumerate(specs):
            button = QPushButton(text)
            label = QLabel("未选择")
            label.setTextInteractionFlags(Qt.TextSelectableByMouse)
            button.clicked.connect(lambda checked=False, k=key, p=pattern: self._choose_file(k, p))
            grid.addWidget(button, row, 0)
            grid.addWidget(label, row, 1)
            self.path_labels[key] = label
        return group

    def _config_group(self) -> QGroupBox:
        group = QGroupBox("检查配置")
        grid = QGridLayout(group)
        self.default_thresholds = QCheckBox("使用默认阈值")
        self.default_thresholds.setChecked(True)
        grid.addWidget(self.default_thresholds, 0, 0)
        labels = [
            "near duplicate threshold: 0.95",
            "high correlation threshold: 0.995",
            "fixed ratio CV threshold: 0.01",
            "equal difference run length: 5",
            "zero/five terminal digit threshold: 0.45/0.60",
            "robust z threshold: 3.5",
            "missingness threshold: 0.5/0.8",
            "extreme value threshold: 1e300",
        ]
        for index, text in enumerate(labels, start=1):
            grid.addWidget(QLabel(text), index // 4, index % 4)
        return group

    def _run_group(self) -> QGroupBox:
        group = QGroupBox("运行控制")
        layout = QVBoxLayout(group)
        buttons = QHBoxLayout()
        self.start_button = QPushButton("开始检查")
        self.stop_button = QPushButton("停止检查")
        self.stop_button.setEnabled(False)
        self.start_button.clicked.connect(self._start)
        self.stop_button.clicked.connect(self._stop)
        buttons.addWidget(self.start_button)
        buttons.addWidget(self.stop_button)
        buttons.addStretch()
        layout.addLayout(buttons)
        self.progress = QProgressBar()
        self.current_step = QLabel("等待开始")
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setMaximumHeight(120)
        layout.addWidget(self.progress)
        layout.addWidget(self.current_step)
        layout.addWidget(self.log)
        return group

    def _result_group(self) -> QGroupBox:
        group = QGroupBox("结果概览")
        grid = QGridLayout(group)
        self.summary_labels = {}
        fields = ["final_status", "red_count", "orange_count", "yellow_count", "file_count", "table_count", "sheet_count", "image_count", "pdf_count", "external_ai_status"]
        for index, field in enumerate(fields):
            label = QLabel("-")
            grid.addWidget(QLabel(field), index // 5, (index % 5) * 2)
            grid.addWidget(label, index // 5, (index % 5) * 2 + 1)
            self.summary_labels[field] = label
        return group

    def _tables_group(self) -> QGroupBox:
        group = QGroupBox("表格预览")
        layout = QVBoxLayout(group)
        filters = QHBoxLayout()
        self.risk_filter = QComboBox()
        self.risk_filter.addItems(["All", "Red", "Orange", "Yellow"])
        self.module_filter = QLineEdit()
        self.module_filter.setPlaceholderText("module filter")
        apply_filter = QPushButton("筛选 issue")
        apply_filter.clicked.connect(self._apply_issue_filter)
        filters.addWidget(self.risk_filter)
        filters.addWidget(self.module_filter)
        filters.addWidget(apply_filter)
        filters.addStretch()
        layout.addLayout(filters)
        tabs = QTabWidget()
        for key, title in [("manifest", "raw_file_manifest"), ("sheets", "sheet_inventory"), ("issues", "QC_issue_log")]:
            view = QTableView()
            view.setModel(self.models[key])
            tabs.addTab(view, title)
        layout.addWidget(tabs)
        return group

    def _output_group(self) -> QGroupBox:
        group = QGroupBox("输出")
        layout = QHBoxLayout(group)
        self.output_buttons: dict[str, QPushButton] = {}
        specs = [
            ("run_dir", "打开输出目录"),
            ("pdf", "打开 PDF 报告"),
            ("issue", "打开 QC_issue_log.xlsx"),
            ("numeric", "打开 numeric_qc_results.xlsx"),
            ("sheets", "打开 sheet_inventory.xlsx"),
            ("image_zip", "打开 image_check_package.zip"),
            ("external", "打开 external_ai_status.xlsx"),
        ]
        for key, text in specs:
            button = QPushButton(text)
            button.setEnabled(False)
            button.clicked.connect(lambda checked=False, k=key: self._open_output(k))
            layout.addWidget(button)
            self.output_buttons[key] = button
        return group

    def _choose_file(self, key: str, pattern: str) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "选择文件", str(self.base_dir), pattern)
        if path:
            self.paths[key] = Path(path)
            self.path_labels[key].setText(path)

    def _open_settings(self) -> None:
        dialog = SettingsDialog(self.base_dir, self)
        if dialog.exec():
            self.session_keys.update(dialog.session_keys)

    def _external_settings(self) -> dict:
        return load_json(self.base_dir / "config" / "external_ai_settings.json")

    def _start(self) -> None:
        if not self.paths["zip_path"]:
            QMessageBox.warning(self, "缺少输入", "请先选择 zip 压缩包。")
            return
        options = {
            "base_dir": self.base_dir,
            "project_name": self.project_name.text().strip(),
            "external_settings": self._external_settings(),
            "session_keys": self.session_keys,
        }
        options.update({k: str(v) for k, v in self.paths.items() if v})
        self.worker = RunWorker(options)
        self.worker.progress.connect(self.progress.setValue)
        self.worker.current_step.connect(self.current_step.setText)
        self.worker.log_message.connect(self.log.append)
        self.worker.result_summary.connect(self._show_summary)
        self.worker.finished_ok.connect(self._finish)
        self.worker.error_message.connect(self._error)
        self.start_button.setEnabled(False)
        self.stop_button.setEnabled(True)
        self.worker.start()

    def _stop(self) -> None:
        if self.worker:
            self.worker.stop()
            self.log.append("已请求停止，正在等待当前步骤结束。")

    def _show_summary(self, summary: dict) -> None:
        for field, label in self.summary_labels.items():
            label.setText(str(summary.get(field, "-")))

    def _finish(self, result: dict) -> None:
        self.result = result
        self.models["manifest"].set_dataframe(result["manifest"])
        self.models["sheets"].set_dataframe(result["sheet_inventory"])
        self.models["issues"].set_dataframe(result["issue_log"])
        self.start_button.setEnabled(True)
        self.stop_button.setEnabled(False)
        for button in self.output_buttons.values():
            button.setEnabled(True)
        QMessageBox.information(self, "完成", f"检查完成：{result['summary']['final_status']}")

    def _error(self, message: str) -> None:
        self.start_button.setEnabled(True)
        self.stop_button.setEnabled(False)
        QMessageBox.critical(self, "检查失败", message)

    def _apply_issue_filter(self) -> None:
        if not self.result:
            return
        df = self.result["issue_log"].copy()
        risk = self.risk_filter.currentText()
        module = self.module_filter.text().strip()
        if risk != "All" and not df.empty:
            df = df[df["risk_level"] == risk]
        if module and not df.empty:
            df = df[df["module"].astype(str).str.contains(module, case=False, na=False)]
        self.models["issues"].set_dataframe(df)

    def _open_output(self, key: str) -> None:
        if not self.result:
            return
        run_dir = Path(self.result["run_dir"])
        paths = {
            "run_dir": run_dir,
            "pdf": run_dir / "outputs" / "reports" / "final_QC_report.pdf",
            "issue": run_dir / "outputs" / "tables" / "QC_issue_log.xlsx",
            "numeric": run_dir / "outputs" / "tables" / "numeric_qc_results.xlsx",
            "sheets": run_dir / "outputs" / "tables" / "sheet_inventory.xlsx",
            "image_zip": run_dir / "outputs" / "image_check" / "image_check_package.zip",
            "external": run_dir / "outputs" / "external_ai" / "external_ai_status.xlsx",
        }
        open_path(paths[key])
