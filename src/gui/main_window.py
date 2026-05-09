from __future__ import annotations

from pathlib import Path

import pandas as pd
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QProgressBar,
    QScrollArea,
    QSplitter,
    QStyle,
    QTableView,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from src import APP_VERSION
from src.core.utils import load_json, open_path
from src.gui.explanations import action_text, evidence_text, highlight_text, issue_title, mechanism_text
from src.gui.dashboard_widgets import DashboardPanel
from src.gui.rule_detail import HighlightTarget, highlight_targets, rule_key, summarize_rules, table_index
from src.gui.run_worker import RunWorker
from src.gui.settings_dialog import SettingsDialog
from src.gui.widgets import DataFrameModel, HighlightDataFrameModel


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
            "rule_issues": DataFrameModel(),
            "source_table": HighlightDataFrameModel(),
        }
        self.path_labels: dict[str, QLabel] = {}
        self.output_buttons: dict[str, QPushButton] = {}
        self.rule_buttons: dict[str, QPushButton] = {}
        self.current_rule_key: str | None = None
        self.current_rule_issues = pd.DataFrame()
        self.current_targets = []
        self.current_table_targets: dict[tuple[str, str], HighlightTarget] = {}
        self.source_tables: dict[tuple[str, str], pd.DataFrame] = {}

        self.setWindowTitle("投稿前 AI 数据质量控制系统")
        self.resize(1380, 920)
        self._build_ui()
        self._apply_style()

    def _build_ui(self) -> None:
        root = QWidget()
        root.setObjectName("appRoot")
        layout = QVBoxLayout(root)
        layout.setContentsMargins(16, 14, 16, 16)
        layout.setSpacing(12)
        layout.addWidget(self._hero())

        splitter = QSplitter(Qt.Horizontal)
        splitter.setChildrenCollapsible(False)
        splitter.addWidget(self._left_panel())
        splitter.addWidget(self._right_panel())
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([260, 1120])
        layout.addWidget(splitter, stretch=1)
        self.setCentralWidget(root)

    def _hero(self) -> QFrame:
        hero = QFrame()
        hero.setObjectName("hero")
        layout = QHBoxLayout(hero)
        layout.setContentsMargins(20, 16, 20, 16)
        layout.setSpacing(18)

        mark = QLabel("QC")
        mark.setObjectName("heroMark")
        title_box = QVBoxLayout()
        title = QLabel("投稿前 AI 数据质控")
        title.setObjectName("heroTitle")
        subtitle = QLabel("上传压缩包后开始检查；右侧按规则查看问题、说明和原始表定位。")
        subtitle.setObjectName("heroSubtitle")
        subtitle.setWordWrap(True)
        title_box.addWidget(title)
        title_box.addWidget(subtitle)

        self.hero_status = QLabel("等待上传")
        self.hero_status.setObjectName("statusPill")
        version = QLabel(f"v{APP_VERSION}")
        version.setObjectName("versionPill")

        layout.addWidget(mark)
        layout.addLayout(title_box, stretch=1)
        layout.addWidget(self.hero_status)
        layout.addWidget(version)
        return hero

    def _left_panel(self) -> QWidget:
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setMinimumWidth(230)
        scroll.setMaximumWidth(300)
        content = QWidget()
        layout = QVBoxLayout(content)
        layout.setContentsMargins(0, 0, 6, 0)
        layout.setSpacing(12)
        layout.addWidget(self._upload_group())
        layout.addWidget(self._api_group())
        layout.addWidget(self._run_group())
        layout.addStretch()
        scroll.setWidget(content)
        return scroll

    def _right_panel(self) -> QWidget:
        splitter = QSplitter(Qt.Vertical)
        splitter.setChildrenCollapsible(False)
        splitter.addWidget(self._rule_result_panel())
        splitter.addWidget(self._rule_detail_panel())
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 3)
        splitter.setSizes([260, 620])
        return splitter

    def _api_group(self) -> QGroupBox:
        group = QGroupBox("API")
        layout = QVBoxLayout(group)
        settings = QPushButton("API Key / 外部工具")
        settings.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_FileDialogDetailedView))
        settings.clicked.connect(self._open_settings)
        layout.addWidget(settings)
        return group

    def _project_group(self) -> QGroupBox:
        group = QGroupBox("项目")
        grid = QGridLayout(group)
        grid.setColumnStretch(1, 1)
        self.project_name = QLineEdit("Demo_Project")
        self.project_name.setPlaceholderText("例如 APOE_Rawdata_QC")
        grid.addWidget(QLabel("项目名称"), 0, 0)
        grid.addWidget(self.project_name, 0, 1)

        settings = QPushButton("API 与外部工具")
        history = QPushButton("历史输出")
        settings.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_FileDialogDetailedView))
        history.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DirOpenIcon))
        settings.clicked.connect(self._open_settings)
        history.clicked.connect(lambda: open_path(self.base_dir / "runs"))
        grid.addWidget(settings, 1, 0)
        grid.addWidget(history, 1, 1)

        base_label = QLabel(f"运行目录：{self.base_dir}")
        base_label.setObjectName("mutedLabel")
        base_label.setWordWrap(True)
        grid.addWidget(base_label, 2, 0, 1, 2)
        return group

    def _upload_group(self) -> QGroupBox:
        group = QGroupBox("上传")
        layout = QVBoxLayout(group)
        specs = [
            ("zip_path", "选择 zip 压缩包", "Zip files (*.zip)", True),
        ]
        for key, text, pattern, primary in specs:
            layout.addWidget(self._file_picker(key, text, pattern, primary))
        return group

    def _file_picker(self, key: str, text: str, pattern: str, primary: bool) -> QWidget:
        row = QFrame()
        row.setObjectName("fileRow")
        layout = QVBoxLayout(row)
        layout.setContentsMargins(12, 10, 12, 10)
        button = QPushButton(text)
        button.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DialogOpenButton))
        if primary:
            button.setObjectName("primaryButton")
        label = QLabel("未选择")
        label.setObjectName("filePathLabel")
        label.setWordWrap(True)
        label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        button.clicked.connect(lambda checked=False, k=key, p=pattern: self._choose_file(k, p))
        layout.addWidget(button)
        layout.addWidget(label)
        self.path_labels[key] = label
        return row

    def _config_group(self) -> QGroupBox:
        group = QGroupBox("检查范围")
        layout = QVBoxLayout(group)
        self.default_thresholds = QCheckBox("使用默认阈值")
        self.default_thresholds.setChecked(True)
        layout.addWidget(self.default_thresholds)
        chips = QWidget()
        grid = QGridLayout(chips)
        grid.setContentsMargins(0, 0, 0, 0)
        labels = ["重复行/列", "固定倍数/差值", "末位数字", "百分比-计数", "跨文件复用", "小数精度", "附表区块审计", "图片重复/相似"]
        for index, text in enumerate(labels):
            chip = QLabel(text)
            chip.setObjectName("ruleChip")
            grid.addWidget(chip, index // 2, index % 2)
        layout.addWidget(chips)
        return group

    def _run_group(self) -> QGroupBox:
        group = QGroupBox("运行")
        layout = QVBoxLayout(group)
        buttons = QHBoxLayout()
        self.start_button = QPushButton("开始检查")
        self.start_button.setObjectName("primaryButton")
        self.start_button.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_MediaPlay))
        self.stop_button = QPushButton("停止")
        self.stop_button.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_MediaStop))
        self.stop_button.setEnabled(False)
        self.start_button.clicked.connect(self._start)
        self.stop_button.clicked.connect(self._stop)
        buttons.addWidget(self.start_button, stretch=2)
        buttons.addWidget(self.stop_button, stretch=1)
        layout.addLayout(buttons)

        self.progress = QProgressBar()
        self.progress.setTextVisible(True)
        self.current_step = QLabel("等待开始")
        self.current_step.setObjectName("currentStep")
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setMaximumHeight(132)
        self.log.setPlaceholderText("运行记录会显示在这里。")
        layout.addWidget(self.progress)
        layout.addWidget(self.current_step)
        layout.addWidget(self.log)
        return group

    def _output_group(self) -> QGroupBox:
        group = QGroupBox("输出文件")
        grid = QGridLayout(group)
        specs = [
            ("run_dir", "输出目录", QStyle.StandardPixmap.SP_DirOpenIcon),
            ("pdf", "PDF 报告", QStyle.StandardPixmap.SP_FileIcon),
            ("html", "HTML 报告", QStyle.StandardPixmap.SP_FileIcon),
            ("json", "JSON 报告", QStyle.StandardPixmap.SP_FileDialogDetailedView),
            ("csv", "CSV 问题表", QStyle.StandardPixmap.SP_FileDialogListView),
            ("issue", "Excel 问题清单", QStyle.StandardPixmap.SP_FileDialogListView),
            ("numeric", "数值审计", QStyle.StandardPixmap.SP_FileDialogDetailedView),
            ("image_qc", "图片规则结果", QStyle.StandardPixmap.SP_FileDialogInfoView),
            ("block_audit", "附表区块审计", QStyle.StandardPixmap.SP_FileDialogInfoView),
            ("sheets", "Sheet 清单", QStyle.StandardPixmap.SP_FileDialogContentsView),
            ("image_zip", "图片待检包", QStyle.StandardPixmap.SP_DriveHDIcon),
            ("external", "外部 AI 状态", QStyle.StandardPixmap.SP_ComputerIcon),
        ]
        for index, (key, text, icon) in enumerate(specs):
            button = QPushButton(text)
            button.setIcon(self.style().standardIcon(icon))
            button.setEnabled(False)
            button.clicked.connect(lambda checked=False, k=key: self._open_output(k))
            grid.addWidget(button, index // 2, index % 2)
            self.output_buttons[key] = button
        return group

    def _rule_result_panel(self) -> DashboardPanel:
        panel = DashboardPanel("规则检测结果")
        top = QHBoxLayout()
        self.rule_status_label = QLabel("运行完成后，这里会按规则显示检测结果。")
        self.rule_status_label.setObjectName("mutedLabel")
        self.rule_status_label.setWordWrap(True)
        self.rule_filter = QComboBox()
        for label, mode in [
            ("重点：Red", "risk_red"),
            ("全部问题", "all"),
            ("风险：Red", "risk_red"),
            ("风险：Orange", "risk_orange"),
            ("风险：Yellow", "risk_yellow"),
        ]:
            self.rule_filter.addItem(label, userData=mode)
        self.rule_filter.currentTextChanged.connect(self._refresh_rule_cards)
        top.addWidget(self.rule_status_label, stretch=1)
        top.addWidget(QLabel("显示"))
        top.addWidget(self.rule_filter)
        panel.body.addLayout(top)

        self.rule_cards_container = QWidget()
        self.rule_cards_grid = QGridLayout(self.rule_cards_container)
        self.rule_cards_grid.setContentsMargins(0, 0, 0, 0)
        self.rule_cards_grid.setSpacing(10)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setWidget(self.rule_cards_container)
        panel.body.addWidget(scroll, stretch=1)
        return panel

    def _rule_detail_panel(self) -> DashboardPanel:
        panel = DashboardPanel("规则明细与原始表高亮")
        self.rule_detail_title = QLabel("请选择上方某个规则")
        self.rule_detail_title.setObjectName("panelTitle")
        self.rule_detail_note = QLabel("点击规则卡片后显示该规则的问题表；能定位到单元格时只高亮单元格，不能精确到单元格时才显示列、行或区块范围。")
        self.rule_detail_note.setObjectName("mutedLabel")
        self.rule_detail_note.setWordWrap(True)
        panel.body.addWidget(self.rule_detail_title)
        panel.body.addWidget(self.rule_detail_note)

        self.problem_description = QTextEdit()
        self.problem_description.setReadOnly(True)
        self.problem_description.setMinimumHeight(90)
        self.problem_description.setObjectName("problemDescription")
        self.problem_description.setPlaceholderText("选择规则和有问题的 sheet 后，这里会用文字说明具体问题。")

        self.source_selector = QComboBox()
        self.source_selector.currentIndexChanged.connect(self._show_selected_source_table)
        panel.body.addWidget(self.source_selector)

        splitter = QSplitter(Qt.Vertical)
        splitter.setChildrenCollapsible(False)
        self.rule_issue_view = QTableView()
        self.rule_issue_view.setModel(self.models["rule_issues"])
        self.rule_issue_view.setAlternatingRowColors(True)
        self.rule_issue_view.setSelectionBehavior(QTableView.SelectRows)
        self.rule_issue_view.clicked.connect(self._select_issue_detail)
        self.source_table_view = QTableView()
        self.source_table_view.setModel(self.models["source_table"])
        self.source_table_view.setAlternatingRowColors(True)
        self.source_table_view.setSelectionBehavior(QTableView.SelectItems)
        splitter.addWidget(self.problem_description)
        splitter.addWidget(self.rule_issue_view)
        splitter.addWidget(self.source_table_view)
        splitter.setSizes([130, 190, 460])
        panel.body.addWidget(splitter, stretch=1)
        return panel

    def _choose_file(self, key: str, pattern: str) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "选择文件", str(self.base_dir), pattern)
        if path:
            self.paths[key] = Path(path)
            self.path_labels[key].setText(path)
            if key == "zip_path":
                self.hero_status.setText("数据已选择，等待开始")

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
            "project_name": Path(self.paths["zip_path"]).stem if self.paths["zip_path"] else "QC_Run",
            "external_settings": self._external_settings(),
            "session_keys": self.session_keys,
        }
        options.update({k: str(v) for k, v in self.paths.items() if v})
        self.worker = RunWorker(options)
        self.worker.progress.connect(self.progress.setValue)
        self.worker.current_step.connect(self._set_current_step)
        self.worker.log_message.connect(self.log.append)
        self.worker.result_summary.connect(self._show_summary)
        self.worker.finished_ok.connect(self._finish)
        self.worker.error_message.connect(self._error)
        self.start_button.setEnabled(False)
        self.stop_button.setEnabled(True)
        self.hero_status.setText("正在检查")
        self.progress.setValue(0)
        self.log.clear()
        self.worker.start()

    def _set_current_step(self, text: str) -> None:
        self.current_step.setText(text)
        self.hero_status.setText(text)

    def _stop(self) -> None:
        if self.worker:
            self.worker.stop()
            self.log.append("已请求停止，正在等待当前步骤结束。")

    def _show_summary(self, summary: dict) -> None:
        self.rule_status_label.setText(
            f"总状态：{summary.get('final_status', '-')}；"
            f"Red {summary.get('red_count', 0)}，Orange {summary.get('orange_count', 0)}，Yellow {summary.get('yellow_count', 0)}。"
        )

    def _finish(self, result: dict) -> None:
        self.result = result
        self.source_tables = table_index(result.get("parsed_sheets", []))
        self._refresh_rule_cards()
        self.start_button.setEnabled(True)
        self.stop_button.setEnabled(False)
        for button in self.output_buttons.values():
            button.setEnabled(True)
        self.hero_status.setText(f"检查完成：{result['summary']['final_status']}")
        QMessageBox.information(self, "完成", f"检查完成：{result['summary']['final_status']}")

    def _error(self, message: str) -> None:
        self.start_button.setEnabled(True)
        self.stop_button.setEnabled(False)
        self.hero_status.setText("检查失败")
        QMessageBox.critical(self, "检查失败", message)

    def _clear_rule_cards(self) -> None:
        while self.rule_cards_grid.count():
            item = self.rule_cards_grid.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()
        self.rule_buttons.clear()

    def _filtered_issue_log(self) -> pd.DataFrame:
        if not self.result:
            return pd.DataFrame()
        issue_log = self.result["issue_log"].copy()
        if issue_log.empty:
            return issue_log
        mode = self.rule_filter.currentData() or "risk_red"
        risk = issue_log["risk_level"].astype(str) if "risk_level" in issue_log else pd.Series("", index=issue_log.index)
        if mode == "all":
            return issue_log
        if mode.startswith("risk_"):
            expected = mode.removeprefix("risk_").capitalize()
            return issue_log[risk.eq(expected)]
        return issue_log

    def _refresh_rule_cards(self) -> None:
        self._clear_rule_cards()
        if not self.result:
            return
        issue_log = self._filtered_issue_log()
        rules = summarize_rules(issue_log)
        self.current_rule_key = self.current_rule_key if self.current_rule_key in {item["key"] for item in rules} else None

        for index, item in enumerate(rules):
            button = QPushButton(
                f"{item['title']}\n"
                f"问题 {item['count']} 项 | 风险 {item['risk']} | 涉及文件 {item['files']} 个"
            )
            button.setObjectName("ruleCard")
            button.setMinimumHeight(82)
            button.setCheckable(True)
            button.clicked.connect(lambda checked=False, key=item["key"]: self._select_rule(key))
            self.rule_cards_grid.addWidget(button, index // 2, index % 2)
            self.rule_buttons[item["key"]] = button
        if rules:
            self._select_rule(self.current_rule_key or rules[0]["key"])
        else:
            self.rule_detail_title.setText("当前筛选条件下没有问题")
            self.rule_detail_note.setText("可以切换为“全部问题”、Orange 或 Yellow 查看完整结果。")
            self.problem_description.clear()
            self.source_selector.clear()
            self.models["rule_issues"].set_dataframe(pd.DataFrame())
            self.models["source_table"].set_dataframe(pd.DataFrame())
            self.models["source_table"].set_highlights()

    def _select_rule(self, key: str) -> None:
        if not self.result:
            return
        self.current_rule_key = key
        for rule, button in self.rule_buttons.items():
            button.setChecked(rule == key)
        issue_log = self._filtered_issue_log()
        filtered = issue_log[issue_log.apply(rule_key, axis=1) == key].copy()
        self.current_rule_issues = filtered
        self.models["rule_issues"].set_dataframe(
            filtered[["risk_level", "rule_id", "issue_type", "file_name", "sheet_or_panel", "sample_or_variable", "evidence"]]
            if not filtered.empty
            else pd.DataFrame()
        )
        title = self.rule_buttons[key].text().splitlines()[0] if key in self.rule_buttons else key
        self.rule_detail_title.setText(title)
        if filtered.empty:
            self.rule_detail_note.setText("该规则当前没有问题。")
            self.problem_description.clear()
            self.source_selector.clear()
            self.models["source_table"].set_dataframe(pd.DataFrame())
            self.models["source_table"].set_highlights()
            return
        self._render_rule_tables(filtered)

    def _select_issue_detail(self, index) -> None:
        if self.current_rule_issues is None or self.current_rule_issues.empty or not index.isValid():
            return
        self._select_issue_sheet(self.current_rule_issues.iloc[index.row()])

    def _render_rule_tables(self, issues: pd.DataFrame) -> None:
        self.current_table_targets = {}
        issue_counts: dict[tuple[str, str], int] = {}
        for _, issue_row in issues.iterrows():
            for target in highlight_targets(issue_row, self.source_tables):
                key = (target.file_name, target.sheet_name)
                if key not in self.source_tables:
                    continue
                issue_counts[key] = issue_counts.get(key, 0) + 1
                if key not in self.current_table_targets:
                    self.current_table_targets[key] = HighlightTarget(
                        file_name=target.file_name,
                        sheet_name=target.sheet_name,
                        note=target.note,
                        scope=target.scope,
                    )
                if target.scope == "row":
                    self.current_table_targets[key].rows.update(target.rows)
                elif target.scope == "column":
                    self.current_table_targets[key].columns.update(target.columns)
                else:
                    self.current_table_targets[key].cells.update(target.cells)

        self.current_targets = list(self.current_table_targets.values())
        self.source_selector.blockSignals(True)
        self.source_selector.clear()
        for target in self.current_targets:
            key = (target.file_name, target.sheet_name)
            highlight_count = len(target.cells) + len(target.rows) + len(target.columns)
            self.source_selector.addItem(
                f"{target.file_name} / {target.sheet_name}  |  问题 {issue_counts.get(key, 0)} 项，高亮对象 {highlight_count} 个",
                userData=target,
            )
        self.source_selector.blockSignals(False)

        if self.current_targets:
            self.rule_detail_note.setText("这里只显示该规则涉及的有问题 sheet。行级规则高亮整行，列关系规则高亮相关列，计算错误等精确问题高亮单元格。")
            self.source_selector.setCurrentIndex(0)
            self._show_selected_source_table()
        else:
            self.rule_detail_note.setText("该规则的问题无法映射到解析后的表格单元格。图片问题或外部报告问题请查看证据字段和输出文件。")
            self.problem_description.setPlainText("该规则没有能精确定位到单元格的表格内容。请查看上方问题列表中的证据字段，或打开输出报告中的详细说明。")
            self.models["source_table"].set_dataframe(pd.DataFrame())
            self.models["source_table"].set_highlights()

    def _select_issue_sheet(self, issue_row) -> None:
        issue_targets = [target for target in highlight_targets(issue_row, self.source_tables) if (target.file_name, target.sheet_name) in self.source_tables]
        if not issue_targets:
            return
        target_keys = {(target.file_name, target.sheet_name) for target in issue_targets}
        for index in range(self.source_selector.count()):
            target = self.source_selector.itemData(index)
            if target is not None and (target.file_name, target.sheet_name) in target_keys:
                self.source_selector.setCurrentIndex(index)
                self._show_selected_source_table()
                return

    def _problem_text_for_target(self, target: HighlightTarget) -> str:
        if self.current_rule_issues is None or self.current_rule_issues.empty:
            return ""
        lines = []
        shown = 0
        for _, issue_row in self.current_rule_issues.iterrows():
            issue_targets = highlight_targets(issue_row, self.source_tables)
            matched_targets = [
                item
                for item in issue_targets
                if item.file_name == target.file_name
                and item.sheet_name == target.sheet_name
                and (not item.cells or not target.cells or item.cells.intersection(target.cells))
            ]
            if not matched_targets:
                continue
            shown += 1
            rule_id = str(issue_row.get("rule_id", "") or "").strip()
            issue_type = str(issue_row.get("issue_type", "") or "").strip()
            risk = str(issue_row.get("risk_level", "") or "").strip()
            sample = str(issue_row.get("sample_or_variable", "") or "").strip()
            heading = f"问题 {shown}：{issue_title(issue_row)}"
            if risk:
                heading += f"（{risk}）"
            lines.append(heading)
            if sample:
                lines.append(f"涉及对象：{sample}")
            lines.append(highlight_text(matched_targets, issue_type))
            lines.append(f"判断依据：{evidence_text(issue_row)}")
            lines.append(mechanism_text(issue_row))
            lines.append(f"建议处理：{action_text(issue_row)}")
            lines.append("")
            if shown >= 8:
                break
        if not lines:
            return target.note or "当前 sheet 有问题，但该规则没有提供可进一步拆分的文字说明。"
        remaining = len(self.current_rule_issues) - shown
        if remaining > 0 and shown >= 8:
            lines.append(f"还有更多问题未展开，请在上方问题列表中逐条查看。")
        return "\n".join(lines).strip()

    def _render_issue(self, issue_row) -> None:
        self.current_targets = highlight_targets(issue_row, self.source_tables)
        self.source_selector.blockSignals(True)
        self.source_selector.clear()
        for target in self.current_targets:
            self.source_selector.addItem(f"{target.file_name} / {target.sheet_name}", userData=target)
        self.source_selector.blockSignals(False)
        if self.current_targets:
            self.rule_detail_note.setText(self.current_targets[0].note or "已定位到原始表。黄色单元格是规则实际定位到的问题数据。")
            self.source_selector.setCurrentIndex(0)
            self._show_selected_source_table()
        else:
            self.rule_detail_note.setText("该问题无法映射到解析后的表格。图片问题或外部报告问题请查看证据字段和输出文件。")
            self.problem_description.setPlainText("该问题没有能精确定位到单元格的表格内容，请查看证据字段和输出报告。")
            self.models["source_table"].set_dataframe(pd.DataFrame())
            self.models["source_table"].set_highlights()

    def _show_selected_source_table(self) -> None:
        target = self.source_selector.currentData()
        if target is None:
            self.problem_description.clear()
            return
        df = self.source_tables.get((target.file_name, target.sheet_name))
        if df is None:
            self.problem_description.setPlainText("没有找到该 sheet 的原始表格内容。")
            self.models["source_table"].set_dataframe(pd.DataFrame())
            self.models["source_table"].set_highlights()
            return
        self.models["source_table"].set_dataframe(df)
        self.models["source_table"].set_highlights(target.rows, target.columns, target.cells)
        self.problem_description.setPlainText(self._problem_text_for_target(target))
        if target.cells:
            first_row, first_col = sorted(target.cells)[0]
            row_index = max(first_row - 2, 0)
            col_index = list(map(str, df.columns)).index(first_col) if first_col in list(map(str, df.columns)) else 0
            self.source_table_view.scrollTo(self.models["source_table"].index(row_index, col_index))
        elif target.rows:
            row_index = max(min(target.rows) - 2, 0)
            self.source_table_view.scrollTo(self.models["source_table"].index(row_index, 0))
        elif target.columns:
            columns = list(map(str, df.columns))
            first_col = sorted(target.columns)[0]
            col_index = columns.index(first_col) if first_col in columns else 0
            self.source_table_view.scrollTo(self.models["source_table"].index(0, col_index))

    def _open_output(self, key: str) -> None:
        if not self.result:
            return
        run_dir = Path(self.result["run_dir"])
        paths = {
            "run_dir": run_dir,
            "pdf": run_dir / "outputs" / "reports" / "final_QC_report.pdf",
            "html": run_dir / "outputs" / "reports" / "final_QC_report.html",
            "json": run_dir / "outputs" / "reports" / "report.json",
            "csv": run_dir / "outputs" / "reports" / "findings.csv",
            "issue": run_dir / "outputs" / "tables" / "QC_issue_log.xlsx",
            "numeric": run_dir / "outputs" / "tables" / "numeric_qc_results.xlsx",
            "image_qc": run_dir / "outputs" / "tables" / "image_qc_results.xlsx",
            "block_audit": run_dir / "outputs" / "tables" / "block_audit_results.xlsx",
            "sheets": run_dir / "outputs" / "tables" / "sheet_inventory.xlsx",
            "image_zip": run_dir / "outputs" / "image_check" / "image_check_package.zip",
            "external": run_dir / "outputs" / "external_ai" / "external_ai_status.xlsx",
        }
        open_path(paths[key])

    def _apply_style(self) -> None:
        self.setStyleSheet(
            """
            QWidget#appRoot { background: #f5f7fb; color: #111827; font-family: "Microsoft YaHei UI", "Segoe UI", Arial; font-size: 13px; }
            QFrame#hero, QGroupBox, QFrame#dashboardPanel { background: #ffffff; border: 1px solid #dde4ee; border-radius: 8px; }
            QFrame#hero { margin: 0; }
            QLabel#heroMark { background: #1d4ed8; color: #ffffff; border-radius: 8px; min-width: 54px; min-height: 54px; font-size: 20px; font-weight: 700; qproperty-alignment: AlignCenter; }
            QLabel#heroTitle { font-size: 24px; font-weight: 700; color: #0f172a; }
            QLabel#heroSubtitle, QLabel#mutedLabel, QLabel#filePathLabel { color: #64748b; }
            QLabel#statusPill, QLabel#versionPill { background: #eef6ff; color: #1d4ed8; border: 1px solid #bfdbfe; border-radius: 8px; padding: 7px 12px; font-weight: 600; }
            QLabel#versionPill { background: #f0fdf4; color: #15803d; border-color: #bbf7d0; }
            QGroupBox { margin-top: 14px; font-weight: 700; }
            QGroupBox::title { subcontrol-origin: margin; left: 12px; padding: 0 6px; color: #334155; }
            QLabel#panelTitle { color: #0f172a; font-size: 16px; font-weight: 700; }
            QFrame#fileRow { background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; }
            QLabel#ruleChip { background: #f1f5f9; border: 1px solid #e2e8f0; border-radius: 8px; padding: 6px 8px; color: #334155; font-weight: 600; }
            QPushButton { background: #ffffff; border: 1px solid #cbd5e1; border-radius: 8px; padding: 8px 10px; color: #1f2937; font-weight: 600; text-align: left; }
            QPushButton:hover { background: #f8fafc; border-color: #94a3b8; }
            QPushButton:checked, QPushButton#ruleCard:checked { background: #dbeafe; border-color: #2563eb; color: #0f172a; }
            QPushButton:disabled { color: #94a3b8; background: #f1f5f9; border-color: #e2e8f0; }
            QPushButton#primaryButton { background: #2563eb; color: #ffffff; border-color: #1d4ed8; text-align: center; }
            QLineEdit, QComboBox, QTextEdit { background: #ffffff; border: 1px solid #cbd5e1; border-radius: 8px; padding: 7px 9px; }
            QTextEdit#problemDescription { background: #fff7ed; border-color: #fed7aa; color: #7c2d12; }
            QProgressBar { background: #e5e7eb; border: 0; border-radius: 8px; height: 18px; text-align: center; color: #111827; font-weight: 700; }
            QProgressBar::chunk { background: #0f766e; border-radius: 8px; }
            QLabel#currentStep { color: #0f766e; font-weight: 700; }
            QSplitter::handle { background: #dbe3ef; }
            QSplitter::handle:hover { background: #94a3b8; }
            QHeaderView::section { background: #f1f5f9; border: 0; border-right: 1px solid #e2e8f0; padding: 7px; font-weight: 700; }
            QTableView { background: #ffffff; alternate-background-color: #f8fafc; border: 1px solid #dde4ee; gridline-color: #e5e7eb; selection-background-color: #dbeafe; selection-color: #0f172a; }
            """
        )
