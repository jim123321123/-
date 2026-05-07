from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QCheckBox,
    QDialog,
    QFormLayout,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from src.core import credential_store
from src.core.utils import load_json, save_json
from src.gui.run_worker import SERVICES


TOOL_LABELS = {
    "proofig": "Proofig AI",
    "imagetwin": "Imagetwin",
    "dataseer": "DataSeer",
    "llm": "LLM",
}

DESCRIPTIONS = {
    "proofig": "Proofig AI 用于科学图片完整性检查。若无官方 API 权限，本软件将生成 image_check_package.zip，用户可手动上传后再导入报告。",
    "imagetwin": "Imagetwin 用于图片重复、复用、篡改、抄袭和AI生成图风险筛查。若无 API 权限，可使用手动报告导入模式。",
    "dataseer": "DataSeer / DataSeer SnapShot 用于数据共享、代码共享、协议共享和投稿政策合规检查，不用于直接判断数值真实性。",
    "llm": "LLM 仅用于总结自动QC结果、生成中文摘要和返查建议，不直接定性研究不端。",
}


class SettingsDialog(QDialog):
    def __init__(self, base_dir: Path, parent=None):
        super().__init__(parent)
        self.base_dir = base_dir
        self.settings_path = base_dir / "config" / "external_ai_settings.json"
        self.session_keys: dict[str, str] = {}
        self.setWindowTitle("设置 API Key / 外部AI工具")
        self.resize(720, 620)
        self.controls: dict[str, dict[str, object]] = {}
        self._build_ui()
        self._load_settings()

    def _build_ui(self) -> None:
        main = QVBoxLayout(self)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        content = QWidget()
        content_layout = QVBoxLayout(content)
        for tool, label in TOOL_LABELS.items():
            group = QGroupBox(label)
            form = QFormLayout(group)
            enabled = QCheckBox("启用")
            api_key = QLineEdit()
            api_key.setEchoMode(QLineEdit.Password)
            endpoint = QLineEdit()
            endpoint.setPlaceholderText("官方 API endpoint")
            model = QLineEdit()
            model.setPlaceholderText("仅 LLM 使用")
            key_state = QLabel("未保存")
            test_button = QPushButton("Test Connection")
            test_button.clicked.connect(lambda checked=False, t=tool: self._test_connection(t))
            form.addRow(enabled)
            form.addRow("API Key", api_key)
            form.addRow("Key 状态", key_state)
            form.addRow("API Endpoint", endpoint)
            if tool == "llm":
                form.addRow("Model Name", model)
            form.addRow(test_button)
            desc = QLabel(DESCRIPTIONS[tool])
            desc.setWordWrap(True)
            form.addRow(desc)
            self.controls[tool] = {
                "enabled": enabled,
                "api_key": api_key,
                "endpoint": endpoint,
                "model": model,
                "key_state": key_state,
            }
            content_layout.addWidget(group)
        scroll.setWidget(content)
        main.addWidget(scroll)

        buttons = QHBoxLayout()
        save_button = QPushButton("保存设置")
        clear_button = QPushButton("清除所有密钥")
        session_button = QPushButton("仅本次会话使用")
        close_button = QPushButton("关闭")
        save_button.clicked.connect(self._save_settings)
        clear_button.clicked.connect(self._clear_keys)
        session_button.clicked.connect(self._use_session)
        close_button.clicked.connect(self.reject)
        buttons.addWidget(save_button)
        buttons.addWidget(clear_button)
        buttons.addWidget(session_button)
        buttons.addStretch()
        buttons.addWidget(close_button)
        main.addLayout(buttons)

    def _load_settings(self) -> None:
        settings = load_json(self.settings_path)
        for tool, controls in self.controls.items():
            data = settings.get(tool, {})
            controls["enabled"].setChecked(bool(data.get("enabled", False)))
            controls["endpoint"].setText(data.get("endpoint", ""))
            controls["model"].setText(data.get("model", ""))
            controls["key_state"].setText("已保存" if credential_store.has_secret(SERVICES[tool], "api_key") else "未保存")

    def settings(self) -> dict:
        data = {}
        for tool, controls in self.controls.items():
            data[tool] = {
                "enabled": controls["enabled"].isChecked(),
                "endpoint": controls["endpoint"].text().strip(),
                "model": controls["model"].text().strip(),
            }
        return data

    def _save_settings(self) -> None:
        save_json(self.settings_path, self.settings())
        if not credential_store.is_keyring_available():
            QMessageBox.warning(self, "凭据存储不可用", "当前系统安全凭据存储不可用。API key 将仅在本次会话中使用，不会保存。")
        for tool, controls in self.controls.items():
            key = controls["api_key"].text().strip()
            if key:
                try:
                    credential_store.save_secret(SERVICES[tool], "api_key", key)
                    controls["key_state"].setText("已保存")
                    controls["api_key"].clear()
                except Exception as exc:
                    QMessageBox.warning(self, "保存失败", str(exc))
        self.accept()

    def _clear_keys(self) -> None:
        for tool, controls in self.controls.items():
            credential_store.delete_secret(SERVICES[tool], "api_key")
            controls["key_state"].setText("未保存")
            controls["api_key"].clear()
        QMessageBox.information(self, "完成", "已请求清除所有密钥。")

    def _use_session(self) -> None:
        self.session_keys = {
            tool: controls["api_key"].text().strip()
            for tool, controls in self.controls.items()
            if controls["api_key"].text().strip()
        }
        save_json(self.settings_path, self.settings())
        self.accept()

    def _test_connection(self, tool: str) -> None:
        endpoint = self.controls[tool]["endpoint"].text().strip()
        if not endpoint:
            QMessageBox.information(self, "Test Connection", "Endpoint 未配置；将使用手动上传/导入模式。")
            return
        QMessageBox.information(self, "Test Connection", "已配置 endpoint。真实 API 测试将在运行检查时执行，API key 不会写入日志或配置文件。")
