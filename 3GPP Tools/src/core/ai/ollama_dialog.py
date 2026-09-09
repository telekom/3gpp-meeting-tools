"""
Configuration dialog for Ollama local LLM connection, proxy mode, and active models.
"""

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QComboBox, QGroupBox, QDialogButtonBox, QMessageBox
)

from .ollama_client import OllamaClient, load_ollama_config, save_ollama_config


class OllamaConfigDialog(QDialog):
    settings_saved = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("🦙 Ollama LLM Configuration")
        self.setMinimumWidth(480)
        self.setStyleSheet("background-color: #FAFAFA;")

        self.cfg = load_ollama_config()
        self.client = OllamaClient(
            host=self.cfg.get("host"),
            proxy_mode=self.cfg.get("proxy_mode", "direct")
        )

        self._setup_ui()
        self._load_current_values()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(12)

        # Title & Info
        title = QLabel("🤖 Local LLM Server (Ollama)")
        title.setStyleSheet("font-size: 13px; font-weight: bold; color: #1E293B;")
        layout.addWidget(title)

        desc = QLabel(
            "Connect to a local Ollama instance for private, offline document summaries, "
            "PlantUML diagram generation, and semantic specification queries."
        )
        desc.setWordWrap(True)
        desc.setStyleSheet("color: #64748B; font-size: 11px; margin-bottom: 4px;")
        layout.addWidget(desc)

        # Server Settings Group
        server_group = QGroupBox("Server & Network")
        server_layout = QVBoxLayout(server_group)

        # Host URL input
        h_layout = QHBoxLayout()
        h_layout.addWidget(QLabel("Server URL:"))
        self.host_input = QLineEdit()
        self.host_input.setPlaceholderText("http://localhost:11434")
        h_layout.addWidget(self.host_input, 1)

        self.test_btn = QPushButton("🔌 Test")
        self.test_btn.setToolTip("Test connection to the Ollama server")
        self.test_btn.clicked.connect(self._test_connection)
        h_layout.addWidget(self.test_btn)
        server_layout.addLayout(h_layout)

        # Proxy mode combo
        p_layout = QHBoxLayout()
        p_layout.addWidget(QLabel("Proxy Routing:"))
        self.proxy_combo = QComboBox()
        self.proxy_combo.addItem("Direct (Bypass Proxy - Recommended)", "direct")
        self.proxy_combo.addItem("Auto-Detect (Direct for LAN, Proxy for Remote)", "auto")
        self.proxy_combo.addItem("Route via App Proxy", "app_proxy")
        p_layout.addWidget(self.proxy_combo, 1)
        server_layout.addLayout(p_layout)

        self.status_label = QLabel("Status: Not checked")
        self.status_label.setStyleSheet("color: #64748B; font-size: 11px; font-weight: bold;")
        server_layout.addWidget(self.status_label)

        layout.addWidget(server_group)

        # Model Selection Group
        model_group = QGroupBox("Model Preference")
        model_layout = QHBoxLayout(model_group)

        model_layout.addWidget(QLabel("Active Model:"))
        self.model_combo = QComboBox()
        self.model_combo.setMinimumWidth(240)
        model_layout.addWidget(self.model_combo, 1)

        layout.addWidget(model_group)

        # Action Buttons
        button_box = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self._save_and_close)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

    def _load_current_values(self):
        self.host_input.setText(self.cfg.get("host", "http://localhost:11434"))

        # Set proxy mode
        current_proxy_mode = self.cfg.get("proxy_mode", "direct")
        idx = self.proxy_combo.findData(current_proxy_mode)
        if idx >= 0:
            self.proxy_combo.setCurrentIndex(idx)

        # Add currently configured model name
        current_model = self.cfg.get("selected_model", "")
        if current_model:
            self.model_combo.addItem(current_model)
            self.model_combo.setCurrentText(current_model)

        # Immediate probe
        self._test_connection()

    def _test_connection(self):
        host = self.host_input.text().strip() or "http://localhost:11434"
        proxy_mode = self.proxy_combo.currentData()
        self.client.reconfigure(host, proxy_mode)

        self.status_label.setText("⏳ Testing connection...")
        self.status_label.setStyleSheet("color: #D97706; font-weight: bold;")
        self.test_btn.setEnabled(False)

        is_online, models, err = self.client.ping_and_get_models()
        self.test_btn.setEnabled(True)

        if is_online:
            self.status_label.setText(f"🟢 Connected! Found {len(models)} installed model(s).")
            self.status_label.setStyleSheet("color: #166534; font-weight: bold;")

            # Refresh model dropdown
            curr_selected = self.model_combo.currentText()
            self.model_combo.clear()
            for m in models:
                self.model_combo.addItem(m)

            if curr_selected in models:
                self.model_combo.setCurrentText(curr_selected)
            elif models:
                self.model_combo.setCurrentIndex(0)
        else:
            self.status_label.setText(f"🔴 Offline: {err}")
            self.status_label.setStyleSheet("color: #DC2626; font-weight: bold;")

    def _save_and_close(self):
        host = self.host_input.text().strip() or "http://localhost:11434"
        selected_model = self.model_combo.currentText().strip()
        proxy_mode = self.proxy_combo.currentData()

        self.cfg["host"] = host
        self.cfg["selected_model"] = selected_model
        self.cfg["proxy_mode"] = proxy_mode

        save_ollama_config(self.cfg)
        self.settings_saved.emit()
        self.accept()