# --- File: src/core/ai/ollama_dialog.py ---
"""
Configuration dialog for Ollama local LLM connection, proxy mode, active models,
and in-app daemon lifecycle control (start/stop/restart).
"""
import os
import webbrowser
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QComboBox, QGroupBox, QDialogButtonBox,
    QSpinBox, QFileDialog, QMessageBox, QFrame
)

from core.ai.ollama_client import (
    OllamaClient,
    OllamaServiceWorker,
    load_ollama_config,
    save_ollama_config,
    find_ollama_binary,
    is_local_host,
    is_ollama_port_open
)


class OllamaConfigDialog(QDialog):
    settings_saved = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("🦙 Ollama LLM Configuration")
        self.setMinimumWidth(560)
        self.setStyleSheet("background-color: #FAFAFA;")

        self.cfg = load_ollama_config()
        self.client = OllamaClient(
            host=self.cfg.get("host"),
            proxy_mode=self.cfg.get("proxy_mode", "direct")
        )
        self.service_worker = None

        self._setup_ui()
        self._load_current_values()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(12)

        # Title & Description
        title = QLabel("🦙 Local LLM Server (Ollama)")
        title.setStyleSheet("font-size: 14px; font-weight: bold; color: #1E293B;")
        layout.addWidget(title)

        desc = QLabel(
            "Manage your local Ollama instance for private, offline document summaries, "
            "PlantUML diagram synthesis, and semantic specification queries."
        )
        desc.setWordWrap(True)
        desc.setStyleSheet("color: #64748B; font-size: 11px; margin-bottom: 2px;")
        layout.addWidget(desc)

        # ==========================================
        # 1. LOCAL SERVICE CONTROL GROUP
        # ==========================================
        self.service_group = QGroupBox("🖥️ Local Service Control")
        service_layout = QVBoxLayout(self.service_group)
        service_layout.setSpacing(8)

        # Status row
        status_row = QHBoxLayout()
        status_row.addWidget(QLabel("Daemon Status:"))
        self.service_status_label = QLabel("Checking...")
        self.service_status_label.setStyleSheet("font-weight: bold; font-size: 11px;")
        status_row.addWidget(self.service_status_label, 1)

        # Service Action Buttons
        self.start_btn = QPushButton("▶️ Start Service")
        self.start_btn.setToolTip("Start local Ollama daemon (ollama serve)")
        self.start_btn.setStyleSheet("""
            QPushButton {
                background-color: #DCFCE7; color: #166534; font-weight: bold;
                border: 1px solid #86EFAC; border-radius: 4px; padding: 4px 12px;
            }
            QPushButton:hover { background-color: #BBF7D0; }
            QPushButton:disabled { background-color: #F1F5F9; color: #94A3B8; border-color: #E2E8F0; }
        """)
        self.start_btn.clicked.connect(lambda: self._execute_service_action("start"))
        status_row.addWidget(self.start_btn)

        self.stop_btn = QPushButton("⏹️ Stop Service")
        self.stop_btn.setToolTip("Stop running local Ollama processes")
        self.stop_btn.setStyleSheet("""
            QPushButton {
                background-color: #FEE2E2; color: #991B1B; font-weight: bold;
                border: 1px solid #FCA5A5; border-radius: 4px; padding: 4px 12px;
            }
            QPushButton:hover { background-color: #FECACA; }
            QPushButton:disabled { background-color: #F1F5F9; color: #94A3B8; border-color: #E2E8F0; }
        """)
        self.stop_btn.clicked.connect(lambda: self._execute_service_action("stop"))
        status_row.addWidget(self.stop_btn)

        self.restart_btn = QPushButton("🔄 Restart")
        self.restart_btn.setToolTip("Restart the local Ollama daemon")
        self.restart_btn.setStyleSheet("padding: 4px 10px;")
        self.restart_btn.clicked.connect(lambda: self._execute_service_action("restart"))
        status_row.addWidget(self.restart_btn)

        service_layout.addLayout(status_row)

        # Executable binary path row
        bin_row = QHBoxLayout()
        bin_row.addWidget(QLabel("Executable:"))
        self.bin_path_input = QLineEdit()
        self.bin_path_input.setPlaceholderText("Auto-detected via PATH or standard directories")
        bin_row.addWidget(self.bin_path_input, 1)

        self.browse_bin_btn = QPushButton("📁 Browse...")
        self.browse_bin_btn.clicked.connect(self._browse_binary)
        bin_row.addWidget(self.browse_bin_btn)

        self.download_btn = QPushButton("🌐 Get Ollama")
        self.download_btn.setToolTip("Open ollama.com in your web browser")
        self.download_btn.clicked.connect(lambda: webbrowser.open("https://ollama.com"))
        bin_row.addWidget(self.download_btn)

        service_layout.addLayout(bin_row)
        layout.addWidget(self.service_group)

        # ==========================================
        # 2. SERVER & NETWORK GROUP
        # ==========================================
        server_group = QGroupBox("🌐 Server & Network")
        server_layout = QVBoxLayout(server_group)
        server_layout.setSpacing(8)

        # Host URL input
        h_layout = QHBoxLayout()
        h_layout.addWidget(QLabel("Server URL:"))
        self.host_input = QLineEdit()
        self.host_input.setPlaceholderText("http://127.0.0.1:11434")
        self.host_input.textChanged.connect(self._on_host_changed)
        h_layout.addWidget(self.host_input, 1)

        self.test_btn = QPushButton("🔌 Test Connection")
        self.test_btn.setToolTip("Test connection to the Ollama REST API")
        self.test_btn.setStyleSheet("font-weight: bold; padding: 4px 12px;")
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

        # ==========================================
        # 3. MODEL SELECTION GROUP
        # ==========================================
        model_group = QGroupBox("🦙 Model Preferences")
        model_layout = QHBoxLayout(model_group)

        model_layout.addWidget(QLabel("Active Model:"))
        self.model_combo = QComboBox()
        self.model_combo.setMinimumWidth(220)
        model_layout.addWidget(self.model_combo, 1)

        self.refresh_models_btn = QPushButton("🔄 Refresh")
        self.refresh_models_btn.setToolTip("Reload installed models from Ollama server")
        self.refresh_models_btn.clicked.connect(self._test_connection)
        model_layout.addWidget(self.refresh_models_btn)

        self.library_btn = QPushButton("🌐 Model Library")
        self.library_btn.setToolTip("Browse and discover models on ollama.com/library")
        self.library_btn.clicked.connect(lambda: webbrowser.open("https://ollama.com/library"))
        model_layout.addWidget(self.library_btn)

        layout.addWidget(model_group)

        # ==========================================
        # 4. ADVANCED SETTINGS GROUP
        # ==========================================
        adv_group = QGroupBox("⚙️ Advanced Settings")
        adv_layout = QHBoxLayout(adv_group)

        # Keep-Alive
        adv_layout.addWidget(QLabel("VRAM Keep-Alive:"))
        self.keep_alive_combo = QComboBox()
        self.keep_alive_combo.addItem("5 minutes (Default)", "5m")
        self.keep_alive_combo.addItem("10 minutes", "10m")
        self.keep_alive_combo.addItem("30 minutes", "30m")
        self.keep_alive_combo.addItem("1 hour", "1h")
        self.keep_alive_combo.addItem("Indefinite (-1)", "-1")
        self.keep_alive_combo.addItem("Unload immediately (0s)", "0s")
        adv_layout.addWidget(self.keep_alive_combo)

        # Request Timeout
        adv_layout.addWidget(QLabel("Timeout (s):"))
        self.timeout_spin = QSpinBox()
        self.timeout_spin.setRange(10, 600)
        self.timeout_spin.setValue(60)
        adv_layout.addWidget(self.timeout_spin)

        # Check interval
        adv_layout.addWidget(QLabel("Heartbeat (s):"))
        self.interval_spin = QSpinBox()
        self.interval_spin.setRange(5, 60)
        self.interval_spin.setValue(10)
        adv_layout.addWidget(self.interval_spin)

        layout.addWidget(adv_group)

        # Action Buttons
        button_box = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self._save_and_close)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

    def _load_current_values(self):
        # Host URL
        host_val = self.cfg.get("host", "http://127.0.0.1:11434")
        self.host_input.setText(host_val)

        # Proxy mode
        current_proxy_mode = self.cfg.get("proxy_mode", "direct")
        idx = self.proxy_combo.findData(current_proxy_mode)
        if idx >= 0:
            self.proxy_combo.setCurrentIndex(idx)

        # Binary path
        saved_bin = self.cfg.get("custom_binary_path", "")
        detected_bin = find_ollama_binary(saved_bin) or ""
        self.bin_path_input.setText(detected_bin)

        # Active Model
        current_model = self.cfg.get("selected_model", "")
        if current_model:
            self.model_combo.addItem(current_model)
            self.model_combo.setCurrentText(current_model)

        # Keep-Alive
        saved_keep_alive = self.cfg.get("keep_alive", "5m")
        ka_idx = self.keep_alive_combo.findData(saved_keep_alive)
        if ka_idx >= 0:
            self.keep_alive_combo.setCurrentIndex(ka_idx)

        # Timeout & Interval
        self.timeout_spin.setValue(int(self.cfg.get("timeout", 60)))
        self.interval_spin.setValue(int(self.cfg.get("check_interval", 10)))

        # Refresh local service and API state
        self._refresh_service_status()
        self._test_connection()

    def _on_host_changed(self, text: str):
        is_local = is_local_host(text.strip())
        self.service_group.setEnabled(is_local)
        if not is_local:
            self.service_status_label.setText("⚪ Remote host (service controls disabled)")
            self.service_status_label.setStyleSheet("color: #64748B;")
        else:
            self._refresh_service_status()

    def _refresh_service_status(self):
        host = self.host_input.text().strip() or "http://127.0.0.1:11434"
        if not is_local_host(host):
            self.service_status_label.setText("⚪ Remote host (service controls disabled)")
            self.service_status_label.setStyleSheet("color: #64748B;")
            self.start_btn.setEnabled(False)
            self.stop_btn.setEnabled(False)
            self.restart_btn.setEnabled(False)
            return

        is_running = is_ollama_port_open(host)
        if is_running:
            self.service_status_label.setText("🟢 Running (Port 11434 open)")
            self.service_status_label.setStyleSheet("color: #166534; font-weight: bold;")
            self.start_btn.setEnabled(False)
            self.stop_btn.setEnabled(True)
            self.restart_btn.setEnabled(True)
        else:
            self.service_status_label.setText("🔴 Stopped")
            self.service_status_label.setStyleSheet("color: #DC2626; font-weight: bold;")
            self.start_btn.setEnabled(True)
            self.stop_btn.setEnabled(False)
            self.restart_btn.setEnabled(False)

    def _browse_binary(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Ollama Executable",
            "",
            "Executable Files (*.exe);;All Files (*.*)" if os.name == "nt" else "All Files (*)"
        )
        if file_path:
            self.bin_path_input.setText(file_path)

    def _set_service_controls_busy(self, is_busy: bool):
        self.start_btn.setEnabled(not is_busy)
        self.stop_btn.setEnabled(not is_busy)
        self.restart_btn.setEnabled(not is_busy)
        self.test_btn.setEnabled(not is_busy)
        self.browse_bin_btn.setEnabled(not is_busy)

    def _execute_service_action(self, action: str):
        self._set_service_controls_busy(True)
        custom_bin = self.bin_path_input.text().strip()
        host = self.host_input.text().strip() or "http://127.0.0.1:11434"

        self.service_worker = OllamaServiceWorker(
            action=action,
            custom_path=custom_bin,
            host=host,
            parent=self
        )
        self.service_worker.progress_status.connect(
            lambda msg: self.service_status_label.setText(f"⏳ {msg}")
        )
        self.service_worker.finished.connect(self._on_service_action_finished)
        self.service_worker.start()

    def _on_service_action_finished(self, success: bool, message: str):
        self._set_service_controls_busy(False)
        self._refresh_service_status()

        if success:
            # Refresh connection and models
            self._test_connection()
        else:
            QMessageBox.warning(self, "Service Management", f"Action failed:\n{message}")

    def _test_connection(self):
        host = self.host_input.text().strip() or "http://127.0.0.1:11434"
        proxy_mode = self.proxy_combo.currentData()
        self.client.reconfigure(host, proxy_mode)

        self.status_label.setText("⏳ Testing connection...")
        self.status_label.setStyleSheet("color: #D97706; font-weight: bold;")
        self.test_btn.setEnabled(False)

        is_online, models, err = self.client.ping_and_get_models()
        self.test_btn.setEnabled(True)

        if is_online:
            self.status_label.setText(f"✅ Connected! Found {len(models)} installed model(s).")
            self.status_label.setStyleSheet("color: #166534; font-weight: bold;")

            curr_selected = self.model_combo.currentText()
            self.model_combo.clear()
            for m in models:
                self.model_combo.addItem(m)

            if curr_selected in models:
                self.model_combo.setCurrentText(curr_selected)
            elif models:
                self.model_combo.setCurrentIndex(0)
        else:
            self.status_label.setText(f"❌ Offline: {err}")
            self.status_label.setStyleSheet("color: #DC2626; font-weight: bold;")

        self._refresh_service_status()

    def _save_and_close(self):
        host = self.host_input.text().strip() or "http://127.0.0.1:11434"
        selected_model = self.model_combo.currentText().strip()
        proxy_mode = self.proxy_combo.currentData()
        custom_bin = self.bin_path_input.text().strip()
        keep_alive = self.keep_alive_combo.currentData()
        timeout = self.timeout_spin.value()
        interval = self.interval_spin.value()

        self.cfg["host"] = host
        self.cfg["selected_model"] = selected_model
        self.cfg["proxy_mode"] = proxy_mode
        self.cfg["custom_binary_path"] = custom_bin
        self.cfg["keep_alive"] = keep_alive
        self.cfg["timeout"] = timeout
        self.cfg["check_interval"] = interval

        save_ollama_config(self.cfg)
        self.settings_saved.emit()
        self.accept()