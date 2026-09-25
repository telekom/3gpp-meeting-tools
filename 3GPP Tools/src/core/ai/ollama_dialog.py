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
    QSpinBox, QFileDialog, QMessageBox, QFrame, QCheckBox, QTableWidget,
    QTableWidgetItem, QHeaderView, QAbstractItemView
)

from core.ai.ollama_client import (
    OllamaClient,
    OllamaServiceWorker,
    OllamaDiagnosticsWorker,
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
        self.diagnostics_worker = None
        self._close_after_service_action = False

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
        # 4. RUNTIME DIAGNOSTICS GROUP
        # ==========================================
        diagnostics_group = QGroupBox("📊 Ollama Runtime Diagnostics")
        diagnostics_layout = QVBoxLayout(diagnostics_group)
        diagnostics_layout.setSpacing(7)

        summary_row = QHBoxLayout()
        self.version_label = QLabel("Version: Checking...")
        self.version_label.setStyleSheet("font-weight: bold;")
        summary_row.addWidget(self.version_label)

        self.runtime_label = QLabel("Runtime: Checking...")
        self.runtime_label.setStyleSheet("font-weight: bold;")
        summary_row.addWidget(self.runtime_label, 1)

        self.refresh_diagnostics_btn = QPushButton("🔄 Refresh Diagnostics")
        self.refresh_diagnostics_btn.setToolTip(
            "Refresh Ollama version, installed-model metadata, and actual CPU/GPU model placement"
        )
        self.refresh_diagnostics_btn.clicked.connect(self._refresh_diagnostics)
        summary_row.addWidget(self.refresh_diagnostics_btn)
        diagnostics_layout.addLayout(summary_row)

        self.acceleration_label = QLabel("Acceleration: Checking...")
        self.acceleration_label.setWordWrap(True)
        self.acceleration_label.setStyleSheet("color: #475569; font-size: 11px;")
        diagnostics_layout.addWidget(self.acceleration_label)

        self.models_table = QTableWidget(0, 6)
        self.models_table.setHorizontalHeaderLabels(
            ["Model", "Size", "Parameters", "Quantization", "Loaded", "Processor"]
        )
        self.models_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.models_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.models_table.setAlternatingRowColors(True)
        self.models_table.verticalHeader().setVisible(False)
        self.models_table.setMinimumHeight(120)
        header = self.models_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        for column in range(1, 6):
            header.setSectionResizeMode(column, QHeaderView.ResizeToContents)
        diagnostics_layout.addWidget(self.models_table)

        diagnostics_note = QLabel(
            "Processor shows observed placement for currently loaded models. "
            "It is derived from Ollama runtime VRAM placement, not from the configured GPU checkboxes."
        )
        diagnostics_note.setWordWrap(True)
        diagnostics_note.setStyleSheet("color: #64748B; font-size: 10px;")
        diagnostics_layout.addWidget(diagnostics_note)

        layout.addWidget(diagnostics_group)

        # ==========================================
        # 5. ADVANCED SETTINGS GROUP
        # ==========================================
        adv_group = QGroupBox("⚙️ Advanced Settings")
        adv_layout = QVBoxLayout(adv_group)

        runtime_row = QHBoxLayout()

        # Keep-Alive
        runtime_row.addWidget(QLabel("VRAM Keep-Alive:"))
        self.keep_alive_combo = QComboBox()
        self.keep_alive_combo.addItem("5 minutes (Default)", "5m")
        self.keep_alive_combo.addItem("10 minutes", "10m")
        self.keep_alive_combo.addItem("30 minutes", "30m")
        self.keep_alive_combo.addItem("1 hour", "1h")
        self.keep_alive_combo.addItem("Indefinite (-1)", "-1")
        self.keep_alive_combo.addItem("Unload immediately (0s)", "0s")
        runtime_row.addWidget(self.keep_alive_combo)

        # Request Timeout
        runtime_row.addWidget(QLabel("Timeout (s):"))
        self.timeout_spin = QSpinBox()
        self.timeout_spin.setRange(10, 600)
        self.timeout_spin.setValue(60)
        runtime_row.addWidget(self.timeout_spin)

        # Check interval
        runtime_row.addWidget(QLabel("Heartbeat (s):"))
        self.interval_spin = QSpinBox()
        self.interval_spin.setRange(5, 60)
        self.interval_spin.setValue(10)
        runtime_row.addWidget(self.interval_spin)
        adv_layout.addLayout(runtime_row)

        hardware_row = QHBoxLayout()
        hardware_label = QLabel("Local GPU:")
        hardware_label.setToolTip(
            "These settings affect only a local Ollama daemon started/restarted by 3GPP Tools."
        )
        hardware_row.addWidget(hardware_label)

        self.vulkan_checkbox = QCheckBox("Enable Vulkan")
        self.vulkan_checkbox.setToolTip(
            "Sets OLLAMA_VULKAN=1 when 3GPP Tools starts the local Ollama service. "
            "Useful for Intel/AMD Vulkan acceleration."
        )
        hardware_row.addWidget(self.vulkan_checkbox)

        self.igpu_checkbox = QCheckBox("Allow integrated GPU (iGPU)")
        self.igpu_checkbox.setToolTip(
            "Sets OLLAMA_IGPU_ENABLE=1 and enables Vulkan when 3GPP Tools starts "
            "the local Ollama service. Required on systems such as Intel Arc integrated GPUs."
        )
        self.igpu_checkbox.toggled.connect(self._on_igpu_toggled)
        hardware_row.addWidget(self.igpu_checkbox)
        hardware_row.addStretch()
        adv_layout.addLayout(hardware_row)

        gpu_note = QLabel(
            "GPU settings apply when 3GPP Tools starts/restarts a local Ollama service. "
            "Restart Ollama after changing them. Existing Windows environment variables remain valid."
        )
        gpu_note.setWordWrap(True)
        gpu_note.setStyleSheet("color: #64748B; font-size: 10px;")
        adv_layout.addWidget(gpu_note)

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

        # Local GPU daemon overrides. Defaults preserve previous behavior.
        self.vulkan_checkbox.setChecked(bool(self.cfg.get("enable_vulkan", False)))
        self.igpu_checkbox.setChecked(bool(self.cfg.get("enable_igpu", False)))
        self._on_igpu_toggled(self.igpu_checkbox.isChecked())

        # Refresh local service and API state
        self._refresh_service_status()
        self._test_connection()
        self._refresh_diagnostics()

    def _on_igpu_toggled(self, checked: bool):
        # Ollama's integrated-GPU path uses Vulkan. Keep the relationship
        # explicit in the UI rather than saving a contradictory configuration.
        if checked:
            self.vulkan_checkbox.setChecked(True)
            self.vulkan_checkbox.setEnabled(False)
        else:
            self.vulkan_checkbox.setEnabled(True)

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
            # Refresh connection, models, and runtime diagnostics
            self._test_connection()
            self._refresh_diagnostics()
            if self._close_after_service_action:
                self._close_after_service_action = False
                self.accept()
        else:
            self._close_after_service_action = False
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

    def _refresh_diagnostics(self):
        """Refreshes runtime diagnostics asynchronously to keep the dialog responsive."""
        if self.diagnostics_worker is not None and self.diagnostics_worker.isRunning():
            return

        host = self.host_input.text().strip() or "http://127.0.0.1:11434"
        proxy_mode = self.proxy_combo.currentData()
        self.client.reconfigure(host, proxy_mode)

        self.refresh_diagnostics_btn.setEnabled(False)
        self.version_label.setText("Version: Checking...")
        self.runtime_label.setText("Runtime: Checking...")
        self.acceleration_label.setText("Acceleration: Checking...")

        self.diagnostics_worker = OllamaDiagnosticsWorker(self.client, self)
        self.diagnostics_worker.finished.connect(self._on_diagnostics_finished)
        self.diagnostics_worker.finished.connect(self.diagnostics_worker.deleteLater)
        self.diagnostics_worker.start()

    def _on_diagnostics_finished(self, info: dict):
        self.refresh_diagnostics_btn.setEnabled(True)
        self.diagnostics_worker = None

        if not info.get("online"):
            error = info.get("error") or "Ollama server is not available."
            self.version_label.setText("Version: Unavailable")
            self.runtime_label.setText("Runtime: Offline")
            self.acceleration_label.setText(f"Acceleration: Unavailable ({error})")
            self.models_table.setRowCount(0)
            return

        version = info.get("version") or "Unknown"
        installed = list(info.get("models") or [])
        running = list(info.get("running_models") or [])
        running_by_name = {m.get("name"): m for m in running if m.get("name")}

        self.version_label.setText(f"Version: {version}")
        if running:
            self.runtime_label.setText(f"Runtime: {len(running)} model(s) loaded")
            processor_text = ", ".join(
                f"{m.get('name')}: {m.get('processor', 'Unknown')}" for m in running
            )
            self.acceleration_label.setText(f"Acceleration: {processor_text}")
        else:
            self.runtime_label.setText("Runtime: No model currently loaded")
            self.acceleration_label.setText(
                "Acceleration: Not observable until a model is loaded. "
                "Run a prompt and refresh diagnostics."
            )

        self.models_table.setRowCount(len(installed))
        for row, model in enumerate(installed):
            name = model.get("name", "")
            runtime = running_by_name.get(name, {})
            values = [
                name,
                model.get("size", "-"),
                model.get("parameter_size", "-"),
                model.get("quantization", "-"),
                "Yes" if runtime else "No",
                runtime.get("processor", "-") if runtime else "-",
            ]
            for column, value in enumerate(values):
                self.models_table.setItem(row, column, QTableWidgetItem(str(value)))

        error = str(info.get("error") or "").strip()
        if error:
            self.acceleration_label.setToolTip(error)
        else:
            self.acceleration_label.setToolTip("")

    def _save_and_close(self):
        host = self.host_input.text().strip() or "http://127.0.0.1:11434"
        selected_model = self.model_combo.currentText().strip()
        proxy_mode = self.proxy_combo.currentData()
        custom_bin = self.bin_path_input.text().strip()
        keep_alive = self.keep_alive_combo.currentData()
        timeout = self.timeout_spin.value()
        interval = self.interval_spin.value()
        enable_vulkan = self.vulkan_checkbox.isChecked()
        enable_igpu = self.igpu_checkbox.isChecked()
        old_enable_vulkan = bool(self.cfg.get("enable_vulkan", False))
        old_enable_igpu = bool(self.cfg.get("enable_igpu", False))

        self.cfg["host"] = host
        self.cfg["selected_model"] = selected_model
        self.cfg["proxy_mode"] = proxy_mode
        self.cfg["custom_binary_path"] = custom_bin
        self.cfg["keep_alive"] = keep_alive
        self.cfg["timeout"] = timeout
        self.cfg["check_interval"] = interval
        self.cfg["enable_vulkan"] = enable_vulkan
        self.cfg["enable_igpu"] = enable_igpu

        save_ollama_config(self.cfg)
        self.settings_saved.emit()

        hardware_changed = (
            enable_vulkan != old_enable_vulkan
            or enable_igpu != old_enable_igpu
        )
        if hardware_changed and is_local_host(host) and is_ollama_port_open(host):
            answer = QMessageBox.question(
                self,
                "Restart Ollama?",
                "The local GPU settings changed. They take effect when the Ollama "
                "daemon is restarted.\n\nRestart Ollama now?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes,
            )
            if answer == QMessageBox.Yes:
                self._close_after_service_action = True
                self._execute_service_action("restart")
                return

        self.accept()