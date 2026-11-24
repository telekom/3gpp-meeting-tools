"""UI construction mixin for `OllamaConfigDialog`."""

import webbrowser

from PyQt5.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QDialogButtonBox,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QPushButton,
    QSpinBox,
    QTableWidget,
    QVBoxLayout,
)

from core.ui.theme import C


class OllamaUiSetupMixin:
    """Builds the grouped UI controls for Ollama daemon, server, model, diagnostics, and GPU settings."""

    def _build_service_group(self) -> QGroupBox:
        self.service_group = QGroupBox("🖥️ Local Service Control")
        service_layout = QVBoxLayout(self.service_group)
        service_layout.setSpacing(8)

        status_row = QHBoxLayout()
        status_row.addWidget(QLabel("Daemon Status:"))
        self.service_status_label = QLabel("Checking...")
        self.service_status_label.setStyleSheet("font-weight: bold; font-size: 11px;")
        status_row.addWidget(self.service_status_label, 1)

        self.start_btn = QPushButton("▶️ Start Service")
        self.start_btn.setToolTip("Start local Ollama daemon (ollama serve)")
        self.start_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {C.SUCCESS_SOFT_ALT}; color: {C.SUCCESS_DARK}; font-weight: bold;
                border: 1px solid {C.SUCCESS_BORDER_MEDIUM}; border-radius: 4px; padding: 4px 12px;
            }}
            QPushButton:hover {{ background-color: {C.SUCCESS_BORDER}; }}
            QPushButton:disabled {{ background-color: {C.BG_CANVAS}; color: {C.TEXT_FAINT}; border-color: {C.BORDER_SUBTLE}; }}
        """)
        self.start_btn.clicked.connect(lambda: self._execute_service_action("start"))
        status_row.addWidget(self.start_btn)

        self.stop_btn = QPushButton("⏹️ Stop Service")
        self.stop_btn.setToolTip("Stop running local Ollama processes")
        self.stop_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {C.DANGER_SOFT_HOVER}; color: {C.DANGER_DARK}; font-weight: bold;
                border: 1px solid {C.DANGER_BORDER_ALT}; border-radius: 4px; padding: 4px 12px;
            }}
            QPushButton:hover {{ background-color: {C.DANGER_BORDER}; }}
            QPushButton:disabled {{ background-color: {C.BG_CANVAS}; color: {C.TEXT_FAINT}; border-color: {C.BORDER_SUBTLE}; }}
        """)
        self.stop_btn.clicked.connect(lambda: self._execute_service_action("stop"))
        status_row.addWidget(self.stop_btn)

        self.restart_btn = QPushButton("🔄 Restart")
        self.restart_btn.setToolTip("Restart the local Ollama daemon")
        self.restart_btn.setStyleSheet("padding: 4px 10px;")
        self.restart_btn.clicked.connect(lambda: self._execute_service_action("restart"))
        status_row.addWidget(self.restart_btn)
        service_layout.addLayout(status_row)

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
        return self.service_group

    def _build_server_group(self) -> QGroupBox:
        server_group = QGroupBox("🌐 Server && Network")
        server_layout = QVBoxLayout(server_group)
        server_layout.setSpacing(8)

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

        p_layout = QHBoxLayout()
        p_layout.addWidget(QLabel("Proxy Routing:"))
        self.proxy_combo = QComboBox()
        self.proxy_combo.addItem("Direct (Bypass Proxy - Recommended)", "direct")
        self.proxy_combo.addItem("Auto-Detect (Direct for LAN, Proxy for Remote)", "auto")
        self.proxy_combo.addItem("Route via App Proxy", "app_proxy")
        p_layout.addWidget(self.proxy_combo, 1)
        server_layout.addLayout(p_layout)

        self.status_label = QLabel("Status: Not checked")
        self.status_label.setStyleSheet(f"color: {C.TEXT_MUTED}; font-size: 11px; font-weight: bold;")
        server_layout.addWidget(self.status_label)
        return server_group

    def _build_model_group(self) -> QGroupBox:
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
        return model_group

    def _build_diagnostics_group(self) -> QGroupBox:
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
        self.acceleration_label.setStyleSheet(f"color: {C.TEXT_SECONDARY}; font-size: 11px;")
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
        diagnostics_note.setStyleSheet(f"color: {C.TEXT_MUTED}; font-size: 10px;")
        diagnostics_layout.addWidget(diagnostics_note)
        return diagnostics_group

    def _build_advanced_group(self) -> QGroupBox:
        adv_group = QGroupBox("⚙️ Advanced Settings")
        adv_layout = QVBoxLayout(adv_group)

        runtime_row = QHBoxLayout()
        runtime_row.addWidget(QLabel("VRAM Keep-Alive:"))
        self.keep_alive_combo = QComboBox()
        self.keep_alive_combo.addItem("5 minutes (Default)", "5m")
        self.keep_alive_combo.addItem("10 minutes", "10m")
        self.keep_alive_combo.addItem("30 minutes", "30m")
        self.keep_alive_combo.addItem("1 hour", "1h")
        self.keep_alive_combo.addItem("Indefinite (-1)", "-1")
        self.keep_alive_combo.addItem("Unload immediately (0s)", "0s")
        runtime_row.addWidget(self.keep_alive_combo)

        runtime_row.addWidget(QLabel("Timeout (s):"))
        self.timeout_spin = QSpinBox()
        self.timeout_spin.setRange(10, 600)
        self.timeout_spin.setValue(60)
        runtime_row.addWidget(self.timeout_spin)

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
        gpu_note.setStyleSheet(f"color: {C.TEXT_MUTED}; font-size: 10px;")
        adv_layout.addWidget(gpu_note)
        return adv_group

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(12)

        title = QLabel("🦙 Local LLM Server (Ollama)")
        title.setStyleSheet(f"font-size: 14px; font-weight: bold; color: {C.TEXT_STRONG};")
        layout.addWidget(title)

        desc = QLabel(
            "Manage your local Ollama instance for private, offline document summaries, "
            "PlantUML diagram synthesis, and semantic specification queries."
        )
        desc.setWordWrap(True)
        desc.setStyleSheet(f"color: {C.TEXT_MUTED}; font-size: 11px; margin-bottom: 2px;")
        layout.addWidget(desc)

        layout.addWidget(self._build_service_group())
        layout.addWidget(self._build_server_group())
        layout.addWidget(self._build_model_group())
        layout.addWidget(self._build_diagnostics_group())
        layout.addWidget(self._build_advanced_group())

        button_box = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self._save_and_close)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
