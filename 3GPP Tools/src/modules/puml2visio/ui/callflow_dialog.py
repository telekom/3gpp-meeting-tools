# --- File: modules/puml2visio/ui/callflow_dialog.py ---
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPlainTextEdit,
    QPushButton, QComboBox, QSplitter, QMessageBox, QFrame
)
from PyQt5.QtCore import Qt, pyqtSignal

from core.ai.ollama_client import OllamaClient, load_ollama_config
from core.ui.ui_components import BUTTON_STYLE_TOOLBAR_SECONDARY
from modules.puml2visio.core.prompt_manager import PromptManager
from modules.puml2visio.core.callflow_thread import CallFlowGeneratorThread
from modules.puml2visio.templates.plantuml_templates import PLANTUML_TYPES

class CallFlowDialog(QDialog):
    """Modeless dialog for generating PlantUML diagrams from specification call flow text."""
    diagram_generated = pyqtSignal(str, bool)  # (puml_code, should_replace)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("🤖 Generate PlantUML from 3GPP Call Flow")
        self.resize(1000, 680)

        # 1. Ensure modeless behavior (never block other windows)
        self.setModal(False)

        # 2. Configure independent top-level window controls (Minimize, Maximize, Close)
        self.setWindowFlags(
            Qt.Window |
            Qt.WindowMinMaxButtonsHint |
            Qt.WindowCloseButtonHint
        )

        self.client = OllamaClient()
        self.prompt_mgr = PromptManager()
        self.worker: CallFlowGeneratorThread = None
        self._raw_stream_buffer = []

        self._setup_ui()
        self._refresh_models()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(14, 14, 14, 14)

        # --- Top Controls Strip ---
        top_strip = QHBoxLayout()
        top_strip.setSpacing(10)

        top_strip.addWidget(QLabel("🦙 Model:"))
        self.model_combo = QComboBox()
        self.model_combo.setMinimumWidth(180)
        top_strip.addWidget(self.model_combo)

        top_strip.addWidget(QLabel("📖 Diagram Type:"))
        self.type_combo = QComboBox()
        self.type_combo.addItems(list(PLANTUML_TYPES.keys()))
        self.type_combo.setCurrentText("Sequence")
        top_strip.addWidget(self.type_combo)

        self.refresh_btn = QPushButton("🔄 Refresh Models")
        self.refresh_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.refresh_btn.clicked.connect(self._refresh_models)
        top_strip.addWidget(self.refresh_btn)

        top_strip.addStretch()

        self.status_lbl = QLabel("Ready.")
        self.status_lbl.setStyleSheet("color: #64748B; font-weight: bold;")
        top_strip.addWidget(self.status_lbl)

        layout.addLayout(top_strip)

        # --- Middle Splitter (Input Text vs. Output Preview) ---
        splitter = QSplitter(Qt.Horizontal)

        # Left: Specification input
        left_widget = QFrame()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_lbl = QLabel("📋 Specification Procedure Text (e.g., from TS 23.502):")
        left_lbl.setStyleSheet("font-weight: bold; color: #334155;")
        self.input_text = QPlainTextEdit()
        self.input_text.setPlaceholderText(
            "Paste procedure text here...\n\n"
            "Example:\n"
            "1. UE sends Registration Request (SUCI, 5G-GUTI) to (R)AN.\n"
            "2. (R)AN selects AMF and forwards the Registration Request.\n"
            "3. AMF requests authentication from AUSF.\n"
            "4. AUSF verifies with UDM and returns result to AMF.\n"
            "5. AMF sends Registration Accept to UE."
        )
        left_layout.addWidget(left_lbl)
        left_layout.addWidget(self.input_text)
        splitter.addWidget(left_widget)

        # Right: Output preview
        right_widget = QFrame()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_lbl = QLabel("✨ Generated PlantUML Code:")
        right_lbl.setStyleSheet("font-weight: bold; color: #334155;")
        self.output_text = QPlainTextEdit()
        self.output_text.setReadOnly(False)
        self.output_text.setPlaceholderText("Generated PlantUML sequence code will stream here...")
        self.output_text.setStyleSheet(
            "font-family: Consolas, monospace; font-size: 12px; background-color: #F8FAFC;"
        )
        right_layout.addWidget(right_lbl)
        right_layout.addWidget(self.output_text)
        splitter.addWidget(right_widget)

        splitter.setSizes([450, 550])
        layout.addWidget(splitter, stretch=1)

        # --- Bottom Action Bar ---
        btn_bar = QHBoxLayout()
        btn_bar.setSpacing(8)

        self.generate_btn = QPushButton("🚀 Generate Diagram")
        self.generate_btn.setObjectName("primaryBtn")
        self.generate_btn.clicked.connect(self._start_generation)
        btn_bar.addWidget(self.generate_btn)

        self.cancel_btn = QPushButton("⏹️ Cancel")
        self.cancel_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.cancel_btn.setEnabled(False)
        self.cancel_btn.clicked.connect(self._cancel_generation)
        btn_bar.addWidget(self.cancel_btn)

        btn_bar.addStretch()

        self.replace_btn = QPushButton("📥 Replace Editor")
        self.replace_btn.setObjectName("primaryBtn")
        self.replace_btn.setEnabled(False)
        self.replace_btn.setToolTip("Overwrite the active diagram in the editor with this code.")
        self.replace_btn.clicked.connect(lambda: self._apply_and_close(replace=True))
        btn_bar.addWidget(self.replace_btn)

        self.append_btn = QPushButton("➕ Append to Editor")
        self.append_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.append_btn.setEnabled(False)
        self.append_btn.setToolTip("Append this generated code to the bottom of the active editor.")
        self.append_btn.clicked.connect(lambda: self._apply_and_close(replace=False))
        btn_bar.addWidget(self.append_btn)

        self.close_btn = QPushButton("Close")
        self.close_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.close_btn.clicked.connect(self.reject)
        btn_bar.addWidget(self.close_btn)

        layout.addLayout(btn_bar)

    def _refresh_models(self):
        """Queries the Ollama client for installed models."""
        is_online, models, err = self.client.ping_and_get_models()
        self.model_combo.clear()
        if is_online and models:
            self.model_combo.addItems(models)
            cfg = load_ollama_config()
            selected = cfg.get("selected_model", "")
            if selected in models:
                self.model_combo.setCurrentText(selected)
            self.status_lbl.setText(f"🟢 Connected ({len(models)} models)")
            self.status_lbl.setStyleSheet("color: #16A34A; font-weight: bold;")
            self.generate_btn.setEnabled(True)
        else:
            self.status_lbl.setText("🔴 Ollama Offline")
            self.status_lbl.setStyleSheet("color: #DC2626; font-weight: bold;")
            self.generate_btn.setEnabled(False)

    def _start_generation(self):
        call_flow = self.input_text.toPlainText().strip()
        if not call_flow:
            QMessageBox.warning(self, "Input Empty", "Please paste or type call flow steps first.")
            return

        model = self.model_combo.currentText()
        if not model:
            QMessageBox.warning(self, "No Model", "No active Ollama model selected.")
            return

        sys_prompt = self.prompt_mgr.get_prompt("callflow_system.txt")
        usr_template = self.prompt_mgr.get_prompt("callflow_user.txt")
        formatted_user_prompt = usr_template.format(call_flow_text=call_flow)

        self.output_text.clear()
        self._raw_stream_buffer = []
        self.status_lbl.setText("⚡ Streaming tokens...")
        self.status_lbl.setStyleSheet("color: #2563EB; font-weight: bold;")
        self.generate_btn.setEnabled(False)
        self.cancel_btn.setEnabled(True)
        self.replace_btn.setEnabled(False)
        self.append_btn.setEnabled(False)

        diag_type = self.type_combo.currentText()
        self.worker = CallFlowGeneratorThread(
            client=self.client,
            model=model,
            system_prompt=sys_prompt,
            user_prompt=formatted_user_prompt,
            diagram_type=diag_type,
            parent=self
        )
        self.worker.token_received.connect(self._handle_token)
        self.worker.generation_completed.connect(self._handle_complete)
        self.worker.generation_failed.connect(self._handle_failed)
        self.worker.start()

    def _handle_token(self, delta: str):
        self._raw_stream_buffer.append(delta)
        self.output_text.insertPlainText(delta)
        self.output_text.ensureCursorVisible()

    def _handle_complete(self, clean_puml: str):
        self.output_text.setPlainText(clean_puml)
        self.status_lbl.setText("✅ Completed")
        self.status_lbl.setStyleSheet("color: #16A34A; font-weight: bold;")
        self.generate_btn.setEnabled(True)
        self.cancel_btn.setEnabled(False)
        self.replace_btn.setEnabled(True)
        self.append_btn.setEnabled(True)

    def _handle_failed(self, error_msg: str):
        self.status_lbl.setText("⚠️ Failed / Cancelled")
        self.status_lbl.setStyleSheet("color: #DC2626; font-weight: bold;")
        self.generate_btn.setEnabled(True)
        self.cancel_btn.setEnabled(False)
        if self.output_text.toPlainText().strip():
            self.replace_btn.setEnabled(True)
            self.append_btn.setEnabled(True)

    def _cancel_generation(self):
        if self.worker and self.worker.isRunning():
            self.worker.cancel()
            self.status_lbl.setText("Cancelling...")

    def _apply_and_close(self, replace: bool):
        final_code = self.output_text.toPlainText().strip()
        if final_code:
            self.diagram_generated.emit(final_code, replace)
            self.accept()

    def closeEvent(self, event):
        if self.worker and self.worker.isRunning():
            self.worker.cancel()
            self.worker.wait(400)
        event.accept()