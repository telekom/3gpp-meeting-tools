# --- File: src/modules/puml2visio/ui/callflow_dialog.py ---
"""Modeless dialog for AI-powered PlantUML sequence diagram generation from 3GPP call flow texts.

Features:
- Streaming token synthesis with real-time PlantUML syntax highlighting
- 3GPP Generation Style Toggles (Box groupings, IEs/Parameters, State transitions, Alt branches)
- Pre-loaded standard 3GPP procedure presets (Registration, PDU Session, Handover, Service Request)
- In-dialog Live SVG Preview dock with smooth zoom and pan
- Automated PlantUML syntax error detection and one-click AI Auto-Fix
- Conversational iterative refinement bar for follow-up prompt modifications
"""

from pathlib import Path
from typing import Optional, List, Dict
import re

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtWidgets import (
    QAction,
    QApplication,
    QCheckBox,
    QComboBox,
    QDialog,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMenu,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QSplitter,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from core.ai.ollama_client import OllamaClient, load_ollama_config
from core.ui.ui_components import BUTTON_STYLE_TOOLBAR_SECONDARY, CARD_FRAME_STYLE
from modules.puml2visio.config.paths import PLANTUML_JAR_NAME
from modules.puml2visio.core.callflow_thread import (
    CallFlowGeneratorThread,
    CallFlowValidationThread,
)
from modules.puml2visio.core.prompt_manager import PromptManager
from modules.puml2visio.templates.plantuml_snippets import SAMPLE_3GPP_PROCEDURES
from modules.puml2visio.templates.plantuml_templates import PLANTUML_TYPES
from modules.puml2visio.ui.svg_preview_widget import EmbeddedSvgPreview
from modules.puml2visio.ui.syntax_highlighter import PlantUMLHighlighter
from modules.puml2visio.utils.paths import get_puml2visio_asset_path


class CallFlowDialog(QDialog):
    """Modeless dialog for generating PlantUML diagrams from specification call flow text."""
    diagram_generated = pyqtSignal(str, bool)  # (puml_code, should_replace)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("🤖 Generate PlantUML from 3GPP Call Flow")
        self.resize(1120, 740)

        # 1. Modeless behavior (does not block other application windows)
        self.setModal(False)

        # 2. Top-level window controls (Minimize, Maximize, Close)
        self.setWindowFlags(
            Qt.Window |
            Qt.WindowMinMaxButtonsHint |
            Qt.WindowCloseButtonHint
        )

        self.client = OllamaClient()
        self.prompt_mgr = PromptManager()
        self.jar_path = get_puml2visio_asset_path(PLANTUML_JAR_NAME)

        self.worker: Optional[CallFlowGeneratorThread] = None
        self.validator: Optional[CallFlowValidationThread] = None
        self.conversation_history: List[Dict[str, str]] = []
        self.last_clean_puml: str = ""
        self.last_syntax_error: str = ""
        self.last_error_line: int = 0

        self._setup_ui()
        self._refresh_models()

    def _setup_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(10)
        main_layout.setContentsMargins(14, 14, 14, 14)

        # ---------------------------------------------------------------------
        # TOP CONTROLS STRIP: Model, Diagram Type, Sample Presets, Status
        # ---------------------------------------------------------------------
        top_strip = QHBoxLayout()
        top_strip.setSpacing(10)

        top_strip.addWidget(QLabel("🦙 Model:"))
        self.model_combo = QComboBox()
        self.model_combo.setMinimumWidth(180)
        top_strip.addWidget(self.model_combo)

        top_strip.addWidget(QLabel("📖 Type:"))
        self.type_combo = QComboBox()
        self.type_combo.addItems(list(PLANTUML_TYPES.keys()))
        self.type_combo.setCurrentText("Sequence")
        top_strip.addWidget(self.type_combo)

        self.refresh_btn = QPushButton("🔄 Refresh")
        self.refresh_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.refresh_btn.setToolTip("Refresh list of locally installed Ollama models.")
        self.refresh_btn.clicked.connect(self._refresh_models)
        top_strip.addWidget(self.refresh_btn)

        # Sample 3GPP Procedures Dropdown Menu
        self.samples_btn = QPushButton("📚 Sample Procedures ▼")
        self.samples_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.samples_btn.setToolTip("Load standard 3GPP specification call flows (TS 23.502, TS 38.300).")
        self._build_samples_menu()
        top_strip.addWidget(self.samples_btn)

        top_strip.addStretch()

        self.status_lbl = QLabel("Ready.")
        self.status_lbl.setStyleSheet("color: #64748B; font-weight: bold;")
        top_strip.addWidget(self.status_lbl)

        main_layout.addLayout(top_strip)

        # ---------------------------------------------------------------------
        # MIDDLE SPLITTER: Input Text & Options (Left) vs Output Preview (Right)
        # ---------------------------------------------------------------------
        splitter = QSplitter(Qt.Horizontal)
        splitter.setStyleSheet("""
            QSplitter::handle {
                background-color: #E2E8F0;
                width: 3px;
                margin: 0px 2px;
            }
        """)

        # LEFT WIDGET: Specification Procedure Text & 3GPP Style Options
        left_widget = QFrame()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(8)

        input_header = QHBoxLayout()
        left_lbl = QLabel("📋 Specification Procedure Text:")
        left_lbl.setStyleSheet("font-weight: bold; color: #1E293B; font-size: 12px;")
        input_header.addWidget(left_lbl)
        input_header.addStretch()

        self.paste_btn = QPushButton("📋 Paste Clipboard")
        self.paste_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.paste_btn.clicked.connect(self._paste_clipboard)
        input_header.addWidget(self.paste_btn)

        self.clear_input_btn = QPushButton("🗑️ Clear")
        self.clear_input_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.clear_input_btn.clicked.connect(self.input_text_clear)
        input_header.addWidget(self.clear_input_btn)

        left_layout.addLayout(input_header)

        self.input_text = QPlainTextEdit()
        self.input_text.setPlaceholderText(
            "Paste procedure steps here, or pick a standard preset from '📚 Sample Procedures ▼' above...\n\n"
            "Example:\n"
            "1. UE sends Registration Request (SUCI, 5G-GUTI) to (R)AN.\n"
            "2. (R)AN selects AMF and forwards the Registration Request.\n"
            "3. AMF requests authentication from AUSF.\n"
            "4. AUSF verifies with UDM and returns result to AMF.\n"
            "5. AMF sends Registration Accept to UE."
        )
        self.input_text.setStyleSheet("""
            QPlainTextEdit {
                font-family: Consolas, 'Courier New', monospace;
                font-size: 12px;
                background-color: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 6px;
                line-height: 1.4;
            }
            QPlainTextEdit:focus {
                border-color: #395396;
            }
        """)
        left_layout.addWidget(self.input_text, stretch=1)

        # 3GPP Generation Style Toggles (Prompt Modifiers)
        style_box = QGroupBox("⚙️ 3GPP Diagram Style & Detail Options")
        style_box.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                color: #334155;
                border: 1px solid #CBD5E1;
                border-radius: 6px;
                margin-top: 8px;
                padding-top: 10px;
                background-color: #F8FAFC;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 4px;
            }
            QCheckBox {
                font-size: 11px;
                color: #475569;
                font-weight: 500;
            }
        """)
        style_grid = QGridLayout(style_box)
        style_grid.setContentsMargins(10, 10, 10, 10)
        style_grid.setSpacing(6)

        self.opt_boxes = QCheckBox("📦 Group NFs into Boxes (RAN / 5GC / User Plane)")
        self.opt_boxes.setChecked(True)
        self.opt_boxes.setToolTip("Enclose participants in colored visual boxes (RAN, Core Control Plane, User Plane).")

        self.opt_ies = QCheckBox("📑 Extract Key Message IEs & Parameters")
        self.opt_ies.setChecked(True)
        self.opt_ies.setToolTip("Extract key parameters (e.g. SUCI, NSSAI, PDU Session ID) into message labels or notes.")

        self.opt_state = QCheckBox("🏷️ Include RRC / NAS States (hnote)")
        self.opt_state.setToolTip("Insert hexagonal state notes (e.g. 5GMM-REGISTERED) before and after procedures.")

        self.opt_alt = QCheckBox("🔀 Model Negative / Failure Branches (alt)")
        self.opt_alt.setChecked(True)
        self.opt_alt.setToolTip("Model alternative branches and rejection cases using 'alt ... else ... end' blocks.")

        self.opt_autonumber = QCheckBox("🔢 Autonumber Steps [01]")
        self.opt_autonumber.setChecked(True)
        self.opt_autonumber.setToolTip("Inject 'autonumber 1 1 \"<b>[00]</b>\"' to automatically number message exchanges.")

        style_grid.addWidget(self.opt_boxes, 0, 0)
        style_grid.addWidget(self.opt_ies, 0, 1)
        style_grid.addWidget(self.opt_state, 1, 0)
        style_grid.addWidget(self.opt_alt, 1, 1)
        style_grid.addWidget(self.opt_autonumber, 2, 0, 1, 2)

        left_layout.addWidget(style_box)
        splitter.addWidget(left_widget)

        # RIGHT WIDGET: Output Tabs (Code vs Rendered SVG) + Refinement Bar
        right_widget = QFrame()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(6)

        # Output Tab Widget (Code vs Diagram Preview)
        self.output_tabs = QTabWidget()
        self.output_tabs.setStyleSheet("""
            QTabWidget::pane {
                border: 1px solid #CBD5E1;
                border-radius: 6px;
                background-color: #FFFFFF;
            }
            QTabBar::tab {
                font-weight: bold;
                padding: 6px 14px;
            }
        """)

        # Tab 1: Code View
        code_tab = QWidget()
        code_layout = QVBoxLayout(code_tab)
        code_layout.setContentsMargins(4, 4, 4, 4)

        self.output_text = QPlainTextEdit()
        self.output_text.setReadOnly(False)
        self.output_text.setPlaceholderText("Generated PlantUML sequence code will stream here...")
        self.output_text.setStyleSheet("""
            QPlainTextEdit {
                font-family: Consolas, 'Cascadia Code', 'Courier New', monospace;
                font-size: 12px;
                background-color: #FFFFFF;
                border: none;
                color: #1E293B;
                line-height: 1.4;
            }
        """)
        # Attach PlantUML syntax highlighter
        self.highlighter = PlantUMLHighlighter(self.output_text.document())
        code_layout.addWidget(self.output_text)
        self.output_tabs.addTab(code_tab, "📝 Code")

        # Tab 2: Visual Diagram Preview
        self.svg_preview = EmbeddedSvgPreview()
        self.output_tabs.addTab(self.svg_preview, "👁️ Diagram Preview")

        right_layout.addWidget(self.output_tabs, stretch=1)

        # PlantUML Syntax Error Banner (Hidden by default)
        self.error_banner = QFrame()
        self.error_banner.setStyleSheet("""
            QFrame {
                background-color: #FEE2E2;
                border: 1px solid #F87171;
                border-radius: 6px;
                padding: 4px;
            }
        """)
        error_layout = QHBoxLayout(self.error_banner)
        error_layout.setContentsMargins(8, 4, 8, 4)
        error_layout.setSpacing(8)

        self.error_banner_lbl = QLabel("")
        self.error_banner_lbl.setStyleSheet("color: #B91C1C; font-weight: bold; font-size: 11px;")
        error_layout.addWidget(self.error_banner_lbl, stretch=1)

        self.autofix_btn = QPushButton("🪄 Auto-Fix Error")
        self.autofix_btn.setStyleSheet("QPushButton { background-color: #DC2626; color: white; border: none; padding: 4px 10px; border-radius: 4px; font-weight: bold; } QPushButton:hover { background-color: #B91C1C; }")
        self.autofix_btn.setToolTip("Prompt local LLM to automatically fix this PlantUML syntax error.")
        self.autofix_btn.clicked.connect(self._autofix_syntax_error)
        error_layout.addWidget(self.autofix_btn)

        self.error_banner.hide()
        right_layout.addWidget(self.error_banner)

        # Conversational Iterative Refinement Bar
        refine_bar = QHBoxLayout()
        refine_bar.setSpacing(6)

        refine_lbl = QLabel("💬 Refine:")
        refine_lbl.setStyleSheet("font-weight: bold; color: #475569; font-size: 11px;")
        refine_bar.addWidget(refine_lbl)

        self.refine_input = QLineEdit()
        self.refine_input.setPlaceholderText("e.g. Add an alt block for authentication failure, or group into VPLMN / HPLMN...")
        self.refine_input.setStyleSheet("""
            QLineEdit {
                border: 1px solid #CBD5E1;
                border-radius: 4px;
                padding: 5px 8px;
                background-color: #FFFFFF;
                font-size: 12px;
            }
            QLineEdit:focus {
                border-color: #395396;
            }
        """)
        self.refine_input.returnPressed.connect(self._refine_diagram)
        refine_bar.addWidget(self.refine_input, stretch=1)

        self.refine_btn = QPushButton("✨ Refine")
        self.refine_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.refine_btn.setEnabled(False)
        self.refine_btn.setToolTip("Send refinement instruction to model to update this diagram.")
        self.refine_btn.clicked.connect(self._refine_diagram)
        refine_bar.addWidget(self.refine_btn)

        right_layout.addLayout(refine_bar)
        splitter.addWidget(right_widget)

        splitter.setSizes([480, 600])
        main_layout.addWidget(splitter, stretch=1)

        # ---------------------------------------------------------------------
        # BOTTOM ACTION BAR: Generate, Cancel, Copy, Apply, Close
        # ---------------------------------------------------------------------
        btn_bar = QHBoxLayout()
        btn_bar.setSpacing(8)

        self.generate_btn = QPushButton("🚀 Generate Diagram")
        self.generate_btn.setObjectName("primaryBtn")
        self.generate_btn.setToolTip("Synthesize PlantUML sequence diagram from the procedure steps.")
        self.generate_btn.clicked.connect(self._start_generation)
        btn_bar.addWidget(self.generate_btn)

        self.cancel_btn = QPushButton("⏹️ Cancel")
        self.cancel_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.cancel_btn.setEnabled(False)
        self.cancel_btn.clicked.connect(self._cancel_generation)
        btn_bar.addWidget(self.cancel_btn)

        btn_bar.addStretch()

        self.copy_code_btn = QPushButton("📋 Copy Code")
        self.copy_code_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.copy_code_btn.setEnabled(False)
        self.copy_code_btn.setToolTip("Copy generated PlantUML source code to clipboard.")
        self.copy_code_btn.clicked.connect(self._copy_code)
        btn_bar.addWidget(self.copy_code_btn)

        self.replace_btn = QPushButton("📥 Replace Editor")
        self.replace_btn.setObjectName("primaryBtn")
        self.replace_btn.setEnabled(False)
        self.replace_btn.setToolTip("Overwrite the active diagram in the main IDE with this code.")
        self.replace_btn.clicked.connect(lambda: self._apply_and_close(replace=True))
        btn_bar.addWidget(self.replace_btn)

        self.append_btn = QPushButton("➕ Append to Editor")
        self.append_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.append_btn.setEnabled(False)
        self.append_btn.setToolTip("Append this generated code to the bottom of the main IDE.")
        self.append_btn.clicked.connect(lambda: self._apply_and_close(replace=False))
        btn_bar.addWidget(self.append_btn)

        self.close_btn = QPushButton("Close")
        self.close_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.close_btn.clicked.connect(self.close)
        btn_bar.addWidget(self.close_btn)

        main_layout.addLayout(btn_bar)

    # -------------------------------------------------------------------------
    # Sample Procedures Menu
    # -------------------------------------------------------------------------
    def _build_samples_menu(self):
        menu = QMenu(self)
        menu.setToolTipsVisible(True)
        for title, text in SAMPLE_3GPP_PROCEDURES.items():
            act = QAction(title.replace("&", "&&"), self)
            act.setToolTip(f"Load '{title}' specification text.")
            act.triggered.connect(lambda checked, t=text: self.input_text.setPlainText(t))
            menu.addAction(act)
        self.samples_btn.setMenu(menu)

    def _paste_clipboard(self):
        cb_text = QApplication.clipboard().text()
        if cb_text:
            self.input_text.setPlainText(cb_text.strip())

    def input_text_clear(self):
        self.input_text.clear()

    # -------------------------------------------------------------------------
    # Ollama Model Inspection
    # -------------------------------------------------------------------------
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

    # -------------------------------------------------------------------------
    # Generation & Prompt Assembly
    # -------------------------------------------------------------------------
    def _build_directive_prompt_additions(self) -> str:
        """Assembles extra prompt steering instructions based on user style checkboxes."""
        directives = []
        if self.opt_boxes.isChecked():
            directives.append(
                "Organize participants into visual boxes where applicable: e.g. box \"Radio Access Network\" #F8FAFC ... end box, "
                "box \"5G Core Control Plane\" #F1F5F9 ... end box, and box \"User Plane\" #F8FAFC ... end box."
            )
        if self.opt_ies.isChecked():
            directives.append(
                "Extract and show key Information Elements (IEs) and parameters inside message descriptions or attached notes "
                "(e.g., SUCI, 5G-GUTI, S-NSSAI, PDU Session ID, Cause values)."
            )
        if self.opt_state.isChecked():
            directives.append(
                "Include hexagonal state notes (e.g., 'hnote over UE: 5GMM-DEREGISTERED' and 'hnote over UE: 5GMM-REGISTERED') "
                "to illustrate RRC/NAS state transitions before and after major procedures."
            )
        if self.opt_alt.isChecked():
            directives.append(
                "If the text describes alternative outcomes, rejections, integrity failures, or timer expiries, model them cleanly using 'alt ... else ... end' blocks."
            )
        if self.opt_autonumber.isChecked():
            directives.append("Include step autonumbering with: autonumber 1 1 \"<b>[00]</b>\"")
        else:
            directives.append("Do not inject autonumber; use explicit step numbering from the text if present.")

        if directives:
            return "\n\nAdditional Technical Directives:\n" + "\n".join(f"- {d}" for d in directives)
        return ""

    def _start_generation(self):
        call_flow = self.input_text.toPlainText().strip()
        if not call_flow:
            QMessageBox.warning(self, "Input Empty", "Please paste or select call flow procedure steps first.")
            return

        model = self.model_combo.currentText()
        if not model:
            QMessageBox.warning(self, "No Model", "No active Ollama model selected.")
            return

        sys_prompt = self.prompt_mgr.get_prompt("callflow_system.txt")
        usr_template = self.prompt_mgr.get_prompt("callflow_user.txt")
        formatted_user_prompt = usr_template.replace("{call_flow_text}", call_flow)
        formatted_user_prompt += self._build_directive_prompt_additions()

        self.output_text.clear()
        self.error_banner.hide()
        self.output_tabs.setCurrentIndex(0)  # Switch to Code tab during streaming
        self.status_lbl.setText("⚡ Streaming tokens...")
        self.status_lbl.setStyleSheet("color: #2563EB; font-weight: bold;")

        self._set_ui_busy(True)

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
        self.output_text.insertPlainText(delta)
        self.output_text.ensureCursorVisible()

    def _handle_complete(self, clean_puml: str):
        self.last_clean_puml = clean_puml
        self.output_text.setPlainText(clean_puml)
        self.conversation_history = getattr(self.worker, "history", [])

        self.status_lbl.setText("⚙️ Validating syntax...")
        self.status_lbl.setStyleSheet("color: #0284C7; font-weight: bold;")
        self._set_ui_busy(False)

        # Background validation & SVG preview generation
        self._validate_and_render_svg(clean_puml)

    def _handle_failed(self, error_msg: str):
        self.status_lbl.setText("⚠️ Failed / Cancelled")
        self.status_lbl.setStyleSheet("color: #DC2626; font-weight: bold;")
        self._set_ui_busy(False)
        if self.output_text.toPlainText().strip():
            self._set_output_ready(True)

    def _cancel_generation(self):
        if self.worker and self.worker.isRunning():
            self.worker.cancel()
            self.status_lbl.setText("Cancelling...")

    # -------------------------------------------------------------------------
    # SVG Validation & Live Preview
    # -------------------------------------------------------------------------
    def _validate_and_render_svg(self, puml_code: str):
        """Asynchronously compiles PlantUML code into SVG to test syntax and update preview."""
        self.validator = CallFlowValidationThread(puml_code, self.jar_path, parent=self)
        self.validator.validation_success.connect(self._on_validation_success)
        self.validator.validation_error.connect(self._on_validation_error)
        self.validator.start()

    def _on_validation_success(self, svg_path: str):
        self.error_banner.hide()
        self.status_lbl.setText("✅ Completed")
        self.status_lbl.setStyleSheet("color: #16A34A; font-weight: bold;")
        self.svg_preview.load_svg(svg_path)
        # Switch to rendered diagram tab so user immediately sees the visual result
        self.output_tabs.setCurrentIndex(1)
        self._set_output_ready(True)

    def _on_validation_error(self, line_no: int, error_msg: str):
        self.last_error_line = line_no
        self.last_syntax_error = error_msg

        clean_err = error_msg.replace("PlantUML Syntax Error:\n", "").strip()
        line_desc = f"on line {line_no}" if line_no > 0 else ""
        self.error_banner_lbl.setText(f"⚠️ PlantUML Syntax Error {line_desc}: {clean_err[:120]}")
        self.error_banner.show()

        self.status_lbl.setText("⚠️ Syntax Error")
        self.status_lbl.setStyleSheet("color: #DC2626; font-weight: bold;")
        self.output_tabs.setCurrentIndex(0)  # Switch back to code tab
        self._set_output_ready(True)

    # -------------------------------------------------------------------------
    # Conversational Refinement & Auto-Fix
    # -------------------------------------------------------------------------
    def _refine_diagram(self):
        instruction = self.refine_input.text().strip()
        if not instruction:
            return

        model = self.model_combo.currentText()
        if not model:
            return

        current_code = self.output_text.toPlainText().strip()
        if not current_code:
            return

        # Build multi-turn conversational message history
        messages = list(self.conversation_history)
        if not messages:
            sys_prompt = self.prompt_mgr.get_prompt("callflow_system.txt")
            messages = [
                {"role": "system", "content": sys_prompt},
                {"role": "user", "content": f"Here is the call flow:\n{self.input_text.toPlainText()}"},
                {"role": "assistant", "content": current_code}
            ]

        refinement_prompt = (
            f"Please update the PlantUML sequence diagram according to this specific instruction:\n"
            f"\"{instruction}\"\n\n"
            f"Return ONLY the complete updated PlantUML code between @startuml and @enduml."
        )
        messages.append({"role": "user", "content": refinement_prompt})

        self.output_text.clear()
        self.error_banner.hide()
        self.refine_input.clear()
        self.output_tabs.setCurrentIndex(0)
        self.status_lbl.setText("✨ Refining diagram...")
        self.status_lbl.setStyleSheet("color: #2563EB; font-weight: bold;")
        self._set_ui_busy(True)

        diag_type = self.type_combo.currentText()
        self.worker = CallFlowGeneratorThread(
            client=self.client,
            model=model,
            system_prompt="",
            user_prompt="",
            diagram_type=diag_type,
            messages=messages,
            parent=self
        )
        self.worker.token_received.connect(self._handle_token)
        self.worker.generation_completed.connect(self._handle_complete)
        self.worker.generation_failed.connect(self._handle_failed)
        self.worker.start()

    def _autofix_syntax_error(self):
        model = self.model_combo.currentText()
        if not model:
            return

        current_code = self.output_text.toPlainText().strip()
        err_msg = self.last_syntax_error

        repair_prompt = (
            f"The following PlantUML code produced a syntax error:\n"
            f"\"\"\"\n{err_msg}\n\"\"\"\n\n"
            f"Please fix the syntax error on line {self.last_error_line} and return the complete, valid PlantUML code between @startuml and @enduml."
        )

        messages = [
            {"role": "system", "content": self.prompt_mgr.get_prompt("callflow_system.txt")},
            {"role": "user", "content": f"PlantUML code with syntax error:\n{current_code}"},
            {"role": "user", "content": repair_prompt}
        ]

        self.output_text.clear()
        self.error_banner.hide()
        self.output_tabs.setCurrentIndex(0)
        self.status_lbl.setText("🪄 Auto-fixing syntax error...")
        self.status_lbl.setStyleSheet("color: #7C3AED; font-weight: bold;")
        self._set_ui_busy(True)

        diag_type = self.type_combo.currentText()
        self.worker = CallFlowGeneratorThread(
            client=self.client,
            model=model,
            system_prompt="",
            user_prompt="",
            diagram_type=diag_type,
            messages=messages,
            parent=self
        )
        self.worker.token_received.connect(self._handle_token)
        self.worker.generation_completed.connect(self._handle_complete)
        self.worker.generation_failed.connect(self._handle_failed)
        self.worker.start()

    # -------------------------------------------------------------------------
    # Helper Actions (Copy, Apply, UI state)
    # -------------------------------------------------------------------------
    def _copy_code(self):
        code = self.output_text.toPlainText().strip()
        if code:
            QApplication.clipboard().setText(code)
            self.copy_code_btn.setText("✅ Copied!")
            self.copy_code_btn.setStyleSheet("")
            from PyQt5.QtCore import QTimer
            QTimer.singleShot(1500, lambda: self.copy_code_btn.setText("📋 Copy Code"))

    def _apply_and_close(self, replace: bool):
        final_code = self.output_text.toPlainText().strip()
        if final_code:
            self.diagram_generated.emit(final_code, replace)
            self.accept()

    def _set_ui_busy(self, busy: bool):
        self.generate_btn.setEnabled(not busy)
        self.cancel_btn.setEnabled(busy)
        self.refine_btn.setEnabled(not busy and bool(self.output_text.toPlainText().strip()))
        if busy:
            self.replace_btn.setEnabled(False)
            self.append_btn.setEnabled(False)
            self.copy_code_btn.setEnabled(False)

    def _set_output_ready(self, ready: bool):
        self.generate_btn.setEnabled(True)
        self.cancel_btn.setEnabled(False)
        self.replace_btn.setEnabled(ready)
        self.append_btn.setEnabled(ready)
        self.copy_code_btn.setEnabled(ready)
        self.refine_btn.setEnabled(ready)

    def closeEvent(self, event):
        if self.worker and self.worker.isRunning():
            self.worker.cancel()
            self.worker.wait(300)
        if self.validator and self.validator.isRunning():
            self.validator.wait(300)
        event.accept()
