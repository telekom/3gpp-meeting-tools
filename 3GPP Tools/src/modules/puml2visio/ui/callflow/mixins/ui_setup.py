"""UI layout, controls strip, sample procedures menu, and panes for `CallFlowDialog`."""

from __future__ import annotations

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QAction,
    QApplication,
    QCheckBox,
    QComboBox,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMenu,
    QPlainTextEdit,
    QPushButton,
    QSplitter,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from core.ui.theme import C
from core.ui.ui_components import BUTTON_STYLE_TOOLBAR_SECONDARY
from modules.puml2visio.templates.plantuml_snippets import SAMPLE_3GPP_PROCEDURES
from modules.puml2visio.templates.plantuml_templates import PLANTUML_TYPES
from modules.puml2visio.ui.svg_preview_widget import EmbeddedSvgPreview
from modules.puml2visio.ui.syntax_highlighter import PlantUMLHighlighter


class CallFlowUiSetupMixin:
    """Mixin building the call-flow generator UI layout, style toggles, and samples menu."""

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
        self.status_lbl.setStyleSheet(f"color: {C.TEXT_MUTED}; font-weight: bold;")
        top_strip.addWidget(self.status_lbl)

        main_layout.addLayout(top_strip)

        # ---------------------------------------------------------------------
        # MIDDLE SPLITTER: Input Text & Options (Left) vs Output Preview (Right)
        # ---------------------------------------------------------------------
        splitter = QSplitter(Qt.Horizontal)
        splitter.setStyleSheet(f"""
            QSplitter::handle {{
                background-color: {C.BORDER_SUBTLE};
                width: 3px;
                margin: 0px 2px;
            }}
        """)

        # LEFT WIDGET: Specification Procedure Text & 3GPP Style Options
        left_widget = QFrame()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(8)

        input_header = QHBoxLayout()
        left_lbl = QLabel("📋 Specification Procedure Text:")
        left_lbl.setStyleSheet(f"font-weight: bold; color: {C.TEXT_STRONG}; font-size: 12px;")
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
        self.input_text.setStyleSheet(f"""
            QPlainTextEdit {{
                font-family: Consolas, 'Courier New', monospace;
                font-size: 12px;
                background-color: {C.BG_CARD};
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 6px;
                line-height: 1.4;
            }}
            QPlainTextEdit:focus {{
                border-color: {C.ACCENT_INDIGO_BLUE};
            }}
        """)
        left_layout.addWidget(self.input_text, stretch=1)

        # 3GPP Generation Style Toggles (Prompt Modifiers)
        style_box = QGroupBox("⚙️ 3GPP Diagram Style && Detail Options")
        style_box.setStyleSheet(f"""
            QGroupBox {{
                font-weight: bold;
                color: {C.TEXT_BODY};
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 6px;
                margin-top: 8px;
                padding-top: 10px;
                background-color: {C.BG_SURFACE};
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 4px;
            }}
            QCheckBox {{
                font-size: 11px;
                color: {C.TEXT_SECONDARY};
                font-weight: 500;
            }}
        """)
        style_grid = QGridLayout(style_box)
        style_grid.setContentsMargins(10, 10, 10, 10)
        style_grid.setSpacing(6)

        self.opt_boxes = QCheckBox("📦 Group NFs into Boxes (RAN / 5GC / User Plane)")
        self.opt_boxes.setChecked(True)
        self.opt_boxes.setToolTip("Enclose participants in colored visual boxes (RAN, Core Control Plane, User Plane).")

        self.opt_ies = QCheckBox("📑 Extract Key Message IEs && Parameters")
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
        self.output_tabs.setStyleSheet(f"""
            QTabWidget::pane {{
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 6px;
                background-color: {C.BG_CARD};
            }}
            QTabBar::tab {{
                font-weight: bold;
                padding: 6px 14px;
            }}
        """)

        # Tab 1: Code View
        code_tab = QWidget()
        code_layout = QVBoxLayout(code_tab)
        code_layout.setContentsMargins(4, 4, 4, 4)

        self.output_text = QPlainTextEdit()
        self.output_text.setReadOnly(False)
        self.output_text.setPlaceholderText("Generated PlantUML sequence code will stream here...")
        self.output_text.setStyleSheet(f"""
            QPlainTextEdit {{
                font-family: Consolas, 'Cascadia Code', 'Courier New', monospace;
                font-size: 12px;
                background-color: {C.BG_CARD};
                border: none;
                color: {C.TEXT_STRONG};
                line-height: 1.4;
            }}
        """)
        self.highlighter = PlantUMLHighlighter(self.output_text.document())
        code_layout.addWidget(self.output_text)
        self.output_tabs.addTab(code_tab, "📝 Code")

        # Tab 2: Visual Diagram Preview
        self.svg_preview = EmbeddedSvgPreview()
        self.output_tabs.addTab(self.svg_preview, "👁️ Diagram Preview")

        right_layout.addWidget(self.output_tabs, stretch=1)

        # PlantUML Syntax Error Banner (Hidden by default)
        self.error_banner = QFrame()
        self.error_banner.setStyleSheet(f"""
            QFrame {{
                background-color: {C.DANGER_SOFT_HOVER};
                border: 1px solid {C.DANGER_BORDER_HOVER};
                border-radius: 6px;
                padding: 4px;
            }}
        """)
        error_layout = QHBoxLayout(self.error_banner)
        error_layout.setContentsMargins(8, 4, 8, 4)
        error_layout.setSpacing(8)

        self.error_banner_lbl = QLabel("")
        self.error_banner_lbl.setStyleSheet(f"color: {C.DANGER}; font-weight: bold; font-size: 11px;")
        error_layout.addWidget(self.error_banner_lbl, stretch=1)

        self.autofix_btn = QPushButton("🪄 Auto-Fix Error")
        self.autofix_btn.setStyleSheet(
            f"QPushButton {{ background-color: {C.DANGER_ALT}; color: white; border: none; "
            f"padding: 4px 10px; border-radius: 4px; font-weight: bold; }} "
            f"QPushButton:hover {{ background-color: {C.DANGER}; }}"
        )
        self.autofix_btn.setToolTip("Prompt local LLM to automatically fix this PlantUML syntax error.")
        self.autofix_btn.clicked.connect(self._autofix_syntax_error)
        error_layout.addWidget(self.autofix_btn)

        self.error_banner.hide()
        right_layout.addWidget(self.error_banner)

        # Conversational Iterative Refinement Bar
        refine_bar = QHBoxLayout()
        refine_bar.setSpacing(6)

        refine_lbl = QLabel("💬 Refine:")
        refine_lbl.setStyleSheet(f"font-weight: bold; color: {C.TEXT_SECONDARY}; font-size: 11px;")
        refine_bar.addWidget(refine_lbl)

        self.refine_input = QLineEdit()
        self.refine_input.setPlaceholderText("e.g. Add an alt block for authentication failure, or group into VPLMN / HPLMN...")
        self.refine_input.setStyleSheet(f"""
            QLineEdit {{
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 4px;
                padding: 5px 8px;
                background-color: {C.BG_CARD};
                font-size: 12px;
            }}
            QLineEdit:focus {{
                border-color: {C.ACCENT_INDIGO_BLUE};
            }}
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
