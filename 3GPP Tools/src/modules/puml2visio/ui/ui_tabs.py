# --- File: src/modules/puml2visio/ui/ui_tabs.py ---
"""Tabs and layouts for Visio & PlantUML tools:

- CodeEditorTab: Interactive PlantUML Mini-IDE with element inserter, syntax palette,
  inline find/replace, error line highlighting, and side-by-side SVG preview
- BatchConvertTab: Drag-and-drop batch conversion queue tab
"""

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtWidgets import (
    QAction,
    QComboBox,
    QHBoxLayout,
    QLabel,
    QMenu,
    QPushButton,
    QSplitter,
    QVBoxLayout,
    QWidget,
)

from core.ui.ui_components import (
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    InteractiveDropLabel,
)
from modules.puml2visio.templates.plantuml_snippets import (
    register_snippets_reloaded_callback,
    reload_snippets,
    PLANTUML_SNIPPETS,
    PlantUMLSnippet,
)
from modules.puml2visio.templates.plantuml_templates import PLANTUML_TYPES
from modules.puml2visio.ui.preset_dialog import (
    UserPresetDialog,
    open_presets_folder,
)
from modules.puml2visio.ui.callflow_dialog import CallFlowDialog
from modules.puml2visio.ui.elements_palette import ElementsPaletteWidget
from modules.puml2visio.ui.svg_preview_widget import EmbeddedSvgPreview
from modules.puml2visio.ui.ui_components import CodeDropTextEdit, FindReplaceBar


class CodeEditorTab(QWidget):
    """Interactive PlantUML IDE tab with element insertion, syntax highlighting,

    smart indentation, inline find/replace, and optional embedded split preview.
    """
    # --- Custom Signals ---
    template_requested = pyqtSignal(str)
    docs_requested = pyqtSignal(str)
    clear_requested = pyqtSignal()
    undo_requested = pyqtSignal()
    redo_requested = pyqtSignal()
    copy_code_requested = pyqtSignal()
    live_view_toggled = pyqtSignal(bool)
    planttext_requested = pyqtSignal()
    copy_path_requested = pyqtSignal()
    open_folder_requested = pyqtSignal()
    export_requested = pyqtSignal(str)
    file_dropped = pyqtSignal(str)
    browser_preview_requested = pyqtSignal()

    def __init__(self):
        super().__init__()
        self._ai_dialog = None
        self._setup_ui()
        register_snippets_reloaded_callback(self._build_insert_element_menu)

    def _setup_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)

        # ---------------------------------------------------------------------
        # TOP TOOLBAR: Templates, Element Inserter, Palette, Find, AI, & Split Preview
        # ---------------------------------------------------------------------
        template_layout = QHBoxLayout()
        template_layout.setSpacing(6)

        template_lbl = QLabel("📖 Templates:")
        template_lbl.setStyleSheet("font-weight: bold; color: #475569; font-size: 12px;")

        self.template_combo = QComboBox()
        self.template_combo.addItems(list(PLANTUML_TYPES.keys()))
        self.template_combo.setToolTip("Select a diagram type boilerplate.")
        self.template_combo.currentTextChanged.connect(self._on_template_type_changed)

        self.insert_tpl_btn = QPushButton("Insert")
        self.insert_tpl_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.insert_tpl_btn.setToolTip("Insert the selected template into the editor.")
        self.insert_tpl_btn.clicked.connect(lambda: self.template_requested.emit(self.template_combo.currentText()))

        # Quick Insert Element Menu Button
        self.insert_elem_btn = QPushButton("🧩 Insert Element ▼")
        self.insert_elem_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.insert_elem_btn.setToolTip("Quickly insert PlantUML syntax elements, arrows, notes, and 3GPP presets at cursor.")
        self._build_insert_element_menu()

        # Syntax Palette Toggle Button
        self.palette_btn = QPushButton("📚 Syntax Palette")
        self.palette_btn.setCheckable(True)
        self.palette_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.palette_btn.setToolTip("Toggle the searchable Elements & Syntax Cheat-Sheet sidebar.")
        self.palette_btn.toggled.connect(self._toggle_palette)

        # Inline Find Button
        self.find_btn = QPushButton("🔍 Find")
        self.find_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.find_btn.setToolTip("Find and Replace in diagram (Ctrl+F).")
        self.find_btn.clicked.connect(self._toggle_find_bar)

        # Official Docs Button
        self.docs_btn = QPushButton("📘 Docs")
        self.docs_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.docs_btn.setToolTip("Open the official PlantUML syntax documentation for this diagram type.")
        self.docs_btn.clicked.connect(lambda: self.docs_requested.emit(self.template_combo.currentText()))

        # AI Call Flow Generator Button
        self.ai_gen_btn = QPushButton("🤖 Generate from Call Flow...")
        self.ai_gen_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.ai_gen_btn.setToolTip(
            "Convert standard 3GPP call flow procedure steps into PlantUML using your local Ollama LLM.")
        self.ai_gen_btn.clicked.connect(self._open_ai_generator)

        # Split Preview Toggle Button
        self.preview_dock_btn = QPushButton("👁️ Split Preview")
        self.preview_dock_btn.setCheckable(True)
        self.preview_dock_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.preview_dock_btn.setToolTip("Toggle side-by-side embedded live SVG preview panel.")
        self.preview_dock_btn.toggled.connect(self._toggle_preview_dock)

        template_layout.addWidget(template_lbl)
        template_layout.addWidget(self.template_combo)
        template_layout.addWidget(self.insert_tpl_btn)
        template_layout.addWidget(self.insert_elem_btn)
        template_layout.addWidget(self.palette_btn)
        template_layout.addWidget(self.find_btn)
        template_layout.addWidget(self.docs_btn)
        template_layout.addWidget(self.ai_gen_btn)
        template_layout.addStretch()
        template_layout.addWidget(self.preview_dock_btn)

        layout.addLayout(template_layout)

        # ---------------------------------------------------------------------
        # INLINE FIND & REPLACE BAR (Hidden by default)
        # ---------------------------------------------------------------------
        self.text_input = CodeDropTextEdit()
        self.text_input.setPlaceholderText(
            "Paste PlantUML code OR drop a generated .vsdx file here to extract its source..."
        )
        self.text_input.file_dropped.connect(self.file_dropped.emit)
        self.text_input.setToolTip(
            "Type or paste PlantUML code here. Drag & drop a generated .vsdx file to retrieve its source code."
        )

        self.find_bar = FindReplaceBar(self.text_input)
        self.text_input.find_requested.connect(self.find_bar.show_and_focus)
        layout.addWidget(self.find_bar)

        # ---------------------------------------------------------------------
        # CENTRAL WORKSPACE SPLITTER (Palette | Code Editor | SVG Preview)
        # ---------------------------------------------------------------------
        self.work_splitter = QSplitter(Qt.Horizontal)
        self.work_splitter.setStyleSheet("""
            QSplitter::handle {
                background-color: #E2E8F0;
                width: 3px;
                margin: 0px 2px;
            }
            QSplitter::handle:hover {
                background-color: #395396;
            }
        """)

        # Left: Elements Palette (Hidden by default)
        self.palette_widget = ElementsPaletteWidget()
        self.palette_widget.snippet_triggered.connect(self.text_input.insert_snippet)
        self.palette_widget.hide()
        self.work_splitter.addWidget(self.palette_widget)

        # Center: Main Code Editor
        self.work_splitter.addWidget(self.text_input)

        # Right: Embedded SVG Preview (Hidden by default)
        self.svg_preview = EmbeddedSvgPreview()
        self.svg_preview.close_requested.connect(lambda: self.preview_dock_btn.setChecked(False))
        self.svg_preview.browser_requested.connect(self.browser_preview_requested.emit)
        self.svg_preview.hide()
        self.work_splitter.addWidget(self.svg_preview)

        # Proportions: 25% Palette, 50% Editor, 25% Preview
        self.work_splitter.setStretchFactor(0, 1)
        self.work_splitter.setStretchFactor(1, 3)
        self.work_splitter.setStretchFactor(2, 2)

        layout.addWidget(self.work_splitter, stretch=1)

        # ---------------------------------------------------------------------
        # BOTTOM ACTION BAR
        # ---------------------------------------------------------------------
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(6)

        self.clear_btn = QPushButton("🗑️ Clear")
        self.clear_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.clear_btn.setToolTip("Clear the text editor.")
        self.clear_btn.clicked.connect(self.clear_requested.emit)

        self.undo_btn = QPushButton("↩️ Undo")
        self.undo_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.undo_btn.setToolTip("Undo the last action (Ctrl+Z).")
        self.undo_btn.clicked.connect(self.undo_requested.emit)

        self.redo_btn = QPushButton("↪️ Redo")
        self.redo_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.redo_btn.setToolTip("Redo the last undone action (Ctrl+Y).")
        self.redo_btn.clicked.connect(self.redo_requested.emit)

        self.copy_code_btn = QPushButton("📄 Copy Code")
        self.copy_code_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.copy_code_btn.setToolTip("Copy the PlantUML source code to your clipboard.")
        self.copy_code_btn.clicked.connect(self.copy_code_requested.emit)

        self.live_view_btn = QPushButton("🌐 Browser Preview")
        self.live_view_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.live_view_btn.setCheckable(True)
        self.live_view_btn.setToolTip("Toggle real-time external browser preview. Auto-updates as you type!")
        self.live_view_btn.clicked.connect(self.live_view_toggled.emit)

        self.planttext_btn = QPushButton("🌐 Show in planttext")
        self.planttext_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.planttext_btn.setToolTip("Open your code in PlantText.com for a quick web preview.")
        self.planttext_btn.clicked.connect(self.planttext_requested.emit)

        self.copy_btn = QPushButton("🔗 Copy Path")
        self.copy_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.copy_btn.setEnabled(False)
        self.copy_btn.setToolTip("Copy the file path of the last generated diagram.")
        self.copy_btn.clicked.connect(self.copy_path_requested.emit)

        self.open_folder_btn = QPushButton("📂 Open Folder")
        self.open_folder_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.open_folder_btn.setToolTip("Open the working directory where files are saved.")
        self.open_folder_btn.clicked.connect(self.open_folder_requested.emit)

        self.export_btn = QPushButton("📤 Export Diagram ▼")
        self.export_btn.setObjectName("primaryBtn")
        self.export_btn.setToolTip("Export your PlantUML code to various formats.")

        export_menu = QMenu(self)
        export_menu.setToolTipsVisible(True)

        visio_action = QAction("To Visio (.vsdx)", self)
        visio_action.setToolTip("Saves to disk and natively opens a fully editable, perfectly aligned Visio diagram.")
        visio_action.triggered.connect(lambda: self.export_requested.emit("vsdx"))
        export_menu.addAction(visio_action)

        pptx_action = QAction("To PowerPoint (.pptx) (New Presentation)", self)
        pptx_action.setToolTip(
            "Generates Office shapes and opens a new unsaved PowerPoint presentation."
        )
        pptx_action.triggered.connect(lambda: self.export_requested.emit("pptx"))
        export_menu.addAction(pptx_action)

        pptx_append_action = QAction("Add to Open PowerPoint Presentation...", self)
        pptx_append_action.setToolTip(
            "Appends the diagram as a new slide at the end of an open presentation."
        )
        pptx_append_action.triggered.connect(lambda: self.export_requested.emit("pptx_append"))
        export_menu.addAction(pptx_append_action)

        svg_action = QAction("To Vector Graphic (.svg)", self)
        svg_action.setToolTip(
            "Saves to disk and opens a standard, scalable vector image (.svg) in your default web browser or viewer."
        )
        svg_action.triggered.connect(lambda: self.export_requested.emit("svg"))
        export_menu.addAction(svg_action)

        ascii_action = QAction("To Text Art (.txt)", self)
        ascii_action.setToolTip("Saves to disk and opens clean Unicode text-art in your default text editor.")
        ascii_action.triggered.connect(lambda: self.export_requested.emit("ascii"))
        export_menu.addAction(ascii_action)

        self.export_btn.setMenu(export_menu)

        btn_layout.addWidget(self.clear_btn)
        btn_layout.addWidget(self.undo_btn)
        btn_layout.addWidget(self.redo_btn)
        btn_layout.addWidget(self.copy_code_btn)
        btn_layout.addWidget(self.live_view_btn)
        btn_layout.addWidget(self.planttext_btn)
        btn_layout.addStretch()
        btn_layout.addWidget(self.copy_btn)
        btn_layout.addWidget(self.open_folder_btn)
        btn_layout.addWidget(self.export_btn)

        layout.addLayout(btn_layout)
        self.setLayout(layout)

    # -------------------------------------------------------------------------
    # Quick Insert Element Menu Builder
    # -------------------------------------------------------------------------
    def _build_insert_element_menu(self):
        menu = QMenu(self)
        menu.setToolTipsVisible(True)

        current_type = self.template_combo.currentText() if hasattr(self, 'template_combo') else "Sequence"

        # 1. Top Section: Relevant to current diagram type
        relevant = [s for s in PLANTUML_SNIPPETS if current_type in s.diagram_scope]
        if relevant:
            fav_menu = menu.addMenu(f"⚡ Common for {current_type}".replace("&", "&&"))
            for s in relevant[:12]:
                act = QAction(s.label.replace("&", "&&"), self)
                act.setToolTip(s.description)
                act.triggered.connect(lambda checked, snip=s: self.text_input.insert_snippet(snip))
                fav_menu.addAction(act)
            menu.addSeparator()

        # 2. All categorized sections
        categories = {}
        for snip in PLANTUML_SNIPPETS:
            categories.setdefault(snip.category, []).append(snip)

        for cat_name, snips in categories.items():
            # Escape '&' as '&&' so Qt does not treat it as a keyboard mnemonic
            cat_menu = menu.addMenu(cat_name.replace("&", "&&"))
            for snip in snips:
                act = QAction(snip.label.replace("&", "&&"), self)
                act.setToolTip(snip.description)
                act.triggered.connect(lambda checked, s=snip: self.text_input.insert_snippet(s))
                cat_menu.addAction(act)

        # 3. User Presets & Configuration Actions
        menu.addSeparator()
        add_preset_act = QAction("➕ Add Custom Preset...", self)
        add_preset_act.triggered.connect(self._open_add_preset_dialog)
        menu.addAction(add_preset_act)

        reload_act = QAction("🔄 Reload Presets from Disk", self)
        reload_act.triggered.connect(self._reload_presets)
        menu.addAction(reload_act)

        folder_act = QAction("📂 Open Presets Config Folder", self)
        folder_act.triggered.connect(open_presets_folder)
        menu.addAction(folder_act)

        self.insert_elem_btn.setMenu(menu)

    def _open_add_preset_dialog(self):
        dialog = UserPresetDialog(parent=self)
        dialog.exec_()

    def _reload_presets(self):
        reload_snippets()

    def _on_template_type_changed(self, new_type: str):
        self._build_insert_element_menu()
        self.palette_widget.set_diagram_scope(new_type)

    def _toggle_palette(self, checked: bool):
        self.palette_widget.setVisible(checked)
        if checked:
            self.work_splitter.setSizes([260, 500, 300])

    def _toggle_preview_dock(self, checked: bool):
        self.svg_preview.setVisible(checked)
        if checked:
            self.work_splitter.setSizes([260, 480, 400])
            self.live_view_toggled.emit(True)

    def _toggle_find_bar(self):
        if self.find_bar.isVisible():
            self.find_bar.hide_and_focus_editor()
        else:
            self.find_bar.show_and_focus()

    # -------------------------------------------------------------------------
    # Public Accessors & Controls
    # -------------------------------------------------------------------------
    def get_text(self):
        return self.text_input.toPlainText().strip()

    def set_copy_path_enabled(self, enabled: bool, out_path: str = ""):
        self.copy_btn.setEnabled(enabled)
        if enabled:
            self.copy_btn.setObjectName("primaryBtn")
            self.copy_btn.setStyleSheet("")
            self.copy_btn.setToolTip(f"Copy path:\n{out_path}")
        else:
            self.copy_btn.setObjectName("")
            self.copy_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
            self.copy_btn.setToolTip("Copy the file path of the last generated diagram.")
        self.copy_btn.style().unpolish(self.copy_btn)
        self.copy_btn.style().polish(self.copy_btn)

    def _open_ai_generator(self):
        """Launches or brings into focus the modeless call flow AI generation window."""
        if self._ai_dialog is None:
            self._ai_dialog = CallFlowDialog(parent=None)
            self._ai_dialog.diagram_generated.connect(self._handle_generated_diagram)

        if self._ai_dialog.isMinimized():
            self._ai_dialog.showNormal()

        self._ai_dialog.show()
        self._ai_dialog.raise_()
        self._ai_dialog.activateWindow()

    def _handle_generated_diagram(self, puml_code: str, replace: bool):
        """Updates the text editor with the generated PlantUML diagram."""
        if replace or not self.text_input.toPlainText().strip():
            self.text_input.setPlainText(puml_code)
        else:
            self.text_input.appendPlainText("\n\n" + puml_code)

    def closeEvent(self, event):
        """Ensure background dialogs and preview processes are cleaned up."""
        ai_dialog = getattr(self, "_ai_dialog", None)
        if ai_dialog is not None:
            ai_dialog.close()
        super().closeEvent(event)


# Re-export modernized BatchConvertTab for backward compatibility
from modules.puml2visio.ui.batch_convert_tab import BatchConvertTab
