from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
                             QLabel, QComboBox, QMenu, QAction, QRadioButton,
                             QGroupBox, QFormLayout, QLineEdit, QSpinBox)
from PyQt5.QtCore import pyqtSignal

from ui_components import CodeDropTextEdit, InteractiveDropLabel
from plantuml_templates import PLANTUML_TYPES


class CodeEditorTab(QWidget):
    # --- Custom Signals ---
    template_requested = pyqtSignal(str)
    docs_requested = pyqtSignal(str)
    clear_requested = pyqtSignal()
    undo_requested = pyqtSignal()
    copy_code_requested = pyqtSignal()
    live_view_toggled = pyqtSignal(bool)
    planttext_requested = pyqtSignal()
    copy_path_requested = pyqtSignal()
    open_folder_requested = pyqtSignal()
    export_requested = pyqtSignal(str)
    file_dropped = pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(15, 15, 15, 15)

        template_layout = QHBoxLayout()
        template_lbl = QLabel("📖 Templates:")
        template_lbl.setStyleSheet("font-weight: bold; color: #555;")

        self.template_combo = QComboBox()
        self.template_combo.addItems(list(PLANTUML_TYPES.keys()))
        self.template_combo.setToolTip("Select a diagram type.")

        self.insert_tpl_btn = QPushButton("Insert")
        self.insert_tpl_btn.setToolTip("Insert the selected boilerplate into the editor.")
        self.insert_tpl_btn.clicked.connect(lambda: self.template_requested.emit(self.template_combo.currentText()))

        self.docs_btn = QPushButton("📘 Docs")
        self.docs_btn.setToolTip("Open the official PlantUML syntax documentation for this diagram type.")
        self.docs_btn.clicked.connect(lambda: self.docs_requested.emit(self.template_combo.currentText()))

        template_layout.addWidget(template_lbl)
        template_layout.addWidget(self.template_combo)
        template_layout.addWidget(self.insert_tpl_btn)
        template_layout.addWidget(self.docs_btn)
        template_layout.addStretch()
        layout.addLayout(template_layout)

        self.text_input = CodeDropTextEdit()
        self.text_input.setPlaceholderText(
            "Paste PlantUML code OR drop a generated .vsdx file here to extract its source...")
        self.text_input.file_dropped.connect(self.file_dropped.emit)
        self.text_input.setToolTip(
            "Type or paste PlantUML code here. Drag & drop a generated .vsdx file to retrieve its source code.")
        layout.addWidget(self.text_input)

        btn_layout = QHBoxLayout()

        self.clear_btn = QPushButton("🗑️ Clear")
        self.clear_btn.clicked.connect(self.clear_requested.emit)
        self.clear_btn.setToolTip("Clear the text editor.")

        self.undo_btn = QPushButton("↩️ Undo")
        self.undo_btn.clicked.connect(self.undo_requested.emit)
        self.undo_btn.setToolTip("Undo the last action (typing, clear, or template insert).")

        self.copy_code_btn = QPushButton("📄 Copy Code")
        self.copy_code_btn.clicked.connect(self.copy_code_requested.emit)
        self.copy_code_btn.setToolTip("Copy the PlantUML source code to your clipboard.")

        self.live_view_btn = QPushButton("👁️ Live Preview")
        self.live_view_btn.setCheckable(True)
        self.live_view_btn.setToolTip("Toggle real-time browser preview. Auto-updates as you type!")
        self.live_view_btn.clicked.connect(self.live_view_toggled.emit)

        self.planttext_btn = QPushButton("🌐 Show in planttext")
        self.planttext_btn.clicked.connect(self.planttext_requested.emit)
        self.planttext_btn.setToolTip("Open your code in PlantText.com for a quick web preview.")

        self.copy_btn = QPushButton("🔗 Copy Path")
        self.copy_btn.setEnabled(False)
        self.copy_btn.clicked.connect(self.copy_path_requested.emit)
        self.copy_btn.setToolTip("Copy the file path of the last generated diagram.")

        self.open_folder_btn = QPushButton("📂 Open Folder")
        self.open_folder_btn.clicked.connect(self.open_folder_requested.emit)
        self.open_folder_btn.setToolTip("Open the working directory where files are saved.")

        self.export_btn = QPushButton("📤 Export Diagram ▼")
        self.export_btn.setObjectName("primaryBtn")
        self.export_btn.setToolTip("Export your PlantUML code to various formats.")

        export_menu = QMenu(self)
        export_menu.setToolTipsVisible(True)

        visio_action = QAction("To Visio (.vsdx)", self)
        visio_action.setToolTip("Saves to disk and natively opens a fully editable, perfectly aligned Visio diagram.")
        visio_action.triggered.connect(lambda: self.export_requested.emit("vsdx"))
        export_menu.addAction(visio_action)

        pptx_action = QAction("To PowerPoint (.pptx)", self)
        pptx_action.setToolTip(
            "Generates Office shapes and leaves PowerPoint open (UNSAVED) so you can instantly copy the slide.")
        pptx_action.triggered.connect(lambda: self.export_requested.emit("pptx"))
        export_menu.addAction(pptx_action)

        svg_action = QAction("To Vector Graphic (.svg)", self)
        svg_action.setToolTip(
            "Saves to disk and opens a standard, scalable vector image (.svg) in your default web browser or viewer.")
        svg_action.triggered.connect(lambda: self.export_requested.emit("svg"))
        export_menu.addAction(svg_action)

        ascii_action = QAction("To Text Art (.txt)", self)
        ascii_action.setToolTip("Saves to disk and opens clean Unicode text-art in your default text editor.")
        ascii_action.triggered.connect(lambda: self.export_requested.emit("ascii"))
        export_menu.addAction(ascii_action)

        self.export_btn.setMenu(export_menu)

        btn_layout.addWidget(self.clear_btn)
        btn_layout.addWidget(self.undo_btn)
        btn_layout.addWidget(self.copy_code_btn)
        btn_layout.addWidget(self.live_view_btn)
        btn_layout.addWidget(self.planttext_btn)
        btn_layout.addStretch()
        btn_layout.addWidget(self.copy_btn)
        btn_layout.addWidget(self.open_folder_btn)
        btn_layout.addWidget(self.export_btn)

        layout.addLayout(btn_layout)
        self.setLayout(layout)

    def get_text(self):
        return self.text_input.toPlainText().strip()

    def set_copy_path_enabled(self, enabled: bool, out_path: str = ""):
        self.copy_btn.setEnabled(enabled)
        if enabled:
            self.copy_btn.setStyleSheet("background-color: #395396; color: white; border: none;")
            self.copy_btn.setToolTip(f"Copy path:\n{out_path}")


class BatchConvertTab(QWidget):
    files_dropped = pyqtSignal(list)

    def __init__(self):
        super().__init__()
        layout = QVBoxLayout()
        layout.setContentsMargins(15, 15, 15, 15)
        self.drop_label = InteractiveDropLabel("⏳ Initializing system checks... Please wait.", ['.puml', '.txt'])
        self.drop_label.file_dropped.connect(self.files_dropped.emit)
        layout.addWidget(self.drop_label)
        self.setLayout(layout)

    def set_state(self, state, text=None):
        self.drop_label.set_state(state, text)


class WordExtractorTab(QWidget):
    extract_visio_requested = pyqtSignal(str)
    split_doc_requested = pyqtSignal(str, str, int)

    def __init__(self):
        super().__init__()
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(15, 15, 15, 15)

        # --- NEW: CONTROL PANEL ---
        controls = QGroupBox("Operation Type")
        controls.setStyleSheet(
            "QGroupBox { font-weight: bold; border: 1px solid #CCC; border-radius: 6px; margin-top: 10px; } QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 3px; }")
        controls_layout = QVBoxLayout()

        self.radio_extract = QRadioButton("Extract Embedded Visio Diagrams")
        self.radio_extract.setChecked(True)
        self.radio_split = QRadioButton("Subtractive Slicing (Split by Clause)")

        controls_layout.addWidget(self.radio_extract)
        controls_layout.addWidget(self.radio_split)

        self.split_opts = QWidget()
        form = QFormLayout()

        self.prefix_input = QLineEdit("6.")
        self.prefix_input.setToolTip("The clause prefix to target (e.g., '6.' or '6.2')")
        self.prefix_input.setStyleSheet("padding: 4px; border: 1px solid #CCC; border-radius: 4px;")

        self.depth_input = QSpinBox()
        self.depth_input.setRange(1, 6)
        self.depth_input.setValue(2)
        self.depth_input.setToolTip("How deep the heading should be. E.g., '6.1' is depth 2. '6.1.4' is depth 3.")
        self.depth_input.setStyleSheet("padding: 4px; border: 1px solid #CCC; border-radius: 4px;")

        form.addRow("Target Clause Prefix:", self.prefix_input)
        form.addRow("Heading Depth:", self.depth_input)
        self.split_opts.setLayout(form)
        self.split_opts.setVisible(False)

        controls_layout.addWidget(self.split_opts)
        controls.setLayout(controls_layout)

        self.radio_split.toggled.connect(self.split_opts.setVisible)
        layout.addWidget(controls)

        # --- DROP ZONE ---
        self.drop_label = InteractiveDropLabel("📥 Drag & Drop your Microsoft Word (.docx) file here", ['.docx'])
        self.drop_label.file_dropped.connect(self._handle_drop)
        layout.addWidget(self.drop_label)

        self.setLayout(layout)

    def _handle_drop(self, files):
        # We must loop through the drop array so batch-dropping Word docs actually works
        for file_path in files:
            if self.radio_extract.isChecked():
                self.extract_visio_requested.emit(file_path)
            else:
                prefix = self.prefix_input.text().strip()
                depth = self.depth_input.value()
                self.split_doc_requested.emit(file_path, prefix, depth)