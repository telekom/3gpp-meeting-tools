from pathlib import Path

from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
                             QLabel, QComboBox, QLineEdit, QSpinBox,
                             QFormLayout, QStackedWidget, QTabWidget, QCheckBox, QMenu, QAction)
from PyQt5.QtCore import pyqtSignal
import win32com.client
import pythoncom

from modules.puml2visio.ui.ui_components import CodeDropTextEdit, InteractiveDropLabel
from modules.puml2visio.templates.plantuml_templates import PLANTUML_TYPES


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


# ==========================================
# --- REUSABLE SYMMETRIC INPUT PANE ---
# ==========================================
class DocumentSelectorPane(QWidget):
    """A symmetric, reusable widget handling Local, Open, and URL inputs."""

    def __init__(self, title: str):
        super().__init__()
        self.title = title
        self.selected_file = ""
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)

        lbl = QLabel(f"<b>{self.title}</b>")
        lbl.setStyleSheet("color: #444; margin-bottom: 5px;")
        layout.addWidget(lbl)

        self.tabs = QTabWidget()
        self.tabs.setObjectName("selector_tabs")

        # Tab 1: Local File Drop
        self.drop_tab = QWidget()
        drop_layout = QVBoxLayout(self.drop_tab)
        self.drop_zone = InteractiveDropLabel("Drop .docx here", ['.docx'])
        self.drop_zone.file_dropped.connect(self._on_drop)
        drop_layout.addWidget(self.drop_zone)
        self.tabs.addTab(self.drop_tab, "📁 Local")

        # Tab 2: Open Documents
        self.open_tab = QWidget()
        open_layout = QVBoxLayout(self.open_tab)

        self.open_combo = QComboBox()
        self.refresh_btn = QPushButton("↻ Refresh Active Documents")
        self.refresh_btn.clicked.connect(self.poll_open_documents)

        open_layout.addWidget(QLabel("Select an open Word document:"))
        open_layout.addWidget(self.open_combo)
        open_layout.addWidget(self.refresh_btn)
        open_layout.addStretch()
        self.tabs.addTab(self.open_tab, "🖥️ Open Docs")

        # Tab 3: URL
        self.url_tab = QWidget()
        url_layout = QVBoxLayout(self.url_tab)

        self.url_input = QLineEdit()
        self.url_input.setPlaceholderText("https://...")

        url_layout.addWidget(QLabel("Paste document URL:"))
        url_layout.addWidget(self.url_input)
        url_layout.addStretch()
        self.tabs.addTab(self.url_tab, "🌐 URL")

        layout.addWidget(self.tabs)
        self.setLayout(layout)
        self.tabs.currentChanged.connect(self._on_tab_changed)

    def _on_drop(self, files):
        if files:
            self.selected_file = files[0]
            self.drop_zone.set_state("ready", f"Ready:\n{Path(self.selected_file).name}")

    def poll_open_documents(self):
        self.open_combo.clear()
        try:
            pythoncom.CoInitialize()
            word = win32com.client.GetActiveObject("Word.Application")
            for doc in word.Documents:
                # Safely extract the name and strip any accidental whitespace
                doc_name = str(doc.Name).strip()

                # Only add the document to the dropdown if it actually has a valid name
                if doc_name:
                    self.open_combo.addItem(doc_name, doc.FullName)

            # If the loop filtered out everything and the list is still empty, show the default message
            if self.open_combo.count() == 0:
                self.open_combo.addItem("No open documents detected.", "")

        except Exception:
            self.open_combo.addItem("No open documents detected.", "")
        finally:
            pythoncom.CoUninitialize()

    def get_input(self) -> str:
        idx = self.tabs.currentIndex()
        if idx == 0:
            return self.selected_file
        elif idx == 1:
            return self.open_combo.currentData() or ""
        elif idx == 2:
            return self.url_input.text().strip()
        return ""

    def _on_tab_changed(self, index):
        """Auto-refreshes the list when the user switches to the 'Open Docs' tab (index 1)."""
        if index == 1:
            self.poll_open_documents()


# ==========================================
# --- WORD EXTRACTOR TAB (DYNAMIC STACK) ---
# ==========================================
class WordExtractorTab(QWidget):
    extract_visio_requested = pyqtSignal(str)
    split_doc_requested = pyqtSignal(str, str, int)
    compare_doc_requested = pyqtSignal(str, str, bool)

    def __init__(self):
        super().__init__()
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(15, 15, 15, 15)

        # Switcher
        switcher_layout = QHBoxLayout()
        switcher_layout.addWidget(QLabel("<b>⚙️ Operation Type:</b>"))

        self.op_combo = QComboBox()
        self.op_combo.setStyleSheet("padding: 4px; font-weight: bold; font-size: 13px;")
        self.op_combo.addItems([
            "Extract Embedded Visio Diagrams",
            "Subtractive Slicing (Split by Clause)",
            "Compare Documents (Native Word Diff)"
        ])
        switcher_layout.addWidget(self.op_combo)
        switcher_layout.addStretch()
        layout.addLayout(switcher_layout)

        # Stack
        self.stack = QStackedWidget()

        # Card 1: Visio Extractor
        self.card_visio = QWidget()
        visio_layout = QVBoxLayout(self.card_visio)
        self.visio_drop = InteractiveDropLabel("📥 Drag & Drop a .docx file here to extract its Visio components",
                                               ['.docx'])
        self.visio_drop.file_dropped.connect(lambda files: [self.extract_visio_requested.emit(f) for f in files])
        visio_layout.addWidget(self.visio_drop)
        self.stack.addWidget(self.card_visio)

        # Card 2: Splitter
        self.card_split = QWidget()
        split_layout = QVBoxLayout(self.card_split)

        form = QFormLayout()
        self.prefix_input = QLineEdit("6.")
        self.prefix_input.setStyleSheet("padding: 4px; border: 1px solid #CCC; border-radius: 4px;")

        self.depth_input = QSpinBox()
        self.depth_input.setRange(1, 6)
        self.depth_input.setValue(2)
        self.depth_input.setStyleSheet("padding: 4px; border: 1px solid #CCC; border-radius: 4px;")

        form.addRow("Target Clause Prefix:", self.prefix_input)
        form.addRow("Heading Depth Hierarchy:", self.depth_input)
        split_layout.addLayout(form)

        self.split_drop = InteractiveDropLabel("📥 Drag & Drop a .docx file here to slice it into chapters", ['.docx'])
        self.split_drop.file_dropped.connect(
            lambda files: [self.split_doc_requested.emit(f, self.prefix_input.text().strip(), self.depth_input.value())
                           for f in files])
        split_layout.addWidget(self.split_drop)
        self.stack.addWidget(self.card_split)

        # Card 3: The Comparator
        self.card_compare = QWidget()
        compare_layout = QVBoxLayout(self.card_compare)

        panes_layout = QHBoxLayout()
        # BECAUSE DocumentSelectorPane IS DEFINED ABOVE, WE CAN USE IT HERE SAFELY:
        self.pane_a = DocumentSelectorPane("📄 DOCUMENT A (Original)")
        self.pane_b = DocumentSelectorPane("📄 DOCUMENT B (Revised)")
        panes_layout.addWidget(self.pane_a)
        panes_layout.addWidget(self.pane_b)

        compare_layout.addLayout(panes_layout)

        self.keep_open_cb = QCheckBox("Keep source documents (A and B) open after comparison")
        self.keep_open_cb.setStyleSheet("color: #444; margin-top: 5px;")
        self.keep_open_cb.setChecked(True)
        compare_layout.addWidget(self.keep_open_cb)

        self.run_compare_btn = QPushButton("⚖️ Run Word Comparison")
        self.run_compare_btn.setStyleSheet(
            "font-weight: bold; padding: 10px; background-color: #395396; color: white; border-radius: 4px;")
        self.run_compare_btn.clicked.connect(self._trigger_comparison)
        compare_layout.addWidget(self.run_compare_btn)

        self.stack.addWidget(self.card_compare)

        layout.addWidget(self.stack)
        self.setLayout(layout)
        self.op_combo.currentIndexChanged.connect(self.stack.setCurrentIndex)

    def _trigger_comparison(self):
        val_a = self.pane_a.get_input()
        val_b = self.pane_b.get_input()
        keep_open = self.keep_open_cb.isChecked()

        if val_a and val_b:
            self.compare_doc_requested.emit(val_a, val_b, keep_open)