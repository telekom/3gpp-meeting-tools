# --- File: src/modules/word_tools/ui/word_tabs.py ---
import webbrowser
from pathlib import Path

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QKeySequence
from PyQt5.QtWidgets import (
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QAbstractItemView,
    QShortcut,
    QMessageBox,
    QPushButton,
    QSpinBox,
    QTreeWidget,
    QTreeWidgetItem,
    QStackedWidget,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)
import win32com.client
import pythoncom
from core.ui.ui_components import (
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    InteractiveDropLabel
)
from modules.word_tools.core.word_config import WordConfig
from modules.word_tools.core.word_excerpt_extractor import (
    WordExcerptThread,
    WordHeadingScanThread,
)
from modules.word_tools.core.libreoffice_converter import (
    LIBREOFFICE_DOWNLOAD_URL,
    is_libreoffice_available,
    find_libreoffice_executable,
    resolve_soffice_binary,
)


class DocumentSelectorPane(QWidget):
    """A symmetric, reusable widget handling Local, Open, and URL inputs."""
    def __init__(self, title: str):
        super().__init__()
        self.title = title
        self.selected_files = []
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(6)

        lbl = QLabel(f"<b>{self.title}</b>")
        lbl.setStyleSheet("color: #1E293B; margin-bottom: 2px;")
        layout.addWidget(lbl)

        self.tabs = QTabWidget()
        self.tabs.setObjectName("selector_tabs")
        self.drop_tab = QWidget()
        drop_layout = QVBoxLayout(self.drop_tab)
        self.drop_zone = InteractiveDropLabel("Drop .docx here", [".docx"])
        self.drop_zone.file_dropped.connect(self._on_drop)
        drop_layout.addWidget(self.drop_zone)
        self.tabs.addTab(self.drop_tab, "📁 Local")

        self.open_tab = QWidget()
        open_layout = QVBoxLayout(self.open_tab)
        open_layout.setSpacing(6)
        self.open_combo = QComboBox()
        self.refresh_btn = QPushButton("↻ Refresh Active Documents")
        self.refresh_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.refresh_btn.clicked.connect(self.poll_open_documents)

        open_layout.addWidget(QLabel("Select an open Word document:"))
        open_layout.addWidget(self.open_combo)
        open_layout.addWidget(self.refresh_btn)
        open_layout.addStretch()
        self.tabs.addTab(self.open_tab, "🖥️ Open Docs")

        self.url_tab = QWidget()
        url_layout = QVBoxLayout(self.url_tab)
        url_layout.setSpacing(6)

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
            self.selected_files = files
            if len(files) == 1:
                self.drop_zone.set_state("ready", f"Ready:\n{Path(files[0]).name}")
            else:
                self.drop_zone.set_state("ready", f"Ready:\n{len(files)} files queued for batch")

    def poll_open_documents(self):
        self.open_combo.clear()
        try:
            pythoncom.CoInitialize()
            word = win32com.client.GetActiveObject("Word.Application")
            for doc in word.Documents:
                doc_name = str(doc.Name).strip()
                if doc_name:
                    self.open_combo.addItem(doc_name, doc.FullName)
            if self.open_combo.count() == 0:
                self.open_combo.addItem("No open documents detected.", "")
        except Exception:
            self.open_combo.addItem("No open documents detected.", "")
        finally:
            pythoncom.CoUninitialize()

    def get_inputs(self) -> list:
        idx = self.tabs.currentIndex()
        if idx == 0:
            return self.selected_files
        elif idx == 1:
            doc = self.open_combo.currentData()
            return [doc] if doc else []
        elif idx == 2:
            url = self.url_input.text().strip()
            return [url] if url else []
        return []

    def _on_tab_changed(self, index):
        if index == 1:
            self.poll_open_documents()


class WordExtractorTab(QWidget):
    extract_visio_requested = pyqtSignal(str)
    split_doc_requested = pyqtSignal(str, str, int)
    compare_doc_requested = pyqtSignal(str, str, bool)
    convert_doc_requested = pyqtSignal(str, str)

    def __init__(self):
        super().__init__()
        self._setup_ui()
        self._check_libreoffice_status()

    def _setup_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(10)

        self.status_banner = QWidget()
        banner_layout = QHBoxLayout(self.status_banner)
        banner_layout.setContentsMargins(12, 8, 12, 8)
        banner_layout.setSpacing(10)

        self.banner_icon = QLabel("⚠️")
        self.banner_text = QLabel()
        self.banner_text.setWordWrap(True)
        self.locate_btn = QPushButton("📂 Locate Executable")
        self.locate_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.locate_btn.setToolTip("Select soffice.exe or LibreOfficePortable.exe from your disk")
        self.locate_btn.clicked.connect(self._browse_for_libreoffice)
        self.download_btn = QPushButton("📥 Download Portable")
        self.download_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.download_btn.clicked.connect(lambda: webbrowser.open(LIBREOFFICE_DOWNLOAD_URL))

        banner_layout.addWidget(self.banner_icon)
        banner_layout.addWidget(self.banner_text, 1)
        banner_layout.addWidget(self.locate_btn)
        banner_layout.addWidget(self.download_btn)
        layout.addWidget(self.status_banner)

        switcher_layout = QHBoxLayout()
        switcher_layout.addWidget(QLabel("<b>⚙️ Operation Type:</b>"))
        self.op_combo = QComboBox()
        self.op_combo.addItems([
            "Convert Legacy .doc to .docx (Explicit LibreOffice)",
            "Extract Embedded Visio Diagrams",
            "Subtractive Slicing (Split by Clause)",
            "Extract Sections by Heading",
            "Compare Documents (Native Word Diff)",
            "Convert Document Format (Auto / Word)",
        ])
        switcher_layout.addWidget(self.op_combo)
        switcher_layout.addStretch()
        layout.addLayout(switcher_layout)

        self.stack = QStackedWidget()

        self.card_lo_convert = QWidget()
        lo_layout = QVBoxLayout(self.card_lo_convert)
        self.lo_drop = InteractiveDropLabel(
            "📥 Drag & Drop legacy Word 97-2003 (.doc) file(s) here to convert to .docx explicitly via LibreOffice",
            [".doc"],
        )
        self.lo_drop.file_dropped.connect(self._on_lo_dropped)
        lo_layout.addWidget(self.lo_drop)
        self.stack.addWidget(self.card_lo_convert)

        self.card_visio = QWidget()
        visio_layout = QVBoxLayout(self.card_visio)
        self.visio_drop = InteractiveDropLabel(
            "📥 Drag & Drop a .docx file here to extract its Visio components", [".docx"]
        )
        self.visio_drop.file_dropped.connect(
            lambda files: [self.extract_visio_requested.emit(f) for f in files]
        )
        visio_layout.addWidget(self.visio_drop)
        self.stack.addWidget(self.card_visio)

        self.card_split = QWidget()
        split_layout = QVBoxLayout(self.card_split)

        form = QFormLayout()
        self.prefix_input = QLineEdit("6.")
        self.depth_input = QSpinBox()
        self.depth_input.setRange(1, 6)
        self.depth_input.setValue(2)

        form.addRow("Target Clause Prefix:", self.prefix_input)
        form.addRow("Heading Depth Hierarchy:", self.depth_input)
        split_layout.addLayout(form)
        self.split_drop = InteractiveDropLabel(
            "📥 Drag & Drop a .docx file here to slice it into chapters", [".docx"]
        )
        self.split_drop.file_dropped.connect(
            lambda files: [
                self.split_doc_requested.emit(
                    f, self.prefix_input.text().strip(), self.depth_input.value()
                )
                for f in files
            ]
        )
        split_layout.addWidget(self.split_drop)
        self.stack.addWidget(self.card_split)

        # Interactive heading browser / excerpt extractor. This card is self-contained
        # so adding it does not require another main-window signal/worker integration.
        self.card_excerpt = QWidget()
        excerpt_layout = QVBoxLayout(self.card_excerpt)
        excerpt_layout.setSpacing(8)

        excerpt_top = QHBoxLayout()
        self.excerpt_order_combo = QComboBox()
        self.excerpt_order_combo.addItem("Filename (natural alphabetical)", "filename")
        self.excerpt_order_combo.addItem("Drop order", "drop")
        excerpt_top.addWidget(QLabel("<b>Source order:</b>"))
        excerpt_top.addWidget(self.excerpt_order_combo)
        excerpt_top.addStretch()
        excerpt_layout.addLayout(excerpt_top)

        self.excerpt_drop = InteractiveDropLabel(
            "📥 Drop one or more .docx files here to scan their headings", [".docx"]
        )
        self.excerpt_drop.file_dropped.connect(self._scan_excerpt_files)
        excerpt_layout.addWidget(self.excerpt_drop)

        source_row = QHBoxLayout()
        self.excerpt_source_list = QListWidget()
        self.excerpt_source_list.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.excerpt_source_list.setMaximumHeight(105)
        self.excerpt_source_list.setToolTip("Select one or more source documents. Press Delete to remove them.")
        source_row.addWidget(self.excerpt_source_list, 1)

        source_buttons = QVBoxLayout()
        self.excerpt_remove_btn = QPushButton("Remove Selected")
        self.excerpt_remove_btn.setEnabled(False)
        self.excerpt_remove_btn.clicked.connect(self._remove_excerpt_sources)
        self.excerpt_clear_sources_btn = QPushButton("Clear All")
        self.excerpt_clear_sources_btn.setEnabled(False)
        self.excerpt_clear_sources_btn.clicked.connect(self._clear_excerpt_sources)
        source_buttons.addWidget(self.excerpt_remove_btn)
        source_buttons.addWidget(self.excerpt_clear_sources_btn)
        source_buttons.addStretch()
        source_row.addLayout(source_buttons)
        excerpt_layout.addLayout(source_row)

        self.excerpt_source_list.itemSelectionChanged.connect(
            lambda: self.excerpt_remove_btn.setEnabled(bool(self.excerpt_source_list.selectedItems()))
        )
        self._excerpt_delete_shortcut = QShortcut(QKeySequence.Delete, self.excerpt_source_list)
        self._excerpt_delete_shortcut.activated.connect(self._remove_excerpt_sources)
        self.excerpt_order_combo.currentIndexChanged.connect(self._rescan_excerpt_sources)

        filter_row = QHBoxLayout()
        self.excerpt_filter = QLineEdit()
        self.excerpt_filter.setPlaceholderText("Filter headings (e.g. registration, NWDAF, 6.3.2)...")
        self.excerpt_filter.textChanged.connect(self._filter_excerpt_tree)
        filter_row.addWidget(self.excerpt_filter, 1)
        self.excerpt_expand_btn = QPushButton("Expand All")
        self.excerpt_collapse_btn = QPushButton("Collapse All")
        self.excerpt_expand_btn.clicked.connect(lambda: self.excerpt_tree.expandAll())
        self.excerpt_collapse_btn.clicked.connect(lambda: self.excerpt_tree.collapseAll())
        filter_row.addWidget(self.excerpt_expand_btn)
        filter_row.addWidget(self.excerpt_collapse_btn)
        excerpt_layout.addLayout(filter_row)

        self.excerpt_tree = QTreeWidget()
        self.excerpt_tree.setHeaderLabels(["Heading", "Source file"])
        self.excerpt_tree.setColumnWidth(0, 560)
        excerpt_layout.addWidget(self.excerpt_tree, 1)

        action_row = QHBoxLayout()
        self.excerpt_clear_btn = QPushButton("Clear Selection")
        self.excerpt_clear_btn.clicked.connect(self._clear_excerpt_selection)
        self.excerpt_extract_btn = QPushButton("✂️ Extract Selected Sections")
        self.excerpt_extract_btn.setObjectName("primaryBtn")
        self.excerpt_extract_btn.setEnabled(False)
        self.excerpt_extract_btn.clicked.connect(self._extract_selected_sections)
        action_row.addWidget(self.excerpt_clear_btn)
        action_row.addStretch()
        action_row.addWidget(self.excerpt_extract_btn)
        excerpt_layout.addLayout(action_row)
        self.stack.addWidget(self.card_excerpt)

        self._excerpt_input_paths = []
        self._excerpt_paths = []
        self._excerpt_headings = []
        self._excerpt_scan_thread = None
        self._excerpt_extract_thread = None

        self.card_compare = QWidget()
        compare_layout = QVBoxLayout(self.card_compare)
        compare_layout.setSpacing(10)

        panes_layout = QHBoxLayout()
        self.pane_a = DocumentSelectorPane("📄 DOCUMENT A (Original)")
        self.pane_b = DocumentSelectorPane("📄 DOCUMENT B (Revised)")
        panes_layout.addWidget(self.pane_a)
        panes_layout.addWidget(self.pane_b)
        compare_layout.addLayout(panes_layout)
        self.keep_open_cb = QCheckBox("Keep source documents (A and B) open after comparison")
        self.keep_open_cb.setChecked(True)
        compare_layout.addWidget(self.keep_open_cb)

        self.run_compare_btn = QPushButton("⚖️ Run Word Comparison")
        self.run_compare_btn.setObjectName("primaryBtn")
        self.run_compare_btn.clicked.connect(self._trigger_comparison)
        compare_layout.addWidget(self.run_compare_btn)
        self.stack.addWidget(self.card_compare)

        self.card_convert = QWidget()
        convert_layout = QVBoxLayout(self.card_convert)
        convert_layout.setSpacing(10)

        self.pane_convert = DocumentSelectorPane("📄 DOCUMENT TO CONVERT")
        convert_layout.addWidget(self.pane_convert)
        conv_form = QFormLayout()
        self.format_combo = QComboBox()
        self.format_combo.addItems(["PDF", "DOCX", "HTML", "XPS", "RTF", "TXT"])
        conv_form.addRow("Target Format:", self.format_combo)
        convert_layout.addLayout(conv_form)
        self.run_convert_btn = QPushButton("🔄 Convert Document")
        self.run_convert_btn.setObjectName("primaryBtn")
        self.run_convert_btn.clicked.connect(self._trigger_conversion)
        convert_layout.addWidget(self.run_convert_btn)
        self.stack.addWidget(self.card_convert)

        layout.addWidget(self.stack)
        self.setLayout(layout)
        self.op_combo.currentIndexChanged.connect(self.stack.setCurrentIndex)

    def _scan_excerpt_files(self, files):
        if not files:
            return
        # A new drop defines the current document set. Preserve this exact list so
        # the optional drop-order mode remains meaningful after removals/rescans.
        self._excerpt_input_paths = list(dict.fromkeys(str(Path(f)) for f in files))
        self._refresh_excerpt_source_list()
        self._rescan_excerpt_sources()

    def _refresh_excerpt_source_list(self):
        self.excerpt_source_list.clear()
        for path in self._excerpt_input_paths:
            self.excerpt_source_list.addItem(Path(path).name)
        has_sources = bool(self._excerpt_input_paths)
        self.excerpt_clear_sources_btn.setEnabled(has_sources)
        self.excerpt_remove_btn.setEnabled(False)

    def _invalidate_excerpt_scan(self):
        self._excerpt_paths = []
        self._excerpt_headings = []
        self.excerpt_tree.clear()
        self.excerpt_extract_btn.setEnabled(False)

    def _rescan_excerpt_sources(self):
        self._invalidate_excerpt_scan()
        if not self._excerpt_input_paths:
            self.excerpt_drop.set_state("idle", "📥 Drop one or more .docx files here to scan their headings")
            return
        if self._excerpt_scan_thread is not None and self._excerpt_scan_thread.isRunning():
            return
        self.excerpt_drop.set_state("ready", f"Scanning {len(self._excerpt_input_paths)} document(s)...")
        mode = self.excerpt_order_combo.currentData()
        self._excerpt_scan_thread = WordHeadingScanThread(list(self._excerpt_input_paths), mode)
        self._excerpt_scan_thread.scanned.connect(self._on_excerpt_scanned)
        self._excerpt_scan_thread.failed.connect(self._on_excerpt_scan_failed)
        self._excerpt_scan_thread.start()

    def _remove_excerpt_sources(self):
        selected_rows = sorted(
            {self.excerpt_source_list.row(item) for item in self.excerpt_source_list.selectedItems()},
            reverse=True,
        )
        if not selected_rows:
            return
        for row in selected_rows:
            if 0 <= row < len(self._excerpt_input_paths):
                del self._excerpt_input_paths[row]
        self._refresh_excerpt_source_list()
        self._rescan_excerpt_sources()

    def _clear_excerpt_sources(self):
        self._excerpt_input_paths = []
        self._refresh_excerpt_source_list()
        self._invalidate_excerpt_scan()
        self.excerpt_drop.set_state("idle", "📥 Drop one or more .docx files here to scan their headings")

    def _on_excerpt_scanned(self, ordered_paths, headings):
        self._excerpt_paths = list(ordered_paths)
        self._excerpt_headings = list(headings)
        self._populate_excerpt_tree()
        self.excerpt_drop.set_state(
            "ready", f"Ready: {len(headings)} headings from {len(ordered_paths)} document(s)"
        )
        self.excerpt_extract_btn.setEnabled(bool(headings))

    def _on_excerpt_scan_failed(self, message):
        self.excerpt_drop.set_state("error", "Heading scan failed")
        QMessageBox.critical(self, "Heading Scan Failed", message)

    def _populate_excerpt_tree(self):
        self.excerpt_tree.clear()
        by_source = {}
        for h in self._excerpt_headings:
            by_source.setdefault(h.source_path, []).append(h)

        for source_path in self._excerpt_paths:
            source_item = QTreeWidgetItem([Path(source_path).name, Path(source_path).name])
            source_item.setFlags(source_item.flags() & ~Qt.ItemIsUserCheckable)
            self.excerpt_tree.addTopLevelItem(source_item)
            parents = {}
            for h in by_source.get(source_path, []):
                item = QTreeWidgetItem([h.text, Path(source_path).name])
                item.setData(0, Qt.UserRole, h.id)
                item.setCheckState(0, Qt.Unchecked)
                parent = None
                for level in range(h.level - 1, 0, -1):
                    if level in parents:
                        parent = parents[level]
                        break
                (parent or source_item).addChild(item)
                parents[h.level] = item
                for level in list(parents):
                    if level > h.level:
                        del parents[level]
        self.excerpt_tree.expandToDepth(1)

    def _filter_excerpt_tree(self, text):
        needle = text.strip().lower()

        def visit(item):
            own_match = not needle or needle in item.text(0).lower()
            child_match = False
            for i in range(item.childCount()):
                child_match = visit(item.child(i)) or child_match
            visible = own_match or child_match
            item.setHidden(not visible)
            if needle and child_match:
                item.setExpanded(True)
            return visible

        for i in range(self.excerpt_tree.topLevelItemCount()):
            visit(self.excerpt_tree.topLevelItem(i))

    def _clear_excerpt_selection(self):
        def clear(item):
            if item.data(0, Qt.UserRole):
                item.setCheckState(0, Qt.Unchecked)
            for i in range(item.childCount()):
                clear(item.child(i))
        for i in range(self.excerpt_tree.topLevelItemCount()):
            clear(self.excerpt_tree.topLevelItem(i))

    def _selected_excerpt_ids(self):
        selected = []
        def collect(item):
            heading_id = item.data(0, Qt.UserRole)
            if heading_id and item.checkState(0) == Qt.Checked:
                selected.append(heading_id)
            for i in range(item.childCount()):
                collect(item.child(i))
        for i in range(self.excerpt_tree.topLevelItemCount()):
            collect(self.excerpt_tree.topLevelItem(i))
        return selected

    def _extract_selected_sections(self):
        selected = self._selected_excerpt_ids()
        if not selected:
            QMessageBox.information(self, "No Selection", "Select at least one heading to extract.")
            return
        default_dir = str(Path(self._excerpt_paths[0]).parent) if self._excerpt_paths else ""
        output, _ = QFileDialog.getSaveFileName(
            self, "Save Extracted Sections", str(Path(default_dir) / "extracted_sections.docx"),
            "Word Document (*.docx)"
        )
        if not output:
            return
        if not output.lower().endswith(".docx"):
            output += ".docx"
        self.excerpt_extract_btn.setEnabled(False)
        self.excerpt_extract_btn.setText("⏳ Extracting...")
        self._excerpt_extract_thread = WordExcerptThread(
            self._excerpt_paths, self._excerpt_headings, selected, output
        )
        self._excerpt_extract_thread.finished_path.connect(self._on_excerpt_finished)
        self._excerpt_extract_thread.failed.connect(self._on_excerpt_extract_failed)
        self._excerpt_extract_thread.finished.connect(self._reset_excerpt_button)
        self._excerpt_extract_thread.start()

    def _on_excerpt_finished(self, path):
        QMessageBox.information(self, "Extraction Complete", f"Created:\n{path}")

    def _on_excerpt_extract_failed(self, message):
        QMessageBox.critical(self, "Extraction Failed", message)

    def _reset_excerpt_button(self):
        self.excerpt_extract_btn.setText("✂️ Extract Selected Sections")
        self.excerpt_extract_btn.setEnabled(bool(self._excerpt_headings))

    def _check_libreoffice_status(self):
        """
        Passive status check only.

        Do not start LibreOffice here. In particular, LibreOffice Portable's
        --version path can open an interactive console that waits for Enter.
        """
        executable = find_libreoffice_executable()
        if executable:
            runtime_type = (
                "Portable"
                if executable.name.lower() == "libreofficeportable.exe"
                else "Installed"
            )
            self.status_banner.setStyleSheet(
                "background-color: #ECFDF5; border: 1px solid #A7F3D0; border-radius: 6px;"
            )
            self.banner_icon.setText("🟢")
            self.banner_text.setText(
                f"<b>LibreOffice configured ({runtime_type}):</b> "
                f"Using <code>{executable}</code>"
            )
            self.banner_text.setStyleSheet("color: #065F46;")
            self.locate_btn.setText("⚙️ Change Path")
            self.download_btn.setVisible(False)
        else:
            self.status_banner.setStyleSheet(
                "background-color: #FFFBEB; border: 1px solid #FDE68A; border-radius: 6px;"
            )
            self.banner_icon.setText("⚠️")
            self.banner_text.setText(
                "<b>LibreOffice not detected.</b> If using LibreOffice Portable, "
                "click 'Locate Executable' and select <code>LibreOfficePortable.exe</code>."
            )
            self.banner_text.setStyleSheet("color: #B45309;")
            self.locate_btn.setText("📂 Locate Executable")
            self.download_btn.setVisible(True)

    def _browse_for_libreoffice(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select LibreOffice Binary or Portable Launcher",
            "",
            "LibreOffice Executable (*soffice.exe *LibreOfficePortable.exe *.exe);;All Files (*.*)",
        )
        if not file_path:
            return

        resolved = resolve_soffice_binary(file_path)
        if resolved:
            # Deliberately do not launch/probe it here. A real conversion is the
            # runtime check; configuration must remain side-effect free.
            WordConfig.set_libreoffice_path(str(resolved))
            self._check_libreoffice_status()
            QMessageBox.information(
                self,
                "LibreOffice Configured",
                f"Successfully linked LibreOffice executable:\n{resolved}",
            )
        else:
            QMessageBox.warning(
                self,
                "Invalid Executable",
                "Could not locate LibreOffice from the selected path.\n"
                "Please select 'soffice.exe' or 'LibreOfficePortable.exe'.",
            )

    def _on_lo_dropped(self, files):
        if not is_libreoffice_available():
            self._check_libreoffice_status()
        for f in files:
            if f:
                self.convert_doc_requested.emit(f, "docx_libreoffice")

    def _trigger_comparison(self):
        val_a_list = self.pane_a.get_inputs()
        val_b_list = self.pane_b.get_inputs()
        val_a = val_a_list[0] if val_a_list else ""
        val_b = val_b_list[0] if val_b_list else ""
        keep_open = self.keep_open_cb.isChecked()

        if val_a and val_b:
            self.compare_doc_requested.emit(val_a, val_b, keep_open)

    def _trigger_conversion(self):
        source_docs = self.pane_convert.get_inputs()
        target_fmt = self.format_combo.currentText().lower()
        for doc in source_docs:
            if doc:
                self.convert_doc_requested.emit(doc, target_fmt)
