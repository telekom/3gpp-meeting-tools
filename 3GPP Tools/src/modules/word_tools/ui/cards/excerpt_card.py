"""Card: Interactive Heading Excerpt & Chapter Slicing."""

import logging
from pathlib import Path
from typing import List, Optional

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QKeySequence
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QHeaderView,
    QShortcut,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from core.ui.ui_components import (
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    CARD_FRAME_STYLE,
    InteractiveDropLabel,
    SEARCH_INPUT_STYLE,
)
from modules.word_tools.core.word_excerpt_extractor import (
    WordExcerptThread,
    WordHeadingScanThread,
)

logger = logging.getLogger(__name__)


class WordExcerptCard(QWidget):
    """Card: Interactive clause heading scanner, hierarchical tree viewer, and clause excerpt generator."""

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self._excerpt_paths: List[str] = []
        self._excerpt_input_paths: List[str] = []
        self._excerpt_headings: list = []
        self._excerpt_scan_thread: Optional[WordHeadingScanThread] = None
        self._excerpt_extract_thread: Optional[WordExcerptThread] = None
        self._active_scan_threads: list = []
        self._excerpt_scan_gen: int = 0
        self._setup_ui()

    def _setup_ui(self):
        excerpt_layout = QVBoxLayout(self)
        excerpt_layout.setContentsMargins(10, 10, 10, 10)
        excerpt_layout.setSpacing(10)

        excerpt_layout.setContentsMargins(12, 10, 12, 10)
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
        self.excerpt_remove_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.excerpt_remove_btn.setEnabled(False)
        self.excerpt_remove_btn.clicked.connect(self._remove_excerpt_sources)
        self.excerpt_clear_sources_btn = QPushButton("Clear All")
        self.excerpt_clear_sources_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.excerpt_clear_sources_btn.setEnabled(False)
        self.excerpt_clear_sources_btn.clicked.connect(self._clear_excerpt_sources)
        source_buttons.addWidget(self.excerpt_remove_btn)
        source_buttons.addWidget(self.excerpt_clear_sources_btn)
        source_buttons.addStretch()
        source_row.addLayout(source_buttons)
        excerpt_layout.addLayout(source_row)

        self.excerpt_progress_label = QLabel("Ready to scan.")
        self.excerpt_progress_label.setVisible(False)
        self.excerpt_progress = QProgressBar()
        self.excerpt_progress.setRange(0, 100)
        self.excerpt_progress.setValue(0)
        self.excerpt_progress.setTextVisible(True)
        self.excerpt_progress.setVisible(False)
        excerpt_layout.addWidget(self.excerpt_progress_label)
        excerpt_layout.addWidget(self.excerpt_progress)

        self.excerpt_source_list.itemSelectionChanged.connect(
            lambda: self.excerpt_remove_btn.setEnabled(bool(self.excerpt_source_list.selectedItems()))
        )
        self._excerpt_delete_shortcut = QShortcut(QKeySequence.Delete, self.excerpt_source_list)
        self._excerpt_delete_shortcut.activated.connect(self._remove_excerpt_sources)
        self.excerpt_order_combo.currentIndexChanged.connect(self._rescan_excerpt_sources)

        filter_row = QHBoxLayout()
        self.excerpt_filter = QLineEdit()
        self.excerpt_filter.setStyleSheet(SEARCH_INPUT_STYLE)
        self.excerpt_filter.setPlaceholderText("Filter headings (e.g. registration, NWDAF, 6.3.2)...")
        self.excerpt_filter.textChanged.connect(self._filter_excerpt_tree)
        filter_row.addWidget(self.excerpt_filter, 1)

        self.lbl_excerpt_match_count = QLabel()
        self.lbl_excerpt_match_count.setStyleSheet("font-size: 11px; color: #475569; padding-right: 4px;")
        filter_row.addWidget(self.lbl_excerpt_match_count)

        self.excerpt_expand_btn = QPushButton("Expand All")
        self.excerpt_expand_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.excerpt_collapse_btn = QPushButton("Collapse All")
        self.excerpt_collapse_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.excerpt_expand_btn.clicked.connect(lambda: self.excerpt_tree.expandAll())
        self.excerpt_collapse_btn.clicked.connect(lambda: self.excerpt_tree.collapseAll())
        filter_row.addWidget(self.excerpt_expand_btn)
        filter_row.addWidget(self.excerpt_collapse_btn)
        excerpt_layout.addLayout(filter_row)

        self.excerpt_tree = QTreeWidget()
        self.excerpt_tree.setHeaderLabels(["Heading", "Source file"])
        header = self.excerpt_tree.header()
        header.setStretchLastSection(False)
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        excerpt_layout.addWidget(self.excerpt_tree, 1)

        action_row = QHBoxLayout()
        self.excerpt_select_all_btn = QPushButton("Select All")
        self.excerpt_select_all_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.excerpt_select_all_btn.clicked.connect(self._select_all_excerpt_headings)
        action_row.addWidget(self.excerpt_select_all_btn)

        self.excerpt_select_l1_btn = QPushButton("Select Level 1 Only")
        self.excerpt_select_l1_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.excerpt_select_l1_btn.clicked.connect(self._select_level1_excerpt_headings)
        action_row.addWidget(self.excerpt_select_l1_btn)

        self.excerpt_clear_btn = QPushButton("Clear Selection")
        self.excerpt_clear_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.excerpt_clear_btn.clicked.connect(self._clear_excerpt_selection)
        action_row.addWidget(self.excerpt_clear_btn)

        action_row.addStretch()

        self.chk_also_markdown = QCheckBox("Also export as Markdown (with figures)")
        self.chk_also_markdown.setToolTip(
            "Automatically convert the extracted sections to Markdown (.md) and extract all embedded figures and diagrams into a companion media folder."
        )
        self.chk_also_markdown.setStyleSheet("color: #334155; font-weight: 500; margin-right: 8px;")
        action_row.addWidget(self.chk_also_markdown)

        self.excerpt_extract_btn = QPushButton("✂️ Extract Selected Sections")
        self.excerpt_extract_btn.setObjectName("primaryBtn")
        self.excerpt_extract_btn.setEnabled(False)
        self.excerpt_extract_btn.clicked.connect(self._extract_selected_sections)
        action_row.addWidget(self.excerpt_extract_btn)
        excerpt_layout.addLayout(action_row)


    def _scan_excerpt_files(self, files):
        if not files:
            return

        # Additional drops extend the current source set instead of replacing it.
        # Compare normalized absolute paths so the same Windows file is not added
        # twice through different path spellings/casing. Preserve first-drop order.
        existing = {
            str(Path(path).resolve()).casefold()
            for path in self._excerpt_input_paths
        }
        added = False
        for file_path in files:
            path = str(Path(file_path).resolve())
            key = path.casefold()
            if key in existing:
                continue
            self._excerpt_input_paths.append(path)
            existing.add(key)
            added = True

        if not added:
            return

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

    def _disconnect_running_scan(self):
        self._excerpt_scan_gen += 1
        if self._excerpt_scan_thread is not None:
            old_thread = self._excerpt_scan_thread
            try:
                old_thread.scanned.disconnect()
            except Exception:
                pass
            try:
                old_thread.progress.disconnect()
            except Exception:
                pass
            try:
                old_thread.failed.disconnect()
            except Exception:
                pass
            if old_thread.isRunning():
                old_thread.requestInterruption()
                self._active_scan_threads.append(old_thread)
                old_thread.finished.connect(
                    lambda t=old_thread: self._active_scan_threads.remove(t) if t in self._active_scan_threads else None
                )
            self._excerpt_scan_thread = None

    def _rescan_excerpt_sources(self):
        self._disconnect_running_scan()
        self._invalidate_excerpt_scan()
        if not self._excerpt_input_paths:
            self.excerpt_drop.set_state("ready", "📥 Drop one or more .docx files here to scan their headings")
            self.excerpt_progress.setVisible(False)
            self.excerpt_progress_label.setVisible(False)
            return
        self.excerpt_drop.set_state("ready", f"Scanning {len(self._excerpt_input_paths)} document(s)...")
        self.excerpt_progress.setValue(0)
        self.excerpt_progress.setVisible(True)
        self.excerpt_progress_label.setText(f"Preparing to parse {len(self._excerpt_input_paths)} document(s)...")
        self.excerpt_progress_label.setVisible(True)
        mode = self.excerpt_order_combo.currentData()
        gen = self._excerpt_scan_gen
        thread = WordHeadingScanThread(list(self._excerpt_input_paths), mode)
        self._excerpt_scan_thread = thread
        thread.scanned.connect(
            lambda ordered_paths, headings, g=gen: self._on_excerpt_scanned(ordered_paths, headings) if g == self._excerpt_scan_gen else None
        )
        thread.progress.connect(
            lambda percent, message, g=gen: self._on_excerpt_scan_progress(percent, message) if g == self._excerpt_scan_gen else None
        )
        thread.failed.connect(
            lambda message, g=gen: self._on_excerpt_scan_failed(message) if g == self._excerpt_scan_gen else None
        )
        thread.start()

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
        self._disconnect_running_scan()
        self._excerpt_input_paths = []
        self._refresh_excerpt_source_list()
        self._invalidate_excerpt_scan()
        self.excerpt_drop.set_state("ready", "📥 Drop one or more .docx files here to scan their headings")
        self.excerpt_progress.setVisible(False)
        self.excerpt_progress_label.setVisible(False)

    def _on_excerpt_scan_progress(self, percent, message):
        self.excerpt_progress.setValue(percent)
        self.excerpt_progress_label.setText(message)

    def _on_excerpt_scanned(self, ordered_paths, headings):
        self._excerpt_paths = list(ordered_paths)
        self._excerpt_headings = list(headings)
        self._populate_excerpt_tree()
        self.excerpt_drop.set_state(
            "ready", f"Ready: {len(headings)} headings from {len(ordered_paths)} document(s)"
        )
        self.excerpt_extract_btn.setEnabled(bool(headings))
        self.excerpt_progress.setValue(100)
        self.excerpt_progress_label.setText(f"Scan complete: {len(headings)} headings found.")

    def _on_excerpt_scan_failed(self, message):
        self.excerpt_drop.set_state("error", "Heading scan failed")
        self.excerpt_progress.setVisible(True)
        self.excerpt_progress_label.setVisible(True)
        self.excerpt_progress_label.setText("Heading scan failed.")
        QMessageBox.critical(self, "Heading Scan Failed", message)

    def _populate_excerpt_tree(self):
        self.excerpt_tree.clear()
        by_source = {}
        for h in self._excerpt_headings:
            by_source.setdefault(h.source_path, []).append(h)

        for source_path in self._excerpt_paths:
            source_item = QTreeWidgetItem([f"📄 {Path(source_path).name}", Path(source_path).name])
            source_item.setFlags(source_item.flags() & ~Qt.ItemIsUserCheckable)
            self.excerpt_tree.addTopLevelItem(source_item)
            parents = {}
            for h in by_source.get(source_path, []):
                item = QTreeWidgetItem([f"[H{h.level}] {h.text}", Path(source_path).name])
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
        matches = 0

        def visit(item):
            nonlocal matches
            is_heading = bool(item.data(0, Qt.UserRole))
            own_match = not needle or needle in item.text(0).lower()
            if needle and is_heading and own_match:
                matches += 1
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

        if hasattr(self, "lbl_excerpt_match_count"):
            if needle:
                self.lbl_excerpt_match_count.setText(f"{matches} match(es)")
            else:
                self.lbl_excerpt_match_count.setText("")

    def _select_all_excerpt_headings(self):
        def select(item):
            if item.data(0, Qt.UserRole) and not item.isHidden():
                item.setCheckState(0, Qt.Checked)
            for i in range(item.childCount()):
                select(item.child(i))
        for i in range(self.excerpt_tree.topLevelItemCount()):
            select(self.excerpt_tree.topLevelItem(i))

    def _select_level1_excerpt_headings(self):
        self._clear_excerpt_selection()
        for i in range(self.excerpt_tree.topLevelItemCount()):
            src_item = self.excerpt_tree.topLevelItem(i)
            for c in range(src_item.childCount()):
                item = src_item.child(c)
                if item.data(0, Qt.UserRole) and not item.isHidden():
                    item.setCheckState(0, Qt.Checked)

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
        also_md = self.chk_also_markdown.isChecked()
        self._excerpt_extract_thread = WordExcerptThread(
            self._excerpt_paths,
            self._excerpt_headings,
            selected,
            output,
            export_markdown=also_md,
        )
        self.excerpt_progress.setRange(0, 100)
        self.excerpt_progress.setValue(0)
        self.excerpt_progress.setVisible(True)
        self.excerpt_progress_label.setText("Preparing extraction...")
        self.excerpt_progress_label.setVisible(True)
        self._excerpt_extract_thread.finished_path.connect(self._on_excerpt_finished)
        self._excerpt_extract_thread.failed.connect(self._on_excerpt_extract_failed)
        self._excerpt_extract_thread.progress.connect(self._on_excerpt_extract_progress)
        self._excerpt_extract_thread.finished.connect(self._reset_excerpt_button)
        self._excerpt_extract_thread.start()

    def _on_excerpt_extract_progress(self, percent, message):
        # Fragment preparation can legitimately take minutes for very large DOCX
        # packages and does not expose meaningful sub-progress.  Use Qt's busy
        # indicator during that isolated stage so the UI visibly remains alive.
        if message.startswith("Preparing excerpt from "):
            self.excerpt_progress.setRange(0, 0)
        else:
            if self.excerpt_progress.minimum() == 0 and self.excerpt_progress.maximum() == 0:
                self.excerpt_progress.setRange(0, 100)
            self.excerpt_progress.setValue(percent)
        self.excerpt_progress_label.setText(message)

    def _on_excerpt_finished(self, path):
        thread = self._excerpt_extract_thread
        if thread and getattr(thread, "export_markdown", False) and getattr(thread, "md_path", None):
            md_path = thread.md_path
            img_count = getattr(thread, "images_count", 0)
            img_msg = f"\n• Figures extracted: {img_count}" if img_count else "\n• Figures extracted: 0"
            msg = (
                f"Successfully extracted selected sections!\n\n"
                f"📄 Word Document:\n{path}\n\n"
                f"📝 Markdown Document:\n{md_path}"
                f"{img_msg}"
            )
            QMessageBox.information(self, "Extraction Complete", msg)
        else:
            QMessageBox.information(self, "Extraction Complete", f"Created:\n{path}")

    def _on_excerpt_extract_failed(self, message):
        QMessageBox.critical(self, "Extraction Failed", message)

    def _reset_excerpt_button(self):
        self.excerpt_extract_btn.setText("✂️ Extract Selected Sections")
        self.excerpt_extract_btn.setEnabled(bool(self._excerpt_headings))

