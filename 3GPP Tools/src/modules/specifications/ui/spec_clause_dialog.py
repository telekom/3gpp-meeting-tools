# --- File: src/modules/specifications/ui/spec_clause_dialog.py ---
import logging
from pathlib import Path
from typing import Optional, List

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtWidgets import (
    QCheckBox,
    QDialog,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
)

from core.ui.ui_components import (
    BADGE_STYLE_INFO,
    BADGE_STYLE_MUTED,
    BADGE_STYLE_PRIMARY,
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    CARD_FRAME_STYLE,
    SEARCH_INPUT_STYLE,
)

from modules.word_tools.core.word_excerpt_extractor import (
    HeadingInfo,
    WordExcerptThread,
)
from modules.word_tools.core.document_opener import open_document
from modules.specifications.ui.threads import SpecExcerptPrepThread

logger = logging.getLogger(__name__)

class SpecClauseExcerptDialog(QDialog):
    """
    Independent dialog enabling interactive clause/subclause extraction directly
    from 3GPP specifications. Handles automated fetching, multi-part ordering,
    legacy .doc conversion, search filtering, and clean Word excerpt generation.
    """

    extraction_completed = pyqtSignal(Path)

    def __init__(
        self,
        spec_num: str,
        spec_type: str,
        version: str,
        fname: str,
        url: str,
        spec_dl_dir: Path,
        parent=None,
    ):
        super().__init__(parent)
        self.spec_num = spec_num.strip()
        self.spec_type = (spec_type or "TS").upper().strip()
        self.version = version.strip()
        self.fname = fname.strip()
        self.url = url.strip()
        self.spec_dl_dir = Path(spec_dl_dir)

        self._prep_thread: Optional[SpecExcerptPrepThread] = None
        self._extract_thread: Optional[WordExcerptThread] = None

        self._ordered_paths: List[str] = []
        self._headings: List[HeadingInfo] = []
        self._is_updating_checks = False

        self.setWindowTitle(f"✂️ Extract Clauses: {self.spec_type} {self.spec_num} v{self.version}")
        self.setWindowFlags(self.windowFlags() | Qt.Window)
        self.resize(900, 700)
        self.setMinimumWidth(740)
        self.setMinimumHeight(540)

        self.setStyleSheet(f"""
            QDialog {{ background-color: #F8F9FA; }}
            {CARD_FRAME_STYLE}
        """)

        self._setup_ui()
        self._start_preparation()

    def _setup_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(16, 16, 16, 16)
        main_layout.setSpacing(10)

        # 1. Header Card
        header_card = QFrame()
        header_card.setObjectName("cardFrame")
        header_layout = QVBoxLayout(header_card)
        header_layout.setContentsMargins(14, 12, 14, 12)
        header_layout.setSpacing(6)

        title_row = QHBoxLayout()
        type_badge = QLabel(f"<b>{self.spec_type}</b>")
        type_badge.setStyleSheet(BADGE_STYLE_INFO)

        num_label = QLabel(f"<b>{self.spec_num}</b>")
        num_label.setStyleSheet("font-size: 16px; color: #1A202C; font-weight: bold;")

        ver_badge = QLabel(f"<b>v{self.version}</b>")
        ver_badge.setStyleSheet(BADGE_STYLE_PRIMARY)

        title_row.addWidget(type_badge)
        title_row.addWidget(num_label)
        title_row.addWidget(ver_badge)
        title_row.addStretch()

        self.source_summary_badge = QLabel("⏳ Initializing...")
        self.source_summary_badge.setStyleSheet(BADGE_STYLE_MUTED)
        title_row.addWidget(self.source_summary_badge)

        header_layout.addLayout(title_row)

        desc_label = QLabel(
            "Select specific clauses or subclauses below to extract into standalone Word (.docx) or Markdown (.md) documents."
        )
        desc_label.setStyleSheet("color: #4A5568; font-size: 12px;")
        header_layout.addWidget(desc_label)
        main_layout.addWidget(header_card)

        # 2. Progress / Status Strip
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setFixedHeight(14)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 1px solid #E2E8F0;
                border-radius: 4px;
                background-color: #EDF2F7;
            }
            QProgressBar::chunk {
                background-color: #0066CC;
                border-radius: 3px;
            }
        """)

        self.status_label = QLabel("Preparing specification documents...")
        self.status_label.setStyleSheet("color: #4A5568; font-size: 11px; font-style: italic;")

        progress_box = QVBoxLayout()
        progress_box.setSpacing(4)
        progress_box.addWidget(self.status_label)
        progress_box.addWidget(self.progress_bar)
        main_layout.addLayout(progress_box)

        # 3. Filter & Tree Controls
        filter_layout = QHBoxLayout()
        filter_layout.setSpacing(6)

        self.filter_input = QLineEdit()
        self.filter_input.setPlaceholderText("🔍 Filter clauses by number or title (e.g. 5.2, registration, NWDAF)...")
        self.filter_input.setStyleSheet(SEARCH_INPUT_STYLE)
        self.filter_input.setClearButtonEnabled(True)
        self.filter_input.textChanged.connect(self._filter_tree)
        filter_layout.addWidget(self.filter_input, 1)

        self.lbl_filter_matches = QLabel()
        self.lbl_filter_matches.setStyleSheet("font-size: 11px; color: #64748B; padding-right: 4px;")
        filter_layout.addWidget(self.lbl_filter_matches)

        self.btn_expand_all = QPushButton("Expand All")
        self.btn_expand_all.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_expand_all.setCursor(Qt.PointingHandCursor)
        self.btn_expand_all.clicked.connect(lambda: self.clause_tree.expandAll())
        filter_layout.addWidget(self.btn_expand_all)

        self.btn_collapse_all = QPushButton("Collapse All")
        self.btn_collapse_all.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_collapse_all.setCursor(Qt.PointingHandCursor)
        self.btn_collapse_all.clicked.connect(lambda: self.clause_tree.collapseAll())
        filter_layout.addWidget(self.btn_collapse_all)

        main_layout.addLayout(filter_layout)

        # 4. Hierarchical Clause Tree
        self.clause_tree = QTreeWidget()
        self.clause_tree.setHeaderLabels(["Clause / Heading", "Source Document"])
        header = self.clause_tree.header()
        header.setStretchLastSection(False)
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.clause_tree.setStyleSheet("""
            QTreeWidget {
                background-color: white;
                border: 1px solid #CBD5E1;
                border-radius: 6px;
                padding: 4px;
                font-size: 12px;
            }
            QTreeWidget::item {
                padding: 3px 0;
            }
            QTreeWidget::item:hover {
                background-color: #F1F5F9;
            }
            QTreeWidget::item:selected {
                background-color: #E2E8F0;
                color: #0F172A;
            }
        """)
        self.clause_tree.itemChanged.connect(self._on_tree_item_changed)
        main_layout.addWidget(self.clause_tree, 1)

        # 5. Bottom Footer (2-Tier Options & Actions)
        footer_layout = QVBoxLayout()
        footer_layout.setSpacing(10)
        footer_layout.setContentsMargins(0, 4, 0, 0)

        # Row 1: Selection Status & Export Options
        options_layout = QHBoxLayout()
        options_layout.setSpacing(10)

        self.btn_clear = QPushButton("Clear Selection")
        self.btn_clear.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_clear.setCursor(Qt.PointingHandCursor)
        self.btn_clear.clicked.connect(self._clear_selection)
        self.btn_clear.setEnabled(False)
        options_layout.addWidget(self.btn_clear)

        self.lbl_selection_count = QLabel("No clauses selected")
        self.lbl_selection_count.setStyleSheet("color: #64748B; font-size: 12px; font-weight: 500;")
        options_layout.addWidget(self.lbl_selection_count)

        options_layout.addStretch()

        self.open_after_cb = QCheckBox("Open when finished")
        self.open_after_cb.setChecked(True)
        self.open_after_cb.setStyleSheet("color: #334155; font-size: 12px;")
        self.open_after_cb.setToolTip("Open the extracted document in Microsoft Word or default viewer when extraction completes.")
        options_layout.addWidget(self.open_after_cb)

        self.chk_also_markdown = QCheckBox("Also export as Markdown (with figures)")
        self.chk_also_markdown.setStyleSheet("color: #334155; font-size: 12px;")
        self.chk_also_markdown.setToolTip(
            "Automatically convert the extracted clause(s) to Markdown (.md) and extract all embedded figures and diagrams into a companion media folder."
        )
        options_layout.addWidget(self.chk_also_markdown)

        footer_layout.addLayout(options_layout)

        # Row 2: Primary & Secondary Action Buttons
        action_layout = QHBoxLayout()
        action_layout.setSpacing(8)
        action_layout.addStretch()

        self.btn_cr_pcr = QPushButton("📝 Draft CR / pCR...")
        self.btn_cr_pcr.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_cr_pcr.setCursor(Qt.PointingHandCursor)
        self.btn_cr_pcr.setEnabled(False)
        self.btn_cr_pcr.clicked.connect(self._open_cr_pcr_from_selection)
        action_layout.addWidget(self.btn_cr_pcr)

        self.btn_extract = QPushButton("✂️ Extract Selected Clauses...")
        self.btn_extract.setCursor(Qt.PointingHandCursor)
        self.btn_extract.setEnabled(False)
        self.btn_extract.setStyleSheet("""
            QPushButton {
                font-weight: bold;
                background-color: #0066CC;
                color: white;
                padding: 7px 20px;
                border-radius: 4px;
                border: 1px solid #0055AA;
            }
            QPushButton:hover {
                background-color: #0052A3;
            }
            QPushButton:disabled {
                background-color: #CBD5E1;
                color: #94A3B8;
                border: 1px solid #CBD5E1;
            }
        """)
        self.btn_extract.clicked.connect(self._extract_selected_clauses)
        action_layout.addWidget(self.btn_extract)

        self.btn_close = QPushButton("Close")
        self.btn_close.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_close.setCursor(Qt.PointingHandCursor)
        self.btn_close.clicked.connect(self.close)
        action_layout.addWidget(self.btn_close)

        footer_layout.addLayout(action_layout)
        main_layout.addLayout(footer_layout)

    def _start_preparation(self):
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.status_label.setText("Starting specification preparation...")

        self._prep_thread = SpecExcerptPrepThread(
            spec_num=self.spec_num,
            version=self.version,
            fname=self.fname,
            url=self.url,
            spec_dl_dir=self.spec_dl_dir,
        )
        self._prep_thread.progress.connect(self._on_prep_progress)
        self._prep_thread.prepared.connect(self._on_prep_finished)
        self._prep_thread.error.connect(self._on_prep_error)
        self._prep_thread.start()

    def _on_prep_progress(self, pct: int, msg: str):
        self.progress_bar.setValue(pct)
        self.status_label.setText(msg)

    def _on_prep_finished(self, ordered_paths: list, headings: list):
        self._ordered_paths = list(ordered_paths)
        self._headings = list(headings)

        parts_count = len(self._ordered_paths)
        parts_text = f"{parts_count} part" if parts_count == 1 else f"{parts_count} parts"
        self.source_summary_badge.setText(f"📄 {parts_text} ({len(self._headings)} clauses)")
        self.source_summary_badge.setStyleSheet(BADGE_STYLE_PRIMARY)

        self.progress_bar.setValue(100)
        self.status_label.setText("Ready: Select clauses and click Extract.")
        self._populate_tree()

    def _on_prep_error(self, err_msg: str):
        self.progress_bar.setVisible(False)
        self.status_label.setText("⚠️ Preparation failed.")
        self.status_label.setStyleSheet("color: #DC2626; font-weight: bold;")
        self.source_summary_badge.setText("Failed")
        self.source_summary_badge.setStyleSheet("color: #DC2626;")
        QMessageBox.critical(self, "Preparation Failed", f"Could not prepare specification for extraction:\n{err_msg}")

    def _populate_tree(self):
        self._is_updating_checks = True
        self.clause_tree.clear()

        by_source = {}
        for h in self._headings:
            by_source.setdefault(h.source_path, []).append(h)

        multi_part = len(self._ordered_paths) > 1

        for source_path in self._ordered_paths:
            source_name = Path(source_path).name
            if multi_part:
                source_item = QTreeWidgetItem([f"📄 {source_name}", source_name])
                source_item.setFlags(source_item.flags() & ~Qt.ItemIsUserCheckable)
                font = source_item.font(0)
                font.setBold(True)
                source_item.setFont(0, font)
                self.clause_tree.addTopLevelItem(source_item)
                parent_root = source_item
            else:
                parent_root = None

            parents = {}
            for h in by_source.get(source_path, []):
                item = QTreeWidgetItem([h.text, source_name])
                item.setData(0, Qt.UserRole, h.id)
                item.setCheckState(0, Qt.Unchecked)

                parent = None
                for level in range(h.level - 1, 0, -1):
                    if level in parents:
                        parent = parents[level]
                        break

                if parent is not None:
                    parent.addChild(item)
                elif parent_root is not None:
                    parent_root.addChild(item)
                else:
                    self.clause_tree.addTopLevelItem(item)

                parents[h.level] = item
                for level in list(parents):
                    if level > h.level:
                        del parents[level]

        self.clause_tree.expandToDepth(1 if multi_part else 0)
        self._is_updating_checks = False
        self._update_extract_button()

    def _on_tree_item_changed(self, item: QTreeWidgetItem, column: int):
        if self._is_updating_checks or column != 0:
            return

        self._is_updating_checks = True
        try:
            state = item.checkState(0)

            # Cascade down to all child items
            def set_child_states(parent_item):
                for i in range(parent_item.childCount()):
                    child = parent_item.child(i)
                    if child.data(0, Qt.UserRole):
                        child.setCheckState(0, state)
                    set_child_states(child)

            set_child_states(item)
        finally:
            self._is_updating_checks = False

        self._update_extract_button()

    def _filter_tree(self, text: str):
        query = text.strip().lower()
        matches = 0

        def visit(item: QTreeWidgetItem) -> bool:
            nonlocal matches
            own_match = not query or query in item.text(0).lower()
            if query and own_match and item.data(0, Qt.UserRole):
                matches += 1
            child_match = False
            for i in range(item.childCount()):
                child_match = visit(item.child(i)) or child_match

            is_visible = own_match or child_match
            item.setHidden(not is_visible)
            if query and child_match:
                item.setExpanded(True)
            return is_visible

        for i in range(self.clause_tree.topLevelItemCount()):
            visit(self.clause_tree.topLevelItem(i))

        if hasattr(self, "lbl_filter_matches"):
            if query:
                self.lbl_filter_matches.setText(f"{matches} match(es)")
            else:
                self.lbl_filter_matches.setText("")

    def _clear_selection(self):
        self._is_updating_checks = True
        try:
            def clear_node(item: QTreeWidgetItem):
                if item.data(0, Qt.UserRole):
                    item.setCheckState(0, Qt.Unchecked)
                for i in range(item.childCount()):
                    clear_node(item.child(i))

            for i in range(self.clause_tree.topLevelItemCount()):
                clear_node(self.clause_tree.topLevelItem(i))
        finally:
            self._is_updating_checks = False

        self._update_extract_button()

    def _get_selected_ids(self) -> List[str]:
        selected = []

        def collect(item: QTreeWidgetItem):
            h_id = item.data(0, Qt.UserRole)
            if h_id and item.checkState(0) == Qt.Checked:
                selected.append(h_id)
            for i in range(item.childCount()):
                collect(item.child(i))

        for i in range(self.clause_tree.topLevelItemCount()):
            collect(self.clause_tree.topLevelItem(i))

        return selected

    def _update_extract_button(self):
        selected_ids = self._get_selected_ids()
        count = len(selected_ids)
        self.btn_extract.setEnabled(count > 0)
        self.btn_cr_pcr.setEnabled(count > 0)
        self.btn_clear.setEnabled(count > 0)
        self.btn_cr_pcr.setText(f"📝 Draft CR / pCR ({count})..." if count > 0 else "📝 Draft CR / pCR...")
        if count > 0:
            self.btn_extract.setText(f"✂️ Extract Selected Clauses ({count})...")
            if hasattr(self, "lbl_selection_count"):
                self.lbl_selection_count.setText(f"<b>{count}</b> clause(s) selected")
                self.lbl_selection_count.setStyleSheet("color: #0284C7; font-size: 12px; font-weight: 500;")
        else:
            self.btn_extract.setText("✂️ Extract Selected Clauses...")
            if hasattr(self, "lbl_selection_count"):
                self.lbl_selection_count.setText("No clauses selected")
                self.lbl_selection_count.setStyleSheet("color: #64748B; font-size: 12px; font-weight: 500;")

    def _extract_selected_clauses(self):
        selected_ids = self._get_selected_ids()
        if not selected_ids:
            QMessageBox.information(self, "No Clauses Selected", "Please select at least one clause to extract.")
            return

        clean_spec = self.spec_num.replace(".", "_")
        default_out_name = f"{self.spec_type}_{clean_spec}_v{self.version}_excerpt.docx"
        default_out_path = str(self.spec_dl_dir / default_out_name)

        output_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Extracted Clauses",
            default_out_path,
            "Word Document (*.docx)",
        )

        if not output_path:
            return

        if not output_path.lower().endswith(".docx"):
            output_path += ".docx"

        self.btn_extract.setEnabled(False)
        self.btn_extract.setText("⏳ Extracting...")
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(True)
        self.status_label.setText("Preparing clause extraction...")

        self._extract_thread = WordExcerptThread(
            source_paths=self._ordered_paths,
            headings=self._headings,
            selected_ids=selected_ids,
            output_path=output_path,
            export_markdown=self.chk_also_markdown.isChecked(),
        )
        self._extract_thread.progress.connect(self._on_extract_progress)
        self._extract_thread.finished_path.connect(self._on_extract_finished)
        self._extract_thread.failed.connect(self._on_extract_failed)
        self._extract_thread.finished.connect(self._on_extract_done)
        self._extract_thread.start()

    def _on_extract_progress(self, percent: int, message: str):
        if message.startswith("Preparing excerpt from "):
            self.progress_bar.setRange(0, 0)
        else:
            if self.progress_bar.minimum() == 0 and self.progress_bar.maximum() == 0:
                self.progress_bar.setRange(0, 100)
            self.progress_bar.setValue(percent)
        self.status_label.setText(message)

    def _on_extract_finished(self, out_path_str: str):
        p = Path(out_path_str)
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(100)
        self.status_label.setText(f"✅ Excerpt saved: {p.name}")

        self.extraction_completed.emit(p)

        if self.open_after_cb.isChecked() and p.exists():
            open_document(p)

        thread = self._extract_thread
        if thread and getattr(thread, "export_markdown", False) and getattr(thread, "md_path", None):
            md_path = thread.md_path
            img_count = getattr(thread, "images_count", 0)
            img_msg = f"\n• Figures extracted: {img_count}" if img_count else "\n• Figures extracted: 0"
            msg = (
                f"Successfully extracted selected clauses!\n\n"
                f"📄 Word Document:\n{p}\n\n"
                f"📝 Markdown Document:\n{md_path}"
                f"{img_msg}"
            )
            QMessageBox.information(self, "Extraction Complete", msg)
        else:
            QMessageBox.information(
                self,
                "Extraction Complete",
                f"Successfully extracted selected clauses to:\n{p}",
            )

    def _on_extract_failed(self, err_msg: str):
        self.progress_bar.setRange(0, 100)
        self.status_label.setText("⚠️ Extraction failed.")
        QMessageBox.critical(self, "Extraction Error", f"Extraction failed:\n{err_msg}")

    def _on_extract_done(self):
        self._update_extract_button()

    def _open_cr_pcr_from_selection(self):
        from modules.word_tools.ui.cr_pcr_dialog import CrPcrDialog
        selected_ids = self._get_selected_ids()
        dlg = CrPcrDialog(
            spec_num=self.spec_num,
            spec_type=self.spec_type,
            version=self.version,
            fname=self.fname,
            url=self.url,
            spec_dl_dir=self.spec_dl_dir,
            ordered_paths=self._ordered_paths,
            headings=self._headings,
            parent=self.parent(),
        )
        dlg._populate_tree()
        dlg._is_updating_checks = True
        try:
            for i in range(dlg.clause_tree.topLevelItemCount()):
                item = dlg.clause_tree.topLevelItem(i)
                def restore_checks(it):
                    h_id = it.data(0, Qt.UserRole)
                    if h_id and h_id in selected_ids:
                        it.setCheckState(0, Qt.Checked)
                    for ci in range(it.childCount()):
                        restore_checks(it.child(ci))
                restore_checks(item)
        finally:
            dlg._is_updating_checks = False
        dlg._update_change_block_summary()
        dlg.show()

    def closeEvent(self, event):
        """Cleanly abort active worker threads if dialog is closed."""
        if self._prep_thread and self._prep_thread.isRunning():
            self._prep_thread.requestInterruption()
            self._prep_thread.wait(200)
        if self._extract_thread and self._extract_thread.isRunning():
            self._extract_thread.requestInterruption()
            self._extract_thread.wait(500)
        super().closeEvent(event)


SpecClauseDialog = SpecClauseExcerptDialog
