"""Card 4: Document Sanitizer & Revision Hygiene Hub."""

import logging
from pathlib import Path
import re
from typing import Optional

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from core.ui.ui_components import (
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    CARD_FRAME_STYLE,
    TABLE_STYLE_CLEAN,
)
from modules.word_tools.core.word_config import WordConfig
from modules.word_tools.core.document_sanitizer import (
    LANGUAGE_PRESETS,
    SanitizeConfig,
    close_word_document_for_update,
    inspect_document,
    is_file_open_in_word,
    sanitize_document,
)
from modules.word_tools.ui.document_selector_pane import DocumentSelectorPane

logger = logging.getLogger(__name__)


class WordSanitizerCard(QWidget):
    """Document Sanitizer, Proofing Language Unification, Language Revision Purge, and Author Consolidation."""

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self._last_author_file: str = ""
        self._last_modified_author_file: str = ""
        self._current_author_stats: list = []
        self._setup_ui()

    def _setup_ui(self):
        outer_layout = QVBoxLayout(self)
        outer_layout.setContentsMargins(0, 0, 0, 0)
        outer_layout.setSpacing(0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setStyleSheet("background: transparent;")

        container = QWidget()
        card_layout = QVBoxLayout(container)
        card_layout.setContentsMargins(10, 10, 10, 10)
        card_layout.setSpacing(10)

        # 1. Header & Description Frame
        hdr_frame = QFrame()
        hdr_frame.setObjectName("cardFrame")
        hdr_frame.setStyleSheet(CARD_FRAME_STYLE)
        hf_layout = QVBoxLayout(hdr_frame)
        hf_layout.setContentsMargins(14, 12, 14, 12)
        hf_layout.setSpacing(6)

        hf_title = QLabel("<b>🛠️ Document Sanitizer &amp; Revision Hygiene Hub</b>")
        hf_title.setStyleSheet("font-size: 13px; color: #1E293B;")
        hf_layout.addWidget(hf_title)

        hf_desc = QLabel(
            "Comprehensive pre-flight sanitation for 3GPP Word contributions: unify proofing language across all "
            "stories and styles, surgically purge Track Changes language marks ('Formatted: Language...'), "
            "consolidate reviewer identities under a single corporate signature, and clean private metadata."
        )
        hf_desc.setWordWrap(True)
        hf_desc.setStyleSheet("color: #475569; font-size: 11px;")
        hf_layout.addWidget(hf_desc)
        card_layout.addWidget(hdr_frame)

        # 2. Document Selector Pane
        self.pane_sanitizer = DocumentSelectorPane("📄 WORD CONTRIBUTION (.DOCX / .DOCM)", accepted_extensions=[".docx", ".docm"])
        self.pane_sanitizer.file_selected.connect(self._on_sanitizer_file_selected)
        card_layout.addWidget(self.pane_sanitizer)

        # 3. Executive Audit Summary Frame
        self.audit_frame = QFrame()
        self.audit_frame.setObjectName("cardFrame")
        self.audit_frame.setStyleSheet(CARD_FRAME_STYLE)
        audit_layout = QVBoxLayout(self.audit_frame)
        audit_layout.setContentsMargins(14, 10, 14, 10)
        audit_layout.setSpacing(8)

        self.lbl_audit_summary = QLabel("No document selected. Drop or select a .docx file above to audit.")
        self.lbl_audit_summary.setStyleSheet("color: #64748B; font-size: 11px; font-style: italic;")
        audit_layout.addWidget(self.lbl_audit_summary)

        # KPI Badges Row
        kpi_row = QHBoxLayout()
        kpi_row.setSpacing(6)

        self.lbl_kpi_lang = QLabel("🌐 Languages: —")
        self.lbl_kpi_lang_rev = QLabel("🏷️ Lang Changes: 0")
        self.lbl_kpi_authors = QLabel("👥 Authors: 0")
        self.lbl_kpi_text_rev = QLabel("📝 Ins: 0 · Del: 0")
        self.lbl_kpi_fmt = QLabel("🎨 Formatting: 0")
        self.lbl_kpi_comments = QLabel("💬 Comments: 0")

        pill_style = "background-color: #F1F5F9; color: #334155; border: 1px solid #CBD5E1; border-radius: 4px; padding: 3px 8px; font-size: 11px; font-weight: 500;"
        for lbl in (self.lbl_kpi_lang, self.lbl_kpi_lang_rev, self.lbl_kpi_authors, self.lbl_kpi_text_rev, self.lbl_kpi_fmt, self.lbl_kpi_comments):
            lbl.setStyleSheet(pill_style)
            kpi_row.addWidget(lbl)
        kpi_row.addStretch()
        audit_layout.addLayout(kpi_row)

        # 1-Click Configuration Presets
        presets_row = QHBoxLayout()
        presets_row.setSpacing(6)
        lbl_presets = QLabel("<b>1-Click Presets:</b>")
        lbl_presets.setStyleSheet("font-size: 11px; color: #475569;")
        presets_row.addWidget(lbl_presets)

        btn_preset_3gpp = QPushButton("⚡ 3GPP Contribution Prep (en-GB + DT + Purge Lang)")
        btn_preset_3gpp.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        btn_preset_3gpp.setToolTip("Sets en-GB, purges language clutter balloons, and unifies authors to corporate signature.")
        btn_preset_3gpp.clicked.connect(self._apply_preset_3gpp)

        btn_preset_lang_only = QPushButton("🌐 Fix Language Only (en-GB + Purge Lang)")
        btn_preset_lang_only.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        btn_preset_lang_only.setToolTip("Unifies language to en-GB and purges language balloons without touching author names.")
        btn_preset_lang_only.clicked.connect(self._apply_preset_lang_only)

        btn_preset_anon = QPushButton("🎭 Full Anonymize (en-GB + Contributor + Clear Metadata)")
        btn_preset_anon.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        btn_preset_anon.setToolTip("Replaces authors with 'Contributor', clears comments and private metadata.")
        btn_preset_anon.clicked.connect(self._apply_preset_anonymize)

        presets_row.addWidget(btn_preset_3gpp)
        presets_row.addWidget(btn_preset_lang_only)
        presets_row.addWidget(btn_preset_anon)
        presets_row.addStretch()
        audit_layout.addLayout(presets_row)

        card_layout.addWidget(self.audit_frame)

        # 4. Operations Configuration Grid
        ops_grid = QGridLayout()
        ops_grid.setSpacing(10)

        # Section A: Language Unification & Purge
        sec_lang_frame = QFrame()
        sec_lang_frame.setObjectName("cardFrame")
        sec_lang_frame.setStyleSheet(CARD_FRAME_STYLE)
        sl_layout = QVBoxLayout(sec_lang_frame)
        sl_layout.setContentsMargins(12, 10, 12, 10)
        sl_layout.setSpacing(6)

        self.chk_sec_lang = QCheckBox("<b>🌐 Proofing Language Unification &amp; Purge</b>")
        self.chk_sec_lang.setChecked(True)
        sl_layout.addWidget(self.chk_sec_lang)

        lang_target_row = QHBoxLayout()
        lang_target_row.addWidget(QLabel("Target Language:"))
        self.combo_target_lang = QComboBox()
        for code, label in LANGUAGE_PRESETS:
            self.combo_target_lang.addItem(label, code)
        self.combo_target_lang.addItem("Other / Custom BCP-47 Code...", "custom")
        self.combo_target_lang.currentIndexChanged.connect(self._on_target_lang_combo_changed)
        lang_target_row.addWidget(self.combo_target_lang)

        self.edit_custom_lang = QLineEdit()
        self.edit_custom_lang.setPlaceholderText("e.g. en-US, de-DE")
        self.edit_custom_lang.setMaximumWidth(100)
        self.edit_custom_lang.setVisible(False)
        lang_target_row.addWidget(self.edit_custom_lang)
        lang_target_row.addStretch()
        sl_layout.addLayout(lang_target_row)

        self.chk_lang_all_stories = QCheckBox("Apply to all stories (body, headers, footers, footnotes, comments)")
        self.chk_lang_all_stories.setChecked(True)
        sl_layout.addWidget(self.chk_lang_all_stories)

        self.chk_lang_styles = QCheckBox("Normalize styles.xml && docDefaults (prevents newly typed text from reverting)")
        self.chk_lang_styles.setChecked(True)
        sl_layout.addWidget(self.chk_lang_styles)

        self.chk_lang_purge_revisions = QCheckBox("<b>Surgically purge language formatting revisions</b> (<w:rPrChange> with <w:lang>)")
        self.chk_lang_purge_revisions.setChecked(True)
        self.chk_lang_purge_revisions.setToolTip(
            "Eliminates trivial 'Formatted: Language...' Track Changes balloons while cleanly preserving "
            "all substantive text insertions and deletions intact."
        )
        sl_layout.addWidget(self.chk_lang_purge_revisions)

        self.chk_clear_no_proof = QCheckBox("Clear <w:noProof> spell-check suppression marks")
        self.chk_clear_no_proof.setChecked(False)
        sl_layout.addWidget(self.chk_clear_no_proof)

        ops_grid.addWidget(sec_lang_frame, 0, 0)

        # Section B: Revision Author Consolidation
        sec_author_frame = QFrame()
        sec_author_frame.setObjectName("cardFrame")
        sec_author_frame.setStyleSheet(CARD_FRAME_STYLE)
        sa_layout = QVBoxLayout(sec_author_frame)
        sa_layout.setContentsMargins(12, 10, 12, 10)
        sa_layout.setSpacing(6)

        self.chk_sec_authors = QCheckBox("<b>👥 Revision Author Consolidation</b>")
        self.chk_sec_authors.setChecked(True)
        sa_layout.addWidget(self.chk_sec_authors)

        author_target_row = QHBoxLayout()
        author_target_row.addWidget(QLabel("Target Author Signature:"))
        self.edit_target_author = QLineEdit()
        default_company = WordConfig.load().get("default_source_company", "Deutsche Telekom")
        self.edit_target_author.setText(default_company)
        self.edit_target_author.textChanged.connect(self._on_target_author_changed)
        author_target_row.addWidget(self.edit_target_author)

        author_target_row.addWidget(QLabel("Initials:"))
        self.edit_target_initials = QLineEdit()
        self.edit_target_initials.setMaximumWidth(60)
        self.edit_target_initials.setText(self._derive_initials(default_company))
        author_target_row.addWidget(self.edit_target_initials)
        sa_layout.addLayout(author_target_row)

        # Author Presets Quick Bar
        ap_bar = QHBoxLayout()
        ap_bar.setSpacing(4)
        ap_bar.addWidget(QLabel("<span style='color:#64748B; font-size:10px;'>Quick Signatures:</span>"))
        for sig in ["Deutsche Telekom", "Editor", "Rapporteur", "Contributor", "Anonymous"]:
            btn_sig = QPushButton(sig)
            btn_sig.setStyleSheet("font-size: 10px; padding: 1px 6px; background-color: #F1F5F9; border: 1px solid #CBD5E1; border-radius: 3px;")
            btn_sig.clicked.connect(lambda checked, s=sig: self._apply_author_preset(s))
            ap_bar.addWidget(btn_sig)
        ap_bar.addStretch()
        sa_layout.addLayout(ap_bar)

        self.chk_author_initials = QCheckBox("Update reviewer initials in comments and callouts")
        self.chk_author_initials.setChecked(True)
        sa_layout.addWidget(self.chk_author_initials)

        self.chk_author_metadata = QCheckBox("Update Dublin Core metadata (dc:creator, cp:lastModifiedBy)")
        self.chk_author_metadata.setChecked(True)
        sa_layout.addWidget(self.chk_author_metadata)

        ops_grid.addWidget(sec_author_frame, 0, 1)
        card_layout.addLayout(ops_grid)

        # 5. Detected Authors Table
        table_hdr = QHBoxLayout()
        table_hdr.addWidget(QLabel("<b>Detected Reviewers in Selected Document:</b>"))
        table_hdr.addStretch()

        btn_sel_all = QPushButton("Select All")
        btn_sel_all.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        btn_sel_all.clicked.connect(lambda: self._select_all_authors(True))
        btn_desel_all = QPushButton("Deselect All")
        btn_desel_all.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        btn_desel_all.clicked.connect(lambda: self._select_all_authors(False))
        table_hdr.addWidget(btn_sel_all)
        table_hdr.addWidget(btn_desel_all)
        card_layout.addLayout(table_hdr)

        self.author_table = QTableWidget(0, 7)
        self.author_table.setHorizontalHeaderLabels(["", "Author Name", "Ins", "Del", "Fmt", "Comments", "Initials"])
        self.author_table.setStyleSheet(TABLE_STYLE_CLEAN)
        self.author_table.setMaximumHeight(150)
        self.author_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.author_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        for col in (2, 3, 4, 5, 6):
            self.author_table.horizontalHeader().setSectionResizeMode(col, QHeaderView.ResizeToContents)
        card_layout.addWidget(self.author_table)

        # 6. Advanced Document Hygiene
        sec_cleanup_frame = QFrame()
        sec_cleanup_frame.setObjectName("cardFrame")
        sec_cleanup_frame.setStyleSheet(CARD_FRAME_STYLE)
        sc_layout = QHBoxLayout(sec_cleanup_frame)
        sc_layout.setContentsMargins(12, 8, 12, 8)
        sc_layout.setSpacing(12)

        self.chk_sec_cleanup = QCheckBox("<b>🧹 Advanced Cleanup:</b>")
        self.chk_sec_cleanup.setChecked(False)
        sc_layout.addWidget(self.chk_sec_cleanup)

        self.chk_purge_all_fmt = QCheckBox("Purge ALL Track Changes formatting revisions (keep substantive text changes)")
        self.chk_remove_comments = QCheckBox("Remove all reviewer comments")
        self.chk_sanitizer_backup = QCheckBox("Create safety backup (.bak) before updating in-place")
        self.chk_sanitizer_backup.setChecked(True)

        sc_layout.addWidget(self.chk_purge_all_fmt)
        sc_layout.addWidget(self.chk_remove_comments)
        sc_layout.addStretch()
        sc_layout.addWidget(self.chk_sanitizer_backup)
        card_layout.addWidget(sec_cleanup_frame)

        # 7. Action Execution Bar
        act_box = QHBoxLayout()
        act_box.setSpacing(8)

        self.btn_run_in_place = QPushButton("⚡ Sanitize && Update In-Place")
        self.btn_run_in_place.setObjectName("primaryBtn")
        self.btn_run_in_place.setToolTip("Applies selected sanitizations directly to the source document (creates .bak backup).")
        self.btn_run_in_place.clicked.connect(lambda: self._trigger_sanitizer(in_place=True))

        self.btn_run_save_as = QPushButton("💾 Save As New Sanitized Document...")
        self.btn_run_save_as.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_run_save_as.clicked.connect(lambda: self._trigger_sanitizer(in_place=False))

        act_box.addWidget(self.btn_run_in_place)
        act_box.addWidget(self.btn_run_save_as)
        act_box.addStretch()
        card_layout.addLayout(act_box)

        # 8. Result Indicator Frame
        self.author_result_frame = QFrame()
        self.author_result_frame.setObjectName("cardFrame")
        self.author_result_frame.setStyleSheet(CARD_FRAME_STYLE + """
            QFrame#cardFrame {
                background-color: #ECFDF5;
                border: 1px solid #6EE7B7;
                border-radius: 6px;
                padding: 6px 12px;
            }
        """)
        self.author_result_frame.setVisible(False)
        rf_layout = QHBoxLayout(self.author_result_frame)
        rf_layout.setContentsMargins(4, 4, 4, 4)

        self.author_result_lbl = QLabel()
        self.author_result_lbl.setStyleSheet("color: #065F46; font-size: 11px; font-weight: bold;")
        rf_layout.addWidget(self.author_result_lbl, 1)

        self.btn_open_author_file = QPushButton("👁️ Open Document")
        self.btn_open_author_file.setStyleSheet("font-size: 11px; padding: 2px 8px;")
        self.btn_open_author_file.clicked.connect(self._open_author_result_file)
        rf_layout.addWidget(self.btn_open_author_file)

        self.btn_open_author_folder = QPushButton("📂 Open Folder")
        self.btn_open_author_folder.setStyleSheet("font-size: 11px; padding: 2px 8px;")
        self.btn_open_author_folder.clicked.connect(self._open_author_result_folder)
        rf_layout.addWidget(self.btn_open_author_folder)

        card_layout.addWidget(self.author_result_frame)
        card_layout.addStretch()

        scroll.setWidget(container)
        outer_layout.addWidget(scroll)

    def _on_target_lang_combo_changed(self, index: int):
        data = self.combo_target_lang.currentData()
        self.edit_custom_lang.setVisible(data == "custom")

    def _apply_preset_3gpp(self):
        self.chk_sec_lang.setChecked(True)
        idx = self.combo_target_lang.findData("en-GB")
        if idx >= 0:
            self.combo_target_lang.setCurrentIndex(idx)
        self.chk_lang_all_stories.setChecked(True)
        self.chk_lang_styles.setChecked(True)
        self.chk_lang_purge_revisions.setChecked(True)
        self.chk_sec_authors.setChecked(True)
        self._apply_author_preset(WordConfig.load().get("default_source_company", "Deutsche Telekom"))
        self.chk_sec_cleanup.setChecked(False)

    def _apply_preset_lang_only(self):
        self.chk_sec_lang.setChecked(True)
        idx = self.combo_target_lang.findData("en-GB")
        if idx >= 0:
            self.combo_target_lang.setCurrentIndex(idx)
        self.chk_lang_all_stories.setChecked(True)
        self.chk_lang_styles.setChecked(True)
        self.chk_lang_purge_revisions.setChecked(True)
        self.chk_sec_authors.setChecked(False)
        self.chk_sec_cleanup.setChecked(False)

    def _apply_preset_anonymize(self):
        self.chk_sec_lang.setChecked(True)
        idx = self.combo_target_lang.findData("en-GB")
        if idx >= 0:
            self.combo_target_lang.setCurrentIndex(idx)
        self.chk_sec_authors.setChecked(True)
        self._apply_author_preset("Contributor", "")
        self.chk_sec_cleanup.setChecked(True)
        self.chk_remove_comments.setChecked(True)

    def _on_sanitizer_file_selected(self, files: list):
        if not files:
            return
        self._inspect_selected_author_document(files[0])

    def _inspect_selected_author_document(self, file_path: str):
        path = Path(file_path)
        if not path.is_file():
            return
        if path.suffix.lower() not in (".docx", ".docm"):
            self.lbl_audit_summary.setText("⚠️ Please select a modern Word document (.docx or .docm). Legacy .doc is not supported.")
            return

        self._last_author_file = str(path.resolve())
        self.author_result_frame.setVisible(False)

        try:
            audit = inspect_document(path)
            self._current_author_stats = audit.authors
            self._populate_author_table(audit.authors)

            # Update KPI Badges
            self.lbl_kpi_lang.setText(f"🌐 Languages: {audit.language_stats.summary_text}")
            self.lbl_kpi_lang_rev.setText(f"🏷️ Lang Changes: {audit.language_stats.language_change_marks}")
            self.lbl_kpi_authors.setText(f"👥 Authors: {audit.total_authors}")
            self.lbl_kpi_text_rev.setText(f"📝 Ins: {audit.total_insertions} · Del: {audit.total_deletions}")
            self.lbl_kpi_fmt.setText(f"🎨 Formatting: {audit.total_formatting}")
            self.lbl_kpi_comments.setText(f"💬 Comments: {audit.total_comments}")

            if audit.total_authors > 0:
                self.chk_sec_authors.setChecked(True)
            else:
                self.chk_sec_authors.setChecked(False)

            self.lbl_audit_summary.setText(
                f"✅ Analyzed {path.name} ({audit.file_size_kb:.1f} KB): "
                f"{audit.language_stats.summary_text} · {audit.language_stats.language_change_marks} lang revision marks · "
                f"{audit.total_authors} author(s) · {audit.total_changes} total changes."
            )

        except Exception as exc:
            logger.error("Failed to inspect document %s: %s", path.name, exc)
            self.lbl_audit_summary.setText(f"❌ Failed to inspect document: {exc}")

    def _populate_author_table(self, stats: list):
        self.author_table.setRowCount(len(stats))
        if not stats:
            return

        for row, s in enumerate(stats):
            chk_item = QTableWidgetItem()
            chk_item.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
            chk_item.setCheckState(Qt.Checked)
            self.author_table.setItem(row, 0, chk_item)

            name_item = QTableWidgetItem(s.name)
            name_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            self.author_table.setItem(row, 1, name_item)

            for col, val in enumerate([s.insertions, s.deletions, s.formatting, s.comments], start=2):
                num_item = QTableWidgetItem(str(val))
                num_item.setTextAlignment(Qt.AlignCenter)
                num_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                self.author_table.setItem(row, col, num_item)

            init_item = QTableWidgetItem(s.initials)
            init_item.setTextAlignment(Qt.AlignCenter)
            init_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            self.author_table.setItem(row, 6, init_item)

    def _select_all_authors(self, select: bool):
        for row in range(self.author_table.rowCount()):
            item = self.author_table.item(row, 0)
            if item:
                item.setCheckState(Qt.Checked if select else Qt.Unchecked)

    def _on_target_author_changed(self, text: str):
        suggested = self._derive_initials(text.strip())
        if suggested:
            self.edit_target_initials.setText(suggested)

    def _derive_initials(self, name: str) -> str:
        words = re.findall(r"[A-Za-z0-9]+", name or "")
        if not words:
            return ""
        if len(words) >= 2:
            return "".join(w[0].upper() for w in words[:3])
        return words[0][:3].title()

    def _apply_author_preset(self, author_name: str, initials: str = None):
        self.edit_target_author.setText(author_name)
        if initials is not None:
            self.edit_target_initials.setText(initials)
        else:
            self.edit_target_initials.setText(self._derive_initials(author_name))

    def _trigger_sanitizer(self, in_place: bool):
        source_inputs = self.pane_sanitizer.get_inputs()
        source_file = source_inputs[0] if source_inputs else self._last_author_file
        if not source_file or not Path(source_file).is_file():
            QMessageBox.warning(self, "No Document Selected", "Please select or drop a valid .docx file first.")
            return

        src_path = Path(source_file).resolve()

        sec_lang = self.chk_sec_lang.isChecked()
        sec_authors = self.chk_sec_authors.isChecked()
        sec_cleanup = self.chk_sec_cleanup.isChecked()

        if not (sec_lang or sec_authors or sec_cleanup):
            QMessageBox.warning(
                self,
                "No Operations Selected",
                "Please enable at least one section (Language, Authors, or Cleanup) to apply.",
            )
            return

        selected_authors = []
        if sec_authors:
            for row in range(self.author_table.rowCount()):
                chk = self.author_table.item(row, 0)
                name_item = self.author_table.item(row, 1)
                if chk and chk.checkState() == Qt.Checked and name_item:
                    selected_authors.append(name_item.text())

            if self.author_table.rowCount() > 0 and not selected_authors:
                QMessageBox.warning(
                    self,
                    "No Authors Selected",
                    "Author replacement is enabled, but no authors are checked in the table.",
                )
                return

        target_author = self.edit_target_author.text().strip()
        new_initials = self.edit_target_initials.text().strip() if self.chk_author_initials.isChecked() else None
        mapping = {old_name: target_author for old_name in selected_authors} if selected_authors else {}

        lang_data = self.combo_target_lang.currentData()
        if lang_data == "custom":
            target_lang = self.edit_custom_lang.text().strip() or "en-GB"
        else:
            target_lang = str(lang_data or "en-GB")

        reopen_after = False

        if in_place:
            dst_path = src_path
            if is_file_open_in_word(src_path):
                reply = QMessageBox.question(
                    self,
                    "Document Open in Word",
                    f"The document '{src_path.name}' is currently open in Microsoft Word.\n\n"
                    "To update it in-place, Word must save and close this document temporarily.\n\n"
                    "Would you like to save & close the document in Word now, apply sanitization, and re-open it?",
                    QMessageBox.Yes | QMessageBox.No,
                    QMessageBox.Yes,
                )
                if reply == QMessageBox.Yes:
                    closed = close_word_document_for_update(src_path, save_changes=True)
                    if not closed:
                        QMessageBox.warning(
                            self,
                            "File Lock",
                            "Could not automatically close the document in Word. Please save and close Word manually, or use 'Save As New Copy...'.",
                        )
                        return
                    reopen_after = True
                else:
                    return
        else:
            default_dst = src_path.parent / f"{src_path.stem}_sanitized{src_path.suffix}"
            chosen, _ = QFileDialog.getSaveFileName(
                self,
                "Save Sanitized Document As",
                str(default_dst),
                "Word Documents (*.docx *.docm)",
            )
            if not chosen:
                return
            dst_path = Path(chosen).resolve()

        config = SanitizeConfig(
            unify_language=sec_lang,
            target_language=target_lang,
            apply_to_all_stories=self.chk_lang_all_stories.isChecked(),
            update_styles_and_defaults=self.chk_lang_styles.isChecked(),
            purge_language_changes=self.chk_lang_purge_revisions.isChecked(),
            clear_no_proof=self.chk_clear_no_proof.isChecked(),
            update_authors=sec_authors and bool(mapping),
            author_mapping=mapping,
            new_initials=new_initials,
            update_metadata=self.chk_author_metadata.isChecked(),
            purge_all_formatting_changes=sec_cleanup and self.chk_purge_all_fmt.isChecked(),
            remove_all_comments=sec_cleanup and self.chk_remove_comments.isChecked(),
            backup=self.chk_sanitizer_backup.isChecked() if in_place else False,
        )

        success, count, msg = sanitize_document(
            docx_path=src_path,
            output_path=dst_path,
            config=config,
        )

        if not success:
            QMessageBox.critical(self, "Sanitization Failed", msg)
            return

        self._last_modified_author_file = str(dst_path)
        self.author_result_lbl.setText(f"✅ {msg}")
        self.author_result_frame.setVisible(True)

        if reopen_after:
            from modules.word_tools.core.document_opener import open_document
            try:
                open_document(dst_path)
            except Exception as exc:
                logger.warning("Could not re-open document: %s", exc)

        self._inspect_selected_author_document(str(dst_path))

    def _open_author_result_file(self):
        if self._last_modified_author_file and Path(self._last_modified_author_file).is_file():
            from modules.word_tools.core.document_opener import open_document
            try:
                open_document(self._last_modified_author_file)
            except Exception as exc:
                QMessageBox.warning(self, "Open Failed", f"Could not open document: {exc}")

    def _open_author_result_folder(self):
        if self._last_modified_author_file:
            folder = Path(self._last_modified_author_file).parent
            if folder.is_dir():
                import os
                if hasattr(os, "startfile"):
                    os.startfile(str(folder))
