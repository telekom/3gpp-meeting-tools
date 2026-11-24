"""2-tier ribbon, workbench cards, and backward-compatible widget alias setup mixin."""

from __future__ import annotations

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QComboBox,
    QFrame,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QSizePolicy,
    QStackedWidget,
    QVBoxLayout,
    QWidget,
)

from core.ui.flow_layout import FlowLayout
from core.ui.theme import C, T
from core.ui.ui_components import (
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    CARD_FRAME_STYLE,
    COMBOBOX_STYLE_TOOLBAR,
    InteractiveDropLabel,
    TOGGLE_PILL_STYLE,
)
from modules.word_tools.ui.cards import (
    WordCompareCard,
    WordConverterCard,
    WordCrPcrCard,
    WordExcerptCard,
    WordMarkdownCard,
    WordSanitizerCard,
    WordVisioCard,
)


class WordTabUiSetupMixin:
    """Constructs the WordExtractorTab 2-tier ribbon, card stack, and legacy aliases."""

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 8, 10, 8)
        layout.setSpacing(6)

        # ---------------------------------------------------------------------
        # 2-Tier Modern Ribbon
        # ---------------------------------------------------------------------
        self.ribbon_frame = QFrame()
        self.ribbon_frame.setObjectName("cardFrame")
        self.ribbon_frame.setStyleSheet(
            CARD_FRAME_STYLE
            + f"""
            QFrame#cardFrame {{
                background-color: {C.BG_SURFACE};
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 6px;
                padding: 4px 8px;
            }}
        """
        )
        ribbon_layout = QVBoxLayout(self.ribbon_frame)
        ribbon_layout.setContentsMargins(8, 4, 8, 4)
        ribbon_layout.setSpacing(6)

        # Row 1: System Monitors & Settings
        row1 = QHBoxLayout()
        row1.setSpacing(8)

        self.btn_active_word = QPushButton("🪟 Active Word: None Detected")
        self.btn_active_word.setStyleSheet(f"""
            QPushButton {{
                background-color: {C.BG_CANVAS};
                color: {C.TEXT_SECONDARY};
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 4px;
                padding: 3px 10px;
                font-weight: 500;
                font-size: 11px;
            }}
        """)
        self.btn_active_word.setToolTip(
            "Refreshes automatically. Click to open the active document folder."
        )
        self.btn_active_word.clicked.connect(self._on_active_word_clicked)
        row1.addWidget(self.btn_active_word)

        sep1 = QFrame()
        sep1.setFrameShape(QFrame.VLine)
        sep1.setStyleSheet(f"color: {C.BORDER_STRONG};")
        row1.addWidget(sep1)

        self.dotx_badge = QLabel("3GPP Macro Template (3gpp_70.dotx): Checking...")
        self.dotx_badge.setStyleSheet(f"font-size: {T.SIZE_TABLE}px; color: {C.TEXT_SECONDARY};")
        row1.addWidget(self.dotx_badge)

        self.dotx_btn = QPushButton("Install Template")
        self.dotx_btn.setStyleSheet("font-size: 11px; padding: 3px 8px;")
        self.dotx_btn.clicked.connect(self._on_toggle_dotx_clicked)
        row1.addWidget(self.dotx_btn)

        sep2 = QFrame()
        sep2.setFrameShape(QFrame.VLine)
        sep2.setStyleSheet(f"color: {C.BORDER_STRONG};")
        row1.addWidget(sep2)

        self.lo_badge = QLabel("LibreOffice: Checking...")
        self.lo_badge.setStyleSheet(f"font-size: {T.SIZE_TABLE}px; color: {C.TEXT_SECONDARY};")
        row1.addWidget(self.lo_badge)

        row1.addStretch()

        self.settings_btn = QPushButton("⚙️ Word Settings ▾")
        self.settings_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self._setup_settings_menu()
        row1.addWidget(self.settings_btn)

        ribbon_layout.addLayout(row1)

        hsep = QFrame()
        hsep.setFrameShape(QFrame.HLine)
        hsep.setStyleSheet(f"color: {C.BORDER_SUBTLE};")
        ribbon_layout.addWidget(hsep)

        # Row 2: Segmented Workbench Navigation
        nav_row_widget = QWidget(self.ribbon_frame)
        nav_row_widget.setObjectName("wordNavFlow")
        nav_row_widget.setStyleSheet(
            "QWidget#wordNavFlow { background: transparent; border: none; }"
        )
        row2 = FlowLayout(nav_row_widget, margin=0, h_spacing=6, v_spacing=6)

        lbl_nav = QLabel("<b>Workbench:</b>")
        lbl_nav.setStyleSheet(
            f"color: {C.TEXT_SECONDARY}; font-size: {T.SIZE_TABLE}px; padding-top: 4px;"
        )
        row2.addWidget(lbl_nav)

        self.nav_btn_compare = QPushButton("⚖️ Compare Documents")
        self.nav_btn_markdown = QPushButton("📝 Markdown ⇄ Word")
        self.nav_btn_markdown.setToolTip(
            "Bidirectional studio: Word to Markdown with media extraction, and Markdown to Word with embedded figures & diagrams."
        )
        self.nav_btn_cr_pcr = QPushButton("🏛️ 3GPP CR && Templates")
        self.nav_btn_author = QPushButton("🛠️ Sanitizer && Revisions")
        self.nav_btn_author.setToolTip(
            "Unify proofing language, purge language Track Changes clutter, consolidate revision authors, and sanitize metadata."
        )
        self.nav_btn_slice = QPushButton("✂️ Slice && Excerpt")
        self.nav_btn_convert = QPushButton("🔄 Format Converter")
        self.nav_btn_visio = QPushButton("🖼️ Visio && Media Extractor")

        self._nav_buttons = [
            (self.nav_btn_compare, 4),
            (self.nav_btn_markdown, 6),
            (self.nav_btn_cr_pcr, 7),
            (self.nav_btn_author, 8),
            (self.nav_btn_slice, 3),
            (self.nav_btn_convert, 5),
            (self.nav_btn_visio, 1),
        ]

        for btn, idx in self._nav_buttons:
            btn.setCheckable(True)
            btn.setSizePolicy(QSizePolicy.Maximum, QSizePolicy.Fixed)
            btn.setStyleSheet(TOGGLE_PILL_STYLE)
            btn.setCursor(Qt.PointingHandCursor)
            btn.clicked.connect(lambda checked, i=idx: self.set_active_operation(i))
            row2.addWidget(btn)

        self.lbl_jump = QLabel("<b>Jump:</b>")
        self.op_combo = QComboBox()
        self.op_combo.setStyleSheet(COMBOBOX_STYLE_TOOLBAR)
        self.op_combo.addItems(
            [
                "Convert Legacy .doc to .docx (Explicit LibreOffice)",
                "Extract Embedded Visio Diagrams",
                "Subtractive Slicing (Split by Clause)",
                "Extract Sections by Heading",
                "Compare Documents (Native Word Diff)",
                "Convert Document Format (Auto / Word)",
                "Convert Word to Markdown (.md)",
                "Draft 3GPP Contribution (CR / pCR)",
                "Consolidate / Anonymize Track Changes Authors",
            ]
        )
        self.op_combo.currentIndexChanged.connect(self._on_op_combo_changed)
        # Keep self.op_combo hidden yet functional for legacy tests and callers
        self.lbl_jump.setVisible(False)
        self.op_combo.setVisible(False)
        row2.addWidget(self.lbl_jump)
        row2.addWidget(self.op_combo)

        ribbon_layout.addWidget(nav_row_widget)
        layout.addWidget(self.ribbon_frame)

        # ---------------------------------------------------------------------
        # Operation Cards (Stacked Widget)
        # ---------------------------------------------------------------------
        self.stack = QStackedWidget()

        self.card_lo_convert = QWidget()
        lo_layout = QVBoxLayout(self.card_lo_convert)
        self.lo_drop = InteractiveDropLabel(
            "📥 Drag & Drop legacy .doc files here for batch conversion via LibreOffice",
            [".doc"],
        )
        self.lo_drop.file_dropped.connect(self._on_lo_dropped)
        lo_layout.addWidget(self.lo_drop)
        self.stack.addWidget(self.card_lo_convert)

        self.card_visio = WordVisioCard()
        self.card_visio.extract_visio_requested.connect(self.extract_visio_requested)
        self.stack.addWidget(self.card_visio)

        self.card_split = QWidget()
        split_layout = QVBoxLayout(self.card_split)
        self.split_drop = InteractiveDropLabel(
            "📥 Drag & Drop Word document (.docx) here to split by major heading clauses",
            [".docx"],
        )
        self.split_drop.file_dropped.connect(self._on_split_dropped)
        split_layout.addWidget(self.split_drop)
        self.stack.addWidget(self.card_split)

        self.card_excerpt = WordExcerptCard()
        self.stack.addWidget(self.card_excerpt)

        self.card_compare = WordCompareCard()
        self.card_compare.compare_doc_requested.connect(self.compare_doc_requested)
        self.stack.addWidget(self.card_compare)

        self.card_convert = WordConverterCard()
        self.card_convert.convert_doc_requested.connect(self.convert_doc_requested)
        self.stack.addWidget(self.card_convert)

        self.card_markdown = WordMarkdownCard()
        self.card_markdown.convert_markdown_requested.connect(self.convert_markdown_requested)
        self.card_markdown.convert_markdown_to_word_requested.connect(
            self.convert_markdown_to_word_requested
        )
        self.stack.addWidget(self.card_markdown)

        self.card_cr_pcr = WordCrPcrCard()
        self.stack.addWidget(self.card_cr_pcr)

        self.card_author_renamer = WordSanitizerCard()
        self.stack.addWidget(self.card_author_renamer)

        # Backward-compatible aliases
        self.card_sanitizer = self.card_author_renamer

        self.btn_swap_slots = self.card_compare.btn_swap_slots
        self.chk_cmp_formatting = self.card_compare.chk_cmp_formatting
        self.pane_a = self.card_compare.pane_a
        self.pane_b = self.card_compare.pane_b

        self.pane_cr = self.card_cr_pcr.pane_cr
        self.cr_mode_combo = self.card_cr_pcr.cr_mode_combo
        self.btn_launch_cr = self.card_cr_pcr.btn_launch_cr
        self.cr_card_dotx_lbl = self.card_cr_pcr.cr_card_dotx_lbl
        self.btn_cr_install_dotx = self.card_cr_pcr.btn_cr_install_dotx
        self.btn_cr_open_tpl = self.card_cr_pcr.btn_cr_open_tpl
        self.btn_cr_install_dotx.clicked.connect(self._on_toggle_dotx_clicked)

        self.pane_visio = self.card_visio.pane_visio
        self.run_visio_btn = self.card_visio.run_visio_btn
        self.visio_result_frame = self.card_visio.visio_result_frame
        self.visio_open_folder_btn = self.card_visio.visio_open_folder_btn

        self.pane_convert = self.card_convert.pane_convert
        self.conv_engine_lbl = self.card_convert.conv_engine_lbl
        self.conv_result_frame = self.card_convert.conv_result_frame
        self.conv_open_folder_btn = self.card_convert.conv_open_folder_btn
        self.format_combo = self.card_convert.format_combo

        self.lbl_excerpt_match_count = self.card_excerpt.lbl_excerpt_match_count
        self.excerpt_select_all_btn = self.card_excerpt.excerpt_select_all_btn
        self.excerpt_select_l1_btn = self.card_excerpt.excerpt_select_l1_btn
        self.excerpt_clear_btn = self.card_excerpt.excerpt_clear_btn
        self.excerpt_clear_sel_btn = self.card_excerpt.excerpt_clear_btn
        self.excerpt_tree = self.card_excerpt.excerpt_tree
        self.excerpt_drop = self.card_excerpt.excerpt_drop
        self.chk_also_markdown = getattr(self.card_excerpt, "chk_also_markdown", None)

        self.pane_md = self.card_markdown.pane_markdown
        self.pane_markdown = self.card_markdown.pane_markdown
        self.pane_m2w = getattr(self.card_markdown, "pane_m2w", None)
        self.run_m2w_btn = getattr(self.card_markdown, "run_m2w_btn", None)
        self.md_revisions_combo = self.card_markdown.md_revisions_combo
        self.md_media_cb = self.card_markdown.md_media_cb
        self.md_header_cb = self.card_markdown.md_header_cb
        self.run_markdown_btn = self.card_markdown.run_markdown_btn
        self.md_progress = self.card_markdown.md_progress
        self.md_result_frame = self.card_markdown.md_result_frame
        self.md_result_lbl = self.card_markdown.md_result_lbl
        self.md_result_icon = self.card_markdown.md_result_icon
        self.md_open_folder_btn = self.card_markdown.md_open_folder_btn
        self.md_open_file_btn = self.card_markdown.md_open_file_btn
        self.md_copy_btn = self.card_markdown.md_copy_btn
        self.md_preview_toggle_btn = self.card_markdown.md_preview_toggle_btn
        self.md_preview_edit = self.card_markdown.md_preview_edit

        self.lbl_kpi_lang = self.card_sanitizer.lbl_kpi_lang
        self.lbl_kpi_authors = self.card_sanitizer.lbl_kpi_authors
        self.lbl_kpi_lang_rev = self.card_sanitizer.lbl_kpi_lang_rev
        self.lbl_kpi_comments = self.card_sanitizer.lbl_kpi_comments
        self.chk_sec_lang = self.card_sanitizer.chk_sec_lang
        self.combo_target_lang = self.card_sanitizer.combo_target_lang
        self.chk_sec_authors = self.card_sanitizer.chk_sec_authors
        self.author_table = self.card_sanitizer.author_table
        self.pane_sanitizer = self.card_sanitizer.pane_sanitizer
        self.pane_author = self.card_sanitizer.pane_sanitizer
        self.edit_target_author = self.card_sanitizer.edit_target_author
        self.edit_target_initials = self.card_sanitizer.edit_target_initials
        self.chk_sec_cleanup = self.card_sanitizer.chk_sec_cleanup
        self.chk_remove_comments = self.card_sanitizer.chk_remove_comments

        layout.addWidget(self.stack, 1)

        # Default to Compare Documents (Index 4)
        self.set_active_operation(4)
