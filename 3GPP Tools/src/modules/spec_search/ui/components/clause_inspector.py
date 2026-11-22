"""SpecClauseInspector widget: two-tier header, action toolbar, and signal emitters."""

from typing import Optional
from urllib.parse import unquote

from PyQt5.QtCore import Qt, QUrl, pyqtSignal
from PyQt5.QtGui import QDesktopServices, QKeySequence
from PyQt5.QtWidgets import (
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QMenu,
    QPushButton,
    QShortcut,
    QTextBrowser,
    QVBoxLayout,
    QWidget,
)

from core.ui.theme import C

from .clause_inspector_render import SpecClauseInspectorRenderMixin
from .styles import (
    INSPECTOR_BTN_AI,
    INSPECTOR_BTN_BASE,
    INSPECTOR_BTN_INFO,
    INSPECTOR_BTN_SUCCESS,
    INSPECTOR_BTN_WARNING,
)


class SpecClauseInspector(SpecClauseInspectorRenderMixin, QGroupBox):
    """
    Renders rich clause content with context excerpts, match navigation, citation copying,
    1-click Word document launching, and deep-linking to the Specifications catalogue.
    Features a clean two-tier header and flexible layout orientation controls.
    """

    compare_clause_requested = pyqtSignal(str, str, str, str)  # spec_number, version, clause_number, clause_title
    open_document_requested = pyqtSignal(str, str, str, str)    # spec_number, version, clause_number, format ("word", "pdf", "folder")
    show_in_specs_tab_requested = pyqtSignal(str, str)          # spec_number, version
    quick_diff_requested = pyqtSignal(str, str, str, str, str)  # spec_number, base_ver, target_ver, clause_number, clause_title
    jump_to_version_requested = pyqtSignal(str, str, str)       # spec_number, target_version, clause_number
    extract_clause_requested = pyqtSignal(str, str, str, str, str, bool)  # spec_number, version, clause_number, clause_title, content, save_as
    export_markdown_requested = pyqtSignal(str, str, str, str, str)       # spec_number, version, clause_number, clause_title, content
    layout_orientation_toggle_requested = pyqtSignal(bool)      # True = side-by-side (horizontal), False = bottom (vertical)
    maximize_toggle_requested = pyqtSignal(bool)                # True = maximized, False = restored
    version_changes_requested = pyqtSignal(str, str)            # spec_number, version
    identify_origin_tdoc_requested = pyqtSignal(str, str, str, str, str)  # spec_number, base_ver, target_ver, clause_number, clause_title
    clause_reference_clicked = pyqtSignal(str, str, str)        # spec_number, version ("" for cross-spec), target_clause

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__("Clause Content Inspector", parent)
        self._current_search_query = ""
        self._last_citation_text = ""
        self._current_spec = ""
        self._current_version = ""
        self._current_clause_num = ""
        self._current_clause_title = ""
        self._current_content = ""
        self._current_release_date = ""
        self._first_added_ver = ""
        self._prev_ver = ""
        self._is_side_by_side = False
        self._is_maximized = False
        self._nav_history: list = []
        self._nav_index: int = -1
        self.setStyleSheet(f"""
            QGroupBox {{
                font-weight: bold;
                font-size: 12px;
                color: {C.TEXT_STRONG};
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 6px;
                margin-top: 6px;
                padding-top: 10px;
                background-color: {C.BG_CARD};
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 4px;
            }}
        """)
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 6, 8, 8)
        layout.setSpacing(5)

        # ---------------------------------------------------------------------
        # Tier 1: Back/Forward History, Context Title, First Added Chip & Toggles
        # ---------------------------------------------------------------------
        tier1_bar = QHBoxLayout()
        tier1_bar.setSpacing(6)

        self.btn_nav_back = QPushButton("⬅")
        self.btn_nav_back.setToolTip("Back to previous clause (Alt+Left)")
        self.btn_nav_back.setCursor(Qt.PointingHandCursor)
        self.btn_nav_back.setStyleSheet(INSPECTOR_BTN_BASE + "QPushButton { min-width: 22px; padding: 2px 6px; font-weight: bold; }")
        self.btn_nav_back.setEnabled(False)
        self.btn_nav_back.clicked.connect(self.navigate_back)
        tier1_bar.addWidget(self.btn_nav_back)

        self.btn_nav_forward = QPushButton("➡")
        self.btn_nav_forward.setToolTip("Forward to next clause (Alt+Right)")
        self.btn_nav_forward.setCursor(Qt.PointingHandCursor)
        self.btn_nav_forward.setStyleSheet(INSPECTOR_BTN_BASE + "QPushButton { min-width: 22px; padding: 2px 6px; font-weight: bold; }")
        self.btn_nav_forward.setEnabled(False)
        self.btn_nav_forward.clicked.connect(self.navigate_forward)
        tier1_bar.addWidget(self.btn_nav_forward)

        self._shortcut_nav_back = QShortcut(QKeySequence("Alt+Left"), self)
        self._shortcut_nav_back.setContext(Qt.WidgetWithChildrenShortcut)
        self._shortcut_nav_back.activated.connect(self.navigate_back)

        self._shortcut_nav_forward = QShortcut(QKeySequence("Alt+Right"), self)
        self._shortcut_nav_forward.setContext(Qt.WidgetWithChildrenShortcut)
        self._shortcut_nav_forward.activated.connect(self.navigate_forward)

        self.title_lbl = QLabel("No clause selected")
        self.title_lbl.setStyleSheet(f"font-weight: bold; color: {C.ACCENT_SKY}; font-size: 12px;")
        tier1_bar.addWidget(self.title_lbl, stretch=1)

        # Clickable First Introduced Chip
        self.first_added_badge = QPushButton("")
        self.first_added_badge.setCursor(Qt.PointingHandCursor)
        self.first_added_badge.setToolTip("Click to jump directly to the version where this clause was first introduced")
        self.first_added_badge.setStyleSheet(f"""
            QPushButton {{
                font-weight: bold;
                font-size: 11px;
                background-color: {C.SUCCESS_SOFT_ALT};
                color: {C.SUCCESS};
                border: 1px solid {C.SUCCESS_BORDER_MEDIUM};
                padding: 1px 8px;
                border-radius: 10px;
            }}
            QPushButton:hover {{
                background-color: {C.SUCCESS_BORDER};
                border-color: {C.EMERALD_ACCENT};
            }}
        """)
        self.first_added_badge.setVisible(False)
        self.first_added_badge.clicked.connect(self._on_first_added_clicked)
        tier1_bar.addWidget(self.first_added_badge)

        # Match navigation group
        self.match_badge = QLabel("")
        self.match_badge.setStyleSheet(f"font-weight: bold; font-size: 11px; color: {C.WARNING}; background-color: {C.WARNING_SOFT_ALT}; border: 1px solid {C.WARNING_BORDER}; padding: 2px 6px; border-radius: 4px;")
        self.match_badge.setVisible(False)
        tier1_bar.addWidget(self.match_badge)

        self.btn_prev = QPushButton("◀")
        self.btn_prev.setToolTip("Jump to previous match in clause")
        self.btn_prev.setCursor(Qt.PointingHandCursor)
        self.btn_prev.setStyleSheet(INSPECTOR_BTN_BASE + "QPushButton { min-width: 20px; font-weight: bold; }")
        self.btn_prev.clicked.connect(self._find_prev)
        self.btn_prev.setEnabled(False)
        tier1_bar.addWidget(self.btn_prev)

        self.btn_next = QPushButton("▶")
        self.btn_next.setToolTip("Jump to next match in clause")
        self.btn_next.setCursor(Qt.PointingHandCursor)
        self.btn_next.setStyleSheet(INSPECTOR_BTN_BASE + "QPushButton { min-width: 20px; font-weight: bold; }")
        self.btn_next.clicked.connect(self._find_next)
        self.btn_next.setEnabled(False)
        tier1_bar.addWidget(self.btn_next)

        # Layout Orientation Toggle (Bottom Split <-> Right Pane)
        self.btn_layout_toggle = QPushButton("◫ Right Pane")
        self.btn_layout_toggle.setToolTip("Toggle inspector layout: Bottom split vs. Right pane side-by-side")
        self.btn_layout_toggle.setCursor(Qt.PointingHandCursor)
        self.btn_layout_toggle.setStyleSheet(INSPECTOR_BTN_BASE + "QPushButton { font-weight: 600; padding: 2px 8px; }")
        self.btn_layout_toggle.clicked.connect(self._on_layout_toggle_clicked)
        tier1_bar.addWidget(self.btn_layout_toggle)

        # Maximize / Restore Toggle
        self.btn_maximize = QPushButton("⤢")
        self.btn_maximize.setToolTip("Maximize reading pane to full height")
        self.btn_maximize.setCursor(Qt.PointingHandCursor)
        self.btn_maximize.setStyleSheet(INSPECTOR_BTN_BASE + "QPushButton { font-weight: bold; min-width: 22px; padding: 2px 6px; }")
        self.btn_maximize.clicked.connect(self._on_maximize_clicked)
        tier1_bar.addWidget(self.btn_maximize)

        layout.addLayout(tier1_bar)

        # ---------------------------------------------------------------------
        # Tier 2: Grouped Action Toolbar (Copy, Extract, Full Doc, Diff, LLM)
        # ---------------------------------------------------------------------
        tier2_bar = QHBoxLayout()
        tier2_bar.setSpacing(6)

        # Group 1: Copy & Extract
        self.btn_copy = QPushButton("📋 Copy ▾")
        self.btn_copy.setToolTip("Copy clause text (Rich HTML for Email/Outlook, Markdown quote block, or Plain Text)")
        self.btn_copy.setCursor(Qt.PointingHandCursor)
        self.btn_copy.setStyleSheet(INSPECTOR_BTN_BASE)
        copy_menu = QMenu(self)
        act_cp_email = copy_menu.addAction("📋 Copy for Email / Outlook (Rich Text)")
        act_cp_email.triggered.connect(lambda: self._copy_formatted("email"))
        act_cp_md = copy_menu.addAction("📝 Copy as Markdown (Quote Block)")
        act_cp_md.triggered.connect(lambda: self._copy_formatted("markdown"))
        act_cp_plain = copy_menu.addAction("📄 Copy Plain Text")
        act_cp_plain.triggered.connect(lambda: self._copy_formatted("text"))
        copy_menu.addSeparator()
        act_cp_cite = copy_menu.addAction("🔖 Copy Citation Only")
        act_cp_cite.triggered.connect(lambda: self._copy_formatted("citation"))
        self.btn_copy.setMenu(copy_menu)
        self.btn_copy.setEnabled(False)
        tier2_bar.addWidget(self.btn_copy)

        self.btn_extract = QPushButton("✂️ Extract ▾")
        self.btn_extract.setToolTip("Extract this clause to a clean Word (.docx) document or export to Markdown (.md)")
        self.btn_extract.setCursor(Qt.PointingHandCursor)
        self.btn_extract.setStyleSheet(INSPECTOR_BTN_SUCCESS)
        extract_menu = QMenu(self)
        act_ext_word = extract_menu.addAction("✂️ Extract Clause in Word (.docx)")
        act_ext_word.triggered.connect(lambda: self._on_extract_clicked(save_as=False))
        act_ext_save = extract_menu.addAction("💾 Save Clause as Word (.docx)...")
        act_ext_save.triggered.connect(lambda: self._on_extract_clicked(save_as=True))
        extract_menu.addSeparator()
        act_ext_md = extract_menu.addAction("📘 Export Clause to Markdown (.md)...")
        act_ext_md.triggered.connect(self._on_export_markdown_clicked)
        self.btn_extract.setMenu(extract_menu)
        self.btn_extract.setEnabled(False)
        tier2_bar.addWidget(self.btn_extract)

        # Separator line
        sep1 = QFrame()
        sep1.setFrameShape(QFrame.VLine)
        sep1.setStyleSheet(f"color: {C.BORDER_SUBTLE};")
        tier2_bar.addWidget(sep1)

        # Group 2: Full Document & Specs Catalogue
        self.btn_open_doc = QPushButton("📖 Full Doc ▾")
        self.btn_open_doc.setToolTip("Open complete specification document in Microsoft Word or PDF")
        self.btn_open_doc.setCursor(Qt.PointingHandCursor)
        self.btn_open_doc.setStyleSheet(INSPECTOR_BTN_BASE)
        doc_menu = QMenu(self)
        act_word = doc_menu.addAction("📝 Open Full Specification in Word (.docx / .doc)")
        act_word.triggered.connect(lambda: self._on_open_doc_clicked("word"))
        act_pdf = doc_menu.addAction("📕 Open Converted PDF (.pdf)")
        act_pdf.triggered.connect(lambda: self._on_open_doc_clicked("pdf"))
        doc_menu.addSeparator()
        act_folder = doc_menu.addAction("📂 Open Containing Folder in Explorer")
        act_folder.triggered.connect(lambda: self._on_open_doc_clicked("folder"))
        self.doc_menu = doc_menu
        self.act_pdf = act_pdf
        self.btn_open_doc.setMenu(doc_menu)
        self.btn_open_doc.setEnabled(False)
        tier2_bar.addWidget(self.btn_open_doc)

        self.btn_show_in_specs = QPushButton("🔍 Specs Tab")
        self.btn_show_in_specs.setToolTip("Switch to Specifications Tab with this spec and version selected")
        self.btn_show_in_specs.setCursor(Qt.PointingHandCursor)
        self.btn_show_in_specs.setStyleSheet(INSPECTOR_BTN_INFO)
        self.btn_show_in_specs.clicked.connect(self._on_show_in_specs_clicked)
        self.btn_show_in_specs.setEnabled(False)
        tier2_bar.addWidget(self.btn_show_in_specs)

        # Separator line
        sep2 = QFrame()
        sep2.setFrameShape(QFrame.VLine)
        sep2.setStyleSheet(f"color: {C.BORDER_SUBTLE};")
        tier2_bar.addWidget(sep2)

        # Group 3: Comparison & LLM Analysis
        self.btn_quick_diff = QPushButton("🔍 Diff Prev")
        self.btn_quick_diff.setToolTip("Compare this clause with the immediately preceding release")
        self.btn_quick_diff.setCursor(Qt.PointingHandCursor)
        self.btn_quick_diff.setStyleSheet(INSPECTOR_BTN_WARNING)
        self.btn_quick_diff.clicked.connect(self._on_quick_diff_clicked)
        self.btn_quick_diff.setEnabled(False)
        tier2_bar.addWidget(self.btn_quick_diff)

        self.btn_llm_diff = QPushButton("🤖 LLM Diff")
        self.btn_llm_diff.setToolTip("Compare this clause with another version and generate an LLM analysis prompt with hierarchical context.")
        self.btn_llm_diff.setCursor(Qt.PointingHandCursor)
        self.btn_llm_diff.setStyleSheet(INSPECTOR_BTN_AI)
        self.btn_llm_diff.clicked.connect(self._on_llm_diff_clicked)
        self.btn_llm_diff.setEnabled(False)
        tier2_bar.addWidget(self.btn_llm_diff)

        self.btn_version_changes = QPushButton("📜 Version Changes")
        self.btn_version_changes.setToolTip("View all integrated CRs, p-CRs, and TDocs for this version, and directly open the source contribution documents.")
        self.btn_version_changes.setCursor(Qt.PointingHandCursor)
        self.btn_version_changes.setStyleSheet(INSPECTOR_BTN_BASE)
        self.btn_version_changes.clicked.connect(self._on_version_changes_clicked)
        self.btn_version_changes.setEnabled(False)
        tier2_bar.addWidget(self.btn_version_changes)

        self.btn_origin_tdoc = QPushButton("🎯 Origin TDoc")
        self.btn_origin_tdoc.setToolTip("Identify which specific CR or p-CR introduced or modified this clause across releases.")
        self.btn_origin_tdoc.setCursor(Qt.PointingHandCursor)
        self.btn_origin_tdoc.setStyleSheet(INSPECTOR_BTN_SUCCESS)
        self.btn_origin_tdoc.clicked.connect(self._on_origin_tdoc_clicked)
        self.btn_origin_tdoc.setEnabled(False)
        tier2_bar.addWidget(self.btn_origin_tdoc)

        tier2_bar.addStretch()
        layout.addLayout(tier2_bar)

        # Document Browser
        self.browser = QTextBrowser()
        self.browser.setReadOnly(True)
        self.browser.setOpenLinks(False)
        self.browser.setOpenExternalLinks(False)
        self.browser.anchorClicked.connect(self._on_anchor_clicked)
        self.browser.setPlaceholderText("Select a cell in the evolution matrix above to inspect the full clause content and match context...")
        self.browser.setStyleSheet(f"""
            QTextBrowser {{
                border: 1px solid {C.BORDER_SUBTLE};
                border-radius: 4px;
                background-color: {C.BG_CARD};
                padding: 6px;
            }}
        """)
        layout.addWidget(self.browser)

    def set_side_by_side(self, side_by_side: bool):
        """Sets the side-by-side orientation state externally."""
        self._is_side_by_side = side_by_side
        self._update_layout_toggle_btn()

    def _on_layout_toggle_clicked(self):
        self._is_side_by_side = not self._is_side_by_side
        self._update_layout_toggle_btn()
        self.layout_orientation_toggle_requested.emit(self._is_side_by_side)

    def _update_layout_toggle_btn(self):
        if self._is_side_by_side:
            self.btn_layout_toggle.setText("⬒ Bottom")
            self.btn_layout_toggle.setToolTip("Switch inspector back to bottom split layout")
        else:
            self.btn_layout_toggle.setText("◫ Right Pane")
            self.btn_layout_toggle.setToolTip("Switch inspector to right-pane side-by-side layout")

    def _on_maximize_clicked(self):
        self._is_maximized = not self._is_maximized
        self._update_maximize_btn()
        self.maximize_toggle_requested.emit(self._is_maximized)

    def _update_maximize_btn(self):
        if self._is_maximized:
            self.btn_maximize.setText("⤡")
            self.btn_maximize.setToolTip("Restore normal inspector height")
        else:
            self.btn_maximize.setText("⤢")
            self.btn_maximize.setToolTip("Maximize reading pane to full height")

    def _on_first_added_clicked(self):
        if self._current_spec and self._first_added_ver:
            self.jump_to_version_requested.emit(
                self._current_spec,
                self._first_added_ver,
                self._current_clause_num,
            )

    def _on_open_doc_clicked(self, fmt="word"):
        if self._current_spec and self._current_version:
            self.open_document_requested.emit(
                self._current_spec,
                self._current_version,
                self._current_clause_num,
                fmt,
            )

    def _on_show_in_specs_clicked(self):
        if self._current_spec:
            self.show_in_specs_tab_requested.emit(
                self._current_spec,
                self._current_version,
            )

    def _on_quick_diff_clicked(self):
        if self._current_spec and self._prev_ver and self._current_version:
            self.quick_diff_requested.emit(
                self._current_spec,
                self._prev_ver,
                self._current_version,
                self._current_clause_num,
                self._current_clause_title,
            )

    def _on_llm_diff_clicked(self):
        if self._current_spec and self._current_clause_num:
            self.compare_clause_requested.emit(
                self._current_spec,
                self._current_version,
                self._current_clause_num,
                self._current_clause_title,
            )

    def _on_extract_clicked(self, save_as: bool = False):
        if self._current_spec and self._current_version and self._current_clause_num:
            self.extract_clause_requested.emit(
                self._current_spec,
                self._current_version,
                self._current_clause_num,
                self._current_clause_title,
                self._current_content,
                save_as,
            )

    def _on_export_markdown_clicked(self):
        if self._current_spec and self._current_version and self._current_clause_num:
            self.export_markdown_requested.emit(
                self._current_spec,
                self._current_version,
                self._current_clause_num,
                self._current_clause_title,
                self._current_content,
            )

    def _on_version_changes_clicked(self):
        if self._current_spec and self._current_version:
            self.version_changes_requested.emit(self._current_spec, self._current_version)

    def _on_origin_tdoc_clicked(self):
        if self._current_spec and self._current_version and self._current_clause_num:
            self.identify_origin_tdoc_requested.emit(
                self._current_spec,
                self._prev_ver or "",
                self._current_version,
                self._current_clause_num,
                self._current_clause_title,
            )

    def _on_anchor_clicked(self, url: QUrl):
        """Routes clicked clause/spec-clause links to SpecSearchTab or opens external URLs."""
        scheme = (url.scheme() or "").lower()
        if scheme == "clause":
            raw = unquote((url.host() or "") + (url.path() or "")).lstrip("/")
            if raw.lower().startswith("annex "):
                raw = "Annex " + raw[6:].upper()
            if raw and self._current_spec and self._current_version:
                self.clause_reference_clicked.emit(self._current_spec, self._current_version, raw)
            return

        if scheme == "spec-clause":
            raw_path = unquote(url.path() or "").lstrip("/")
            if url.host():
                spec = unquote(url.host()).strip()
                clause = raw_path.strip()
            elif "/" in raw_path:
                spec, clause = (part.strip() for part in raw_path.split("/", 1))
            else:
                spec, clause = "", ""
            if spec and clause:
                self.clause_reference_clicked.emit(spec, "", clause)
            return

        if scheme in ("http", "https", "mailto"):
            QDesktopServices.openUrl(url)
