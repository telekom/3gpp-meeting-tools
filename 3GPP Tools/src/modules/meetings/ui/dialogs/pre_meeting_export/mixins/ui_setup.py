"""UI layout construction mixin for PreMeetingExportDialog."""

import datetime
import re

from PyQt5.QtWidgets import (
    QButtonGroup,
    QCheckBox,
    QComboBox,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QRadioButton,
    QSizePolicy,
    QStackedWidget,
    QTreeWidget,
    QVBoxLayout,
    QWidget,
)

from core.ui.theme import C
from core.ui.ui_components import BUTTON_STYLE_TOOLBAR_SECONDARY
from modules.meetings.ui.dialogs.pre_meeting_export.styles import (
    STYLE_COMPACT_BTN,
    STYLE_TOGGLE_ACTIVE,
    STYLE_TOGGLE_INACTIVE,
    STYLE_TREE,
)


class PreMeetingUiSetupMixin:
    """Builds the header card, scope selector, company filter, options, and action bar."""

    def _setup_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(8)
        main_layout.setContentsMargins(12, 10, 12, 10)

        # --- Header Card ---
        header_frame = QFrame()
        header_frame.setStyleSheet(f"""
            QFrame {{
                background-color: {C.TEAL_SURFACE};
                border: 1px solid {C.TEAL_BORDER};
                border-radius: 6px;
                padding: 6px;
            }}
        """)
        header_layout = QVBoxLayout(header_frame)
        header_layout.setContentsMargins(8, 4, 8, 4)
        header_layout.setSpacing(1)

        wg = str(self.mtg_info.get("wg_name", "3GPP")).strip()
        mtg = str(self.mtg_info.get("meeting_number", "")).strip()
        title_text = f"<b>🧭 {wg} #{mtg} Pre-Meeting Briefing & Dossier Export</b>" if (wg or mtg) else "<b>🧭 Pre-Meeting Briefing & Dossier Export</b>"
        lbl_title = QLabel(title_text)
        lbl_title.setStyleSheet(f"font-size: 13px; color: {C.TEAL_DARK}; font-weight: bold; border: none;")
        lbl_sub = QLabel(
            "Generates a structured pre-meeting briefing covering incoming LS, proposed candidate LS replies, "
            "draft CRs/pCRs, and company discussion papers with revision deduplication and contributor analytics."
        )
        lbl_sub.setStyleSheet(f"font-size: 11px; color: {C.TEAL_DARK}; border: none;")
        lbl_sub.setWordWrap(True)

        header_layout.addWidget(lbl_title)
        header_layout.addWidget(lbl_sub)
        main_layout.addWidget(header_frame)

        # --- 2-Column Split Body ---
        body_layout = QHBoxLayout()
        body_layout.setSpacing(10)

        # --- Left Column: Group 1: Scope Selection ---
        scope_group = QGroupBox("1. Select Export Scope")
        scope_group.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        scope_layout = QVBoxLayout(scope_group)
        scope_layout.setSpacing(6)
        scope_layout.setContentsMargins(10, 12, 10, 8)

        self.scope_btn_group = QButtonGroup(self)

        n_followed = len(self.followed_wis)
        n_starred = len(self.starred_wis)
        n_tracked = len(self.tracked_wis)
        n_visible = len(self.visible_rows)
        n_total = len(self.all_rows)
        n_all_wis = len(self.meeting_wis)

        self.rb_followed = QRadioButton(f"🎯 Followed WIs only ({n_followed} WIs)")
        self.rb_whole = QRadioButton(f"🌐 Whole Meeting ({n_all_wis} WIs, {n_total} docs)")
        self.rb_starred = QRadioButton(f"⭐ Starred WIs only ({n_starred} WIs)")
        self.rb_visible = QRadioButton(f"🔍 Filtered View ({n_visible} docs)")
        self.rb_tracked = QRadioButton(f"🎯 Followed && ⭐ Starred ({n_tracked} WIs)")
        self.rb_custom = QRadioButton("📋 Custom Selection...")

        self.scope_btn_group.addButton(self.rb_followed)
        self.scope_btn_group.addButton(self.rb_whole)
        self.scope_btn_group.addButton(self.rb_starred)
        self.scope_btn_group.addButton(self.rb_visible)
        self.scope_btn_group.addButton(self.rb_tracked)
        self.scope_btn_group.addButton(self.rb_custom)

        rb_grid = QGridLayout()
        rb_grid.setHorizontalSpacing(10)
        rb_grid.setVerticalSpacing(4)
        rb_grid.addWidget(self.rb_followed, 0, 0)
        rb_grid.addWidget(self.rb_whole, 0, 1)
        rb_grid.addWidget(self.rb_starred, 1, 0)
        rb_grid.addWidget(self.rb_visible, 1, 1)
        rb_grid.addWidget(self.rb_tracked, 2, 0)
        rb_grid.addWidget(self.rb_custom, 2, 1)
        scope_layout.addLayout(rb_grid)

        if n_followed > 0:
            self.rb_followed.setChecked(True)
        else:
            self.rb_whole.setChecked(True)

        self.rb_followed.toggled.connect(self._on_scope_radio_toggled)
        self.rb_whole.toggled.connect(self._on_scope_radio_toggled)
        self.rb_starred.toggled.connect(self._on_scope_radio_toggled)
        self.rb_tracked.toggled.connect(self._on_scope_radio_toggled)
        self.rb_visible.toggled.connect(self._on_scope_radio_toggled)
        self.rb_custom.toggled.connect(self._on_scope_radio_toggled)

        # --- Custom Selection Widget ---
        self.custom_wi_widget = QWidget()
        self.custom_wi_widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        custom_layout = QVBoxLayout(self.custom_wi_widget)
        custom_layout.setContentsMargins(0, 4, 0, 0)
        custom_layout.setSpacing(6)

        self.lbl_custom_hint = QLabel()
        self.lbl_custom_hint.setStyleSheet("font-size: 11px;")
        custom_layout.addWidget(self.lbl_custom_hint)

        ctrl_layout = QHBoxLayout()
        ctrl_layout.setSpacing(4)

        self.btn_view_wi = QPushButton("🗂️ Work Items")
        self.btn_view_wi.setToolTip("Group by Work Item (WI ➔ Agenda Items)")
        self.btn_view_wi.setStyleSheet(STYLE_TOGGLE_ACTIVE)
        self.btn_view_wi.clicked.connect(lambda: self._switch_custom_view("wi"))
        ctrl_layout.addWidget(self.btn_view_wi)

        self.btn_view_ai = QPushButton("📋 Agenda Items")
        self.btn_view_ai.setToolTip("Group by Agenda Item (AI ➔ Work Items)")
        self.btn_view_ai.setStyleSheet(STYLE_TOGGLE_INACTIVE)
        self.btn_view_ai.clicked.connect(lambda: self._switch_custom_view("ai"))
        ctrl_layout.addWidget(self.btn_view_ai)

        ctrl_layout.addStretch()

        self.btn_expand_all = QPushButton("⊞ Expand")
        self.btn_expand_all.setStyleSheet(STYLE_COMPACT_BTN)
        self.btn_expand_all.clicked.connect(self._expand_active_tree)
        ctrl_layout.addWidget(self.btn_expand_all)

        self.btn_collapse_all = QPushButton("⊟ Collapse")
        self.btn_collapse_all.setStyleSheet(STYLE_COMPACT_BTN)
        self.btn_collapse_all.clicked.connect(self._collapse_active_tree)
        ctrl_layout.addWidget(self.btn_collapse_all)

        self.btn_select_all = QPushButton("☑ All")
        self.btn_select_all.setStyleSheet(STYLE_COMPACT_BTN)
        self.btn_select_all.clicked.connect(self._select_all_custom)
        ctrl_layout.addWidget(self.btn_select_all)

        self.btn_clear_all = QPushButton("☐ None")
        self.btn_clear_all.setStyleSheet(STYLE_COMPACT_BTN)
        self.btn_clear_all.clicked.connect(self._clear_all_custom)
        ctrl_layout.addWidget(self.btn_clear_all)

        custom_layout.addLayout(ctrl_layout)

        self.custom_search = QLineEdit()
        self.custom_search.setPlaceholderText("Filter by Work Item, AI # (e.g. 5.1), or topic keyword...")
        self.custom_search.setClearButtonEnabled(True)
        self.custom_search.textChanged.connect(self._filter_custom_tree)
        custom_layout.addWidget(self.custom_search)

        self.tree_stack = QStackedWidget()
        self.tree_stack.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        self.tree_wi = QTreeWidget()
        self.tree_wi.setHeaderHidden(True)
        self.tree_wi.setStyleSheet(STYLE_TREE)
        self.tree_wi.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self._populate_tree_wi()

        self.tree_ai = QTreeWidget()
        self.tree_ai.setHeaderHidden(True)
        self.tree_ai.setStyleSheet(STYLE_TREE)
        self.tree_ai.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self._populate_tree_ai()

        self.tree_stack.addWidget(self.tree_wi)
        self.tree_stack.addWidget(self.tree_ai)
        custom_layout.addWidget(self.tree_stack, stretch=1)

        scope_layout.addWidget(self.custom_wi_widget, stretch=1)

        # Optional Company Filter Bar
        comp_filter_layout = QHBoxLayout()
        comp_filter_layout.setSpacing(6)
        comp_lbl = QLabel("🏢 Focus on Contributor:")
        comp_lbl.setStyleSheet(f"font-size: 11px; font-weight: bold; color: {C.TEAL_DARK};")
        comp_filter_layout.addWidget(comp_lbl)

        self.combo_company = QComboBox()
        self.combo_company.setStyleSheet(f"padding: 3px 6px; font-size: 11px; background: white; border: 1px solid {C.BORDER_STRONG}; border-radius: 4px;")
        self.combo_company.addItem("All Contributors (No Filter)")
        for c in self.meeting_companies:
            self.combo_company.addItem(f"🏢 {c}", c)
        self.combo_company.currentIndexChanged.connect(self._update_estimates)
        comp_filter_layout.addWidget(self.combo_company, stretch=1)

        scope_layout.addLayout(comp_filter_layout)
        body_layout.addWidget(scope_group, stretch=12)

        # --- Right Column: Sections 2 & 3 ---
        right_column = QVBoxLayout()
        right_column.setSpacing(8)

        # --- Group 2: Document Types & Statuses ---
        docs_group = QGroupBox("2. Target Contribution Proposals && Pre-Meeting Filters")
        docs_layout = QVBoxLayout(docs_group)
        docs_layout.setSpacing(4)
        docs_layout.setContentsMargins(10, 12, 10, 8)

        self.chk_ls_in = QCheckBox("📥 Incoming Liaison Statements (LS in)")
        self.chk_ls_in.setChecked(True)
        self.chk_ls_in.stateChanged.connect(self._update_estimates)

        self.chk_ls_out = QCheckBox("📤 Proposed Outgoing LS Replies (LS out)")
        self.chk_ls_out.setChecked(True)
        self.chk_ls_out.stateChanged.connect(self._update_estimates)

        self.chk_pcr = QCheckBox("📝 Proposed Pseudo-CRs (pCR)")
        self.chk_pcr.setChecked(True)
        self.chk_pcr.stateChanged.connect(self._update_estimates)

        self.chk_cr = QCheckBox("📄 Proposed Change Requests (CR)")
        self.chk_cr.setChecked(True)
        self.chk_cr.stateChanged.connect(self._update_estimates)

        self.chk_disc = QCheckBox("💬 Discussion, Evaluation && WID Papers (Disc)")
        self.chk_disc.setChecked(True)  # True by default for Pre-Meeting!
        self.chk_disc.setToolTip("Include company discussion papers, study reports, Key Issue solutions, and Work Item Descriptions.")
        self.chk_disc.stateChanged.connect(self._update_estimates)

        self.chk_general_ls = QCheckBox("📦 Include unassigned && general LS (without followed WI)")
        self.chk_general_ls.setChecked(True)
        self.chk_general_ls.stateChanged.connect(self._update_estimates)

        # Pre-Meeting Lineage & Reservation Controls
        sep_line = QFrame()
        sep_line.setFrameShape(QFrame.HLine)
        sep_line.setStyleSheet(f"color: {C.BORDER_SUBTLE}; margin: 4px 0;")
        docs_layout.addWidget(self.chk_ls_in)
        docs_layout.addWidget(self.chk_ls_out)
        docs_layout.addWidget(self.chk_pcr)
        docs_layout.addWidget(self.chk_cr)
        docs_layout.addWidget(self.chk_disc)
        docs_layout.addWidget(self.chk_general_ls)
        docs_layout.addWidget(sep_line)

        self.chk_dedup_revs = QCheckBox("🔄 Deduplicate revisions (export only latest active revision per chain)")
        self.chk_dedup_revs.setChecked(True)
        self.chk_dedup_revs.setToolTip(
            "When enabled, obsolete predecessor drafts (e.g. S2-260100 revised to S2-260500 prior to meeting) "
            "are superseded so only the latest active revision is exported, annotating predecessor provenance."
        )
        self.chk_dedup_revs.stateChanged.connect(self._update_estimates)
        docs_layout.addWidget(self.chk_dedup_revs)

        self.chk_include_reserved = QCheckBox("⏳ List unuploaded 'Reserved' TDocs in index tables (skip download)")
        self.chk_include_reserved.setChecked(True)
        self.chk_include_reserved.setToolTip(
            "Lists agenda reservations in index tables with a 'Reserved' status badge without attempting futile FTP downloads."
        )
        self.chk_include_reserved.stateChanged.connect(self._update_estimates)
        docs_layout.addWidget(self.chk_include_reserved)

        right_column.addWidget(docs_group)

        # --- Group 3: Options & Destination ---
        opts_group = QGroupBox("3. Briefing Organization && Output Options")
        opts_layout = QVBoxLayout(opts_group)
        opts_layout.setSpacing(4)
        opts_layout.setContentsMargins(10, 12, 10, 8)

        self.chk_user_notes = QCheckBox("📌 Embed personal 'My Status' && 'My Notes' into dossiers && index")
        self.chk_user_notes.setChecked(True)
        self.chk_user_notes.setToolTip("Embeds your triage status (Support, Object, My TDoc) and saved notes into the exported dossiers.")

        self.chk_local_indices = QCheckBox("Generate local index.md briefing inside each Work Item directory")
        self.chk_local_indices.setChecked(True)

        self.chk_download = QCheckBox("Download missing TDoc ZIP packages automatically from FTP")
        self.chk_download.setChecked(True)

        self.chk_force = QCheckBox("Force reconvert (overwrite existing local markdown caches)")
        self.chk_force.setChecked(False)

        self.chk_compile_dossier = QCheckBox("Also compile consolidated Navigable Dossier (HTML && PDF)")
        self.chk_compile_dossier.setChecked(False)
        self.chk_compile_dossier.setToolTip(
            "Synthesizes an interactive HTML dossier and bookmarked vector PDF containing all exported deliverables."
        )

        self.chk_auto_open = QCheckBox("Open export directory in Windows Explorer upon completion")
        self.chk_auto_open.setChecked(True)

        opts_layout.addWidget(self.chk_user_notes)
        opts_layout.addWidget(self.chk_local_indices)
        opts_layout.addWidget(self.chk_download)
        opts_layout.addWidget(self.chk_force)
        opts_layout.addWidget(self.chk_compile_dossier)
        opts_layout.addWidget(self.chk_auto_open)

        opts_layout.addWidget(QLabel("Destination Directory:"))
        dest_layout = QHBoxLayout()
        dest_layout.setSpacing(6)
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S")
        clean_wg = re.sub(r'[^\w\-]', '_', wg)
        clean_mtg = re.sub(r'[^\w\-]', '_', mtg)
        folder_name = f"PreMeeting_{clean_wg}_{clean_mtg}_{timestamp}" if (clean_wg or clean_mtg) else f"PreMeeting_{timestamp}"
        default_dir = self.meeting_dir / "Export" / folder_name
        self.dest_edit = QLineEdit(str(default_dir))
        self.dest_edit.setReadOnly(False)
        self.dest_edit.textChanged.connect(self._update_destination_status)
        dest_layout.addWidget(self.dest_edit)

        self.btn_browse = QPushButton("📁 Browse...")
        self.btn_browse.setStyleSheet(STYLE_COMPACT_BTN)
        self.btn_browse.clicked.connect(self._browse_destination)
        dest_layout.addWidget(self.btn_browse)

        self.btn_recent = QPushButton("🕒 Recent...")
        self.btn_recent.setStyleSheet(STYLE_COMPACT_BTN)
        self.btn_recent.setToolTip("Quickly select an existing export directory from this meeting to append missing TDocs")
        self.btn_recent.clicked.connect(self._show_recent_exports_menu)
        dest_layout.addWidget(self.btn_recent)

        opts_layout.addLayout(dest_layout)

        # Status badge for detected existing export folder
        self.lbl_dest_status = QLabel()
        self.lbl_dest_status.setWordWrap(True)
        self.lbl_dest_status.setVisible(False)
        opts_layout.addWidget(self.lbl_dest_status)

        right_column.addWidget(opts_group)
        body_layout.addLayout(right_column, stretch=9)
        main_layout.addLayout(body_layout, stretch=1)

        # --- Estimated Scope Banner ---
        self.lbl_estimate = QLabel()
        self.lbl_estimate.setWordWrap(True)
        main_layout.addWidget(self.lbl_estimate)

        # --- Action Buttons ---
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()

        btn_cancel = QPushButton("Cancel")
        btn_cancel.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        btn_cancel.clicked.connect(self.reject)
        btn_layout.addWidget(btn_cancel)

        btn_start = QPushButton("🚀 Start Pre-Meeting Export")
        btn_start.setStyleSheet(f"""
            QPushButton {{
                background-color: {C.TEAL_DARK};
                color: {C.BG_CARD};
                font-weight: bold;
                border-radius: 4px;
                padding: 6px 16px;
                border: 1px solid {C.TEAL_DARK};
            }}
            QPushButton:hover {{
                background-color: {C.EMERALD_DARK};
            }}
            QPushButton:pressed {{
                background-color: {C.WG_RAN_TEXT};
            }}
        """)
        btn_start.setDefault(True)
        btn_start.clicked.connect(self._on_start_clicked)
        btn_layout.addWidget(btn_start)

        main_layout.addLayout(btn_layout)

        self._sync_scope_ui_state()

        self.tree_wi.itemChanged.connect(lambda item, col: self._on_tree_item_changed(self.tree_wi, item, col))
        self.tree_ai.itemChanged.connect(lambda item, col: self._on_tree_item_changed(self.tree_ai, item, col))
