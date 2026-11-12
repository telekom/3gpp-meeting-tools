# --- File: src/modules/meetings/ui/dialogs/pre_meeting_export_dialog.py ---
"""
pre_meeting_export_dialog.py - Configuration dialog for the Pre-Meeting Report & Dossier Exporter.

Features:
- Responsive 2-column layout (Left: Scope selection and full-height WI/AI tree; Right: Document types and conversion options)
- Pre-checks Followed & Starred WIs automatically
- Dual-Mode synchronized custom tree (WI -> AI and AI -> WI)
- Pre-Meeting document types (Discussion papers, WIDs, proposed CRs, proposed pCRs, incoming LS, proposed outgoing LS)
- Pre-Meeting options: Revision deduplication, reserved TDocs listing, personal stance/notes embedding, optional company filtering
- Live real-time estimated scope banner with dynamic category counts
"""

import datetime
import re
from pathlib import Path
from typing import Dict, List, Set, Optional, Tuple, Any

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QGridLayout, QLabel, QGroupBox,
    QRadioButton, QCheckBox, QPushButton, QLineEdit, QFileDialog,
    QTreeWidget, QTreeWidgetItem, QStackedWidget, QFrame, QComboBox,
    QWidget, QButtonGroup, QMessageBox, QSizePolicy, QMenu
)

from core.ui.ui_components import BUTTON_STYLE_TOOLBAR_SECONDARY
from core.utils.company_sanitizer import CompanySanitizer
from modules.meetings.core.agenda_manager import AgendaItem
from modules.meetings.ui.tdocs_models import TDocsTableModel
from core.ui.theme import C


STYLE_TOGGLE_ACTIVE = f"""
    QPushButton {{
        background-color: #0D9488;
        color: {C.BG_CARD};
        font-weight: bold;
        border: 1px solid {C.TEAL_DARK};
        border-radius: 4px;
        padding: 3px 8px;
        font-size: 11px;
    }}
"""

STYLE_TOGGLE_INACTIVE = f"""
    QPushButton {{
        background-color: {C.BG_CARD};
        color: {C.TEXT_SECONDARY};
        border: 1px solid {C.BORDER_STRONG};
        border-radius: 4px;
        padding: 3px 8px;
        font-size: 11px;
    }}
    QPushButton:hover {{
        background-color: {C.TEAL_SURFACE};
        color: {C.TEAL_DARK};
    }}
"""

STYLE_COMPACT_BTN = f"""
    QPushButton {{
        background-color: {C.BG_SURFACE};
        color: {C.TEXT_STRONG};
        border: 1px solid {C.BORDER_STRONG};
        border-radius: 4px;
        padding: 3px 8px;
        font-size: 11px;
        font-weight: 500;
    }}
    QPushButton:hover {{
        background-color: {C.BG_CANVAS};
        border-color: {C.TEXT_FAINT};
    }}
    QPushButton:pressed {{
        background-color: {C.BORDER_SUBTLE};
    }}
"""

STYLE_TREE = f"""
    QTreeWidget {{
        background-color: {C.BG_CARD};
        border: 1px solid {C.BORDER_STRONG};
        border-radius: 6px;
        padding: 4px;
        font-size: 11px;
        color: {C.TEXT_STRONG};
    }}
    QTreeWidget::item {{
        padding: 3px 2px;
        border-radius: 3px;
    }}
    QTreeWidget::item:hover {{
        background-color: {C.TEAL_SURFACE};
    }}
    QTreeWidget::item:selected {{
        background-color: {C.TEAL_SOFT};
        color: #134E4A;
    }}
"""


def natural_ai_sort_key(ai_str: str) -> list:
    """Natural numerical sort key for 3GPP Agenda Item numbers (e.g. 5.1 before 5.10)."""
    parts = []
    for chunk in re.split(r'(\d+)', str(ai_str).strip()):
        if chunk.isdigit():
            parts.append(int(chunk))
        else:
            parts.append(chunk.lower())
    return parts


class PreMeetingExportDialog(QDialog):
    """
    Configuration dialog to customize and trigger pre-meeting Markdown briefing dossiers.
    Supports scoping by Followed/Starred Work Items, filtered view, whole meeting,
    or custom Work Item & Agenda Item selection (with dual WI->AI and AI->WI tree views),
    with live scope estimation and per-release/per-WI directory organization.
    """

    def __init__(
        self,
        parent,
        meeting_dir: Path,
        all_rows: list,
        visible_rows: list,
        agenda_map: dict,
        mtg_info: dict,
        followed_tokens: set,
        starred_tokens: set,
    ):
        super().__init__(parent)
        self.meeting_dir = Path(meeting_dir)
        self.all_rows = all_rows or []
        self.visible_rows = visible_rows or []
        self.agenda_map = agenda_map or {}
        self.mtg_info = mtg_info or {}
        self.followed_tokens = set(followed_tokens or [])
        self.starred_tokens = set(starred_tokens or [])
        self.tracked_tokens = self.followed_tokens | self.starred_tokens

        self._updating_tree = False
        self.selected_wi_ais: Set[Tuple[str, str]] = set()

        self.setWindowTitle("🧭 Pre-Meeting Briefing & Markdown Dossier Export")
        self.resize(980, 680)
        self.setMinimumSize(900, 580)
        self.setStyleSheet(f"""
            QDialog {{ background-color: {C.BG_SURFACE}; }}
            QGroupBox {{
                font-weight: bold;
                color: #134E4A;
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 6px;
                margin-top: 10px;
                padding-top: 14px;
                background-color: {C.BG_CARD};
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 6px;
                background-color: {C.BG_CARD};
            }}
            QLabel {{ color: {C.TEXT_BODY}; }}
            QRadioButton, QCheckBox {{
                color: {C.TEXT_STRONG};
                padding: 2px;
            }}
        """)

        # Extract available meeting Work Items, Agenda Items, companies, and lineage
        self._analyze_meeting_contributions()

        # Pre-populate custom selection with Followed & Starred WIs automatically
        self._initialize_default_custom_selection()

        self._setup_ui()

        # Update initial tree check states and live estimation banner
        self._sync_tree_from_selection(self.tree_wi)
        self._sync_tree_from_selection(self.tree_ai)
        self._expand_checked_parents(self.tree_wi)
        self._update_estimates()

    def _extract_row_wis(self, row: dict) -> List[str]:
        """Resolves canonical Work Items associated with a row."""
        raw_wis = TDocsTableModel._split_related_wis(row.get("Related WIs", ""))
        wis = []
        for rw in raw_wis:
            clean = rw.strip()
            if clean and clean.lower() not in {"(no wi)", "none", "nan", "-"}:
                if clean not in wis:
                    wis.append(clean)
        if not wis:
            ai_num = str(row.get("Agenda Item", "")).strip()
            if ai_num and self.agenda_map:
                item = self.agenda_map.get(ai_num)
                if item and getattr(item, "acronym", None):
                    acr = item.acronym.strip()
                    if acr and acr not in wis:
                        wis.append(acr)
            if not wis:
                desc = str(row.get("Agenda item description", "")).strip()
                if desc:
                    cand = AgendaItem._extract_acronym(desc)
                    if cand and cand not in wis:
                        wis.append(cand.strip())
        if not wis:
            wis = ["General"]
        return wis

    def _analyze_meeting_contributions(self):
        """Builds indices of Work Items, Agenda Items, TDoc distributions, and contributor companies."""
        wis_set = set()
        ais_set = set()
        companies_set = set()
        self.ai_descriptions: Dict[str, str] = {}
        self.wi_to_ais: Dict[str, Dict[str, int]] = {}
        self.ai_to_wis: Dict[str, Dict[str, int]] = {}
        self.wi_counts: Dict[str, int] = {}
        self.ai_counts: Dict[str, int] = {}

        # Precompute revision lineage to support live estimation of deduplication
        tdoc_map = {str(r.get("TDoc", "")).strip().upper(): r for r in self.all_rows if r.get("TDoc")}
        forward_edges: Dict[str, Set[str]] = {}
        for tdoc_id, r in tdoc_map.items():
            rev_to = str(r.get("Revised to", "")).strip()
            if rev_to and rev_to.lower() != "none":
                for cand in re.findall(r'[A-Za-z0-9]+-\d+(?:r\d+[a-zA-Z]?)?', rev_to):
                    cand_upper = cand.upper()
                    if cand_upper in tdoc_map and cand_upper != tdoc_id:
                        forward_edges.setdefault(tdoc_id, set()).add(cand_upper)

            rev_of = str(r.get("Is revision of", "")).strip()
            if rev_of and rev_of.lower() != "none":
                for cand in re.findall(r'[A-Za-z0-9]+-\d+(?:r\d+[a-zA-Z]?)?', rev_of):
                    cand_upper = cand.upper()
                    if cand_upper in tdoc_map and cand_upper != tdoc_id:
                        forward_edges.setdefault(cand_upper, set()).add(tdoc_id)

            base_match = re.search(r'^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$', tdoc_id, re.IGNORECASE)
            if base_match:
                base_tdoc = base_match.group(1).upper()
                if base_tdoc in tdoc_map and base_tdoc != tdoc_id:
                    forward_edges.setdefault(base_tdoc, set()).add(tdoc_id)

        self.superseded_tdocs: Set[str] = set()
        for tdoc_id, r in tdoc_map.items():
            st = str(r.get("TDoc Status", "")).strip().lower()
            if "withdrawn" in st:
                continue
            curr = tdoc_id
            visited = {curr}
            while True:
                children = forward_edges.get(curr, set())
                active_children = [c for c in children if c not in visited and "withdrawn" not in str(tdoc_map[c].get("TDoc Status", "")).strip().lower()]
                if not active_children:
                    break
                active_children.sort()
                next_node = active_children[-1]
                visited.add(next_node)
                curr = next_node
            if curr != tdoc_id:
                self.superseded_tdocs.add(tdoc_id)

        for r in self.all_rows:
            ai_num = str(r.get("Agenda Item", "")).strip() or "Unassigned"
            ais_set.add(ai_num)

            # Store best description for AI
            if ai_num not in self.ai_descriptions:
                desc = ""
                if ai_num in self.agenda_map:
                    item = self.agenda_map[ai_num]
                    desc = getattr(item, "description", "") or ""
                if not desc:
                    desc = str(r.get("Agenda item description", "")).strip()
                self.ai_descriptions[ai_num] = desc

            # Collect companies
            companies = r.get("_All_Sanitized_Companies", r.get("_Sanitized_Companies", []))
            if not companies:
                confirmed, tentative = CompanySanitizer.parse_contributors_detailed(r.get("Source", ""))
                companies = list(confirmed) or ["Other"]
            for comp in companies:
                if comp and comp != "Other":
                    companies_set.add(comp)

            row_wis = self._extract_row_wis(r)
            for w in row_wis:
                if w != "General":
                    wis_set.add(w)

                if w not in self.wi_to_ais:
                    self.wi_to_ais[w] = {}
                self.wi_to_ais[w][ai_num] = self.wi_to_ais[w].get(ai_num, 0) + 1
                self.wi_counts[w] = self.wi_counts.get(w, 0) + 1

                if ai_num not in self.ai_to_wis:
                    self.ai_to_wis[ai_num] = {}
                self.ai_to_wis[ai_num][w] = self.ai_to_wis[ai_num].get(w, 0) + 1
                self.ai_counts[ai_num] = self.ai_counts.get(ai_num, 0) + 1

        if "General" in self.wi_to_ais:
            wis_set.add("General")

        self.meeting_wis = sorted(list(wis_set), key=lambda x: (x == "General", x.lower()))
        self.meeting_ais = sorted(list(ais_set), key=natural_ai_sort_key)
        self.meeting_companies = sorted(list(companies_set), key=lambda x: x.lower())

        self.followed_wis = [w for w in self.meeting_wis if (TDocsTableModel._extract_base_tokens(w) & self.followed_tokens)]
        self.starred_wis = [w for w in self.meeting_wis if (TDocsTableModel._extract_base_tokens(w) & self.starred_tokens)]
        self.tracked_wis = [w for w in self.meeting_wis if (TDocsTableModel._extract_base_tokens(w) & self.tracked_tokens)]

    def _initialize_default_custom_selection(self):
        """Pre-checks Followed & Starred WIs (and all their child AIs) automatically."""
        self.selected_wi_ais.clear()
        has_tracked = bool(self.tracked_wis)

        for wi in self.meeting_wis:
            wi_tokens = TDocsTableModel._extract_base_tokens(wi)
            is_tracked = bool(wi_tokens & self.tracked_tokens)
            if is_tracked or (not has_tracked):
                for ai in self.wi_to_ais.get(wi, {}):
                    self.selected_wi_ais.add((wi, ai))

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
        lbl_sub.setStyleSheet("font-size: 11px; color: #115E59; border: none;")
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
        comp_lbl.setStyleSheet("font-size: 11px; font-weight: bold; color: #134E4A;")
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
                background-color: #0D9488;
                color: {C.BG_CARD};
                font-weight: bold;
                border-radius: 4px;
                padding: 6px 16px;
                border: 1px solid {C.TEAL_DARK};
            }}
            QPushButton:hover {{
                background-color: {C.TEAL_DARK};
            }}
            QPushButton:pressed {{
                background-color: #115E59;
            }}
        """)
        btn_start.setDefault(True)
        btn_start.clicked.connect(self._on_start_clicked)
        btn_layout.addWidget(btn_start)

        main_layout.addLayout(btn_layout)

        self._sync_scope_ui_state()

        self.tree_wi.itemChanged.connect(lambda item, col: self._on_tree_item_changed(self.tree_wi, item, col))
        self.tree_ai.itemChanged.connect(lambda item, col: self._on_tree_item_changed(self.tree_ai, item, col))

    # -------------------------------------------------------------------------
    # Tree Population & Synchronization
    # -------------------------------------------------------------------------

    def _populate_tree_wi(self):
        """Populates Page 0: Work Item -> Agenda Items."""
        self.tree_wi.blockSignals(True)
        try:
            self.tree_wi.clear()
            for wi in self.meeting_wis:
                ai_dict = self.wi_to_ais.get(wi, {})
                n_ais = len(ai_dict)
                total_docs = sum(ai_dict.values())

                parent_item = QTreeWidgetItem(self.tree_wi)
                parent_item.setText(0, f"{wi}  ({n_ais} AIs, {total_docs} TDocs)")
                parent_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsUserCheckable)
                parent_item.setCheckState(0, Qt.Unchecked)
                parent_item.setData(0, Qt.UserRole, wi)

                sorted_ais = sorted(ai_dict.keys(), key=natural_ai_sort_key)
                for ai in sorted_ais:
                    docs_count = ai_dict[ai]
                    desc = self.ai_descriptions.get(ai, "")
                    desc_str = f": {desc}" if desc else ""
                    child_label = f"AI {ai}{desc_str} ({docs_count} docs)"

                    child_item = QTreeWidgetItem(parent_item)
                    child_item.setText(0, child_label)
                    child_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsUserCheckable)
                    child_item.setCheckState(0, Qt.Unchecked)
                    child_item.setData(0, Qt.UserRole, (wi, ai))
        finally:
            self.tree_wi.blockSignals(False)

    def _populate_tree_ai(self):
        """Populates Page 1: Agenda Item -> Work Items."""
        self.tree_ai.blockSignals(True)
        try:
            self.tree_ai.clear()
            for ai in self.meeting_ais:
                wi_dict = self.ai_to_wis.get(ai, {})
                n_wis = len(wi_dict)
                total_docs = sum(wi_dict.values())
                desc = self.ai_descriptions.get(ai, "")
                desc_str = f" · {desc}" if desc else ""

                parent_item = QTreeWidgetItem(self.tree_ai)
                parent_item.setText(0, f"AI {ai}{desc_str}  ({n_wis} WIs, {total_docs} TDocs)")
                parent_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsUserCheckable)
                parent_item.setCheckState(0, Qt.Unchecked)
                parent_item.setData(0, Qt.UserRole, ai)

                sorted_wis = sorted(wi_dict.keys(), key=lambda x: (x == "General", x.lower()))
                for wi in sorted_wis:
                    docs_count = wi_dict[wi]
                    child_label = f"{wi} ({docs_count} docs)"

                    child_item = QTreeWidgetItem(parent_item)
                    child_item.setText(0, child_label)
                    child_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsUserCheckable)
                    child_item.setCheckState(0, Qt.Unchecked)
                    child_item.setData(0, Qt.UserRole, (wi, ai))
        finally:
            self.tree_ai.blockSignals(False)

    def _on_tree_item_changed(self, tree: QTreeWidget, item: QTreeWidgetItem, column: int):
        if self._updating_tree:
            return

        if not self.rb_custom.isChecked():
            self.rb_custom.setChecked(True)

        self._updating_tree = True
        try:
            is_parent = (item.parent() is None)
            new_state = item.checkState(0)

            if is_parent:
                target_state = Qt.Checked if new_state == Qt.Checked else Qt.Unchecked
                item.setCheckState(0, target_state)
                for i in range(item.childCount()):
                    child = item.child(i)
                    child.setCheckState(0, target_state)
                    pair = child.data(0, Qt.UserRole)
                    if isinstance(pair, tuple):
                        if target_state == Qt.Checked:
                            self.selected_wi_ais.add(pair)
                        else:
                            self.selected_wi_ais.discard(pair)
            else:
                pair = item.data(0, Qt.UserRole)
                if isinstance(pair, tuple):
                    if new_state == Qt.Checked:
                        self.selected_wi_ais.add(pair)
                    else:
                        self.selected_wi_ais.discard(pair)
                self._update_parent_check_state(item.parent())
        finally:
            self._updating_tree = False

        self._update_estimates()

    def _update_parent_check_state(self, parent: QTreeWidgetItem):
        if not parent or parent.childCount() == 0:
            return
        n = parent.childCount()
        n_checked = sum(1 for i in range(n) if parent.child(i).checkState(0) == Qt.Checked)
        n_partial = sum(1 for i in range(n) if parent.child(i).checkState(0) == Qt.PartiallyChecked)

        if n_checked == n:
            parent.setCheckState(0, Qt.Checked)
        elif n_checked == 0 and n_partial == 0:
            parent.setCheckState(0, Qt.Unchecked)
        else:
            parent.setCheckState(0, Qt.PartiallyChecked)

    def _sync_tree_from_selection(self, tree: QTreeWidget):
        tree.blockSignals(True)
        self._updating_tree = True
        try:
            root = tree.invisibleRootItem()
            for i in range(root.childCount()):
                parent = root.child(i)
                for j in range(parent.childCount()):
                    child = parent.child(j)
                    pair = child.data(0, Qt.UserRole)
                    if isinstance(pair, tuple):
                        child.setCheckState(0, Qt.Checked if pair in self.selected_wi_ais else Qt.Unchecked)
                self._update_parent_check_state(parent)
        finally:
            self._updating_tree = False
            tree.blockSignals(False)

    def _expand_checked_parents(self, tree: QTreeWidget):
        root = tree.invisibleRootItem()
        for i in range(root.childCount()):
            parent = root.child(i)
            parent.setExpanded(parent.checkState(0) in (Qt.Checked, Qt.PartiallyChecked))

    def _switch_custom_view(self, mode: str):
        if mode == "wi":
            self.btn_view_wi.setStyleSheet(STYLE_TOGGLE_ACTIVE)
            self.btn_view_ai.setStyleSheet(STYLE_TOGGLE_INACTIVE)
            self._sync_tree_from_selection(self.tree_wi)
            self.tree_stack.setCurrentWidget(self.tree_wi)
            self._filter_tree(self.tree_wi, self.custom_search.text())
        else:
            self.btn_view_wi.setStyleSheet(STYLE_TOGGLE_INACTIVE)
            self.btn_view_ai.setStyleSheet(STYLE_TOGGLE_ACTIVE)
            self._sync_tree_from_selection(self.tree_ai)
            self.tree_stack.setCurrentWidget(self.tree_ai)
            self._filter_tree(self.tree_ai, self.custom_search.text())

    def _filter_custom_tree(self, text: str):
        active_tree = self.tree_stack.currentWidget()
        if isinstance(active_tree, QTreeWidget):
            self._filter_tree(active_tree, text)

    def _filter_tree(self, tree: QTreeWidget, query: str):
        q = query.strip().lower()
        root = tree.invisibleRootItem()
        for i in range(root.childCount()):
            parent = root.child(i)
            parent_matches = bool(q and q in parent.text(0).lower())
            any_child_matched = False

            for j in range(parent.childCount()):
                child = parent.child(j)
                child_matches = bool(not q or q in child.text(0).lower() or parent_matches)
                child.setHidden(not child_matches)
                if child_matches:
                    any_child_matched = True

            if not q:
                parent.setHidden(False)
            else:
                parent.setHidden(not (parent_matches or any_child_matched))
                if any_child_matched:
                    parent.setExpanded(True)

    def _expand_active_tree(self):
        active_tree = self.tree_stack.currentWidget()
        if isinstance(active_tree, QTreeWidget):
            active_tree.expandAll()

    def _collapse_active_tree(self):
        active_tree = self.tree_stack.currentWidget()
        if isinstance(active_tree, QTreeWidget):
            active_tree.collapseAll()

    def _select_all_custom(self):
        if not self.rb_custom.isChecked():
            self.rb_custom.setChecked(True)
        for wi in self.meeting_wis:
            for ai in self.wi_to_ais.get(wi, {}):
                self.selected_wi_ais.add((wi, ai))
        self._sync_tree_from_selection(self.tree_wi)
        self._sync_tree_from_selection(self.tree_ai)
        self._update_estimates()

    def _clear_all_custom(self):
        if not self.rb_custom.isChecked():
            self.rb_custom.setChecked(True)
        self.selected_wi_ais.clear()
        self._sync_tree_from_selection(self.tree_wi)
        self._sync_tree_from_selection(self.tree_ai)
        self._update_estimates()

    # -------------------------------------------------------------------------
    # Live Scope Estimation
    # -------------------------------------------------------------------------

    def _sync_scope_ui_state(self):
        is_custom = self.rb_custom.isChecked()
        if is_custom:
            self.lbl_custom_hint.setText("⚡ <b>Custom Selection:</b> Choose specific Work Items and Agenda Items to export:")
            self.lbl_custom_hint.setStyleSheet(f"font-size: 11px; color: {C.TEAL_DARK};")
        else:
            self.lbl_custom_hint.setText("💡 <i>Check any item below or select 'Custom Selection' to customize scope:</i>")
            self.lbl_custom_hint.setStyleSheet(f"font-size: 11px; color: {C.TEXT_MUTED};")

    def _on_scope_radio_toggled(self, checked: bool):
        self._sync_scope_ui_state()
        if self.rb_custom.isChecked():
            active_tree = self.tree_stack.currentWidget()
            if isinstance(active_tree, QTreeWidget):
                self._sync_tree_from_selection(active_tree)
        self._update_estimates()

    def _update_estimates(self):
        scope_mode = self.get_scope_mode()
        include_ls_in = self.chk_ls_in.isChecked()
        include_ls_out = self.chk_ls_out.isChecked()
        include_pcr = self.chk_pcr.isChecked()
        include_cr = self.chk_cr.isChecked()
        include_disc = self.chk_disc.isChecked()
        include_general_ls = self.chk_general_ls.isChecked()
        dedup_revs = self.chk_dedup_revs.isChecked()
        include_reserved = self.chk_include_reserved.isChecked()

        selected_comp = self.combo_company.currentData()
        company_filter = {selected_comp.lower()} if selected_comp else set()

        matching_tdocs = set()
        superseded_hidden_count = 0
        scope_counts = {"LS_IN": 0, "LS_OUT": 0, "pCR": 0, "CR": 0, "DISC": 0}
        general_ls_count = 0
        matching_wis = set()
        matching_ais = set()
        matching_companies = set()

        visible_ids = {str(r.get("TDoc", "")).strip().upper() for r in self.visible_rows if r.get("TDoc")}

        for row in self.all_rows:
            tdoc_id = str(row.get("TDoc", "")).strip().upper()
            if not tdoc_id:
                continue

            status = str(row.get("TDoc Status", "")).strip().lower()
            if "withdrawn" in status:
                continue

            if dedup_revs and tdoc_id in self.superseded_tdocs:
                superseded_hidden_count += 1
                continue

            is_reserved = any(res in status for res in ["reserved", "not uploaded", "not yet uploaded"])
            if is_reserved and not include_reserved:
                continue

            doc_type = str(row.get("Type", "")).strip().lower()

            category = None
            if "ls in" in doc_type or doc_type == "ls":
                category = "LS_IN"
            elif "ls out" in doc_type or "ls_out" in doc_type:
                category = "LS_OUT"
            elif "pcr" in doc_type or "p-cr" in doc_type:
                category = "pCR"
            elif "cr" in doc_type:
                category = "CR"
            elif any(d in doc_type for d in ["disc", "discussion", "report", "wid", "sid", "other", "draft", "evaluation"]):
                category = "DISC"

            if not category:
                continue

            cat_enabled = (
                (category == "LS_IN" and include_ls_in) or
                (category == "LS_OUT" and include_ls_out) or
                (category == "pCR" and include_pcr) or
                (category == "CR" and include_cr) or
                (category == "DISC" and include_disc)
            )

            if scope_mode == "visible" and tdoc_id not in visible_ids:
                continue

            # Check company filter
            row_comps = row.get("_All_Sanitized_Companies", row.get("_Sanitized_Companies", []))
            row_comps_lower = {str(c).strip().lower() for c in row_comps}
            if company_filter and not (row_comps_lower & company_filter):
                continue

            row_wis = self._extract_row_wis(row)
            ai_num = str(row.get("Agenda Item", "")).strip() or "Unassigned"

            matches_scope = False
            matched_wis_for_row = []

            for wi in row_wis:
                if wi == "General":
                    if scope_mode in ["whole", "visible"] or ((category in ("LS_IN", "LS_OUT")) and include_general_ls):
                        matches_scope = True
                        matched_wis_for_row.append(wi)
                    elif scope_mode == "custom" and (
                        ("General", ai_num) in self.selected_wi_ais
                        or ("General", "Unassigned") in self.selected_wi_ais
                        or (include_general_ls and category in ("LS_IN", "LS_OUT"))
                    ):
                        matches_scope = True
                        matched_wis_for_row.append(wi)
                    continue

                wi_tokens = TDocsTableModel._extract_base_tokens(wi)
                if scope_mode == "followed":
                    if wi_tokens & self.followed_tokens:
                        matches_scope = True
                        matched_wis_for_row.append(wi)
                elif scope_mode == "starred":
                    if wi_tokens & self.starred_tokens:
                        matches_scope = True
                        matched_wis_for_row.append(wi)
                elif scope_mode == "tracked":
                    if wi_tokens & self.tracked_tokens:
                        matches_scope = True
                        matched_wis_for_row.append(wi)
                elif scope_mode == "custom":
                    if (wi, ai_num) in self.selected_wi_ais:
                        matches_scope = True
                        matched_wis_for_row.append(wi)
                    else:
                        for sel_wi, sel_ai in self.selected_wi_ais:
                            if sel_ai == ai_num and (TDocsTableModel._extract_base_tokens(sel_wi) & wi_tokens):
                                matches_scope = True
                                matched_wis_for_row.append(wi)
                                break
                else:  # whole
                    matches_scope = True
                    matched_wis_for_row.append(wi)

            if not matches_scope and (category in ("LS_IN", "LS_OUT")) and include_general_ls:
                matches_scope = True
                matched_wis_for_row.append("General")

            if matches_scope:
                scope_counts[category] += 1
                if "General" in matched_wis_for_row and category in ("LS_IN", "LS_OUT"):
                    general_ls_count += 1
                if cat_enabled:
                    matching_tdocs.add(tdoc_id)
                    for w in matched_wis_for_row:
                        matching_wis.add(w)
                    matching_ais.add(ai_num)
                    for c in row_comps:
                        if c and c != "Other":
                            matching_companies.add(c)

        # Update Section 2 checkbox labels with dynamic counts
        self.chk_ls_in.setText(f"📥 Incoming Liaison Statements (LS in) — {scope_counts['LS_IN']} docs")
        self.chk_ls_out.setText(f"📤 Proposed Outgoing LS Replies (LS out) — {scope_counts['LS_OUT']} docs")
        self.chk_pcr.setText(f"📝 Proposed Pseudo-CRs (pCR) — {scope_counts['pCR']} docs")
        self.chk_cr.setText(f"📄 Proposed Change Requests (CR) — {scope_counts['CR']} docs")
        self.chk_disc.setText(f"💬 Discussion, Evaluation && WID Papers (Disc) — {scope_counts['DISC']} docs")
        self.chk_general_ls.setText(f"📦 Include unassigned && general LS — {general_ls_count} docs")

        total = len(matching_tdocs)
        if total > 0:
            parts = []
            if self.chk_pcr.isChecked() and scope_counts['pCR']:
                parts.append(f"{scope_counts['pCR']} pCRs")
            if self.chk_cr.isChecked() and scope_counts['CR']:
                parts.append(f"{scope_counts['CR']} CRs")
            if self.chk_disc.isChecked() and scope_counts['DISC']:
                parts.append(f"{scope_counts['DISC']} Disc")
            if self.chk_ls_in.isChecked() and scope_counts['LS_IN']:
                parts.append(f"{scope_counts['LS_IN']} LS in")
            if self.chk_ls_out.isChecked() and scope_counts['LS_OUT']:
                parts.append(f"{scope_counts['LS_OUT']} LS out")
            summary_str = f" ({', '.join(parts)})" if parts else ""

            dedup_note = f" [{superseded_hidden_count} older revisions hidden]" if (dedup_revs and superseded_hidden_count > 0) else ""
            n_wis = len(matching_wis)
            n_ais = len(matching_ais)
            n_comps = len(matching_companies)
            self.lbl_estimate.setText(
                f"📊 <b>Estimated Pre-Meeting Scope:</b> {total} Active Proposals across {n_wis} Work Items, "
                f"{n_ais} Agenda Items, and {n_comps} Companies{summary_str}{dedup_note}"
            )
            self.lbl_estimate.setStyleSheet(f"""
                QLabel {{
                    background-color: {C.SUCCESS_SOFT};
                    border: 1px solid {C.SUCCESS_BORDER};
                    border-radius: 6px;
                    padding: 6px 10px;
                    font-size: 11px;
                    color: {C.SUCCESS_DARK};
                    font-weight: 500;
                }}
            """)
        else:
            self.lbl_estimate.setText(
                "⚠️ <b>Estimated Scope:</b> 0 TDocs match current criteria. Adjust scope, filters, or document types."
            )
            self.lbl_estimate.setStyleSheet(f"""
                QLabel {{
                    background-color: {C.WARNING_SOFT};
                    border: 1px solid {C.WARNING_BORDER};
                    border-radius: 6px;
                    padding: 6px 10px;
                    font-size: 11px;
                    color: {C.WARNING};
                    font-weight: 500;
                }}
            """)

    def _browse_destination(self):
        d = QFileDialog.getExistingDirectory(self, "Select Export Directory", str(self.meeting_dir / "Export"))
        if d:
            self.dest_edit.setText(d)

    def _update_destination_status(self):
        text = self.dest_edit.text().strip()
        if not text:
            self.lbl_dest_status.setVisible(False)
            return

        p = Path(text)
        if p.is_dir() and ((p / "index.md").is_file() or (p / ".export_manifest.json").is_file()):
            from modules.meetings.core.exporters.export_manifest import load_manifest
            m = load_manifest(p)
            if m:
                n_docs = m.get_total_tdocs()
                n_updates = len(m.sessions) - 1
                up_str = f", {n_updates} previous update(s)" if n_updates > 0 else ""
                msg = (
                    f"🔄 <b>Existing Export Detected:</b> Incremental Append Mode active "
                    f"({n_docs} existing TDocs{up_str} will be preserved and merged into the index)."
                )
            else:
                md_count = len([f for f in p.glob("**/*.md") if f.name.lower() != "index.md"])
                msg = (
                    f"🔄 <b>Existing Export Detected:</b> Incremental Append Mode active "
                    f"({md_count} existing TDocs will be preserved and merged into the index)."
                )
            self.lbl_dest_status.setText(msg)
            self.lbl_dest_status.setStyleSheet(f"""
                QLabel {{
                    background-color: {C.SUCCESS_SOFT};
                    border: 1px solid {C.SUCCESS_BORDER};
                    border-radius: 4px;
                    padding: 4px 8px;
                    font-size: 11px;
                    color: {C.SUCCESS_DARK};
                }}
            """)
            self.lbl_dest_status.setVisible(True)
        else:
            self.lbl_dest_status.setVisible(False)

    def _show_recent_exports_menu(self):
        export_root = self.meeting_dir / "Export"
        if not export_root.is_dir():
            QMessageBox.information(self, "No Recent Exports", "No existing export folders found in this meeting directory.")
            return

        candidates = []
        for item in export_root.iterdir():
            if item.is_dir() and ((item / "index.md").is_file() or (item / ".export_manifest.json").is_file()):
                try:
                    candidates.append((item.stat().st_mtime, item))
                except Exception:
                    pass

        if not candidates:
            QMessageBox.information(self, "No Recent Exports", "No previous export folders with index.md found in Export/.")
            return

        candidates.sort(key=lambda x: x[0], reverse=True)
        menu = QMenu(self)
        for mtime, folder in candidates[:10]:
            dt_str = datetime.datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M")
            action = menu.addAction(f"📁 {folder.name} ({dt_str})")
            action.triggered.connect(lambda checked=False, p=folder: self._select_recent_export(p))

        menu.exec_(self.btn_recent.mapToGlobal(self.btn_recent.rect().bottomLeft()))

    def _select_recent_export(self, folder: Path):
        self.dest_edit.setText(str(folder))

    def _on_start_clicked(self):
        if not any([
            self.chk_ls_in.isChecked(),
            self.chk_ls_out.isChecked(),
            self.chk_pcr.isChecked(),
            self.chk_cr.isChecked(),
            self.chk_disc.isChecked(),
        ]):
            QMessageBox.warning(self, "No Types Selected", "Please select at least one document category to export.")
            return

        if self.rb_custom.isChecked():
            if not self.selected_wi_ais:
                QMessageBox.warning(
                    self,
                    "No Items Selected",
                    "Please select at least one Work Item or Agenda Item from the custom selection tree."
                )
                return

        self.accept()

    def get_scope_mode(self) -> str:
        if self.rb_followed.isChecked():
            return "followed"
        if self.rb_whole.isChecked():
            return "whole"
        if self.rb_starred.isChecked():
            return "starred"
        if self.rb_tracked.isChecked():
            return "tracked"
        if self.rb_visible.isChecked():
            return "visible"
        if self.rb_custom.isChecked():
            return "custom"
        return "followed"

    def get_custom_wis(self) -> Set[str]:
        return {wi.lower() for wi, ai in self.selected_wi_ais}

    def get_custom_ais(self) -> Set[str]:
        return {ai for wi, ai in self.selected_wi_ais}

    def get_custom_wi_ais(self) -> Dict[str, Set[str]]:
        wi_map: Dict[str, Set[str]] = {}
        for wi, ai in self.selected_wi_ais:
            k = wi.lower()
            if k not in wi_map:
                wi_map[k] = set()
            wi_map[k].add(ai)
        return wi_map

    def get_company_filter(self) -> Optional[Set[str]]:
        c = self.combo_company.currentData()
        return {str(c).strip().lower()} if c else None

    def get_export_config(self) -> dict:
        dest_text = self.dest_edit.text().strip()
        output_dir = Path(dest_text) if dest_text else None

        return {
            "scope_mode": self.get_scope_mode(),
            "custom_wis": self.get_custom_wis(),
            "custom_ais": self.get_custom_ais(),
            "custom_wi_ais": self.get_custom_wi_ais(),
            "include_ls_in": self.chk_ls_in.isChecked(),
            "include_ls_out": self.chk_ls_out.isChecked(),
            "include_pcr": self.chk_pcr.isChecked(),
            "include_cr": self.chk_cr.isChecked(),
            "include_disc": self.chk_disc.isChecked(),
            "include_general_ls": self.chk_general_ls.isChecked(),
            "deduplicate_revisions": self.chk_dedup_revs.isChecked(),
            "include_reserved_in_index": self.chk_include_reserved.isChecked(),
            "company_filter": self.get_company_filter(),
            "include_user_notes": self.chk_user_notes.isChecked(),
            "generate_local_indices": self.chk_local_indices.isChecked(),
            "download_missing": self.chk_download.isChecked(),
            "force_reconvert": self.chk_force.isChecked(),
            "compile_dossier": self.chk_compile_dossier.isChecked(),
            "auto_open": self.chk_auto_open.isChecked(),
            "output_dir": output_dir,
        }
