# --- File: src/modules/meetings/ui/post_meeting_export_dialog.py ---
import datetime
import re
from pathlib import Path
from typing import Dict, List, Set, Optional, Tuple, Any

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QGridLayout, QLabel, QGroupBox,
    QRadioButton, QCheckBox, QPushButton, QLineEdit, QFileDialog,
    QTreeWidget, QTreeWidgetItem, QStackedWidget, QFrame, QScrollArea,
    QWidget, QButtonGroup, QMessageBox, QSizePolicy
)

from core.ui.ui_components import BUTTON_STYLE_TOOLBAR_SECONDARY
from modules.meetings.core.agenda_manager import AgendaItem
from modules.meetings.ui.tdocs_models import TDocsTableModel


STYLE_TOGGLE_ACTIVE = """
    QPushButton {
        background-color: #2563EB;
        color: #FFFFFF;
        font-weight: bold;
        border: 1px solid #1D4ED8;
        border-radius: 4px;
        padding: 3px 8px;
        font-size: 11px;
    }
"""

STYLE_TOGGLE_INACTIVE = """
    QPushButton {
        background-color: #FFFFFF;
        color: #475569;
        border: 1px solid #CBD5E1;
        border-radius: 4px;
        padding: 3px 8px;
        font-size: 11px;
    }
    QPushButton:hover {
        background-color: #F1F5F9;
        color: #1E293B;
    }
"""

STYLE_COMPACT_BTN = """
    QPushButton {
        background-color: #F8FAFC;
        color: #1E293B;
        border: 1px solid #CBD5E1;
        border-radius: 4px;
        padding: 3px 8px;
        font-size: 11px;
        font-weight: 500;
    }
    QPushButton:hover {
        background-color: #F1F5F9;
        border-color: #94A3B8;
    }
    QPushButton:pressed {
        background-color: #E2E8F0;
    }
"""

STYLE_TREE = """
    QTreeWidget {
        background-color: #FFFFFF;
        border: 1px solid #CBD5E1;
        border-radius: 6px;
        padding: 4px;
        font-size: 11px;
        color: #1E293B;
    }
    QTreeWidget::item {
        padding: 3px 2px;
        border-radius: 3px;
    }
    QTreeWidget::item:hover {
        background-color: #F1F5F9;
    }
    QTreeWidget::item:selected {
        background-color: #E0E7FF;
        color: #1E293B;
    }
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


class PostMeetingExportDialog(QDialog):
    """
    Configuration dialog to customize and trigger post-meeting Markdown dossiers.
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

        self.setWindowTitle("📑 Post-Meeting Report & Markdown Dossier Export")
        self.resize(960, 650)
        self.setMinimumSize(880, 560)
        self.setStyleSheet("""
            QDialog { background-color: #F8FAFC; }
            QGroupBox {
                font-weight: bold;
                color: #1E293B;
                border: 1px solid #CBD5E1;
                border-radius: 6px;
                margin-top: 10px;
                padding-top: 14px;
                background-color: #FFFFFF;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 6px;
                background-color: #FFFFFF;
            }
            QLabel { color: #334155; }
            QRadioButton, QCheckBox {
                color: #1E293B;
                padding: 2px;
            }
        """)

        # Extract available meeting Work Items, Agenda Items, and their cross-correlations
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
        """Builds indices of Work Items, Agenda Items, and TDoc distributions."""
        wis_set = set()
        ais_set = set()
        self.ai_descriptions: Dict[str, str] = {}
        self.wi_to_ais: Dict[str, Dict[str, int]] = {}
        self.ai_to_wis: Dict[str, Dict[str, int]] = {}
        self.wi_counts: Dict[str, int] = {}
        self.ai_counts: Dict[str, int] = {}

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

            row_wis = self._extract_row_wis(r)
            for w in row_wis:
                if w != "General":
                    wis_set.add(w)

                # wi -> ai distribution
                if w not in self.wi_to_ais:
                    self.wi_to_ais[w] = {}
                self.wi_to_ais[w][ai_num] = self.wi_to_ais[w].get(ai_num, 0) + 1
                self.wi_counts[w] = self.wi_counts.get(w, 0) + 1

                # ai -> wi distribution
                if ai_num not in self.ai_to_wis:
                    self.ai_to_wis[ai_num] = {}
                self.ai_to_wis[ai_num][w] = self.ai_to_wis[ai_num].get(w, 0) + 1
                self.ai_counts[ai_num] = self.ai_counts.get(ai_num, 0) + 1

        # If General has contributions, include it in the WI catalog
        if "General" in self.wi_to_ais:
            wis_set.add("General")

        self.meeting_wis = sorted(list(wis_set), key=lambda x: (x == "General", x.lower()))
        self.meeting_ais = sorted(list(ais_set), key=natural_ai_sort_key)

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
        header_frame.setStyleSheet("""
            QFrame {
                background-color: #EFF6FF;
                border: 1px solid #BFDBFE;
                border-radius: 6px;
                padding: 6px;
            }
        """)
        header_layout = QVBoxLayout(header_frame)
        header_layout.setContentsMargins(8, 4, 8, 4)
        header_layout.setSpacing(1)

        wg = str(self.mtg_info.get("wg_name", "3GPP")).strip()
        mtg = str(self.mtg_info.get("meeting_number", "")).strip()
        title_text = f"<b>{wg} #{mtg} Post-Meeting Report</b>" if (wg or mtg) else "<b>Post-Meeting Report</b>"
        lbl_title = QLabel(title_text)
        lbl_title.setStyleSheet("font-size: 13px; color: #1E40AF; font-weight: bold; border: none;")
        lbl_sub = QLabel("Generates a structured, per-release and per-WI Markdown dossier with an executive summary cover page.")
        lbl_sub.setStyleSheet("font-size: 11px; color: #3B82F6; border: none;")
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

        # 2-column grid for radio buttons
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

        # Default selection: Followed if available, else Whole Meeting
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

        # --- Custom Selection Widget (Expandable Tree with Dual Views) ---
        self.custom_wi_widget = QWidget()
        self.custom_wi_widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        custom_layout = QVBoxLayout(self.custom_wi_widget)
        custom_layout.setContentsMargins(0, 4, 0, 0)
        custom_layout.setSpacing(6)

        # Dynamic Status / Hint Label
        self.lbl_custom_hint = QLabel()
        self.lbl_custom_hint.setStyleSheet("font-size: 11px;")
        custom_layout.addWidget(self.lbl_custom_hint)

        # Header bar: View switcher + Tree controls
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

        # Filter Line Edit
        self.custom_search = QLineEdit()
        self.custom_search.setPlaceholderText("Filter by Work Item, AI # (e.g. 5.1), or topic keyword...")
        self.custom_search.setClearButtonEnabled(True)
        self.custom_search.textChanged.connect(self._filter_custom_tree)
        custom_layout.addWidget(self.custom_search)

        # Stacked Trees: Page 0 = WI->AI, Page 1 = AI->WI
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
        body_layout.addWidget(scope_group, stretch=12)

        # --- Right Column: Sections 2 & 3 ---
        right_column = QVBoxLayout()
        right_column.setSpacing(8)

        # --- Group 2: Document Types & Statuses ---
        docs_group = QGroupBox("2. Target Document Types && Statuses")
        docs_layout = QVBoxLayout(docs_group)
        docs_layout.setSpacing(4)
        docs_layout.setContentsMargins(10, 12, 10, 8)

        self.chk_ls_in = QCheckBox("📥 Incoming Liaison Statements (LS in)")
        self.chk_ls_in.setChecked(True)
        self.chk_ls_in.stateChanged.connect(self._update_estimates)

        self.chk_pcr = QCheckBox("📝 Agreed / Approved Pseudo-CRs (pCR)")
        self.chk_pcr.setChecked(True)
        self.chk_pcr.stateChanged.connect(self._update_estimates)

        self.chk_cr = QCheckBox("📄 Agreed / Approved Change Requests (CR)")
        self.chk_cr.setChecked(True)
        self.chk_cr.stateChanged.connect(self._update_estimates)

        self.chk_ls_out = QCheckBox("📤 Agreed / Approved Outgoing Liaison Statements (LS out)")
        self.chk_ls_out.setChecked(True)
        self.chk_ls_out.stateChanged.connect(self._update_estimates)

        self.chk_disc = QCheckBox("💬 Agreed / Noted Discussion && Status Papers (Disc / WID)")
        self.chk_disc.setChecked(False)
        self.chk_disc.setToolTip(
            "Include discussion papers, study reports, and Work Item Descriptions (WID) noted or agreed by the meeting."
        )
        self.chk_disc.stateChanged.connect(self._update_estimates)

        self.chk_general_ls = QCheckBox("📦 Include unassigned && general LS in (without followed WI)")
        self.chk_general_ls.setChecked(True)
        self.chk_general_ls.setToolTip(
            "Include incoming Liaison Statements (LS in) for which no specific Work Item was assigned "
            "or whose topic is not followed, placing them under 'General/LS_IN/'."
        )
        self.chk_general_ls.stateChanged.connect(self._update_estimates)

        docs_layout.addWidget(self.chk_ls_in)
        docs_layout.addWidget(self.chk_pcr)
        docs_layout.addWidget(self.chk_cr)
        docs_layout.addWidget(self.chk_ls_out)
        docs_layout.addWidget(self.chk_disc)
        docs_layout.addWidget(self.chk_general_ls)

        right_column.addWidget(docs_group)

        # --- Group 3: Conversion & Output Options ---
        opts_group = QGroupBox("3. Organization && Conversion Options")
        opts_layout = QVBoxLayout(opts_group)
        opts_layout.setSpacing(4)
        opts_layout.setContentsMargins(10, 12, 10, 8)

        self.chk_local_indices = QCheckBox("Generate local index.md cover inside each Work Item directory")
        self.chk_local_indices.setChecked(True)

        self.chk_download = QCheckBox("Download missing TDoc ZIP packages automatically from FTP")
        self.chk_download.setChecked(True)

        self.chk_force = QCheckBox("Force reconvert (overwrite existing local markdown caches)")
        self.chk_force.setChecked(False)

        self.chk_compile_dossier = QCheckBox("Also compile consolidated Navigable Dossier (HTML && PDF)")
        self.chk_compile_dossier.setChecked(True)
        self.chk_compile_dossier.setToolTip(
            "Synthesizes an interactive HTML dossier and bookmarked vector PDF containing all exported deliverables."
        )

        self.chk_auto_open = QCheckBox("Open export directory in Windows Explorer upon completion")
        self.chk_auto_open.setChecked(True)

        opts_layout.addWidget(self.chk_local_indices)
        opts_layout.addWidget(self.chk_download)
        opts_layout.addWidget(self.chk_force)
        opts_layout.addWidget(self.chk_compile_dossier)
        opts_layout.addWidget(self.chk_auto_open)

        # Destination Folder row
        opts_layout.addWidget(QLabel("Destination Directory:"))
        dest_layout = QHBoxLayout()
        dest_layout.setSpacing(6)
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S")
        clean_wg = re.sub(r'[^\w\-]', '_', wg)
        clean_mtg = re.sub(r'[^\w\-]', '_', mtg)
        folder_name = f"PostMeeting_{clean_wg}_{clean_mtg}_{timestamp}" if (clean_wg or clean_mtg) else f"PostMeeting_{timestamp}"
        default_dir = self.meeting_dir / "Export" / folder_name
        self.dest_edit = QLineEdit(str(default_dir))
        self.dest_edit.setReadOnly(False)
        dest_layout.addWidget(self.dest_edit)

        btn_browse = QPushButton("📁 Browse...")
        btn_browse.setStyleSheet(STYLE_COMPACT_BTN)
        btn_browse.clicked.connect(self._browse_destination)
        dest_layout.addWidget(btn_browse)

        opts_layout.addLayout(dest_layout)

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

        btn_start = QPushButton("🚀 Start Export")
        btn_start.setStyleSheet("""
            QPushButton {
                background-color: #2563EB;
                color: #FFFFFF;
                font-weight: bold;
                border-radius: 4px;
                padding: 6px 16px;
                border: 1px solid #1D4ED8;
            }
            QPushButton:hover {
                background-color: #1D4ED8;
            }
            QPushButton:pressed {
                background-color: #1E40AF;
            }
        """)
        btn_start.setDefault(True)
        btn_start.clicked.connect(self._on_start_clicked)
        btn_layout.addWidget(btn_start)

        main_layout.addLayout(btn_layout)

        # Initial label hint state
        self._sync_scope_ui_state()

        # Connect tree itemChanged after all widgets and estimation banners exist
        self.tree_wi.itemChanged.connect(lambda item, col: self._on_tree_item_changed(self.tree_wi, item, col))
        self.tree_ai.itemChanged.connect(lambda item, col: self._on_tree_item_changed(self.tree_ai, item, col))

    # -------------------------------------------------------------------------
    # Tree Population & Management
    # -------------------------------------------------------------------------

    def _populate_tree_wi(self):
        """Populates Page 0: Work Item (parent) -> Agenda Items (children)."""
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
        """Populates Page 1: Agenda Item (parent) -> Work Items (children)."""
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
        """Synchronizes checks between parent and child nodes and updates the shared selection model."""
        if self._updating_tree:
            return

        # Auto-activate Custom Selection mode when user modifies selection
        if not self.rb_custom.isChecked():
            self.rb_custom.setChecked(True)

        self._updating_tree = True
        try:
            is_parent = (item.parent() is None)
            new_state = item.checkState(0)

            if is_parent:
                # Parent clicked: cascade state to all children
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
                # Child clicked: update selection model and recompute parent tri-state
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
        """Calculates Checked, PartiallyChecked, or Unchecked state for a parent node."""
        if not parent:
            return
        n = parent.childCount()
        if n == 0:
            return

        n_checked = 0
        n_partial = 0
        for i in range(n):
            st = parent.child(i).checkState(0)
            if st == Qt.Checked:
                n_checked += 1
            elif st == Qt.PartiallyChecked:
                n_partial += 1

        if n_checked == n:
            parent.setCheckState(0, Qt.Checked)
        elif n_checked == 0 and n_partial == 0:
            parent.setCheckState(0, Qt.Unchecked)
        else:
            parent.setCheckState(0, Qt.PartiallyChecked)

    def _sync_tree_from_selection(self, tree: QTreeWidget):
        """Refreshes all check states in a tree from self.selected_wi_ais."""
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
        """Expands top-level items that are checked or partially checked."""
        root = tree.invisibleRootItem()
        for i in range(root.childCount()):
            parent = root.child(i)
            if parent.checkState(0) in (Qt.Checked, Qt.PartiallyChecked):
                parent.setExpanded(True)
            else:
                parent.setExpanded(False)

    def _switch_custom_view(self, mode: str):
        """Switches between 'By Work Item' and 'By Agenda Item' views seamlessly."""
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
    # Live Scope & Document Counts Estimation
    # -------------------------------------------------------------------------

    def _sync_scope_ui_state(self):
        """Updates hint label and visual style when scope mode changes."""
        is_custom = self.rb_custom.isChecked()
        if is_custom:
            self.lbl_custom_hint.setText("⚡ <b>Custom Selection:</b> Choose specific Work Items and Agenda Items to export:")
            self.lbl_custom_hint.setStyleSheet("font-size: 11px; color: #1D4ED8;")
        else:
            self.lbl_custom_hint.setText("💡 <i>Check any item below or select 'Custom Selection' to customize scope:</i>")
            self.lbl_custom_hint.setStyleSheet("font-size: 11px; color: #64748B;")

    def _on_scope_radio_toggled(self, checked: bool):
        self._sync_scope_ui_state()
        if self.rb_custom.isChecked():
            active_tree = self.tree_stack.currentWidget()
            if isinstance(active_tree, QTreeWidget):
                self._sync_tree_from_selection(active_tree)
        self._update_estimates()

    def _update_estimates(self):
        """Calculates matching TDocs in real time and updates summary text and Section 2 counts."""
        scope_mode = self.get_scope_mode()
        include_ls_in = self.chk_ls_in.isChecked()
        include_pcr = self.chk_pcr.isChecked()
        include_cr = self.chk_cr.isChecked()
        include_ls_out = self.chk_ls_out.isChecked()
        include_disc = self.chk_disc.isChecked()
        include_general_ls = self.chk_general_ls.isChecked()

        matching_tdocs = set()
        scope_counts = {"LS_IN": 0, "pCR": 0, "CR": 0, "LS_OUT": 0, "DISC": 0}
        general_ls_count = 0
        matching_wis = set()
        matching_ais = set()

        visible_ids = {str(r.get("TDoc", "")).strip().upper() for r in self.visible_rows if r.get("TDoc")}

        for row in self.all_rows:
            tdoc_id = str(row.get("TDoc", "")).strip().upper()
            if not tdoc_id:
                continue

            doc_type = str(row.get("Type", "")).strip().lower()
            status = str(row.get("TDoc Status", "")).strip().lower()
            if "withdrawn" in status:
                continue

            is_not_agreed = any(neg in status for neg in ["not agreed", "not approved", "rejected", "postponed", "withdrawn"])
            is_agreed = (not is_not_agreed) and any(s in status for s in ["agreed", "approved"])

            category = None
            if "ls in" in doc_type or doc_type == "ls":
                category = "LS_IN"
            elif "ls out" in doc_type or "ls_out" in doc_type:
                if is_agreed:
                    category = "LS_OUT"
            elif "pcr" in doc_type or "p-cr" in doc_type:
                if is_agreed:
                    category = "pCR"
            elif "cr" in doc_type:
                if is_agreed:
                    category = "CR"
            elif any(d in doc_type for d in ["disc", "discussion", "report", "wid", "sid"]):
                is_not_final = any(neg in status for neg in ["not agreed", "not approved", "rejected", "postponed", "withdrawn", "revised", "not discussed"])
                if (not is_not_final) and (status == "" or any(s in status for s in ["noted", "agreed", "approved"])):
                    category = "DISC"

            if not category:
                continue

            cat_enabled = (
                (category == "LS_IN" and include_ls_in) or
                (category == "pCR" and include_pcr) or
                (category == "CR" and include_cr) or
                (category == "LS_OUT" and include_ls_out) or
                (category == "DISC" and include_disc)
            )

            if scope_mode == "visible" and tdoc_id not in visible_ids:
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

        # Update Section 2 checkbox labels with dynamic document counts
        self.chk_ls_in.setText(f"📥 Incoming Liaison Statements (LS in) — {scope_counts['LS_IN']} docs")
        self.chk_pcr.setText(f"📝 Agreed / Approved Pseudo-CRs (pCR) — {scope_counts['pCR']} docs")
        self.chk_cr.setText(f"📄 Agreed / Approved Change Requests (CR) — {scope_counts['CR']} docs")
        self.chk_ls_out.setText(f"📤 Agreed / Approved Outgoing Liaison Statements (LS out) — {scope_counts['LS_OUT']} docs")
        self.chk_disc.setText(f"💬 Agreed / Noted Discussion && Status Papers (Disc / WID) — {scope_counts['DISC']} docs")
        self.chk_general_ls.setText(f"📦 Include unassigned && general LS in — {general_ls_count} docs")

        # Update Estimated Scope status banner
        total = len(matching_tdocs)
        if total > 0:
            parts = []
            if self.chk_cr.isChecked() and scope_counts['CR']:
                parts.append(f"{scope_counts['CR']} CRs")
            if self.chk_pcr.isChecked() and scope_counts['pCR']:
                parts.append(f"{scope_counts['pCR']} pCRs")
            if self.chk_ls_in.isChecked() and scope_counts['LS_IN']:
                parts.append(f"{scope_counts['LS_IN']} LS in")
            if self.chk_ls_out.isChecked() and scope_counts['LS_OUT']:
                parts.append(f"{scope_counts['LS_OUT']} LS out")
            if self.chk_disc.isChecked() and scope_counts['DISC']:
                parts.append(f"{scope_counts['DISC']} Disc")
            summary_str = f" ({', '.join(parts)})" if parts else ""
            n_wis = len(matching_wis)
            n_ais = len(matching_ais)
            self.lbl_estimate.setText(
                f"📊 <b>Estimated Scope:</b> {total} TDocs across {n_wis} Work Items and {n_ais} Agenda Items{summary_str}"
            )
            self.lbl_estimate.setStyleSheet("""
                QLabel {
                    background-color: #F0FDF4;
                    border: 1px solid #BBF7D0;
                    border-radius: 6px;
                    padding: 6px 10px;
                    font-size: 11px;
                    color: #166534;
                    font-weight: 500;
                }
            """)
        else:
            self.lbl_estimate.setText(
                "⚠️ <b>Estimated Scope:</b> 0 TDocs match current criteria. Adjust scope or document types."
            )
            self.lbl_estimate.setStyleSheet("""
                QLabel {
                    background-color: #FFFBEB;
                    border: 1px solid #FDE68A;
                    border-radius: 6px;
                    padding: 6px 10px;
                    font-size: 11px;
                    color: #B45309;
                    font-weight: 500;
                }
            """)

    # -------------------------------------------------------------------------
    # Dialog Output & Control
    # -------------------------------------------------------------------------

    def _browse_destination(self):
        d = QFileDialog.getExistingDirectory(self, "Select Export Directory", str(self.meeting_dir / "Export"))
        if d:
            self.dest_edit.setText(d)

    def _on_start_clicked(self):
        # Validate that at least one document type is selected
        if not any([
            self.chk_ls_in.isChecked(),
            self.chk_pcr.isChecked(),
            self.chk_cr.isChecked(),
            self.chk_ls_out.isChecked(),
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
        """Set of selected lowercase Work Item names (for backwards compatibility)."""
        return {wi.lower() for wi, ai in self.selected_wi_ais}

    def get_custom_ais(self) -> Set[str]:
        """Set of selected Agenda Item numbers."""
        return {ai for wi, ai in self.selected_wi_ais}

    def get_custom_wi_ais(self) -> Dict[str, Set[str]]:
        """Map of lowercase Work Item name to set of selected Agenda Item numbers."""
        wi_map: Dict[str, Set[str]] = {}
        for wi, ai in self.selected_wi_ais:
            k = wi.lower()
            if k not in wi_map:
                wi_map[k] = set()
            wi_map[k].add(ai)
        return wi_map

    def get_export_config(self) -> dict:
        dest_text = self.dest_edit.text().strip()
        output_dir = Path(dest_text) if dest_text else None

        return {
            "scope_mode": self.get_scope_mode(),
            "custom_wis": self.get_custom_wis(),
            "custom_ais": self.get_custom_ais(),
            "custom_wi_ais": self.get_custom_wi_ais(),
            "include_ls_in": self.chk_ls_in.isChecked(),
            "include_pcr": self.chk_pcr.isChecked(),
            "include_cr": self.chk_cr.isChecked(),
            "include_ls_out": self.chk_ls_out.isChecked(),
            "include_disc": self.chk_disc.isChecked(),
            "include_general_ls": self.chk_general_ls.isChecked(),
            "generate_local_indices": self.chk_local_indices.isChecked(),
            "download_missing": self.chk_download.isChecked(),
            "force_reconvert": self.chk_force.isChecked(),
            "compile_dossier": self.chk_compile_dossier.isChecked(),
            "auto_open": self.chk_auto_open.isChecked(),
            "output_dir": output_dir,
        }
