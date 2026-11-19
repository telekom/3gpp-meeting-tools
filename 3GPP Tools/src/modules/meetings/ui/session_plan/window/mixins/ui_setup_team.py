# --- File: src/modules/meetings/ui/session_plan/window/mixins/ui_setup_team.py ---
"""
Widget construction for the Team / Delegation view page.

Mixin of SessionPlanWindow: relies on attributes set in SessionPlanWindow.__init__.
"""
from PyQt5.QtCore import QSettings, Qt
from PyQt5.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QPushButton,
    QSizePolicy,
    QTableWidget,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)
from core.ui.flow_layout import FlowLayout
from core.ui.font_scale import register_scaled_row_height, scale_px
from core.ui.theme import C


class TeamSetupMixin:
    """Widget construction for the Team / Delegation view page."""

    # --- View 5: Team Allocations (HoD Dispatcher & Topic Ownership) ---

    def _setup_team_view(self):
        page = QWidget()
        page_layout = QVBoxLayout(page)
        page_layout.setContentsMargins(0, 0, 0, 0)
        page_layout.setSpacing(8)

        # 1. Roster Card
        self.team_roster_card = QFrame()
        rc_sp = QSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)
        rc_sp.setHeightForWidth(True)
        self.team_roster_card.setSizePolicy(rc_sp)
        self.team_roster_card.setStyleSheet(
            f"QFrame {{ background: {C.BG_CARD}; border: 1px solid {C.BORDER_SUBTLE}; border-radius: 8px; }}"
        )
        rc_lay = FlowLayout(self.team_roster_card, margin=0, h_spacing=8, v_spacing=6)
        rc_lay.setContentsMargins(10, 6, 10, 6)

        roster_lbl = QLabel("<b>👥 Delegation Roster:</b>")
        roster_lbl.setStyleSheet("border: none; background: transparent;")
        rc_lay.addWidget(roster_lbl)

        self.team_chips_container = QWidget()
        self.team_chips_container.setStyleSheet(".QWidget { background: transparent; border: none; }")
        self.team_member_chips_layout = QHBoxLayout(self.team_chips_container)
        self.team_member_chips_layout.setContentsMargins(0, 0, 0, 0)
        self.team_member_chips_layout.setSpacing(6)
        rc_lay.addWidget(self.team_chips_container)

        self.btn_add_member = QPushButton("➕ Add Member...")
        self.btn_add_member.setStyleSheet(f"padding: 3px 10px; font-size: 11px; font-weight: 600; border: 1px dashed {C.INFO_ACCENT}; color: {C.ACCENT_ALT}; border-radius: 6px; background: {C.INFO_SOFT};")
        self.btn_add_member.clicked.connect(self._on_add_delegation_member)
        rc_lay.addWidget(self.btn_add_member)

        # Action Buttons
        self.btn_export_team_email = QPushButton("📧 Export Email Summary...")
        self.btn_export_team_email.setStyleSheet(f"padding: 5px 12px; font-weight: 600; font-size: 11px; border-radius: 6px; background: {C.ACCENT_ALT}; color: {C.BG_CARD};")
        self.btn_export_team_email.setToolTip("Generate an executive delegation summary email with color codes, topic distribution, and personal agendas.")
        self.btn_export_team_email.clicked.connect(self._on_export_delegation_email)
        rc_lay.addWidget(self.btn_export_team_email)

        self.btn_copy_team_json = QPushButton("📋 Copy Plan JSON")
        self.btn_copy_team_json.setStyleSheet(f"padding: 5px 10px; font-size: 11px; border-radius: 6px; border: 1px solid {C.BORDER_STRONG}; background: {C.BG_SURFACE};")
        self.btn_copy_team_json.setToolTip("Copy machine-readable JSON representation of this team plan to clipboard.")
        self.btn_copy_team_json.clicked.connect(self._on_copy_delegation_json)
        rc_lay.addWidget(self.btn_copy_team_json)

        self.btn_import_team_plan = QPushButton("📥 Import Team Plan...")
        self.btn_import_team_plan.setStyleSheet(f"padding: 5px 10px; font-size: 11px; font-weight: 600; border-radius: 6px; border: 1px solid {C.BORDER_STRONG}; background: {C.BG_CANVAS};")
        self.btn_import_team_plan.setToolTip("Import delegation plan from pasted email or exported .json file.")
        self.btn_import_team_plan.clicked.connect(self._on_import_delegation_plan)
        rc_lay.addWidget(self.btn_import_team_plan)

        page_layout.addWidget(self.team_roster_card)

        # 2. Tabs: Topic Ownership vs Clash & Dispatch Board
        self.team_tabs = QTabWidget()

        # --- Tab 1: Topic Ownership Matrix ---
        tab_topics = QWidget()
        tt_lay = QVBoxLayout(tab_topics)
        tt_lay.setContentsMargins(4, 6, 4, 4)
        tt_lay.setSpacing(6)

        tt_filter_widget = QWidget(tab_topics)
        tt_filter_widget.setStyleSheet(".QWidget { background: transparent; border: none; }")
        tt_f_sp = QSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)
        tt_f_sp.setHeightForWidth(True)
        tt_filter_widget.setSizePolicy(tt_f_sp)
        tt_filter_box = FlowLayout(tt_filter_widget, margin=0, h_spacing=8, v_spacing=5)

        self.team_ai_search = QLineEdit()
        self.team_ai_search.setPlaceholderText("🔍 Filter AI number, topic, or owner...")
        self.team_ai_search.setClearButtonEnabled(True)
        self.team_ai_search.setMinimumWidth(scale_px(210))
        self.team_ai_search.setMaximumWidth(scale_px(300))
        self.team_ai_search.setStyleSheet(f"padding: 5px 10px; border: 1px solid {C.BORDER_STRONG}; border-radius: 6px; font-size: 12px;")
        self.team_ai_search.textChanged.connect(self._filter_topic_ownership_table)
        tt_filter_box.addWidget(self.team_ai_search)

        settings = QSettings("3GPP_Tools", "SessionPlanWindow")
        saved_scheduled = settings.value("team_scheduled_only", False, type=bool)

        self.btn_toggle_scheduled_only = QPushButton("📅 Scheduled Only")
        self.btn_toggle_scheduled_only.setCheckable(True)
        self.btn_toggle_scheduled_only.setChecked(saved_scheduled)
        self.btn_toggle_scheduled_only.setCursor(Qt.PointingHandCursor)
        self.btn_toggle_scheduled_only.setToolTip(
            "Toggle visibility of unscheduled Agenda Items (hide rows with no allocated sessions in the session plan)."
        )
        self.btn_toggle_scheduled_only.toggled.connect(self._on_scheduled_only_toggled)
        self._style_scheduled_toggle(saved_scheduled)
        tt_filter_box.addWidget(self.btn_toggle_scheduled_only)
        self.chk_scheduled_only = self.btn_toggle_scheduled_only

        self.btn_auto_assign_wis = QPushButton("🎯 Auto-Assign Followed WIs...")
        self.btn_auto_assign_wis.setCursor(Qt.PointingHandCursor)
        self.btn_auto_assign_wis.setStyleSheet(
            f"QPushButton {{ padding: 5px 12px; font-size: 11.5px; font-weight: 600; border-radius: 6px; border: 1px solid {C.BORDER_STRONG}; background: {C.BG_SURFACE}; color: {C.TEXT_STRONG}; }}"
            f"QPushButton:hover {{ background: {C.INFO_SOFT}; border-color: {C.INFO_ACCENT}; color: {C.INFO}; }}"
        )
        self.btn_auto_assign_wis.setToolTip("Assign topics matching your followed Work Items watchlist to a delegate.")
        self.btn_auto_assign_wis.clicked.connect(self._auto_assign_followed_wis_to_self)
        tt_filter_box.addWidget(self.btn_auto_assign_wis)

        self.btn_reset_session_overrides = QPushButton("↺ Reset Session Overrides")
        self.btn_reset_session_overrides.setCursor(Qt.PointingHandCursor)
        self.btn_reset_session_overrides.setStyleSheet(
            f"QPushButton {{ padding: 5px 11px; font-size: 11.5px; font-weight: 600; border-radius: 6px; border: 1px solid {C.WARNING_BORDER}; background: {C.WARNING_SOFT}; color: {C.WARNING_DARK}; }}"
            f"QPushButton:hover {{ background: {C.WARNING_SOFT_ALT}; border-color: {C.WARNING_ACCENT}; }}"
        )
        self.btn_reset_session_overrides.setToolTip(
            "Clear all session-level coverage overrides for this meeting while keeping persistent Default AI Assignments intact."
        )
        self.btn_reset_session_overrides.clicked.connect(self._on_reset_all_session_overrides)
        tt_filter_box.addWidget(self.btn_reset_session_overrides)

        lbl_persist_hint = QLabel("📌 <i>Default AI Assignments persist across meetings · Session overrides apply to individual slots</i>")
        lbl_persist_hint.setStyleSheet(f"font-size: 11px; color: {C.TEXT_MUTED};")
        tt_filter_box.addWidget(lbl_persist_hint)
        tt_lay.addWidget(tt_filter_widget)

        self.team_topics_table = QTableWidget()
        self.team_topics_table.setColumnCount(5)
        self.team_topics_table.setHorizontalHeaderLabels([
            "AI #", "Topic Description", "📅 Scheduled Slots", "👤 Topic Owner", "🛡️ Coverage Status"
        ])
        h_header = self.team_topics_table.horizontalHeader()
        h_header.setSectionResizeMode(0, QHeaderView.Interactive)
        h_header.setSectionResizeMode(1, QHeaderView.Stretch)
        h_header.setSectionResizeMode(2, QHeaderView.Interactive)
        h_header.setSectionResizeMode(3, QHeaderView.Interactive)
        h_header.setSectionResizeMode(4, QHeaderView.Interactive)

        self.team_topics_table.setColumnWidth(0, 65)    # AI #
        self.team_topics_table.setColumnWidth(2, 220)   # 📅 Scheduled Slots
        self.team_topics_table.setColumnWidth(3, 195)   # 👤 Topic Owner
        self.team_topics_table.setColumnWidth(4, 185)   # 🛡️ Coverage Status
        if self.team_topics_table.horizontalHeaderItem(3):
            self.team_topics_table.horizontalHeaderItem(3).setToolTip(
                "Default AI Assignment — persists across meetings automatically and can be updated between meetings."
            )
        if self.team_topics_table.horizontalHeaderItem(4):
            self.team_topics_table.horizontalHeaderItem(4).setToolTip(
                "Coverage status across scheduled sessions, including any session-level collision overrides."
            )

        register_scaled_row_height(self.team_topics_table, 36)
        self.team_topics_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.team_topics_table.setAlternatingRowColors(True)
        self.team_topics_table.setStyleSheet(f"""
            QHeaderView::section {{
                background-color: {C.BG_SURFACE};
                color: {C.TEXT_BODY};
                font-weight: 700;
                font-size: 11.5px;
                padding: 6px 8px;
                border: 1px solid {C.BORDER_SUBTLE};
                border-top: none;
                border-left: none;
            }}
            QTableWidget {{
                gridline-color: {C.BG_CANVAS};
                selection-background-color: {C.INFO_SOFT};
                selection-color: {C.TEXT_PRIMARY};
                font-size: 12px;
            }}
            QTableWidget::item {{
                padding: 4px 8px;
            }}
        """)
        tt_lay.addWidget(self.team_topics_table)
        self.team_tabs.addTab(tab_topics, "📋 Topic Ownership Matrix")

        # --- Tab 2: Clash & Dispatch Board ---
        tab_clashes = QWidget()
        tc_lay = QVBoxLayout(tab_clashes)
        tc_lay.setContentsMargins(4, 6, 4, 4)
        tc_lay.setSpacing(6)

        self.clash_kpi_lbl = QLabel()
        self.clash_kpi_lbl.setStyleSheet("font-size: 12px; font-weight: 600; padding: 6px 10px; border-radius: 6px;")
        tc_lay.addWidget(self.clash_kpi_lbl)

        self.team_clashes_table = QTableWidget()
        self.team_clashes_table.setColumnCount(7)
        self.team_clashes_table.setHorizontalHeaderLabels([
            "📅 Slot & TU", "🕒 Time", "👤 Colliding Delegate", "Stream A (Assigned)", "Parallel Stream B", "🟢 Free Colleagues", "🛡️ Action / Coverage"
        ])
        h_clashes = self.team_clashes_table.horizontalHeader()
        h_clashes.setSectionResizeMode(0, QHeaderView.Interactive)
        h_clashes.setSectionResizeMode(1, QHeaderView.Interactive)
        h_clashes.setSectionResizeMode(2, QHeaderView.Interactive)
        h_clashes.setSectionResizeMode(3, QHeaderView.Stretch)
        h_clashes.setSectionResizeMode(4, QHeaderView.Stretch)
        h_clashes.setSectionResizeMode(5, QHeaderView.Interactive)
        h_clashes.setSectionResizeMode(6, QHeaderView.Interactive)

        self.team_clashes_table.setColumnWidth(0, 125)
        self.team_clashes_table.setColumnWidth(1, 145)
        self.team_clashes_table.setColumnWidth(2, 140)
        self.team_clashes_table.setColumnWidth(5, 165)
        self.team_clashes_table.setColumnWidth(6, 185)

        register_scaled_row_height(self.team_clashes_table, 42)
        self.team_clashes_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.team_clashes_table.setAlternatingRowColors(True)
        self.team_clashes_table.cellDoubleClicked.connect(self._on_clash_cell_double_clicked)
        self.team_clashes_table.setStyleSheet(f"""
            QHeaderView::section {{
                background-color: {C.BG_SURFACE};
                color: {C.TEXT_BODY};
                font-weight: 700;
                font-size: 11.5px;
                padding: 6px 8px;
                border: 1px solid {C.BORDER_SUBTLE};
                border-top: none;
                border-left: none;
            }}
            QTableWidget {{
                gridline-color: {C.BG_CANVAS};
                selection-background-color: {C.INFO_SOFT};
                selection-color: {C.TEXT_PRIMARY};
                font-size: 12px;
            }}
            QTableWidget::item {{
                padding: 4px 8px;
            }}
        """)
        tc_lay.addWidget(self.team_clashes_table)

        tc_lay.addWidget(QLabel("💡 <i>Double-click any row to jump to that slot in the Overview Grid. Click '⚡ Appoint Coverage ▾' to 1-click assign a free colleague or open the dispatch dialog.</i>"))
        self.team_tabs.addTab(tab_clashes, "⚡ Clash && Dispatch Board")

        page_layout.addWidget(self.team_tabs, 1)
        self.team_page = page
        self.view_stack.addWidget(page)
