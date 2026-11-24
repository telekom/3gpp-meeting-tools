# --- File: src/modules/meetings/ui/session_plan/window/mixins/ui_setup_ribbons.py ---
"""
Widget construction: top ribbon, navigation ribbon and the Overview / NOW / Agenda /
Drafting / Tracked view pages (called once from ``_setup_ui``).

Mixin of SessionPlanWindow: relies on attributes set in SessionPlanWindow.__init__.
"""
from typing import List, Optional
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QButtonGroup,
    QCheckBox,
    QComboBox,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMenu,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QTableWidget,
    QVBoxLayout,
    QWidget,
)
from core.ui.flow_layout import FlowLayout
from core.ui.font_scale import register_scaled_row_height, scale_px
from core.ui.theme import C
from modules.meetings.ui.session_plan.window.constants import DAY_NAMES
from modules.meetings.ui.session_plan.window.grid_headers import SessionPlanVerticalHeader, SessionPlanHorizontalHeader, SessionPlanGridTable


class RibbonSetupMixin:
    """Widget construction: top ribbon, navigation ribbon and the Overview / NOW / Agenda /"""

    def _setup_top_ribbon(self, parent_layout: QVBoxLayout):
        top_bar = QWidget(self)
        top_bar.setStyleSheet(".QWidget { background: transparent; border: none; }")
        top_sp = QSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)
        top_sp.setHeightForWidth(True)
        top_bar.setSizePolicy(top_sp)
        ribbon = FlowLayout(top_bar, margin=0, h_spacing=8, v_spacing=6)
        self._top_ribbon_widget = top_bar

        # Identity + Snapshot group
        id_grp = QWidget(top_bar)
        id_grp.setStyleSheet(".QWidget { background: transparent; border: none; }")
        id_lay = QHBoxLayout(id_grp)
        id_lay.setContentsMargins(0, 0, 0, 0)
        id_lay.setSpacing(8)

        title_lbl = QLabel(f"<b>SA2 #{self.meeting_number}</b>")
        title_lbl.setStyleSheet(f"font-size: 16px; color: {C.TEXT_PRIMARY};")
        title_lbl.setToolTip(f"SA WG2 Meeting #{self.meeting_number} ({self.location})")
        id_lay.addWidget(title_lbl)

        self.snap_combo = QComboBox()
        self.snap_combo.setMinimumWidth(scale_px(210))
        self.snap_combo.setMaximumWidth(scale_px(320))
        self.snap_combo.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
        self.snap_combo.setStyleSheet(f"padding: 4px 8px; font-weight: 600; border-radius: 6px; border: 1px solid {C.BORDER_STRONG};")
        self.snap_combo.setToolTip("Select or switch between imported session plan snapshots and revisions.")
        self.snap_combo.currentIndexChanged.connect(self._on_snapshot_changed)
        id_lay.addWidget(self.snap_combo)
        ribbon.addWidget(id_grp)

        # Prominent Dual-Clock Pill + Timezone Selector Group
        clk_grp = QWidget(top_bar)
        clk_grp.setStyleSheet(".QWidget { background: transparent; border: none; }")
        clk_lay = QHBoxLayout(clk_grp)
        clk_lay.setContentsMargins(0, 0, 0, 0)
        clk_lay.setSpacing(6)

        self.clock_lbl = QLabel()
        self.clock_lbl.setStyleSheet(
            f"background: {C.INFO_SOFT}; color: {C.INFO_NAVY}; border: 1px solid {C.INFO_BORDER}; "
            f"border-radius: 999px; padding: 4px 12px; font-size: 12.5px; font-weight: 600;"
        )
        self.clock_lbl.setToolTip("Current physical time at the meeting venue and on your corporate laptop.")
        clk_lay.addWidget(self.clock_lbl)

        self.tz_btn = QPushButton(f"🌍 {self.tz_name}")
        self.tz_btn.setStyleSheet(f"""
            QPushButton {{
                background: {C.BG_CANVAS}; color: {C.TEXT_BODY}; font-size: 11.5px; font-weight: 600;
                padding: 4px 10px; border: 1px solid {C.BORDER_STRONG}; border-radius: 999px;
            }}
            QPushButton:hover {{ background: {C.BORDER_SUBTLE}; }}
        """)
        self.tz_btn.setToolTip("Meeting Venue Timezone. Click to change or verify against your corporate laptop clock.")
        self.tz_btn.clicked.connect(self._open_timezone_dialog)
        clk_lay.addWidget(self.tz_btn)
        ribbon.addWidget(clk_grp)

        # Action Buttons Group
        act_grp = QWidget(top_bar)
        act_grp.setStyleSheet(".QWidget { background: transparent; border: none; }")
        act_lay = QHBoxLayout(act_grp)
        act_lay.setContentsMargins(0, 0, 0, 0)
        act_lay.setSpacing(6)

        self.import_btn = QPushButton("📥 Import ▾")
        self.import_btn.setStyleSheet(f"padding: 5px 12px; font-weight: 600; border-radius: 6px; background: {C.ACCENT_ALT}; color: {C.BG_CARD};")
        self.import_btn.setToolTip("Import a session plan document from URL, TDoc#, local file, or clipboard paste, or drag & drop files directly onto this window.")
        import_menu = QMenu(self)
        import_menu.addAction("🔗 Enter URL or TDoc#...", self._import_via_url).setToolTip("Fetch via 3GPP URL or type a TDoc number (e.g. S2-2610195).")
        import_menu.addAction("📂 Browse Local Document...", self._import_via_file).setToolTip("Select a .docx, .doc, .zip, or HTML table from your disk.")
        import_menu.addAction("📋 Paste Markdown/Text Table...", self._import_via_paste).setToolTip("Paste table text directly from an email, TOHRU chat, or notes.")
        self.import_btn.setMenu(import_menu)
        act_lay.addWidget(self.import_btn)

        self.diff_btn = QPushButton("⇄ Compare")
        self.diff_btn.setStyleSheet(f"padding: 5px 10px; font-weight: 500; border-radius: 6px; background: {C.BG_CANVAS}; border: 1px solid {C.BORDER_STRONG};")
        self.diff_btn.setToolTip("Compare the active session plan against an earlier revision to see added, removed, rescheduled, or volume-changed sessions.")
        self.diff_btn.clicked.connect(self._open_diff_dialog)
        act_lay.addWidget(self.diff_btn)

        self.export_btn = QPushButton("📤 Export ▾")
        self.export_btn.setStyleSheet(f"padding: 5px 10px; font-weight: 500; border-radius: 6px; background: {C.BG_CANVAS}; border: 1px solid {C.BORDER_STRONG};")
        self.export_btn.setToolTip("Export to responsive offline Mobile HTML (no PDF table separator issues!), strict UTC .ics Calendar, or Outlook Email summary.")
        export_menu = QMenu(self)
        export_menu.addAction("🌐 Export Interactive HTML (Overview && Mobile)...", self._export_mobile_html).setToolTip("Generate an interactive, responsive standalone HTML file featuring the 5-Day Overview Grid and Mobile Day Cards with delegate toggling.")
        export_menu.addAction("🌐 Open Interactive HTML in Default Browser", self._open_mobile_html_in_browser).setToolTip("Export and instantly open the interactive session plan (Overview Grid & Mobile) in your browser.")
        export_menu.addSeparator()
        export_menu.addAction("📅 Export Full Meeting .ics Calendar...", self._export_ics).setToolTip("Export calendar events with strict UTC timestamps for Outlook, Apple, or Google Calendar.")
        export_menu.addAction("📅 Export Delegate .ics (Personal / Colleague)...", self._export_delegate_ics).setToolTip("Export personal calendar events scoped to a specific delegate with assigned topics.")
        export_menu.addSeparator()
        export_menu.addAction("✉️ Copy Personal/Tracked Email Summary", self._copy_email_summary).setToolTip("Copy rich HTML email summary of personal tracked topics to clipboard.")
        export_menu.addAction("📧 Copy Complete Delegation Email Summary...", self._on_export_delegation_email).setToolTip("Copy rich HTML delegation plan and machine-readable interchange comment for team email sharing.")
        export_menu.addSeparator()
        export_menu.addAction("📂 Open Session Plans Folder", self._open_plans_folder).setToolTip("Open <meeting_dir>/Agenda/SessionPlans/ in Windows Explorer.")
        self.export_btn.setMenu(export_menu)
        act_lay.addWidget(self.export_btn)
        ribbon.addWidget(act_grp)

        parent_layout.addWidget(top_bar)

    def _setup_navigation_ribbon(self, parent_layout: QVBoxLayout):
        nav_bar = QWidget(self)
        nav_bar.setStyleSheet(".QWidget { background: transparent; border: none; }")
        nav_sp = QSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)
        nav_sp.setHeightForWidth(True)
        nav_bar.setSizePolicy(nav_sp)
        sub_layout = FlowLayout(nav_bar, margin=0, h_spacing=6, v_spacing=5)
        self._nav_ribbon_widget = nav_bar

        # Segmented view mode buttons
        self.mode_group = QButtonGroup(self)
        self.mode_group.setExclusive(True)

        self.btn_grid = QPushButton("🗓️ Overview Grid")
        self.btn_grid.setToolTip("Full 5-day spatial overview showing parallel Streams 1, 2, and 3 across all Quarters and Time Units.")
        self.btn_now = QPushButton("🔴 Now")
        self.btn_now.setToolTip("See what is happening right now across all streams & physical rooms, plus upcoming periods and custom room directions.")
        self.btn_agenda = QPushButton("📋 Agenda (0 Sched / 0)")
        self.btn_agenda.setToolTip("Cross-references the official meeting agenda against scheduled sessions. Highlights which Agenda Items are scheduled vs unscheduled (missing from plan).")
        self.btn_drafting = QPushButton("🛠️ Drafting Radar (0)")
        self.btn_drafting.setToolTip("Chronological tracking of early morning (08:00), lunch (13:00), and evening (18:00) informal and formal drafting sessions.")
        self.btn_tracked = QPushButton("🎯 Delegate Agenda (0)")
        self.btn_tracked.setToolTip("View personal schedules by delegate (or your own agenda) and the entire delegation timeline.")
        self.btn_team = QPushButton("👥 Team Allocations (Clean)")
        self.btn_team.setToolTip("Distribute Agenda Items across team members, detect double-booking clashes, and appoint available delegates on a session-by-session basis.")

        for idx, btn in enumerate([self.btn_grid, self.btn_now, self.btn_agenda, self.btn_drafting, self.btn_tracked, self.btn_team]):
            btn.setCheckable(True)
            btn.setStyleSheet(f"""
                QPushButton {{
                    padding: 5px 10px; font-size: 12px; font-weight: 600; border: 1px solid {C.BORDER_STRONG};
                    background: {C.BG_CARD}; color: {C.TEXT_SECONDARY}; border-radius: 6px;
                }}
                QPushButton:checked {{
                    background: {C.INFO_ACCENT}; color: {C.BG_CARD}; border-color: {C.INFO_ACCENT};
                }}
            """)
            self.mode_group.addButton(btn, idx)
            sub_layout.addWidget(btn)

        self.btn_grid.setChecked(True)
        self.mode_group.buttonClicked[int].connect(self._on_view_mode_changed)

        # Quick-Jump Conflicts Button
        self.btn_conflicts_quick = QPushButton("⚠️ Conflicts (0)")
        self.btn_conflicts_quick.setCursor(Qt.PointingHandCursor)
        self.btn_conflicts_quick.setStyleSheet(f"""
            QPushButton {{
                padding: 5px 10px; font-size: 12px; font-weight: 600; border: 1px solid {C.DANGER_BORDER};
                background: {C.DANGER_SOFT}; color: {C.DANGER}; border-radius: 6px;
            }}
            QPushButton:hover {{ background: {C.DANGER_SOFT_HOVER}; }}
        """)
        self.btn_conflicts_quick.setToolTip("Jump directly to the Parallel Stream Conflicts & Coverage Dispatch Board.")
        self.btn_conflicts_quick.clicked.connect(self._jump_to_conflicts_board)
        sub_layout.addWidget(self.btn_conflicts_quick)

        # Quick Track Topics manager
        self.btn_manage_topics = QPushButton("⭐ Track Topics...")
        self.btn_manage_topics.setStyleSheet(f"""
            QPushButton {{
                padding: 5px 10px; font-size: 12px; font-weight: 600; border: 1px solid {C.VIOLET_BORDER};
                background: {C.VIOLET_SURFACE}; color: {C.VIOLET_PRIMARY}; border-radius: 6px;
            }}
            QPushButton:hover {{ background: {C.VIOLET_SOFT}; }}
        """)
        self.btn_manage_topics.setToolTip("Select which Agenda Items or Work Items to follow in 'My Topics' and highlight in purple.")
        self.btn_manage_topics.clicked.connect(self._open_manage_topics_dialog)
        sub_layout.addWidget(self.btn_manage_topics)

        # Search Bar
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("🔍 Filter AI, topic, room, convener...")
        self.search_input.setToolTip("Filter sessions and agenda items by AI number (e.g. 20.6.18), topic keyword, room, or convener.")
        self.search_input.setClearButtonEnabled(True)
        self.search_input.setMinimumWidth(scale_px(185))
        self.search_input.setMaximumWidth(scale_px(280))
        self.search_input.setStyleSheet(f"padding: 5px 10px; border: 1px solid {C.BORDER_STRONG}; border-radius: 6px;")
        self.search_input.textChanged.connect(self._apply_search_filter)
        sub_layout.addWidget(self.search_input)

        parent_layout.addWidget(nav_bar)

    # --- View 1: Overview Grid ---

    def _setup_overview_view(self):
        page = QWidget()
        page_layout = QVBoxLayout(page)
        page_layout.setContentsMargins(0, 0, 0, 0)
        page_layout.setSpacing(6)

        # Room & Convener Guide & Stream Toggle Strip
        self.room_strip = QFrame()
        room_sp = QSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)
        room_sp.setHeightForWidth(True)
        self.room_strip.setSizePolicy(room_sp)
        self.room_strip.setStyleSheet(
            f"QFrame {{ background: {C.BG_SURFACE}; border: 1px solid {C.BORDER_SUBTLE}; border-radius: 6px; }}"
        )
        self.room_strip_layout = FlowLayout(self.room_strip, margin=0, h_spacing=10, v_spacing=4)
        self.room_strip_layout.setContentsMargins(8, 3, 8, 3)

        self.btn_toggle_s1 = QCheckBox("🏢 Str 1: Main Room")
        self.btn_toggle_s1.setChecked(True)
        self.btn_toggle_s1.setStyleSheet(f"QCheckBox {{ color: {C.INDIGO_DARK}; font-weight: 600; font-size: 11px; padding: 1px 4px; }}")
        self.btn_toggle_s1.setToolTip("Toggle visibility of Stream 1 (Main Room / Plenary / Single Stream).")
        self.btn_toggle_s1.toggled.connect(self._update_row_visibility)
        self.room_strip_layout.addWidget(self.btn_toggle_s1)

        self.btn_toggle_s2 = QCheckBox("🏢 Str 2: Breakout 1")
        self.btn_toggle_s2.setChecked(True)
        self.btn_toggle_s2.setStyleSheet(f"QCheckBox {{ color: {C.EMERALD_DARK}; font-weight: 600; font-size: 11px; padding: 1px 4px; }}")
        self.btn_toggle_s2.setToolTip("Toggle visibility of Stream 2 (Breakout 1 / Mike).")
        self.btn_toggle_s2.toggled.connect(self._update_row_visibility)
        self.room_strip_layout.addWidget(self.btn_toggle_s2)

        self.btn_toggle_s3 = QCheckBox("🏢 Str 3: Breakout 2")
        self.btn_toggle_s3.setChecked(True)
        self.btn_toggle_s3.setStyleSheet(f"QCheckBox {{ color: {C.WARNING}; font-weight: 600; font-size: 11px; padding: 1px 4px; }}")
        self.btn_toggle_s3.setToolTip("Toggle visibility of Stream 3 (Breakout 2 / Wanqiang).")
        self.btn_toggle_s3.toggled.connect(self._update_row_visibility)
        self.room_strip_layout.addWidget(self.btn_toggle_s3)

        sep = QFrame()
        sep.setFrameShape(QFrame.VLine)
        sep.setStyleSheet(f"color: {C.BORDER_STRONG};")
        self.room_strip_layout.addWidget(sep)

        self.chk_breaks = QCheckBox("☕ Show Breaks")
        self.chk_breaks.setChecked(True)
        self.chk_breaks.setStyleSheet(f"QCheckBox {{ color: {C.TEXT_SECONDARY}; font-weight: 600; font-size: 11px; padding: 2px 4px; }}")
        self.chk_breaks.setToolTip("Toggle visibility of Coffee Breaks, Lunch Breaks, and Short Breaks.")
        self.chk_breaks.toggled.connect(self._update_row_visibility)
        self.room_strip_layout.addWidget(self.chk_breaks)

        sep_days = QFrame()
        sep_days.setFrameShape(QFrame.VLine)
        sep_days.setStyleSheet(f"color: {C.BORDER_STRONG};")
        self.room_strip_layout.addWidget(sep_days)

        day_grp = QWidget(self.room_strip)
        day_grp.setStyleSheet(".QWidget { background: transparent; border: none; }")
        day_lay = QHBoxLayout(day_grp)
        day_lay.setContentsMargins(0, 0, 0, 0)
        day_lay.setSpacing(3)

        lbl_days = QLabel("<b>Days:</b>")
        lbl_days.setStyleSheet(f"font-size: 11px; color: {C.TEXT_SECONDARY}; border: none; background: transparent;")
        day_lay.addWidget(lbl_days)

        self.btn_day_all = QPushButton("All 5")
        self.btn_day_all.setCheckable(True)
        self.btn_day_all.setChecked(True)
        self.btn_day_all.setCursor(Qt.PointingHandCursor)
        self.btn_day_all.setToolTip("Show all 5 meeting days (Monday – Friday)")
        self.btn_day_all.clicked.connect(lambda: self._on_day_filter_clicked("ALL"))
        day_lay.addWidget(self.btn_day_all)

        self.btn_day_today_plus = QPushButton("Today+")
        self.btn_day_today_plus.setCheckable(True)
        self.btn_day_today_plus.setCursor(Qt.PointingHandCursor)
        self.btn_day_today_plus.setToolTip("Hide past days and show Today through Friday")
        self.btn_day_today_plus.clicked.connect(lambda: self._on_day_filter_clicked("TODAY_PLUS"))
        day_lay.addWidget(self.btn_day_today_plus)

        self.day_filter_btns: List[QPushButton] = []
        for d_idx, d_name in enumerate(DAY_NAMES):
            btn_d = QPushButton(d_name[:3])
            btn_d.setCheckable(True)
            btn_d.setCursor(Qt.PointingHandCursor)
            btn_d.setToolTip(f"Focus on {d_name} (click multiple days to compare side-by-side)")
            btn_d.clicked.connect(lambda _, i=d_idx: self._on_day_filter_clicked(i))
            day_lay.addWidget(btn_d)
            self.day_filter_btns.append(btn_d)

        self.room_strip_layout.addWidget(day_grp)
        self._style_day_filter_buttons()

        sep_loc = QFrame()
        sep_loc.setFrameShape(QFrame.VLine)
        sep_loc.setStyleSheet(f"color: {C.BORDER_STRONG};")
        self.room_strip_layout.addWidget(sep_loc)

        self.btn_room_locations = QPushButton("📍 Room Locations...")
        self.btn_room_locations.setCursor(Qt.PointingHandCursor)
        self.btn_room_locations.setStyleSheet(
            f"QPushButton {{ padding: 2px 8px; font-size: 11px; font-weight: 600; border: 1px solid {C.BORDER_STRONG}; "
            f"border-radius: 5px; background: {C.BG_CARD}; color: {C.TEXT_BODY}; }}"
            f"QPushButton:hover {{ background: {C.INFO_SOFT}; border-color: {C.INFO_ACCENT}; color: {C.INFO}; }}"
        )
        self.btn_room_locations.setToolTip("Add hand-made floor/wing/hall directions for physical rooms (saved locally in meeting_locations.json).")
        self.btn_room_locations.clicked.connect(lambda: self._open_room_locations_dialog())
        self.room_strip_layout.addWidget(self.btn_room_locations)

        page_layout.addWidget(self.room_strip)

        # Delegate Focus / Heatmap Highlight Strip
        self.delegate_strip = QFrame()
        del_sp = QSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)
        del_sp.setHeightForWidth(True)
        self.delegate_strip.setSizePolicy(del_sp)
        self.delegate_strip.setStyleSheet(
            f"QFrame {{ background: {C.BG_CARD}; border: 1px solid {C.BORDER_SUBTLE}; border-radius: 6px; }}"
        )
        self.delegate_strip_layout = QHBoxLayout(self.delegate_strip)
        self.delegate_strip_layout.setContentsMargins(8, 3, 8, 3)
        self.delegate_strip_layout.setSpacing(6)

        lbl_f = QLabel("<b>Delegate Focus:</b>")
        lbl_f.setStyleSheet(f"font-size: 11px; color: {C.TEXT_SECONDARY}; border: none; background: transparent;")
        lbl_f.setToolTip("Tick one or more delegates to highlight their sessions; all other sessions are dimmed.")
        self.delegate_strip_layout.addWidget(lbl_f)

        self.btn_chip_all = QPushButton("🌈 All Team")
        self.btn_chip_all.setToolTip("Clear the focus selection and show every session at full contrast.")
        self.btn_chip_all.setCursor(Qt.PointingHandCursor)
        self.btn_chip_all.clicked.connect(lambda: self._on_delegate_highlight_clicked("ALL"))
        self.delegate_strip_layout.addWidget(self.btn_chip_all)

        # Dynamic per-delegate focus checkboxes live in a responsive FlowLayout
        self.delegate_chips_widget = QWidget(self.delegate_strip)
        self.delegate_chips_widget.setStyleSheet(".QWidget { background: transparent; border: none; }")
        chips_sp = QSizePolicy(QSizePolicy.Expanding, QSizePolicy.Maximum)
        chips_sp.setHeightForWidth(True)
        self.delegate_chips_widget.setSizePolicy(chips_sp)
        self.delegate_chips_layout = FlowLayout(self.delegate_chips_widget, margin=0, h_spacing=6, v_spacing=4)
        self.delegate_strip_layout.addWidget(self.delegate_chips_widget, 1)

        self.btn_chip_me: Optional[QPushButton] = None
        self.dynamic_delegate_chips: List[QCheckBox] = []

        self.lbl_badge_legend = QLabel(f"<span style='color:{C.TEXT_MUTED};'>ℹ️ Legend</span>")
        self.lbl_badge_legend.setStyleSheet(
            f"font-size: 10.5px; font-weight: 600; background: {C.BG_SURFACE}; "
            f"border: 1px solid {C.BORDER_SUBTLE}; border-radius: 9px; padding: 1px 7px;"
        )
        self.lbl_badge_legend.setToolTip(
            "<b>Cell &amp; Badge Legend</b><br>"
            "• <b>🕒 Pill</b> = Sub-slot or custom time override (blue = half-quarter, amber = custom)<br>"
            "• <b>Solid pill</b> = Topic owner<br>"
            "• <b>Dashed pill</b> = Ad-hoc session cover<br>"
            f"• <b style='color:{C.DANGER_ALT};'>Red ring / ⚠</b> = Parallel-stream clash<br>"
            "• <b>❓</b> = Unassigned Agenda Item<br>"
            "• <b>⛔</b> = Explicitly Not Followed (Out of scope)"
        )
        self.delegate_strip_layout.addWidget(self.lbl_badge_legend)
        page_layout.addWidget(self.delegate_strip)
        self._style_focus_buttons()

        # Overview Table (10 columns: 2 per day)
        self.grid_table = SessionPlanGridTable(self)
        self.grid_table.setColumnCount(10)
        self.grid_table.setCornerButtonEnabled(False)
        self.horizontal_day_header = SessionPlanHorizontalHeader(self, self.grid_table)
        self.grid_table.setHorizontalHeader(self.horizontal_day_header)
        self.horizontal_day_header.setSectionResizeMode(QHeaderView.Stretch)
        self.horizontal_day_header.setFixedHeight(scale_px(44))
        self.vertical_time_header = SessionPlanVerticalHeader(self, self.grid_table)
        self.grid_table.setVerticalHeader(self.vertical_time_header)
        self.grid_table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        register_scaled_row_height(self.grid_table, 46)
        self.grid_table.setAlternatingRowColors(False)
        self.grid_table.setStyleSheet(
            f"QTableWidget {{ border: 1px solid {C.BORDER_STRONG}; selection-background-color: {C.INDIGO_SOFT}; }}"
        )
        self.grid_table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.grid_table.customContextMenuRequested.connect(self._on_grid_context_menu)
        self.grid_table.cellDoubleClicked.connect(self._on_grid_cell_double_clicked)
        page_layout.addWidget(self.grid_table)

        page_layout.addWidget(QLabel("💡 <i>Double-click any cell to add personal notes / priority. Right-click to track topic or view details.</i>"))

        self.overview_page = page
        self.view_stack.addWidget(page)

    # --- View 2: Live "Now" & Upcoming Rooms Dashboard ---

    def _setup_now_view(self):
        page = QWidget()
        page_layout = QVBoxLayout(page)
        page_layout.setContentsMargins(0, 0, 0, 0)
        page_layout.setSpacing(8)

        # Top Live Status & Manual Time-Travel Stepper Bar
        ctrl_card = QFrame()
        ctrl_sp = QSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)
        ctrl_sp.setHeightForWidth(True)
        ctrl_card.setSizePolicy(ctrl_sp)
        ctrl_card.setStyleSheet(
            f"QFrame {{ background: {C.BG_CARD}; border: 1px solid {C.BORDER_SUBTLE}; border-radius: 8px; }}"
        )
        ctrl_lay = FlowLayout(ctrl_card, margin=0, h_spacing=8, v_spacing=6)
        ctrl_lay.setContentsMargins(10, 6, 10, 6)

        self.now_status_pill = QLabel("⏳ Initializing...")
        self.now_status_pill.setStyleSheet(
            f"background: {C.INFO_SOFT}; color: {C.INFO}; border: 1px solid {C.INFO_BORDER}; "
            f"border-radius: 999px; padding: 3px 10px; font-weight: 700; font-size: 11.5px;"
        )
        ctrl_lay.addWidget(self.now_status_pill)

        self.now_banner_lbl = QLabel("")
        self.now_banner_lbl.setWordWrap(True)
        self.now_banner_lbl.setStyleSheet(f"color: {C.TEXT_STRONG}; font-size: 12px; font-weight: 600;")
        ctrl_lay.addWidget(self.now_banner_lbl)

        self.btn_now_prev = QPushButton("◀ Prev")
        self.btn_now_prev.setCursor(Qt.PointingHandCursor)
        self.btn_now_prev.setToolTip("Inspect the previous scheduled time unit")
        self.btn_now_prev.clicked.connect(lambda: self._step_now_unit(-1))
        ctrl_lay.addWidget(self.btn_now_prev)

        self.now_slot_combo = QComboBox()
        self.now_slot_combo.setMinimumWidth(scale_px(210))
        self.now_slot_combo.setMaximumWidth(scale_px(320))
        self.now_slot_combo.setStyleSheet(f"padding: 3px 8px; font-size: 11.5px; border: 1px solid {C.BORDER_STRONG}; border-radius: 5px;")
        self.now_slot_combo.currentIndexChanged.connect(self._on_now_combo_changed)
        ctrl_lay.addWidget(self.now_slot_combo)

        self.btn_now_next = QPushButton("Next ▶")
        self.btn_now_next.setCursor(Qt.PointingHandCursor)
        self.btn_now_next.setToolTip("Inspect the next scheduled time unit")
        self.btn_now_next.clicked.connect(lambda: self._step_now_unit(1))
        ctrl_lay.addWidget(self.btn_now_next)

        self.btn_now_live = QPushButton("🔴 Live Now")
        self.btn_now_live.setCursor(Qt.PointingHandCursor)
        self.btn_now_live.setStyleSheet(
            f"QPushButton {{ padding: 4px 10px; font-size: 11.5px; font-weight: 700; border-radius: 6px; "
            f"background: {C.SUCCESS_SOFT}; color: {C.SUCCESS}; border: 1px solid {C.SUCCESS_BORDER}; }}"
            f"QPushButton:hover {{ background: {C.SUCCESS_SOFT_ALT}; }}"
        )
        self.btn_now_live.setToolTip("Sync back to the current real-time Venue clock")
        self.btn_now_live.clicked.connect(self._reset_now_to_live)
        ctrl_lay.addWidget(self.btn_now_live)

        self.btn_now_locations = QPushButton("📍 Room Locations...")
        self.btn_now_locations.setCursor(Qt.PointingHandCursor)
        self.btn_now_locations.setStyleSheet(
            f"QPushButton {{ padding: 4px 10px; font-size: 11.5px; font-weight: 600; border-radius: 6px; "
            f"background: {C.BG_SURFACE}; color: {C.TEXT_BODY}; border: 1px solid {C.BORDER_STRONG}; }}"
            f"QPushButton:hover {{ background: {C.INFO_SOFT}; border-color: {C.INFO_ACCENT}; color: {C.INFO}; }}"
        )
        self.btn_now_locations.setToolTip("Edit hand-made floor/wing/hall directions stored locally in meeting_locations.json")
        self.btn_now_locations.clicked.connect(lambda: self._open_room_locations_dialog())
        ctrl_lay.addWidget(self.btn_now_locations)

        page_layout.addWidget(ctrl_card)

        # Scrollable Content Container
        self.now_scroll = QScrollArea()
        self.now_scroll.setWidgetResizable(True)
        self.now_scroll.setStyleSheet("QScrollArea { border: none; background: transparent; }")
        self.now_content_widget = QWidget()
        self.now_content_layout = QVBoxLayout(self.now_content_widget)
        self.now_content_layout.setContentsMargins(0, 0, 4, 8)
        self.now_content_layout.setSpacing(12)
        self.now_scroll.setWidget(self.now_content_widget)
        page_layout.addWidget(self.now_scroll, 1)

        self.now_page = page
        self.view_stack.addWidget(page)

    # --- View 3: Agenda Mode ---

    def _setup_agenda_view(self):
        page = QWidget()
        page_layout = QVBoxLayout(page)
        page_layout.setContentsMargins(0, 0, 0, 0)

        self.agenda_table = QTableWidget()
        self.agenda_table.setColumnCount(6)
        self.agenda_table.setHorizontalHeaderLabels([
            "⭐", "AI #", "Topic Description", "Status", "Allocated Quarters & Rooms", "Actions"
        ])
        self.agenda_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.agenda_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.agenda_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.agenda_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        self.agenda_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.agenda_table.setAlternatingRowColors(True)
        page_layout.addWidget(self.agenda_table)

        self.agenda_page = page
        self.view_stack.addWidget(page)

    # --- View 3: Drafting Radar ---

    def _setup_drafting_view(self):
        page = QWidget()
        page_layout = QVBoxLayout(page)
        page_layout.setContentsMargins(0, 0, 0, 0)
        page_layout.setSpacing(6)

        self.drafting_table = QTableWidget()
        self.drafting_table.setColumnCount(6)
        self.drafting_table.setHorizontalHeaderLabels([
            "Day", "Quarter", "Time (Venue)", "Stream / Room", "Convener", "Drafting Topic & AIs"
        ])
        self.drafting_table.horizontalHeader().setSectionResizeMode(5, QHeaderView.Stretch)
        self.drafting_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.drafting_table.setAlternatingRowColors(True)
        self.drafting_table.cellDoubleClicked.connect(self._on_drafting_cell_double_clicked)
        page_layout.addWidget(self.drafting_table)
        page_layout.addWidget(QLabel("💡 <i>Double-click any drafting session to view and highlight in the 5-day Overview Grid.</i>"))

        self.drafting_page = page
        self.view_stack.addWidget(page)

    # --- View 4: Delegate Agenda (formerly My Topics) ---

    def _setup_tracked_view(self):
        page = QWidget()
        page_layout = QVBoxLayout(page)
        page_layout.setContentsMargins(0, 0, 0, 0)
        page_layout.setSpacing(6)

        # Delegate Switcher Bar
        top_bar = QHBoxLayout()
        top_bar.setSpacing(8)
        top_bar.addWidget(QLabel("<b>Delegate Agenda:</b>"))
        self.delegate_agenda_combo = QComboBox()
        self.delegate_agenda_combo.setMinimumWidth(230)
        self.delegate_agenda_combo.setStyleSheet(f"padding: 3px 8px; font-weight: 600; border: 1px solid {C.BORDER_STRONG}; border-radius: 6px;")
        self.delegate_agenda_combo.currentIndexChanged.connect(self._on_delegate_agenda_changed)
        top_bar.addWidget(self.delegate_agenda_combo)

        top_bar.addStretch()
        page_layout.addLayout(top_bar)

        self.tracked_table = QTableWidget()
        self.tracked_table.setColumnCount(7)
        self.tracked_table.setHorizontalHeaderLabels([
            "Day", "Quarter", "Time (Venue)", "Stream / Room", "Delegate", "Matched Topic", "My Notes / Priority / Conflict"
        ])
        self.tracked_table.horizontalHeader().setSectionResizeMode(5, QHeaderView.Stretch)
        self.tracked_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.tracked_table.setAlternatingRowColors(True)
        self.tracked_table.cellDoubleClicked.connect(self._on_tracked_cell_double_clicked)
        self.tracked_table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tracked_table.customContextMenuRequested.connect(self._on_tracked_context_menu)
        page_layout.addWidget(self.tracked_table)
        page_layout.addWidget(QLabel("💡 <i>Double-click any session to jump to the Overview Grid. Right-click to edit notes or copy details.</i>"))

        self.tracked_page = page
        self.view_stack.addWidget(page)
