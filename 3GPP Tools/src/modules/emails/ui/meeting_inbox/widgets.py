"""Shared custom widgets, item delegates, toolbar builders, and QSS helpers for MeetingEmailInboxDialog."""

from __future__ import annotations

from PyQt5.QtCore import QItemSelectionModel, Qt, pyqtSignal
from PyQt5.QtGui import QColor, QKeySequence, QPainter
from PyQt5.QtWidgets import (
    QButtonGroup,
    QComboBox,
    QFrame,
    QHBoxLayout,
    QLineEdit,
    QMenu,
    QPlainTextEdit,
    QPushButton,
    QShortcut,
    QSizePolicy,
    QStyledItemDelegate,
    QTableWidget,
    QToolButton,
    QWidget,
)

from core.ui.flow_layout import FlowLayout
from core.ui.font_scale import create_help_tips_pill, scale_px
from core.ui.theme import C
from core.ui.ui_components import CLEAR_ALL_BTN_STYLE
from modules.meetings.ui.tdocs_components import CheckableComboBox


_STATUS_COLOR_MAP = {
    "🔵 My TDoc": C.ACCENT_ALT,
    "🟢 Support": C.SUCCESS,
    "🔴 Object": C.DANGER_ALT,
    "🟡 Monitor": C.WARNING_ACCENT,
}

_STATUS_BADGE_BG_MAP = {
    "🔵 My TDoc": C.INFO_SOFT,
    "🟢 Support": C.SUCCESS_SOFT,
    "🔴 Object": C.DANGER_SOFT,
    "🟡 Monitor": C.WARNING_SOFT,
}


def _status_color_hex(status: str) -> str:
    return _STATUS_COLOR_MAP.get(str(status or "⚪ Neutral").strip(), "")


def get_compact_action_bar_qss(container_name: str = "InboxActionBar") -> str:
    """Shared compact button styling for email thread/TDoc reading panels."""
    return (
        f"#{container_name} QPushButton {{ padding: 3px 8px; border-radius: 4px; "
        f"border: 1px solid {C.BORDER_STRONG}; background-color: {C.BG_CARD}; color: {C.TEXT_BODY}; font-size: 12px; }} "
        f"#{container_name} QPushButton:hover {{ background-color: {C.BG_HOVER}; border-color: {C.BORDER_HOVER}; }} "
        f"#{container_name} QPushButton:checked {{ background-color: {C.INFO_SOFT}; border-color: {C.ACCENT_ALT}; color: {C.INFO}; font-weight: bold; }} "
        f"#{container_name} QPushButton:checked:hover {{ background-color: {C.ACCENT_SOFT_HOVER}; }} "
        f"#{container_name} QPushButton:disabled {{ background-color: {C.BG_SURFACE}; color: {C.TEXT_FAINT}; border-color: {C.BORDER_SUBTLE}; }}"
    )


def get_status_combo_qss(status_val: str) -> str:
    """Shared compact tint QSS for My Status selector combos."""
    st = str(status_val or "⚪ Neutral").strip() or "⚪ Neutral"
    accent = _STATUS_COLOR_MAP.get(st, C.BORDER_STRONG)
    bg = _STATUS_BADGE_BG_MAP.get(st, C.BG_CARD)
    return (
        f"QComboBox {{ background-color: {bg}; color: {C.TEXT_PRIMARY}; "
        f"border: 1px solid {accent}; border-radius: 4px; padding: 2px 6px; font-size: 12px; "
        f"font-weight: {'bold' if st != '⚪ Neutral' else 'normal'}; }}"
    )


def get_inbox_toolbar_button_qss(compact: bool = False) -> str:
    """Uniform secondary toolbar button QSS aligned with TDocsWindow controls."""
    pad = "2px 8px" if compact else "3px 9px"
    font_px = 11 if compact else 12
    return (
        f"QPushButton {{ background-color: {C.BG_SURFACE}; color: {C.TEXT_STRONG}; "
        f"border: 1px solid {C.BORDER_STRONG}; border-radius: 4px; padding: {pad}; font-size: {font_px}px; font-weight: 500; }} "
        f"QPushButton:hover {{ background-color: {C.BG_CANVAS}; border-color: {C.BORDER_HOVER}; }} "
        f"QPushButton:pressed {{ background-color: {C.BG_MUTED}; }} "
        f"QPushButton:checked {{ background-color: {C.ACCENT_SOFT}; color: {C.ACCENT_PRIMARY}; "
        f"border: 1.5px solid {C.ACCENT_PRIMARY}; font-weight: bold; }} "
        f"QPushButton:checked:hover {{ background-color: {C.ACCENT_SOFT_HOVER}; border-color: {C.ACCENT_HOVER}; }} "
        f"QPushButton:disabled {{ background-color: {C.BG_SURFACE}; color: {C.TEXT_FAINT}; border-color: {C.BORDER_SUBTLE}; }}"
    )


def get_sync_toolbutton_qss(scoped: bool = False) -> str:
    """Uniform split QToolButton QSS for the Outlook Sync button (normal vs filter-scoped)."""
    if scoped:
        return (
            f"QToolButton {{ background-color: {C.INFO_SOFT}; color: {C.INFO}; "
            f"border: 1px solid {C.INFO_BORDER}; border-radius: 4px; padding: 3px 20px 3px 9px; "
            f"font-size: 12px; font-weight: bold; }} "
            f"QToolButton:hover {{ background-color: {C.ACCENT_SOFT_HOVER}; border-color: {C.ACCENT_ALT}; }} "
            f"QToolButton:disabled {{ background-color: {C.BG_MUTED}; color: {C.TEXT_FAINT}; border: 1px solid {C.BORDER_STRONG}; }} "
            f"QToolButton::menu-button {{ border-left: 1px solid {C.INFO_BORDER}; width: 17px; }}"
        )
    return (
        f"QToolButton {{ background-color: {C.BG_SURFACE}; color: {C.TEXT_STRONG}; "
        f"border: 1px solid {C.BORDER_STRONG}; border-radius: 4px; padding: 3px 20px 3px 9px; "
        f"font-size: 12px; font-weight: 500; }} "
        f"QToolButton:hover {{ background-color: {C.BG_CANVAS}; border-color: {C.BORDER_HOVER}; }} "
        f"QToolButton:disabled {{ background-color: {C.BG_MUTED}; color: {C.TEXT_FAINT}; border: 1px solid {C.BORDER_STRONG}; }} "
        f"QToolButton::menu-button {{ border-left: 1px solid {C.BORDER_SUBTLE}; width: 17px; }}"
    )


def get_live_sync_pill_qss() -> str:
    """Returns QSS styling for the checkable continuous Live Auto-Sync toolbar button."""
    return (
        f"QPushButton {{ padding: 3px 9px; border-radius: 4px; border: 1px solid {C.BORDER_STRONG}; "
        f"background-color: {C.BG_SURFACE}; color: {C.TEXT_STRONG}; font-size: 12px; font-weight: 500; }} "
        f"QPushButton:hover {{ background-color: {C.SUCCESS_SOFT}; border-color: {C.SUCCESS_BORDER_MEDIUM}; }} "
        f"QPushButton:checked {{ background-color: {C.SUCCESS_SOFT_ALT}; border: 1.5px solid {C.SUCCESS}; "
        f"color: {C.SUCCESS_DARK}; font-weight: bold; }} "
        f"QPushButton:checked:hover {{ background-color: {C.SUCCESS_BORDER}; }}"
    )


def get_hide_objected_button_qss() -> str:
    """Returns QSS styling for the checkable Hide Objected filter button."""
    return (
        f"QPushButton {{ padding: 3px 9px; border-radius: 4px; border: 1px solid {C.BORDER_STRONG}; "
        f"background-color: {C.BG_SURFACE}; color: {C.TEXT_STRONG}; font-size: 12px; font-weight: 500; }} "
        f"QPushButton:hover {{ background-color: {C.DANGER_SOFT}; border-color: {C.DANGER_BORDER_HOVER}; }} "
        f"QPushButton:checked {{ background-color: {C.DANGER_SOFT_HOVER}; border: 1.5px solid {C.DANGER_ALT}; "
        f"color: {C.DANGER}; font-weight: bold; }} "
        f"QPushButton:checked:hover {{ background-color: {C.DANGER_BORDER}; }}"
    )


def get_scope_pill_qss(variant: str = "info") -> str:
    """Returns uniform pill QSS for the 5 triage mode scope pills."""
    if variant == "warning":
        hover_bg = C.WARNING_SOFT
        chk_bg = C.WARNING_SOFT_ALT
        chk_border = C.WARNING_ACCENT
        chk_fg = C.WARNING_DARK
        chk_hover = C.WARNING_BORDER
    elif variant == "danger":
        hover_bg = C.DANGER_SOFT
        chk_bg = C.DANGER_SOFT_HOVER
        chk_border = C.DANGER_ALT
        chk_fg = C.DANGER
        chk_hover = C.DANGER_BORDER
    else:
        hover_bg = C.BG_CANVAS
        chk_bg = C.INFO_SOFT_STRONG
        chk_border = C.ACCENT_ALT
        chk_fg = C.INFO
        chk_hover = C.ACCENT_SOFT_HOVER
    return (
        f"QPushButton {{ padding: 3px 9px; border-radius: 12px; border: 1px solid {C.BORDER_STRONG}; "
        f"background-color: {C.BG_SURFACE}; color: {C.TEXT_BODY}; font-size: 12px; font-weight: 500; }} "
        f"QPushButton:hover {{ background-color: {hover_bg}; border-color: {C.BORDER_HOVER}; }} "
        f"QPushButton:checked {{ background-color: {chk_bg}; border: 1.5px solid {chk_border}; "
        f"color: {chk_fg}; font-weight: bold; }} "
        f"QPushButton:checked:hover {{ background-color: {chk_hover}; }}"
    )


def make_inbox_toolbar_divider(parent: QWidget | None = None) -> QFrame:
    """Creates a subtle vertical divider line between toolbar button groups."""
    div = QFrame(parent)
    div.setFrameShape(QFrame.VLine)
    div.setFixedWidth(1)
    div.setFixedHeight(scale_px(18))
    div.setStyleSheet(f"QFrame {{ background-color: {C.BORDER_STRONG}; border: none; }}")
    return div


def make_inbox_toolbar_group(parent: QWidget | None = None, spacing: int = 4) -> tuple[QWidget, QHBoxLayout]:
    """Creates an atomic horizontal sub-group container so related controls stay together in FlowLayout."""
    box = QWidget(parent)
    box.setStyleSheet(".QWidget { background: transparent; border: none; }")
    lay = QHBoxLayout(box)
    lay.setContentsMargins(0, 0, 0, 0)
    lay.setSpacing(spacing)
    return box, lay


def build_inbox_sync_controls(dialog):
    """Creates and attaches the Sync Outlook split button + Live (2m) toggle button to `dialog`."""
    ctrl_h = scale_px(26)
    sync_btn = QToolButton(dialog)
    sync_btn.setText("⚡ Sync New")
    sync_btn.setCursor(Qt.PointingHandCursor)
    sync_btn.setMinimumHeight(ctrl_h)
    sync_btn.setToolButtonStyle(Qt.ToolButtonTextOnly)
    sync_btn.setPopupMode(QToolButton.MenuButtonPopup)
    sync_btn.setStyleSheet(get_sync_toolbutton_qss(scoped=False))
    sync_btn.setToolTip(
        "Fast incremental Outlook sync of new Tracked Thread replies, TDoc emails, and read/flag updates since your last sync."
    )
    sync_menu = QMenu(sync_btn)
    act_selected = sync_menu.addAction("🔄 Sync Selected Thread")
    act_selected.setEnabled(False)
    act_selected.triggered.connect(dialog._on_sync_selected_clicked)
    act_filtered = sync_menu.addAction("⚡ Sync Filtered Only")
    act_filtered.triggered.connect(lambda: dialog._start_outlook_sync(scoped=True))
    sync_menu.addSeparator()
    act_full = sync_menu.addAction("📅 Full Meeting Sync (Date Range)...")
    act_full.triggered.connect(dialog._on_full_meeting_sync_clicked)
    sync_btn.setMenu(sync_menu)
    sync_btn.clicked.connect(dialog._on_sync_outlook_clicked)

    btn_live = QPushButton("🟢 Live (2m)", dialog)
    btn_live.setCheckable(True)
    btn_live.setCursor(Qt.PointingHandCursor)
    btn_live.setMinimumHeight(ctrl_h)
    btn_live.setChecked(bool(getattr(dialog, "_live_sync_enabled", False)))
    btn_live.setToolTip(
        "Continuous eMeeting Update Mode (every 2 minutes).\n"
        "Silently pulls newly arrived emails using the persisted watermark without rescanning old history.\n"
        "Persists across restarts and catches up the missing gap automatically when reopened."
    )
    btn_live.setStyleSheet(get_live_sync_pill_qss())
    btn_live.toggled.connect(dialog._on_live_sync_toggled)

    dialog.sync_outlook_btn = sync_btn
    dialog._sync_menu = sync_menu
    dialog._act_sync_selected = act_selected
    dialog._act_sync_filtered = act_filtered
    dialog._act_sync_full = act_full
    dialog._default_sync_scoped = False
    dialog.btn_live_sync = btn_live
    return sync_btn, btn_live


def build_inbox_top_toolbar(dialog) -> QFrame:
    """Builds Row 1 of MeetingEmailInboxDialog: grouped Scope Pills | View Switcher | Sync & Refresh | Tools."""
    ctrl_h = scale_px(26)
    row1_frame = QFrame(dialog)
    row1_frame.setObjectName("inboxTopBar")
    row1_sp = QSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)
    row1_sp.setHeightForWidth(True)
    row1_frame.setSizePolicy(row1_sp)
    row1_frame.setStyleSheet(
        f"QFrame#inboxTopBar {{ background-color: {C.BG_CARD}; border: 1px solid {C.BORDER_STRONG}; border-radius: 6px; }}"
    )
    top_row = FlowLayout(row1_frame, margin=0, h_spacing=6, v_spacing=4)
    top_row.setContentsMargins(6, 4, 6, 4)

    # Group 1: Triage Scope Pills
    scope_box, scope_lay = make_inbox_toolbar_group(row1_frame, spacing=4)
    dialog.mode_group = QButtonGroup(dialog)
    dialog.mode_group.setExclusive(True)

    pill_specs = (
        ("pill_priority", "🔥 Priority (0)", "PRIORITY", "warning",
         "Unread activity on your Important/Watched TDocs (non-Neutral My Status), Flagged items, "
         "or inside your Followed AIs (excluding parked 🛑 Objected items unless a newer revision arrived)."),
        ("pill_unread", "🔵 Unread Only (0)", "UNREAD", "info",
         "Show only unread emails/threads."),
        ("pill_flagged", "🚩 Flagged (0)", "FLAGGED", "danger",
         "Show only emails/threads flagged for follow-up."),
        ("pill_recent", "🕒 Last 24h (0)", "RECENT", "info",
         "Show emails/threads active within the last 24 hours."),
        ("pill_all", "📋 All (0)", "ALL", "info",
         "Show all indexed meeting emails and tracked threads."),
    )
    for attr_name, label, mode_key, variant, tip in pill_specs:
        btn = QPushButton(label, scope_box)
        btn.setCheckable(True)
        btn.setCursor(Qt.PointingHandCursor)
        btn.setMinimumHeight(ctrl_h)
        btn.setChecked(dialog._active_mode == mode_key)
        btn.setToolTip(tip)
        btn.setStyleSheet(get_scope_pill_qss(variant))
        btn.clicked.connect(lambda _=False, m=mode_key: dialog._set_mode(m))
        dialog.mode_group.addButton(btn)
        setattr(dialog, attr_name, btn)
        scope_lay.addWidget(btn)
    top_row.addWidget(scope_box)
    top_row.addWidget(make_inbox_toolbar_divider(row1_frame))

    # Group 2: Joined Segmented View Mode Switcher (By Message | By Thread | By TDoc)
    view_track = QFrame(row1_frame)
    view_track.setObjectName("inboxViewSegment")
    view_track.setMinimumHeight(ctrl_h)
    view_track.setStyleSheet(
        f"QFrame#inboxViewSegment {{ background-color: {C.BG_CANVAS}; border: 1px solid {C.BORDER_STRONG}; border-radius: 5px; }}"
    )
    view_lay = QHBoxLayout(view_track)
    view_lay.setContentsMargins(2, 1, 2, 1)
    view_lay.setSpacing(2)
    seg_btn_qss = (
        f"QPushButton {{ padding: 2px 8px; border-radius: 4px; border: 1px solid transparent; "
        f"background-color: transparent; color: {C.TEXT_SECONDARY}; font-size: 12px; font-weight: 500; }} "
        f"QPushButton:hover {{ background-color: {C.BG_HOVER}; color: {C.TEXT_PRIMARY}; }} "
        f"QPushButton:checked {{ background-color: {C.BG_CARD}; border: 1px solid {C.INDIGO_DARK}; "
        f"color: {C.INDIGO_DARK}; font-weight: bold; }} "
        f"QPushButton:checked:hover {{ background-color: {C.WG_SA_BG}; }}"
    )
    dialog.view_mode_group = QButtonGroup(dialog)
    dialog.view_mode_group.setExclusive(True)
    view_specs = (
        ("btn_view_messages", "📨 By Message", "MESSAGES",
         "Show individual email messages sorted newest first."),
        ("btn_view_threads", "🧵 By Thread", "THREADS",
         "Group emails by discussion thread for fast AI/WI assignment and thread triage."),
        ("btn_view_tdocs", "📄 By TDoc", "TDOCS",
         "Group all meeting emails by Base TDoc ID (ideal for e-Meetings) with Author, Revisions, "
         "Objection signals, and direct TDocsWindow status/notes synchronization."),
    )
    for attr_name, label, vm_key, tip in view_specs:
        btn = QPushButton(label, view_track)
        btn.setCheckable(True)
        btn.setCursor(Qt.PointingHandCursor)
        btn.setChecked(dialog._view_mode == vm_key)
        btn.setToolTip(tip)
        btn.setStyleSheet(seg_btn_qss)
        btn.clicked.connect(lambda _=False, vm=vm_key: dialog._set_view_mode(vm))
        dialog.view_mode_group.addButton(btn)
        setattr(dialog, attr_name, btn)
        view_lay.addWidget(btn)
    top_row.addWidget(view_track)
    top_row.addWidget(make_inbox_toolbar_divider(row1_frame))

    # Group 3: Outlook Sync & Local Reload
    sync_box, sync_lay = make_inbox_toolbar_group(row1_frame, spacing=4)
    sync_btn, btn_live = build_inbox_sync_controls(dialog)
    sync_lay.addWidget(sync_btn)
    sync_lay.addWidget(btn_live)

    dialog.refresh_btn = QPushButton("🔄", sync_box)
    dialog.refresh_btn.setCursor(Qt.PointingHandCursor)
    dialog.refresh_btn.setMinimumHeight(ctrl_h)
    dialog.refresh_btn.setToolTip("Reload latest messages already saved in local meeting databases.")
    dialog.refresh_btn.setStyleSheet(get_inbox_toolbar_button_qss())
    dialog.refresh_btn.clicked.connect(dialog.reload_from_dbs)
    sync_lay.addWidget(dialog.refresh_btn)
    top_row.addWidget(sync_box)
    top_row.addWidget(make_inbox_toolbar_divider(row1_frame))

    # Group 4: Analytics, Layout & Help
    tools_box, tools_lay = make_inbox_toolbar_group(row1_frame, spacing=4)
    dialog.btn_stats = QPushButton("📊 Stats", tools_box)
    dialog.btn_stats.setCursor(Qt.PointingHandCursor)
    dialog.btn_stats.setMinimumHeight(ctrl_h)
    dialog.btn_stats.setStyleSheet(get_inbox_toolbar_button_qss())
    dialog.btn_stats.setToolTip("Generate an interactive HTML Email Analytics Dashboard for this meeting.")
    dialog.btn_stats.clicked.connect(dialog._on_generate_email_statistics)
    tools_lay.addWidget(dialog.btn_stats)

    dialog.btn_briefings = QPushButton("🤖 Briefings", tools_box)
    dialog.btn_briefings.setCursor(Qt.PointingHandCursor)
    dialog.btn_briefings.setMinimumHeight(ctrl_h)
    dialog.btn_briefings.setStyleSheet(get_inbox_toolbar_button_qss())
    dialog.btn_briefings.setToolTip("Open Per-AI cross-thinking briefings and saved selection summaries.")
    dialog.btn_briefings.clicked.connect(lambda _=False: dialog._open_briefings_dialog())
    tools_lay.addWidget(dialog.btn_briefings)

    dialog.btn_layout_toggle = QPushButton("⬒ List on Top", tools_box)
    dialog.btn_layout_toggle.setCursor(Qt.PointingHandCursor)
    dialog.btn_layout_toggle.setMinimumHeight(ctrl_h)
    dialog.btn_layout_toggle.setStyleSheet(get_inbox_toolbar_button_qss())
    dialog.btn_layout_toggle.setToolTip("Switch between Side-by-Side (desktop) and Top/Bottom (laptop) layout (Ctrl+L).")
    dialog.btn_layout_toggle.clicked.connect(dialog._toggle_layout_orientation)
    tools_lay.addWidget(dialog.btn_layout_toggle)
    dialog._layout_shortcut = QShortcut(QKeySequence("Ctrl+L"), dialog)
    dialog._layout_shortcut.activated.connect(dialog._toggle_layout_orientation)

    dialog.help_pill_btn = create_help_tips_pill(
        [
            "⚡ Sync New & 🟢 Live (2m): incremental MAPI Table.Restrict sync using a persisted watermark (survives app closes & reboots cleanly)",
            "🔥 Priority: unread activity on your Watched TDocs (non-Neutral My Status), Flagged items, or Followed AIs",
            "📄 By TDoc: rolls up all emails per Base TDoc ID with Author, Revisions (🆕 rXX), <<START>> (📋), and Objection signals (⚠️ / 🛑)",
            "🛑 Objected: parks an objected thread/TDoc out of Priority; automatically resurfaces if a newer revision (r02 > r01) is posted!",
            "Ctrl + Mouse Wheel in reading pane: zoom text larger/smaller (Ctrl+0 to reset; saved automatically)",
            "Ctrl + L: toggle Side-by-Side (large monitor) vs. List on Top (laptop) split layout",
        ],
        parent=tools_box,
        title="Email & e-Meeting Inbox Tips",
    )
    dialog.help_pill_btn.setMinimumHeight(ctrl_h)
    dialog.help_pill_btn.setStyleSheet(get_inbox_toolbar_button_qss())
    tools_lay.addWidget(dialog.help_pill_btn)
    top_row.addWidget(tools_box)
    return row1_frame


def build_inbox_filter_bar(dialog) -> QFrame:
    """Builds Row 2 of MeetingEmailInboxDialog: Search | Priority Radar Scope | Facet Filter Combos."""
    ctrl_h = scale_px(26)
    row2_frame = QFrame(dialog)
    row2_frame.setObjectName("inboxFilterBar")
    row2_sp = QSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)
    row2_sp.setHeightForWidth(True)
    row2_frame.setSizePolicy(row2_sp)
    row2_frame.setStyleSheet(
        f"QFrame#inboxFilterBar {{ background-color: {C.BG_CARD}; border: 1px solid {C.BORDER_STRONG}; border-radius: 6px; }} "
        f"QFrame#inboxFilterBar QLineEdit {{ padding: 3px 8px; border: 1px solid {C.BORDER_STRONG}; border-radius: 4px; "
        f"background-color: {C.BG_SURFACE}; color: {C.TEXT_STRONG}; font-size: 12px; }} "
        f"QFrame#inboxFilterBar QLineEdit:focus {{ background-color: {C.BG_CARD}; border: 1px solid {C.ACCENT_PRIMARY}; }} "
        f"QFrame#inboxFilterBar QComboBox {{ padding: 3px 8px; border: 1px solid {C.BORDER_STRONG}; border-radius: 4px; "
        f"background-color: {C.BG_SURFACE}; color: {C.TEXT_STRONG}; font-size: 12px; }} "
        f"QFrame#inboxFilterBar QComboBox:hover {{ background-color: {C.BG_CANVAS}; border-color: {C.BORDER_HOVER}; }} "
        f"QFrame#inboxFilterBar QComboBox[activeFilter=\"true\"] {{ background-color: {C.INFO_SOFT}; "
        f"border: 1.5px solid {C.INFO_ACCENT}; color: {C.INFO}; font-weight: bold; }}"
    )
    filter_row = FlowLayout(row2_frame, margin=0, h_spacing=6, v_spacing=4)
    filter_row.setContentsMargins(6, 4, 6, 4)

    # Group 1: Search Box
    search_box, search_lay = make_inbox_toolbar_group(row2_frame, spacing=4)
    dialog.search_edit = QLineEdit(search_box)
    dialog.search_edit.setMinimumHeight(ctrl_h)
    dialog.search_edit.setMinimumWidth(scale_px(220))
    dialog.search_edit.setClearButtonEnabled(True)
    dialog.search_edit.setPlaceholderText("🔍 Search subject, sender, company, TDoc, or body...")
    if dialog._saved_search_text:
        dialog.search_edit.setText(dialog._saved_search_text)
    dialog.search_edit.textChanged.connect(dialog._apply_filters)
    search_lay.addWidget(dialog.search_edit)
    filter_row.addWidget(search_box)
    filter_row.addWidget(make_inbox_toolbar_divider(row2_frame))

    # Group 2: Radar & Objection Scope
    radar_box, radar_lay = make_inbox_toolbar_group(row2_frame, spacing=4)
    dialog.followed_ais_combo = CheckableComboBox("★ Followed AIs", radar_box)
    dialog.followed_ais_combo.setMinimumHeight(ctrl_h)
    dialog.followed_ais_combo.setMinimumWidth(scale_px(145))
    dialog.followed_ais_combo.setMaximumWidth(scale_px(210))
    dialog.followed_ais_combo.setToolTip(
        "Agenda Items on your Priority Radar (auto-inherited from Followed/Starred WIs, Session Plan, and My TDocs).\n"
        "Check/uncheck to override for this meeting, or right-click to reset to inherited defaults."
    )
    dialog.followed_ais_combo.selectionChanged.connect(dialog._on_followed_ais_combo_changed)
    dialog.followed_ais_combo.setContextMenuPolicy(Qt.CustomContextMenu)
    dialog.followed_ais_combo.customContextMenuRequested.connect(dialog._on_followed_ais_context_menu)
    radar_lay.addWidget(dialog.followed_ais_combo)

    dialog.btn_hide_objected = QPushButton("🛑 Hide Objected", radar_box)
    dialog.btn_hide_objected.setCheckable(True)
    dialog.btn_hide_objected.setCursor(Qt.PointingHandCursor)
    dialog.btn_hide_objected.setMinimumHeight(ctrl_h)
    dialog.btn_hide_objected.setChecked(bool(getattr(dialog, "_hide_objected", False)))
    dialog.btn_hide_objected.setToolTip(
        "Hide threads/TDocs marked as 🛑 Objected (Parked), unless a newer revision arrived after the objection."
    )
    dialog.btn_hide_objected.setStyleSheet(get_hide_objected_button_qss())
    dialog.btn_hide_objected.toggled.connect(dialog._on_hide_objected_toggled)
    radar_lay.addWidget(dialog.btn_hide_objected)
    filter_row.addWidget(radar_box)
    filter_row.addWidget(make_inbox_toolbar_divider(row2_frame))

    # Group 3: Facet Filter Combos + Clear Button
    facets_box, facets_lay = make_inbox_toolbar_group(row2_frame, spacing=4)
    dialog.ai_combo = CheckableComboBox("AI", facets_box)
    dialog.ai_combo.setMinimumHeight(ctrl_h)
    dialog.ai_combo.setMinimumWidth(scale_px(115))
    dialog.ai_combo.setMaximumWidth(scale_px(185))
    dialog.ai_combo.setToolTip("Filter by one or more Agenda Items (or Unassigned).")
    dialog.ai_combo.selectionChanged.connect(dialog._apply_filters)
    facets_lay.addWidget(dialog.ai_combo)

    dialog.company_combo = CheckableComboBox("Company", facets_box)
    dialog.company_combo.setMinimumHeight(ctrl_h)
    dialog.company_combo.setMinimumWidth(scale_px(120))
    dialog.company_combo.setMaximumWidth(scale_px(175))
    dialog.company_combo.setToolTip("Filter by one or more participating/sender companies.")
    dialog.company_combo.selectionChanged.connect(dialog._apply_filters)
    facets_lay.addWidget(dialog.company_combo)

    dialog.status_combo = CheckableComboBox("Status", facets_box)
    dialog.status_combo.setMinimumHeight(ctrl_h)
    dialog.status_combo.setMinimumWidth(scale_px(110))
    dialog.status_combo.setMaximumWidth(scale_px(165))
    dialog.status_combo.setToolTip("Filter by thread/TDoc status (My Status).")
    dialog.status_combo.selectionChanged.connect(dialog._apply_filters)
    facets_lay.addWidget(dialog.status_combo)

    dialog.source_combo = QComboBox(facets_box)
    dialog.source_combo.setMinimumHeight(ctrl_h)
    dialog.source_combo.addItem("All Sources", "")
    dialog.source_combo.addItem("⚠️ Needs Assignment (No AI / Unpinned)", "NEEDS_TRIAGE")
    dialog.source_combo.addItem("🧵 Tracked Threads Only", "THREAD")
    dialog.source_combo.addItem("📄 Unpinned Only", "UNPINNED")
    if dialog._saved_source_filter:
        src_idx = dialog.source_combo.findData(dialog._saved_source_filter)
        if src_idx >= 0:
            dialog.source_combo.setCurrentIndex(src_idx)
    dialog.source_combo.setToolTip(
        "Filter by Tracked Thread vs. Unpinned TDoc Email, or isolate items needing assignment."
    )
    dialog.source_combo.currentIndexChanged.connect(dialog._apply_filters)
    facets_lay.addWidget(dialog.source_combo)

    dialog.btn_clear_filters = QPushButton("✕ Clear", facets_box)
    dialog.btn_clear_filters.setCursor(Qt.PointingHandCursor)
    dialog.btn_clear_filters.setMinimumHeight(ctrl_h)
    dialog.btn_clear_filters.setStyleSheet(CLEAR_ALL_BTN_STYLE)
    dialog.btn_clear_filters.setToolTip(
        "Clear search text and reset AI, Company, Status, Source, and Hide Objected filters."
    )
    dialog.btn_clear_filters.clicked.connect(dialog._clear_all_filters)
    facets_lay.addWidget(dialog.btn_clear_filters)

    filter_row.addWidget(facets_box)
    return row2_frame


class _InboxStatusStripDelegate(QStyledItemDelegate):
    """Paints a crisp left-edge status color stripe on the State column without widening the table."""

    def paint(self, painter: QPainter, option, index):
        super().paint(painter, option, index)
        color_hex = str(index.data(Qt.UserRole + 1) or "").strip()
        if not color_hex:
            return
        painter.save()
        bar_w = scale_px(5, minimum=4)
        rect = option.rect
        painter.fillRect(rect.left() + 1, rect.top() + 2, bar_w, max(4, rect.height() - 4), QColor(color_hex))
        painter.restore()


class _NotePlainTextEdit(QPlainTextEdit):
    """Compact inline note editor that emits save signals on blur or Ctrl+Enter."""

    focusLost = pyqtSignal()
    saveRequested = pyqtSignal()

    def focusOutEvent(self, event):
        super().focusOutEvent(event)
        self.focusLost.emit()

    def keyPressEvent(self, event):
        if (event.modifiers() & Qt.ControlModifier) and event.key() in (Qt.Key_Return, Qt.Key_Enter):
            self.saveRequested.emit()
            event.accept()
            return
        super().keyPressEvent(event)


class _InboxTableWidget(QTableWidget):
    """QTableWidget with ExtendedSelection, deterministic selectRow(), and viewport resize signal."""

    viewportResized = pyqtSignal()

    def selectRow(self, row: int):
        sel_model = self.selectionModel()
        model = self.model()
        if sel_model is None or model is None or row < 0 or row >= self.rowCount():
            return
        idx = model.index(row, 0)
        sel_model.setCurrentIndex(idx, QItemSelectionModel.ClearAndSelect | QItemSelectionModel.Rows)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.viewportResized.emit()
