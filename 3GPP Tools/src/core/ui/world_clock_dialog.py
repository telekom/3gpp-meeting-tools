# --- File: src/core/ui/world_clock_dialog.py ---
"""
Interactive World Clock & 3GPP Venue Timezone Management Dialog.

Allows delegates to:
- View Laptop Local Time (IT-locked), Official 3GPP UTC, and all active 3GPP meeting venues
- Filter displayed venues by Working Group(s) (e.g., all active WGs vs. only SA2)
- Sync venue timezones from public geocoding/time servers (VPN-proof city lookup)
- Detect physical exit-IP timezone via public IP server
- Apply or reset per-venue / per-meeting-folder manual UTC offsets & DST nudges
"""

from __future__ import annotations

import logging
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from PyQt5.QtCore import Qt, QTimer, pyqtSignal
from PyQt5.QtWidgets import (
    QButtonGroup,
    QCheckBox,
    QComboBox,
    QDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QRadioButton,
    QScrollArea,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from core.ui.theme import C, R, S, T
from core.ui.ui_components import (
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    CARD_FRAME_STYLE,
    SPEC_FILTER_COMBO_STYLE,
)
from core.ui.world_clock_service import (
    VenueClockState,
    WebTimezoneSyncWorker,
    WorldClockService,
    WorldClockSnapshot,
    _format_wg_summary,
)
from modules.meetings.core.session_plan.session_plan_timezone import (
    SessionPlanTimezone,
    format_utc_offset,
)
from modules.meetings.ui.tdocs_components import CheckableComboBox

logger = logging.getLogger(__name__)


def _build_utc_offset_options() -> List[Tuple[int, str]]:
    """Builds UTC-12:00 .. UTC+14:00 options in 30-minute steps."""
    opts: List[Tuple[int, str]] = []
    for mins in range(-12 * 60, 14 * 60 + 1, 30):
        opts.append((mins, format_utc_offset(mins)))
    return opts


def _badge_style(bg: str, fg: str, bd: str) -> str:
    return (
        f"background-color: {bg}; color: {fg}; border: 1px solid {bd}; "
        f"border-radius: {R.SM}px; padding: 2px 7px; font-size: {T.SIZE_CAPTION}px; font-weight: 600;"
    )


def _source_badge_spec(source: str) -> Tuple[str, str, str, str]:
    if source == "manual_override":
        return ("✏️ Manual Override", C.WARNING_SOFT, C.WARNING_DARK, C.WARNING_BORDER)
    if source == "web_api":
        return ("🌐 Web Verified", C.SUCCESS_SOFT, C.SUCCESS_DARK, C.SUCCESS_BORDER)
    if source == "meeting_folder":
        return ("📁 Meeting Folder", C.INFO_SOFT, C.INFO_DARK, C.INFO_BORDER)
    if source == "builtin":
        return ("📚 Built-in TZ", C.BG_SURFACE, C.TEXT_SECONDARY, C.BORDER_STRONG)
    return ("⚠️ Unresolved", C.DANGER_SOFT, C.DANGER, C.DANGER_BORDER)


def _format_delta_vs_laptop(diff_min: int) -> str:
    if diff_min == 0:
        return "Same as Laptop"
    sign = "+" if diff_min > 0 else "-"
    abs_m = abs(diff_min)
    h, m = divmod(abs_m, 60)
    if m == 0:
        return f"{sign}{h}h vs Laptop"
    return f"{sign}{h}h{m:02d}m vs Laptop"


class _VenueClockCard(QFrame):
    """Card widget displaying live time & override controls for one 3GPP venue."""

    offset_override_requested = pyqtSignal(str, int, list, object)  # loc, offset_min, meetings, explicit_dir
    nudge_requested = pyqtSignal(str, int, list, object)  # loc, delta_min, meetings, explicit_dir
    web_sync_requested = pyqtSignal(str, list, object)  # loc, meetings, explicit_dir
    reset_requested = pyqtSignal(str, list, object)  # loc, meetings, explicit_dir

    def __init__(
        self,
        venue: VenueClockState,
        explicit_meeting_dir: Optional[Path] = None,
        is_focused_meeting: bool = False,
        parent: Optional[QWidget] = None,
    ):
        super().__init__(parent)
        self.venue = venue
        self.explicit_meeting_dir = explicit_meeting_dir
        self.is_focused_meeting = is_focused_meeting
        self._updating_combo = False

        self.setObjectName("venueClockCard")
        border_col = C.ACCENT_PRIMARY if is_focused_meeting else C.BORDER_SUBTLE
        bg_col = C.ACCENT_SOFT if is_focused_meeting else C.BG_CARD
        self.setStyleSheet(
            f"QFrame#venueClockCard {{"
            f" background-color: {bg_col}; border: 1px solid {border_col}; border-radius: {R.LG}px;"
            f"}}"
        )

        root = QVBoxLayout(self)
        root.setContentsMargins(12, 10, 12, 10)
        root.setSpacing(6)

        # --- Row 1: Location + Live Time ---
        top_row = QHBoxLayout()
        top_row.setSpacing(8)

        prefix = "📌" if is_focused_meeting else ("⏭️" if venue.is_upcoming else "📍")
        tag = " (Current Meeting)" if is_focused_meeting else (" (Upcoming)" if venue.is_upcoming else "")
        self.title_lbl = QLabel(f"<b>{prefix} {venue.location}{tag}</b>")
        self.title_lbl.setStyleSheet(f"font-size: {T.SIZE_SECTION_TITLE}px; color: {C.TEXT_PRIMARY};")
        top_row.addWidget(self.title_lbl)

        self.source_badge = QLabel()
        top_row.addWidget(self.source_badge)

        self.day_badge = QLabel()
        self.day_badge.setVisible(False)
        top_row.addWidget(self.day_badge)

        top_row.addStretch(1)

        self.time_lbl = QLabel("--:--:--")
        self.time_lbl.setStyleSheet(
            f"font-family: {T.FONT_FAMILY_MONO}; font-size: 18px; font-weight: 700; color: {C.TEXT_PRIMARY};"
        )
        top_row.addWidget(self.time_lbl)
        root.addLayout(top_row)

        # --- Row 2: Metadata & Delta ---
        meta_row = QHBoxLayout()
        meta_row.setSpacing(10)

        wg_summary = _format_wg_summary(venue.wgs, max_items=16) or "—"
        date_txt = ""
        if venue.start_date and venue.end_date:
            date_txt = f"  ·  {venue.start_date} → {venue.end_date}"
        elif venue.start_date:
            date_txt = f"  ·  {venue.start_date}"

        self.meta_lbl = QLabel(f"WGs: {wg_summary}{date_txt}")
        self.meta_lbl.setWordWrap(True)
        self.meta_lbl.setStyleSheet(f"font-size: {T.SIZE_CAPTION}px; color: {C.TEXT_SECONDARY};")
        meta_row.addWidget(self.meta_lbl, 1)

        self.date_delta_lbl = QLabel()
        self.date_delta_lbl.setStyleSheet(f"font-size: {T.SIZE_CAPTION}px; color: {C.TEXT_MUTED}; font-weight: 600;")
        meta_row.addWidget(self.date_delta_lbl)
        root.addLayout(meta_row)

        # --- Row 3: Manual Override & Web Sync Controls ---
        ctrl_row = QHBoxLayout()
        ctrl_row.setSpacing(6)

        tz_caption = QLabel("Timezone:")
        tz_caption.setStyleSheet(f"font-size: {T.SIZE_CAPTION}px; color: {C.TEXT_MUTED};")
        ctrl_row.addWidget(tz_caption)

        self.offset_combo = QComboBox()
        self.offset_combo.setMinimumWidth(120)
        for mins, lbl in _build_utc_offset_options():
            self.offset_combo.addItem(lbl, mins)
        self.offset_combo.currentIndexChanged.connect(self._on_combo_changed)
        ctrl_row.addWidget(self.offset_combo)

        self.btn_minus_1h = QPushButton("-1h")
        self.btn_minus_1h.setToolTip("Subtract 1 hour (manual DST override)")
        self.btn_minus_1h.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_minus_1h.setCursor(Qt.PointingHandCursor)
        self.btn_minus_1h.clicked.connect(
            lambda: self.nudge_requested.emit(
                self.venue.location, -60, self.venue.meetings, self.explicit_meeting_dir
            )
        )
        ctrl_row.addWidget(self.btn_minus_1h)

        self.btn_plus_1h = QPushButton("+1h")
        self.btn_plus_1h.setToolTip("Add 1 hour (manual DST override)")
        self.btn_plus_1h.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_plus_1h.setCursor(Qt.PointingHandCursor)
        self.btn_plus_1h.clicked.connect(
            lambda: self.nudge_requested.emit(
                self.venue.location, 60, self.venue.meetings, self.explicit_meeting_dir
            )
        )
        ctrl_row.addWidget(self.btn_plus_1h)

        ctrl_row.addStretch(1)

        self.btn_sync_web = QPushButton("🔄 Sync Web")
        self.btn_sync_web.setToolTip("Query public geocoding & timezone server for this city")
        self.btn_sync_web.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_sync_web.setCursor(Qt.PointingHandCursor)
        self.btn_sync_web.clicked.connect(
            lambda: self.web_sync_requested.emit(
                self.venue.location, self.venue.meetings, self.explicit_meeting_dir
            )
        )
        ctrl_row.addWidget(self.btn_sync_web)

        self.btn_reset = QPushButton("↺ Reset")
        self.btn_reset.setToolTip("Clear manual/saved override and revert to automatic resolution")
        self.btn_reset.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_reset.setCursor(Qt.PointingHandCursor)
        self.btn_reset.clicked.connect(
            lambda: self.reset_requested.emit(
                self.venue.location, self.venue.meetings, self.explicit_meeting_dir
            )
        )
        ctrl_row.addWidget(self.btn_reset)

        root.addLayout(ctrl_row)

        self.update_state(venue)

    def _on_combo_changed(self, idx: int) -> None:
        if self._updating_combo or idx < 0:
            return
        mins = self.offset_combo.itemData(idx)
        if mins is not None:
            self.offset_override_requested.emit(
                self.venue.location, int(mins), self.venue.meetings, self.explicit_meeting_dir
            )

    def set_syncing(self, syncing: bool) -> None:
        self.btn_sync_web.setEnabled(not syncing)
        self.btn_sync_web.setText("⏳ Syncing..." if syncing else "🔄 Sync Web")

    def update_state(self, venue: VenueClockState) -> None:
        self.venue = venue
        b_text, b_bg, b_fg, b_bd = _source_badge_spec(venue.source)
        self.source_badge.setText(b_text)
        self.source_badge.setStyleSheet(_badge_style(b_bg, b_fg, b_bd))
        self.source_badge.setToolTip(venue.tz_label)

        if venue.venue_dt is not None:
            self.time_lbl.setText(venue.venue_dt.strftime("%H:%M:%S"))
            delta_txt = _format_delta_vs_laptop(venue.offset_diff_vs_laptop_min)
            self.date_delta_lbl.setText(f"{venue.venue_dt:%a %d %b %Y}  ·  {delta_txt}")
            if venue.day_diff_vs_laptop != 0:
                day_lbl = "Tomorrow (+1d)" if venue.day_diff_vs_laptop > 0 else "Yesterday (-1d)"
                self.day_badge.setText(day_lbl)
                self.day_badge.setStyleSheet(_badge_style(C.INFO_SOFT, C.INFO_DARK, C.INFO_BORDER))
                self.day_badge.setVisible(True)
            else:
                self.day_badge.setVisible(False)
        else:
            self.time_lbl.setText("--:--:--")
            self.date_delta_lbl.setText("Timezone unresolved — select offset or Sync Web")
            self.day_badge.setVisible(False)

        self.btn_reset.setEnabled(venue.source in ("manual_override", "web_api", "meeting_folder"))

        self._updating_combo = True
        try:
            if venue.tz_offset_minutes is not None:
                idx = self.offset_combo.findData(int(venue.tz_offset_minutes))
                if idx < 0:
                    self.offset_combo.addItem(
                        format_utc_offset(int(venue.tz_offset_minutes)),
                        int(venue.tz_offset_minutes),
                    )
                    idx = self.offset_combo.count() - 1
                self.offset_combo.setCurrentIndex(idx)
        finally:
            self._updating_combo = False


class WorldClockDialog(QDialog):
    """Interactive World Clock, WG Filter, and Venue Timezone Override dialog."""

    settings_changed = pyqtSignal()

    def __init__(
        self,
        service: WorldClockService,
        focus_meeting: Optional[dict] = None,
        focus_meeting_dir: Optional[Path] = None,
        parent: Optional[QWidget] = None,
    ):
        super().__init__(parent)
        self.service = service
        self.focus_meeting = focus_meeting
        self.focus_meeting_dir = Path(focus_meeting_dir) if focus_meeting_dir else None
        self._active_workers: List[WebTimezoneSyncWorker] = []
        self._venue_cards: List[Tuple[str, bool, _VenueClockCard]] = []  # (loc_key, is_focus, card)
        self._detected_ip_info: Optional[Tuple[str, int, str, str]] = None  # city, off, lbl, iana
        self._last_layout_signature: Optional[tuple] = None

        self.setWindowTitle("🕒 3GPP World Clock & Active Meeting Venues")
        self.setMinimumSize(680, 560)
        self.resize(760, 660)

        self._build_ui()
        self._load_preferences_into_ui()
        self._refresh_snapshot(rebuild_cards=True)

        self._tick_timer = QTimer(self)
        self._tick_timer.setInterval(1000)
        self._tick_timer.timeout.connect(self._on_tick)
        self._tick_timer.start()

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(14, 14, 14, 14)
        root.setSpacing(10)

        # --- 1. Reference Clocks Card (Laptop + UTC + Physical IP Detector) ---
        ref_card = QFrame()
        ref_card.setStyleSheet(CARD_FRAME_STYLE)
        ref_layout = QVBoxLayout(ref_card)
        ref_layout.setContentsMargins(14, 12, 14, 12)
        ref_layout.setSpacing(8)

        clocks_row = QHBoxLayout()
        clocks_row.setSpacing(18)

        # Laptop clock column
        lap_box = QVBoxLayout()
        lap_box.setSpacing(2)
        lap_hdr = QLabel("💻 Laptop Local Time (IT Policy)")
        lap_hdr.setStyleSheet(f"font-size: {T.SIZE_CAPTION}px; font-weight: 600; color: {C.TEXT_MUTED};")
        self.laptop_time_lbl = QLabel("--:--:--")
        self.laptop_time_lbl.setStyleSheet(
            f"font-family: {T.FONT_FAMILY_MONO}; font-size: 22px; font-weight: 700; color: {C.TEXT_PRIMARY};"
        )
        self.laptop_sub_lbl = QLabel("")
        self.laptop_sub_lbl.setStyleSheet(f"font-size: {T.SIZE_CAPTION}px; color: {C.TEXT_SECONDARY};")
        lap_box.addWidget(lap_hdr)
        lap_box.addWidget(self.laptop_time_lbl)
        lap_box.addWidget(self.laptop_sub_lbl)
        clocks_row.addLayout(lap_box, 1)

        sep1 = QFrame()
        sep1.setFrameShape(QFrame.VLine)
        sep1.setStyleSheet(f"color: {C.BORDER_SUBTLE};")
        clocks_row.addWidget(sep1)

        # UTC clock column
        utc_box = QVBoxLayout()
        utc_box.setSpacing(2)
        utc_hdr = QLabel("🌐 Official 3GPP Time (UTC)")
        utc_hdr.setStyleSheet(f"font-size: {T.SIZE_CAPTION}px; font-weight: 600; color: {C.TEXT_MUTED};")
        self.utc_time_lbl = QLabel("--:--:-- UTC")
        self.utc_time_lbl.setStyleSheet(
            f"font-family: {T.FONT_FAMILY_MONO}; font-size: 22px; font-weight: 700; color: {C.ACCENT_PRIMARY};"
        )
        self.utc_sub_lbl = QLabel("")
        self.utc_sub_lbl.setStyleSheet(f"font-size: {T.SIZE_CAPTION}px; color: {C.TEXT_SECONDARY};")
        utc_box.addWidget(utc_hdr)
        utc_box.addWidget(self.utc_time_lbl)
        utc_box.addWidget(self.utc_sub_lbl)
        clocks_row.addLayout(utc_box, 1)

        ref_layout.addLayout(clocks_row)

        # Physical IP detection bar
        ip_row = QHBoxLayout()
        ip_row.setSpacing(8)
        self.btn_detect_ip = QPushButton("📡 Detect My Physical Timezone (via Public IP)")
        self.btn_detect_ip.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_detect_ip.setCursor(Qt.PointingHandCursor)
        self.btn_detect_ip.setToolTip(
            "Queries a public Geo-IP timezone server to detect your current physical timezone "
            "(useful when not on a corporate full-tunnel VPN)."
        )
        self.btn_detect_ip.clicked.connect(self._on_detect_ip_clicked)
        ip_row.addWidget(self.btn_detect_ip)

        self.ip_status_lbl = QLabel("")
        self.ip_status_lbl.setStyleSheet(f"font-size: {T.SIZE_CAPTION}px; color: {C.TEXT_SECONDARY};")
        ip_row.addWidget(self.ip_status_lbl, 1)

        self.btn_apply_ip = QPushButton("Apply to Venue")
        self.btn_apply_ip.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_apply_ip.setCursor(Qt.PointingHandCursor)
        self.btn_apply_ip.setVisible(False)
        self.btn_apply_ip.clicked.connect(self._on_apply_ip_to_primary)
        ip_row.addWidget(self.btn_apply_ip)

        ref_layout.addLayout(ip_row)
        root.addWidget(ref_card)

        # --- 2. Scrollable Active Venues List ---
        venues_hdr_row = QHBoxLayout()
        venues_hdr = QLabel("<b>📍 3GPP Meeting Venues</b>")
        venues_hdr.setStyleSheet(f"font-size: {T.SIZE_SECTION_TITLE}px; color: {C.TEXT_PRIMARY};")
        venues_hdr_row.addWidget(venues_hdr)
        venues_hdr_row.addStretch(1)

        self.btn_sync_all = QPushButton("🔄 Sync All Venues from Web")
        self.btn_sync_all.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_sync_all.setCursor(Qt.PointingHandCursor)
        self.btn_sync_all.clicked.connect(self._on_sync_all_venues)
        venues_hdr_row.addWidget(self.btn_sync_all)
        root.addLayout(venues_hdr_row)

        self.venues_scroll = QScrollArea()
        self.venues_scroll.setWidgetResizable(True)
        self.venues_scroll.setFrameShape(QFrame.NoFrame)
        self.venues_scroll.setStyleSheet("QScrollArea { background: transparent; }")

        self.venues_container = QWidget()
        self.venues_container.setStyleSheet("background: transparent;")
        self.venues_layout = QVBoxLayout(self.venues_container)
        self.venues_layout.setContentsMargins(0, 0, 4, 0)
        self.venues_layout.setSpacing(8)
        self.venues_scroll.setWidget(self.venues_container)
        root.addWidget(self.venues_scroll, 1)

        # --- 3. Working Group Filter & Bar Display Preferences ---
        pref_card = QFrame()
        pref_card.setStyleSheet(CARD_FRAME_STYLE)
        pref_layout = QVBoxLayout(pref_card)
        pref_layout.setContentsMargins(14, 10, 14, 10)
        pref_layout.setSpacing(6)

        pref_title = QLabel("<b>🎯 Status Bar Display & Working Group Filter</b>")
        pref_title.setStyleSheet(f"font-size: {T.SIZE_BODY}px; color: {C.TEXT_PRIMARY};")
        pref_layout.addWidget(pref_title)

        wg_row = QHBoxLayout()
        wg_row.setSpacing(10)
        self.wg_mode_group = QButtonGroup(self)
        self.rb_all_wgs = QRadioButton("All Active Working Groups")
        self.rb_selected_wgs = QRadioButton("Only Selected WGs:")
        self.wg_mode_group.addButton(self.rb_all_wgs, 0)
        self.wg_mode_group.addButton(self.rb_selected_wgs, 1)

        self.wg_combo = CheckableComboBox("WGs", self)
        self.wg_combo.setStyleSheet(SPEC_FILTER_COMBO_STYLE)
        self.wg_combo.setMinimumWidth(180)
        self.wg_combo.addItems(self.service.get_available_working_groups())

        wg_row.addWidget(self.rb_all_wgs)
        wg_row.addWidget(self.rb_selected_wgs)
        wg_row.addWidget(self.wg_combo)
        wg_row.addStretch(1)
        pref_layout.addLayout(wg_row)

        opts_row = QHBoxLayout()
        opts_row.setSpacing(16)
        self.chk_show_utc = QCheckBox("Show UTC clock in status bar")
        self.chk_show_upcoming = QCheckBox("Show next upcoming venue during off-weeks")
        opts_row.addWidget(self.chk_show_utc)
        opts_row.addWidget(self.chk_show_upcoming)
        opts_row.addStretch(1)
        pref_layout.addLayout(opts_row)

        root.addWidget(pref_card)

        # --- 4. Footer ---
        footer = QHBoxLayout()
        self.footer_hint = QLabel("Overrides save to both the meeting folder (meeting_timezone.json) and app config.")
        self.footer_hint.setStyleSheet(f"font-size: {T.SIZE_CAPTION}px; color: {C.TEXT_MUTED};")
        footer.addWidget(self.footer_hint, 1)

        self.btn_close = QPushButton("Close")
        self.btn_close.setObjectName("primaryBtn")
        self.btn_close.setCursor(Qt.PointingHandCursor)
        self.btn_close.clicked.connect(self.accept)
        footer.addWidget(self.btn_close)
        root.addLayout(footer)

        # Wire preference change signals after widgets exist
        self.rb_all_wgs.toggled.connect(self._on_preferences_edited)
        self.rb_selected_wgs.toggled.connect(self._on_preferences_edited)
        self.wg_combo.selectionChanged.connect(self._on_wg_combo_changed)
        self.chk_show_utc.toggled.connect(self._on_preferences_edited)
        self.chk_show_upcoming.toggled.connect(self._on_preferences_edited)

    def _load_preferences_into_ui(self) -> None:
        cfg = self.service.settings.get_clock_config()
        wg_filter = [str(w).strip() for w in (cfg.get("wg_filter") or []) if str(w).strip()]

        for w in (self.rb_all_wgs, self.rb_selected_wgs, self.wg_combo, self.chk_show_utc, self.chk_show_upcoming):
            w.blockSignals(True)

        self.chk_show_utc.setChecked(bool(cfg.get("show_utc_in_bar", True)))
        self.chk_show_upcoming.setChecked(bool(cfg.get("show_upcoming_when_idle", True)))

        if wg_filter:
            self.rb_selected_wgs.setChecked(True)
            self.wg_combo.setEnabled(True)
            selected_upper = {w.upper() for w in wg_filter}
            model = self.wg_combo.model()
            all_checked = True
            for i in range(1, model.rowCount()):
                item = model.item(i)
                if item:
                    is_sel = str(item.data(Qt.UserRole) or item.text()).strip().upper() in selected_upper
                    item.setCheckState(Qt.Checked if is_sel else Qt.Unchecked)
                    if not is_sel:
                        all_checked = False
            root_item = model.item(0)
            if root_item:
                root_item.setCheckState(Qt.Checked if all_checked else Qt.Unchecked)
            self.wg_combo._refresh_text_and_tooltip()
            self.wg_combo.update()
        else:
            self.rb_all_wgs.setChecked(True)
            self.wg_combo.setEnabled(False)

        for w in (self.rb_all_wgs, self.rb_selected_wgs, self.wg_combo, self.chk_show_utc, self.chk_show_upcoming):
            w.blockSignals(False)

    def _on_wg_combo_changed(self, _checked: list) -> None:
        if not self.rb_selected_wgs.isChecked():
            self.rb_selected_wgs.blockSignals(True)
            self.rb_selected_wgs.setChecked(True)
            self.rb_selected_wgs.blockSignals(False)
        self._on_preferences_edited()

    def _on_preferences_edited(self) -> None:
        use_selected = self.rb_selected_wgs.isChecked()
        self.wg_combo.setEnabled(use_selected)

        cfg = self.service.settings.get_clock_config()
        if use_selected:
            checked = self.wg_combo.getCheckedItems()
            total_wgs = max(0, self.wg_combo.model().rowCount() - 1)
            if len(checked) == total_wgs and total_wgs > 0:
                cfg["wg_filter"] = []
            else:
                cfg["wg_filter"] = list(checked)
        else:
            cfg["wg_filter"] = []

        cfg["show_utc_in_bar"] = self.chk_show_utc.isChecked()
        cfg["show_upcoming_when_idle"] = self.chk_show_upcoming.isChecked()
        self.service.settings.save_clock_config(cfg)
        self.service.invalidate_cache()
        self._refresh_snapshot(rebuild_cards=True)
        self.settings_changed.emit()

    def _on_tick(self) -> None:
        self._refresh_snapshot(rebuild_cards=False)

    def _refresh_snapshot(self, rebuild_cards: bool = False) -> None:
        snapshot = self.service.get_snapshot()
        self._update_reference_clocks(snapshot)

        focused_state: Optional[VenueClockState] = None
        if self.focus_meeting is not None:
            focused_state = self.service.resolve_single_meeting(
                self.focus_meeting, meeting_dir=self.focus_meeting_dir, ref_utc=snapshot.utc_dt
            )

        sig = (
            focused_state.location.lower() if focused_state else "",
            tuple((v.location.lower(), v.is_upcoming) for v in snapshot.primary_venues),
            tuple((v.location.lower(), v.is_upcoming) for v in snapshot.other_venues),
        )

        if rebuild_cards or sig != self._last_layout_signature:
            self._rebuild_venue_cards(snapshot, focused_state)
            self._last_layout_signature = sig
        else:
            self._update_existing_cards(snapshot, focused_state)

    def _update_reference_clocks(self, snapshot: WorldClockSnapshot) -> None:
        self.laptop_time_lbl.setText(snapshot.laptop_dt.strftime("%H:%M:%S"))
        self.laptop_sub_lbl.setText(
            f"{snapshot.laptop_dt:%a %d %b %Y}  ·  {format_utc_offset(snapshot.laptop_offset_minutes)}"
        )
        self.utc_time_lbl.setText(f"{snapshot.utc_dt:%H:%M:%S} UTC")
        self.utc_sub_lbl.setText(f"{snapshot.utc_dt:%a %d %b %Y}")

    def _clear_venues_layout(self) -> None:
        self._venue_cards.clear()
        while self.venues_layout.count():
            item = self.venues_layout.takeAt(0)
            w = item.widget()
            if w is not None:
                w.deleteLater()

    def _make_card(
        self,
        venue: VenueClockState,
        explicit_dir: Optional[Path] = None,
        is_focused: bool = False,
    ) -> _VenueClockCard:
        card = _VenueClockCard(
            venue=venue,
            explicit_meeting_dir=explicit_dir,
            is_focused_meeting=is_focused,
            parent=self.venues_container,
        )
        card.offset_override_requested.connect(self._on_offset_override)
        card.nudge_requested.connect(self._on_nudge_offset)
        card.web_sync_requested.connect(self._on_web_sync_venue)
        card.reset_requested.connect(self._on_reset_venue)
        self._venue_cards.append((venue.location.lower(), is_focused, card))
        return card

    def _rebuild_venue_cards(
        self,
        snapshot: WorldClockSnapshot,
        focused_state: Optional[VenueClockState],
    ) -> None:
        self._clear_venues_layout()

        if focused_state is not None:
            card = self._make_card(focused_state, explicit_dir=self.focus_meeting_dir, is_focused=True)
            self.venues_layout.addWidget(card)

        if snapshot.primary_venues:
            for v in snapshot.primary_venues:
                if (
                    focused_state is not None
                    and v.location.lower() == focused_state.location.lower()
                    and len(snapshot.primary_venues) == 1
                ):
                    continue
                self.venues_layout.addWidget(self._make_card(v))
        elif focused_state is None:
            empty_lbl = QLabel(
                "No in-person 3GPP meetings are active right now for the selected Working Group filter."
            )
            empty_lbl.setWordWrap(True)
            empty_lbl.setAlignment(Qt.AlignCenter)
            empty_lbl.setStyleSheet(
                f"color: {C.TEXT_MUTED}; font-size: {T.SIZE_BODY}px; padding: 24px; "
                f"background: {C.BG_SURFACE}; border: 1px dashed {C.BORDER_STRONG}; border-radius: {R.LG}px;"
            )
            self.venues_layout.addWidget(empty_lbl)

        if snapshot.other_venues:
            other_hdr = QLabel("<b>Other Active 3GPP Venues (Hidden from Status Bar by WG Filter)</b>")
            other_hdr.setStyleSheet(
                f"font-size: {T.SIZE_CAPTION}px; color: {C.TEXT_MUTED}; margin-top: 8px;"
            )
            self.venues_layout.addWidget(other_hdr)
            for v in snapshot.other_venues:
                self.venues_layout.addWidget(self._make_card(v))

        self.venues_layout.addStretch(1)

    def _update_existing_cards(
        self,
        snapshot: WorldClockSnapshot,
        focused_state: Optional[VenueClockState],
    ) -> None:
        by_loc: Dict[str, VenueClockState] = {}
        for v in snapshot.primary_venues + snapshot.other_venues:
            by_loc[v.location.lower()] = v

        for loc_key, is_focused, card in self._venue_cards:
            if is_focused and focused_state is not None:
                card.update_state(focused_state)
            elif loc_key in by_loc:
                card.update_state(by_loc[loc_key])

    def _apply_manual_offset(
        self,
        location: str,
        offset_min: int,
        meetings: list,
        explicit_dir: Optional[Path],
    ) -> None:
        label = f"{format_utc_offset(offset_min)} (Manual Override)"
        if explicit_dir is not None and self.focus_meeting is not None:
            SessionPlanTimezone.save_meeting_folder_timezone(
                meeting_dir=explicit_dir,
                tz_offset_minutes=offset_min,
                tz_label=label,
                source="manual_override",
                location=location,
                wg=str(self.focus_meeting.get("wg_name", "")),
                meeting_number=str(self.focus_meeting.get("meeting_number", "")),
            )
        self.service.settings.set_venue_timezone_override(
            location=location,
            tz_offset_minutes=offset_min,
            tz_label=label,
            source="manual_override",
            matching_meetings=meetings,
        )
        self.service.invalidate_cache()
        self._refresh_snapshot(rebuild_cards=False)
        self.settings_changed.emit()

    def _on_offset_override(self, location: str, offset_min: int, meetings: list, explicit_dir) -> None:
        self._apply_manual_offset(location, offset_min, meetings, explicit_dir)

    def _on_nudge_offset(self, location: str, delta_min: int, meetings: list, explicit_dir) -> None:
        cur_off, _lbl, _iana, _src = self.service.resolve_venue_timezone(
            location, meetings=meetings, explicit_meeting_dir=explicit_dir
        )
        base = cur_off if cur_off is not None else 0
        new_off = max(-12 * 60, min(14 * 60, int(base) + int(delta_min)))
        self._apply_manual_offset(location, new_off, meetings, explicit_dir)

    def _on_reset_venue(self, location: str, meetings: list, explicit_dir) -> None:
        if explicit_dir is not None:
            SessionPlanTimezone.clear_meeting_folder_timezone(Path(explicit_dir))
        self.service.settings.clear_venue_timezone_override(
            location=location, matching_meetings=meetings
        )
        self.service.invalidate_cache()
        self._refresh_snapshot(rebuild_cards=False)
        self.settings_changed.emit()

    def _on_web_sync_venue(self, location: str, meetings: list, explicit_dir) -> None:
        for loc_key, _isf, card in self._venue_cards:
            if loc_key == location.lower():
                card.set_syncing(True)

        worker = WebTimezoneSyncWorker(mode="city", location=location, meetings=meetings, parent=self)
        worker.city_resolved.connect(
            lambda loc, off, lbl, iana, mtgs, ed=explicit_dir: self._on_city_sync_done(
                loc, off, lbl, iana, mtgs, ed
            )
        )
        worker.sync_failed.connect(self._on_city_sync_failed)
        worker.finished.connect(lambda w=worker: self._cleanup_worker(w))
        self._active_workers.append(worker)
        worker.start()

    def _on_sync_all_venues(self) -> None:
        seen = set()
        for loc_key, _isf, card in self._venue_cards:
            if loc_key in seen:
                continue
            seen.add(loc_key)
            self._on_web_sync_venue(card.venue.location, card.venue.meetings, card.explicit_meeting_dir)

    def _on_city_sync_done(
        self,
        location: str,
        offset_min: int,
        label: str,
        iana_tz: str,
        meetings: list,
        explicit_dir: Optional[Path],
    ) -> None:
        for loc_key, _isf, card in self._venue_cards:
            if loc_key == location.lower():
                card.set_syncing(False)

        # Explicit user Web Sync replaces any prior manual lock so fresh web verification takes effect
        self.service.settings.clear_venue_timezone_override(
            location=location, matching_meetings=meetings
        )
        if explicit_dir is not None:
            SessionPlanTimezone.save_meeting_folder_timezone(
                meeting_dir=Path(explicit_dir),
                tz_offset_minutes=offset_min,
                tz_label=label,
                source="web_api",
                location=location,
                wg=str((self.focus_meeting or {}).get("wg_name", "")),
                meeting_number=str((self.focus_meeting or {}).get("meeting_number", "")),
                iana_tz=iana_tz,
            )

        self.service.settings.set_venue_timezone_override(
            location=location,
            tz_offset_minutes=offset_min,
            tz_label=label,
            source="web_api",
            iana_tz=iana_tz,
            matching_meetings=meetings,
        )
        self.service.invalidate_cache()
        self._refresh_snapshot(rebuild_cards=False)
        self.settings_changed.emit()

    def _on_city_sync_failed(self, target: str, err: str) -> None:
        for loc_key, _isf, card in self._venue_cards:
            if loc_key == (target or "").lower():
                card.set_syncing(False)
        self.ip_status_lbl.setText(f"⚠️ Web sync failed for {target}: {err}")

    def _on_detect_ip_clicked(self) -> None:
        self.btn_detect_ip.setEnabled(False)
        self.btn_detect_ip.setText("⏳ Detecting...")
        self.btn_apply_ip.setVisible(False)
        self.ip_status_lbl.setText("Querying public Geo-IP timezone server...")

        worker = WebTimezoneSyncWorker(mode="ip", parent=self)
        worker.ip_resolved.connect(self._on_ip_detected)
        worker.sync_failed.connect(self._on_ip_failed)
        worker.finished.connect(lambda w=worker: self._cleanup_worker(w))
        self._active_workers.append(worker)
        worker.start()

    def _on_ip_detected(self, city: str, offset_min: int, label: str, iana_tz: str) -> None:
        self.btn_detect_ip.setEnabled(True)
        self.btn_detect_ip.setText("📡 Detect My Physical Timezone (via Public IP)")
        self._detected_ip_info = (city, offset_min, label, iana_tz)
        self.ip_status_lbl.setText(f"Detected: {city} — {label}")
        self.btn_apply_ip.setVisible(bool(self._venue_cards))

    def _on_ip_failed(self, _target: str, err: str) -> None:
        self.btn_detect_ip.setEnabled(True)
        self.btn_detect_ip.setText("📡 Detect My Physical Timezone (via Public IP)")
        self.ip_status_lbl.setText(f"⚠️ Could not detect IP timezone: {err}")

    def _on_apply_ip_to_primary(self) -> None:
        if not self._detected_ip_info or not self._venue_cards:
            return
        _city, offset_min, _label, _iana = self._detected_ip_info
        _loc_key, _isf, card = self._venue_cards[0]
        self._apply_manual_offset(
            card.venue.location, offset_min, card.venue.meetings, card.explicit_meeting_dir
        )

    def _cleanup_worker(self, worker: WebTimezoneSyncWorker) -> None:
        if worker in self._active_workers:
            self._active_workers.remove(worker)

    def closeEvent(self, event) -> None:
        if hasattr(self, "_tick_timer") and self._tick_timer.isActive():
            self._tick_timer.stop()
        for w in list(self._active_workers):
            if w.isRunning():
                w.requestInterruption()
                w.wait(250)
        self._active_workers.clear()
        super().closeEvent(event)

    def done(self, result: int) -> None:
        if hasattr(self, "_tick_timer") and self._tick_timer.isActive():
            self._tick_timer.stop()
        for w in list(self._active_workers):
            if w.isRunning():
                w.requestInterruption()
                w.wait(250)
        self._active_workers.clear()
        super().done(result)
