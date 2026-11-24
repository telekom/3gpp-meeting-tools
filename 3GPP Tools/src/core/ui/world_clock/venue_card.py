"""Venue clock card widget and timezone badge/delta formatting helpers."""

from __future__ import annotations

from pathlib import Path
from typing import List, Optional, Tuple

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtWidgets import (
    QComboBox,
    QFrame,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from core.ui.theme import C, R, T
from core.ui.ui_components import BUTTON_STYLE_TOOLBAR_SECONDARY
from core.ui.world_clock_service import VenueClockState, _format_wg_summary
from modules.meetings.core.session_plan.session_plan_timezone import format_utc_offset


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
