"""WorldClockDialog main class assembling UI setup and sync/override mixins."""

from __future__ import annotations

from pathlib import Path
from typing import List, Optional, Tuple

from PyQt5.QtCore import QTimer, pyqtSignal
from PyQt5.QtWidgets import QDialog, QWidget

from core.ui.world_clock.mixins import (
    WorldClockSyncActionsMixin,
    WorldClockUiSetupMixin,
)
from core.ui.world_clock.venue_card import _VenueClockCard
from core.ui.world_clock_service import WebTimezoneSyncWorker, WorldClockService


class WorldClockDialog(
    WorldClockUiSetupMixin,
    WorldClockSyncActionsMixin,
    QDialog,
):
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
