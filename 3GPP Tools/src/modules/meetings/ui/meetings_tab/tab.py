"""Main MeetingsTab widget composition and lifecycle shutdown."""

from pathlib import Path
from typing import Optional

from PyQt5.QtCore import QTimer, pyqtSignal
from PyQt5.QtWidgets import QWidget

from core.ui.font_scale import get_font_scale_notifier
from modules.meetings.core.meetings_db import MeetingsDatabase
from modules.meetings.core.settings import MeetingsSettings
from modules.meetings.ui.search_controller import GlobalSearchController

from .mixins import (
    MeetingsTabFiltersDataMixin,
    MeetingsTabQuickTdocMenusMixin,
    MeetingsTabSyncMixin,
    MeetingsTabUiSetupMixin,
    MeetingsTabWindowsActionsMixin,
)


class MeetingsTab(
    MeetingsTabUiSetupMixin,
    MeetingsTabSyncMixin,
    MeetingsTabFiltersDataMixin,
    MeetingsTabQuickTdocMenusMixin,
    MeetingsTabWindowsActionsMixin,
    QWidget,
):
    update_db_requested = pyqtSignal(bool, bool, bool)
    update_specific_requested = pyqtSignal(list, bool, bool, bool)
    update_meetings_requested = pyqtSignal(dict)
    abort_sync_requested = pyqtSignal()

    def __init__(self, db_path: Path, config_file: Optional[Path] = None):
        super().__init__()
        self.db = MeetingsDatabase(db_path)
        self.settings = MeetingsSettings(config_file=config_file)
        self.search_controller = GlobalSearchController(self)

        self.tdoc_windows = {}
        self._active_dl_threads = {}
        self._current_tsg = ""

        self.save_filters_timer = QTimer()
        self.save_filters_timer.setSingleShot(True)
        self.save_filters_timer.setInterval(1000)
        self.save_filters_timer.timeout.connect(self._save_filters)

        self._setup_ui()
        get_font_scale_notifier().scale_changed.connect(self.on_font_scale_changed)
        self.search_controller.connect_signals()

        self._load_filters()
        self._update_last_meeting_btn()
        self.refresh_table()

    def shutdown(self):
        """Cleanly shuts down all child windows, dialogs, and background threads."""
        if hasattr(self, 'save_filters_timer') and self.save_filters_timer.isActive():
            self.save_filters_timer.stop()
            self._save_filters()
        if hasattr(self, 'contribution_dialog') and self.contribution_dialog:
            try:
                self.contribution_dialog.close()
            except Exception:
                pass
        if hasattr(self, 'agreement_tracker_dialog') and self.agreement_tracker_dialog:
            try:
                self.agreement_tracker_dialog.close()
            except Exception:
                pass

        for win in list(self.tdoc_windows.values()):
            if win:
                if getattr(win, '_session_plan_window', None):
                    try:
                        win._session_plan_window.close()
                    except Exception:
                        pass
                try:
                    win.close()
                except Exception:
                    pass
        self.tdoc_windows.clear()

        tab_threads = [
            getattr(self, 'agenda_dl_thread', None),
            getattr(self, 'cacher_thread', None),
            getattr(self, 'merger_thread', None),
            getattr(self, 'cmp_thread', None),
        ]
        if hasattr(self, '_active_dl_threads'):
            tab_threads.extend(self._active_dl_threads.values())

        for th in filter(None, tab_threads):
            if hasattr(th, 'is_cancelled'):
                th.is_cancelled = True
            if th.isRunning():
                th.requestInterruption()
                th.quit()
                th.wait(150)

    def closeEvent(self, event):
        self.shutdown()
        super().closeEvent(event)
