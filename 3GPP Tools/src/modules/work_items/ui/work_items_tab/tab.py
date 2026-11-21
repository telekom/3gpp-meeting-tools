"""WorkItemsTab shell widget combining Work Items mixins and lifecycle hooks."""

import logging
from pathlib import Path
from typing import Optional

from PyQt5.QtCore import QTimer, pyqtSignal
from PyQt5.QtWidgets import QWidget

from core.ui.font_scale import get_font_scale_notifier
from modules.work_items.core.wi_database import WorkItemsDatabase
from modules.work_items.core.wi_settings import WorkItemsSettings
from modules.work_items.ui.work_items_tab.mixins import (
    WorkItemsDataSyncMixin,
    WorkItemsDialogsMenusMixin,
    WorkItemsFiltersTrackingMixin,
    WorkItemsUiSetupMixin,
)

logger = logging.getLogger("modules.work_items.ui.ui_tabs")


class WorkItemsTab(
    WorkItemsUiSetupMixin,
    WorkItemsFiltersTrackingMixin,
    WorkItemsDataSyncMixin,
    WorkItemsDialogsMenusMixin,
    QWidget,
):
    """Main 3GPP Work Items browser, tracker, and synchronization tab."""

    global_action_requested = pyqtSignal(str, str)

    def __init__(self, db_path: Path, config_file: Optional[Path] = None):
        super().__init__()
        self.db_path = db_path
        self.db = WorkItemsDatabase(db_path)
        self.settings = WorkItemsSettings(config_file=config_file)

        self._chunk_size = 60
        self._current_offset = 0
        self._total_count = 0
        self._loaded_data = []
        self._is_loading_chunk = False

        # Tracking state (Personal Interest & Delegation Assignment)
        self.starred, self.following = self.settings.get_tracking_sets()
        self.starred_only = False
        self.following_only = False
        self._current_tsg = ""
        self.tsg_buttons = {}
        self._rendered_tag_labels = None

        self.search_timer = QTimer()
        self.search_timer.setSingleShot(True)
        self.search_timer.setInterval(350)
        self.search_timer.timeout.connect(lambda: self.refresh_table(reset_pagination=True))

        self.save_filters_timer = QTimer()
        self.save_filters_timer.setSingleShot(True)
        self.save_filters_timer.setInterval(1000)
        self.save_filters_timer.timeout.connect(self._save_filters)

        self._setup_ui()
        get_font_scale_notifier().scale_changed.connect(self.on_font_scale_changed)
        self._populate_filters()
        self._load_filters()
        self._update_chip_styles()
        self.refresh_table(reset_pagination=True)

    def shutdown(self):
        if hasattr(self, "search_timer"):
            self.search_timer.stop()
        if hasattr(self, "save_filters_timer"):
            self.save_filters_timer.stop()
        for th in (getattr(self, "scraper_thread", None), getattr(self, "targeted_thread", None)):
            if th and th.isRunning():
                th.requestInterruption()
                th.quit()
                th.wait(150)

    def closeEvent(self, event):
        self.shutdown()
        super().closeEvent(event)
