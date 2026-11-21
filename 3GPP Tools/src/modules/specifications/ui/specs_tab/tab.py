"""3GPP Specifications Catalogue tab composed from modular feature mixins."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Optional

from PyQt5.QtCore import QTimer, pyqtSignal
from PyQt5.QtWidgets import QAbstractItemView, QComboBox, QLabel, QWidget

from core.ui.font_scale import get_font_scale_notifier
from core.utils.paths import get_project_root
from modules.spec_search.core.spec_search_db import SpecSearchDatabase
from modules.specifications.core.database import SpecsDatabase
from modules.specifications.core.spec_history_db import SpecHistoryDatabase
from .mixins import (
    SpecsTabActionsMixin,
    SpecsTabFiltersMixin,
    SpecsTabRowDocumentsCellMixin,
    SpecsTabTableRenderMixin,
    SpecsTabUiSetupMixin,
)


class SpecificationsTab(
    SpecsTabUiSetupMixin,
    SpecsTabFiltersMixin,
    SpecsTabTableRenderMixin,
    SpecsTabRowDocumentsCellMixin,
    SpecsTabActionsMixin,
    QWidget,
):
    update_db_requested = pyqtSignal(bool)
    update_specific_requested = pyqtSignal(list, bool)
    log_msg = pyqtSignal(str, int)
    search_in_spec_search_requested = pyqtSignal(str, str)  # spec_number, version
    index_in_spec_search_requested = pyqtSignal(str, str)   # spec_number, version
    batch_index_in_spec_search_requested = pyqtSignal(list) # list of (spec_number, version) tuples
    smart_index_in_spec_search_requested = pyqtSignal(str)  # spec_number

    def __init__(self, db_path: Path, config_file: Optional[Path] = None):
        super().__init__()
        self.db = SpecsDatabase(db_path)
        self.spec_search_db = SpecSearchDatabase(get_project_root() / "db" / "3gpp_spec_search.db")
        self.history_db = SpecHistoryDatabase(get_project_root() / "db" / "3gpp_spec_history.db")
        self._download_threads = []

        self.config_file = (
            Path(config_file)
            if config_file is not None
            else (get_project_root() / "config" / "specs_config.json")
        )

        # Defaults
        self.default_dl_dir = str(Path.home() / "3GPP_SA2_Meeting_Helper" / "specs")
        self.download_dir = self.default_dl_dir
        self.table_filters = {"series": "", "tech": "", "group": "", "spec_type": "Any"}
        self._current_tsg = ""
        self._show_all_wgs = False
        self.saved_search = ""
        self.saved_version = ""
        self.saved_downloaded_only = False
        self.favorites = set()
        self.favorites_only = False
        self.tag_filter = None

        self.filter_options = self.db.get_filter_options()

        # Load persisted settings
        self._load_settings()

        # Debounce Timers
        self.search_timer = QTimer(self)
        self.search_timer.setSingleShot(True)
        self.search_timer.setInterval(300)
        self.search_timer.timeout.connect(self.refresh_table)

        self.filter_timer = QTimer(self)
        self.filter_timer.setSingleShot(True)
        self.filter_timer.setInterval(150)
        self.filter_timer.timeout.connect(self.refresh_table)

        self.save_settings_timer = QTimer(self)
        self.save_settings_timer.setSingleShot(True)
        self.save_settings_timer.setInterval(800)
        self.save_settings_timer.timeout.connect(self._save_settings)

        self._setup_ui()
        get_font_scale_notifier().scale_changed.connect(self.on_font_scale_changed)

        # Lazy loading flag: table populates only when the tab becomes active
        self._is_loaded = False

    def ensure_loaded(self):
        """Called automatically when the user selects the Specifications tab."""
        if not self._is_loaded:
            self._is_loaded = True
            self.refresh_table()

    def select_and_show_spec(self, spec_number: str, version: str = None):
        """
        Focuses on a specific specification and version in the specifications table,
        clearing conflicting filters and scrolling directly to the row.
        """
        clean_spec = re.sub(r"^(?:TS|TR)\s*", "", str(spec_number or "").strip(), flags=re.IGNORECASE)
        clean_ver = version.lstrip("vV").strip() if version else None

        self._is_loaded = True
        self.spec_search_input.blockSignals(True)
        self.spec_search_input.setText(clean_spec)
        self.spec_search_input.blockSignals(False)

        self.refresh_table()

        def _find_matching_row() -> int:
            for r in range(self.table.rowCount()):
                spec_widget = self.table.cellWidget(r, 0)
                if spec_widget:
                    spec_label = spec_widget.findChild(QLabel, "specNumberLabel")
                    if spec_label and spec_label.property("specNumber") == clean_spec:
                        return r
            return -1

        matched_row = _find_matching_row()
        if matched_row < 0 and hasattr(self, "_clear_all_filters"):
            # Active Favorites / Downloaded / WG / Series / Tech filters hid the target spec
            self._clear_all_filters()
            self.spec_search_input.blockSignals(True)
            self.spec_search_input.setText(clean_spec)
            self.spec_search_input.blockSignals(False)
            self.refresh_table()
            matched_row = _find_matching_row()

        if matched_row >= 0:
            self.table.selectRow(matched_row)
            item_for_scroll = self.table.item(matched_row, 1)
            if item_for_scroll:
                self.table.scrollToItem(item_for_scroll, QAbstractItemView.PositionAtCenter)

            if clean_ver:
                cell_widget = self.table.cellWidget(matched_row, 2)
                if cell_widget:
                    combo = cell_widget.findChild(QComboBox)
                    if combo:
                        for i in range(combo.count()):
                            c_data = combo.itemData(i) or {}
                            if c_data.get("version") == clean_ver or f"v{clean_ver}" in combo.itemText(i):
                                combo.setCurrentIndex(i)
                                break
