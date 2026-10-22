# --- File: src/modules/specifications/ui/ui_tabs.py ---
import json
import logging
import os
import re
import webbrowser
import zipfile
from pathlib import Path

from PyQt5 import sip
from PyQt5.QtCore import QTimer, Qt, pyqtSignal
from PyQt5.QtGui import QKeySequence
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QAction,
    QCheckBox,
    QComboBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMenu,
    QMessageBox,
    QPushButton,
    QShortcut,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget, QDialog,
)

from core.ui.ui_components import BUTTON_STYLE_TOOLBAR_SECONDARY, BUTTON_STYLE_TOOLBAR_DANGER
from core.utils.paths import get_project_root, to_extended_path
from modules.meetings.ui.tdocs_components import CheckableComboBox
from modules.specifications.core.database import SpecsDatabase
from modules.specifications.ui.components import HoverMenuButton
from modules.specifications.ui.dialogs import (
    AddSpecDialog,
    AdvancedSyncDialog,
    SpecClauseExcerptDialog,
    SpecInfoDialog,
    SpecsConfigDialog,
    TableFilterDialog,
    TargetedSyncDialog,
)
from modules.specifications.ui.spec_annotations_dialog import (
    BatchTagDialog,
    EditSpecAnnotationsDialog,
)
from modules.specifications.ui.spec_action_controller import (
    handle_document_action,
    handle_zip_action,
)
from modules.specifications.ui.spec_menu import (
    copy_specs_default,
    get_saved_copy_mode,
    populate_spec_copy_menu,
    save_copy_mode,
    show_spec_context_menu,
)
from modules.specifications.ui.threads import SpecDownloadThread
from modules.specifications.utils.utils import build_dynareport_url
from modules.word_tools.core.document_opener import open_document

logger = logging.getLogger(__name__)

PART_NUM_REGEX = re.compile(r"\d{4,5}-(\d{1,2})(?:[-_.]|$)")

# --- STYLES FOR SPECIFICATIONS FILTER CONTROLS ---
SPEC_FILTER_COMBO_STYLE = """
QComboBox {
    padding: 3px 8px;
    border: 1px solid #CBD5E1;
    border-radius: 4px;
    font-size: 12px;
    background-color: #FFFFFF;
    color: #1E293B;
}
QComboBox:hover {
    border-color: #94A3B8;
}
QComboBox:focus {
    border-color: #1E5C99;
}
"""

SPEC_FILTER_INPUT_STYLE = """
QLineEdit {
    padding: 4px 8px;
    border: 1px solid #CBD5E1;
    border-radius: 4px;
    font-size: 12px;
    background-color: #FFFFFF;
    color: #1E293B;
}
QLineEdit:focus {
    border: 1px solid #1E5C99;
}
"""

SEGMENT_BTN_STYLE = """
QPushButton {
    padding: 3px 9px;
    font-size: 11px;
    font-weight: 500;
    background-color: #F8FAFC;
    color: #334155;
    border: 1px solid #CBD5E1;
    border-radius: 4px;
}
QPushButton:hover {
    background-color: #F1F5F9;
    border-color: #94A3B8;
}
QPushButton:checked {
    background-color: #0066CC;
    color: #FFFFFF;
    font-weight: bold;
    border: 1px solid #0055AA;
}
"""

TOGGLE_PILL_STYLE = """
QPushButton {
    padding: 3px 10px;
    font-size: 11px;
    font-weight: 500;
    background-color: #F8FAFC;
    color: #334155;
    border: 1px solid #CBD5E1;
    border-radius: 12px;
}
QPushButton:hover {
    background-color: #F1F5F9;
    border-color: #94A3B8;
}
QPushButton:checked {
    background-color: #EBF3FC;
    color: #1E5C99;
    font-weight: bold;
    border: 1.5px solid #1E5C99;
}
QPushButton:checked:hover {
    background-color: #DFEDFA;
}
"""

FAV_PILL_STYLE = """
QPushButton {
    padding: 3px 10px;
    font-size: 11px;
    font-weight: 500;
    background-color: #F8FAFC;
    color: #334155;
    border: 1px solid #CBD5E1;
    border-radius: 12px;
}
QPushButton:hover {
    background-color: #F1F5F9;
    border-color: #94A3B8;
}
QPushButton:checked {
    background-color: #FEF3C7;
    color: #92400E;
    font-weight: bold;
    border: 1.5px solid #F59E0B;
}
QPushButton:checked:hover {
    background-color: #FDE68A;
}
"""

DISMISSIBLE_TAG_STYLE = """
QPushButton {
    padding: 2px 8px;
    font-size: 11px;
    font-weight: 500;
    background-color: #EFF6FF;
    color: #1E40AF;
    border: 1px solid #BFDBFE;
    border-radius: 10px;
}
QPushButton:hover {
    background-color: #FEE2E2;
    color: #991B1B;
    border-color: #FCA5A5;
}
"""

CLEAR_ALL_BTN_STYLE = """
QPushButton {
    padding: 2px 8px;
    font-size: 11px;
    font-weight: bold;
    background-color: #FEF2F2;
    color: #B91C1C;
    border: 1px solid #FECACA;
    border-radius: 4px;
}
QPushButton:hover {
    background-color: #FEE2E2;
    border-color: #F87171;
}
"""

WG_DESCRIPTIONS = {
    # SA
    "SA": "Service and System Aspects (Plenary)",
    "SA1": "Services",
    "SA2": "Architecture",
    "SA3": "Security",
    "SA4": "Codec",
    "SA5": "Telecom Management",
    "SA6": "Mission Critical & Vertical Applications",
    # RAN
    "RAN": "Radio Access Network (Plenary)",
    "RAN1": "Radio Layer 1 (Physical layer)",
    "RAN2": "Radio Layer 2 & 3",
    "RAN3": "Architecture & Network Interfaces",
    "RAN4": "Radio Performance & Protocol Aspects",
    "RAN5": "Mobile Terminal Conformance Testing",
    "RAN6": "GSM/EDGE Radio Access Network",
    # CT
    "CT": "Core Network and Terminals (Plenary)",
    "CT1": "User Equipment - Core Network Protocols",
    "CT3": "Interworking with External Networks",
    "CT4": "Core Network Protocols",
    "CT6": "Smart Card Application Aspects",
    # Legacy / Other
    "GERAN": "GSM EDGE Radio Access Network",
    "GERAN1": "GERAN Radio Aspects",
    "GERAN2": "GERAN Protocol Aspects",
    "GERAN3": "GERAN Base Station Testing",
    "CN1": "Core Network 1 (Legacy)",
    "CN3": "Core Network 3 (Legacy)",
    "CN4": "Core Network 4 (Legacy)",
    "CN5": "Open Service Access (Legacy)",
    "T2": "Mobile Terminal Services (Legacy)",
    "T3": "USIM / Smart Cards (Legacy)",
    "SMG01": "Special Mobile Group 1 (Legacy)",
    "SMG03": "Special Mobile Group 3 (Legacy)",
}


def get_wg_badge_style_and_desc(wg: str) -> tuple:
    wg_upper = str(wg or "").strip().upper()
    desc = WG_DESCRIPTIONS.get(wg_upper, "Working Group")

    # SA (Soft Indigo / Blue)
    if wg_upper.startswith("SA"):
        style = """
            background-color: #EEF2FF;
            color: #3730A3;
            border: 1px solid #C7D2FE;
            border-radius: 3px;
            padding: 1px 0px;
            font-size: 10px;
            font-weight: bold;
        """
    # RAN (Soft Emerald / Green)
    elif wg_upper.startswith("RAN") or wg_upper.startswith("GERAN"):
        style = """
            background-color: #ECFDF5;
            color: #065F46;
            border: 1px solid #A7F3D0;
            border-radius: 3px;
            padding: 1px 0px;
            font-size: 10px;
            font-weight: bold;
        """
    # CT (Soft Amber / Warm Orange)
    elif wg_upper.startswith("CT") or wg_upper.startswith("CN"):
        style = """
            background-color: #FFF7ED;
            color: #9A3412;
            border: 1px solid #FED7AA;
            border-radius: 3px;
            padding: 1px 0px;
            font-size: 10px;
            font-weight: bold;
        """
    # Default Slate
    else:
        style = """
            background-color: #F1F5F9;
            color: #475569;
            border: 1px solid #CBD5E1;
            border-radius: 3px;
            padding: 1px 0px;
            font-size: 10px;
            font-weight: bold;
        """
    return style, desc


class SpecificationsTab(QWidget):
    update_db_requested = pyqtSignal(bool)
    update_specific_requested = pyqtSignal(list, bool)
    log_msg = pyqtSignal(str, int)
    search_in_spec_search_requested = pyqtSignal(str, str)  # spec_number, version
    index_in_spec_search_requested = pyqtSignal(str, str)   # spec_number, version

    def __init__(self, db_path: Path):
        super().__init__()
        self.db = SpecsDatabase(db_path)
        self._download_threads = []

        self.config_file = get_project_root() / "config" / "specs_config.json"

        # Defaults
        self.default_dl_dir = str(Path.home() / "3GPP_SA2_Meeting_Helper" / "specs")
        self.download_dir = self.default_dl_dir
        self.table_filters = {"series": "", "tech": "", "group": "", "spec_type": "Any"}
        self._current_tsg = ""
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
        clean_spec = spec_number.strip()
        clean_ver = version.lstrip("v").strip() if version else None

        self._is_loaded = True
        self.spec_search_input.blockSignals(True)
        self.spec_search_input.setText(clean_spec)
        self.spec_search_input.blockSignals(False)

        self.refresh_table()

        # Find matching row in self.table
        for r in range(self.table.rowCount()):
            spec_widget = self.table.cellWidget(r, 0)
            if spec_widget:
                spec_label = spec_widget.findChild(QLabel, "specNumberLabel")
                if spec_label and spec_label.property("specNumber") == clean_spec:
                    self.table.selectRow(r)
                    item_for_scroll = self.table.item(r, 1)
                    if item_for_scroll:
                        self.table.scrollToItem(item_for_scroll, QAbstractItemView.PositionAtCenter)

                    if clean_ver:
                        cell_widget = self.table.cellWidget(r, 2)
                        if cell_widget:
                            combo = cell_widget.findChild(QComboBox)
                            if combo:
                                for i in range(combo.count()):
                                    c_data = combo.itemData(i) or {}
                                    if c_data.get("version") == clean_ver or f"v{clean_ver}" in combo.itemText(i):
                                        combo.setCurrentIndex(i)
                                        break
                    break

    # ==========================================
    # --- CONFIG / PERSISTENCE ---
    # ==========================================
    def _load_settings(self):
        if self.config_file.exists():
            try:
                with open(self.config_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    self.download_dir = data.get("download_dir", self.default_dl_dir)
                    self.saved_search = data.get("search_query", "")
                    self.saved_version = data.get("version_query", "")
                    self.saved_downloaded_only = data.get("downloaded_only", False)
                    self.favorites = set(data.get("favorites", []))
                    self.favorites_only = data.get("favorites_only", False)
                    self.copy_link_mode = data.get("copy_link_mode", "both")
                    saved_filters = data.get("table_filters", {})
                    if isinstance(saved_filters, dict):
                        self.table_filters.update(saved_filters)
            except Exception as e:
                logger.error(f"Error loading specs_config.json: {e}", exc_info=True)
        else:
            self.download_dir = self.default_dl_dir
            self.copy_link_mode = "both"

    def _save_settings(self):
        try:
            data = {
                "download_dir": self.download_dir,
                "search_query": self.spec_search_input.text().strip(),
                "version_query": self.version_search_input.text().strip(),
                "downloaded_only": self.downloaded_only_cb.isChecked(),
                "favorites": sorted(list(self.favorites)),
                "favorites_only": self.favorites_only,
                "copy_link_mode": getattr(self, "copy_link_mode", "both"),
                "table_filters": self.table_filters,
            }
            with open(self.config_file, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=4)
        except Exception as e:
            logger.error(f"Error saving specs_config.json: {e}", exc_info=True)

    # ==========================================
    # --- UI SETUP ---
    # ==========================================
    def _setup_ui(self):
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(8, 8, 8, 8)
        main_layout.setSpacing(6)

        # --- ROW 1: Network Sync & Configuration Toolbar ---
        toolbar_layout = QHBoxLayout()
        toolbar_layout.setSpacing(8)

        toolbar_layout.addWidget(QLabel("<b>🌐 Network Sync:</b>"))

        self.sync_menu_btn = QPushButton("🌐 Sync && Fetch ▾")
        self.sync_menu_btn.setCursor(Qt.PointingHandCursor)
        self.sync_menu_btn.setStyleSheet("""
                    QPushButton {
                        font-weight: bold;
                        background-color: #0066CC;
                        color: #FFFFFF;
                        border: 1px solid #0055AA;
                        border-radius: 4px;
                        padding: 5px 12px;
                    }
                    QPushButton:hover {
                        background-color: #0052A3;
                    }
                """)

        sync_menu = QMenu(self)
        sync_menu.setStyleSheet("""
                    QMenu { background-color: #FAFAFA; border: 1px solid #CCC; }
                    QMenu::item { padding: 6px 20px 6px 15px; color: #333333; }
                    QMenu::item:selected { background-color: #E1F0FF; color: #0078D7; }
                """)

        quick_action = sync_menu.addAction("🎯 Quick Fetch (By Spec or Series)...")
        quick_action.setToolTip("Instantly download specific specs or series (e.g., '23.501' or '23') directly.")
        quick_action.triggered.connect(self._open_targeted_sync)

        scoped_action = sync_menu.addAction("⚙️ Scoped Sync (By WG / Series)...")
        scoped_action.setToolTip("Sync only specific portions of the database based on Series, Tech, or Group.")
        scoped_action.triggered.connect(self._open_advanced_sync)

        full_action = sync_menu.addAction("🔄 Full Database Sync (All Series)...")
        full_action.setToolTip("Synchronize the entire 3GPP specification database from the FTP server.")
        full_action.triggered.connect(lambda: self.update_db_requested.emit(self.force_meta_action.isChecked()))

        sync_menu.addSeparator()

        self.force_meta_action = QAction("Force Re-download Metadata", self)
        self.force_meta_action.setCheckable(True)
        self.force_meta_action.setChecked(False)
        self.force_meta_action.setToolTip("Force the scraper to re-download HTML metadata even if it already exists.")
        sync_menu.addAction(self.force_meta_action)

        sync_menu.addSeparator()
        config_action = sync_menu.addAction("⚙️ Download Settings...")
        config_action.triggered.connect(self._open_settings_dialog)

        open_folder_action = sync_menu.addAction("📂 Open Base Download Directory")
        open_folder_action.triggered.connect(self._open_download_dir)

        self.sync_menu_btn.setMenu(sync_menu)
        toolbar_layout.addWidget(self.sync_menu_btn)

        # ➕ Add Individual Specification Button
        self.add_spec_btn = QPushButton("➕ Add Spec")
        self.add_spec_btn.setToolTip("Add and inspect an individual specification without running a full sync.")
        self.add_spec_btn.setCursor(Qt.PointingHandCursor)
        self.add_spec_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.add_spec_btn.clicked.connect(self._open_add_spec_dialog)
        toolbar_layout.addWidget(self.add_spec_btn)

        # ⚙️ Settings Button
        self.settings_btn = QPushButton("⚙️ Settings")
        self.settings_btn.setToolTip("Configure the local specifications download folder and paths.")
        self.settings_btn.setCursor(Qt.PointingHandCursor)
        self.settings_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.settings_btn.clicked.connect(self._open_settings_dialog)
        toolbar_layout.addWidget(self.settings_btn)

        # ↗️ Open Local Folder Button
        self.open_dir_btn = QPushButton("↗️ Open Folder")
        self.open_dir_btn.setToolTip("Open the local specifications storage directory in Explorer.")
        self.open_dir_btn.setCursor(Qt.PointingHandCursor)
        self.open_dir_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.open_dir_btn.clicked.connect(self._open_download_dir)
        toolbar_layout.addWidget(self.open_dir_btn)

        # ⚠️ Wipe Database Button
        self.wipe_db_btn = QPushButton("⚠️ Wipe DB")
        self.wipe_db_btn.setToolTip("Clear all specifications, file mappings, and reset the local database.")
        self.wipe_db_btn.setCursor(Qt.PointingHandCursor)
        self.wipe_db_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_DANGER)
        self.wipe_db_btn.clicked.connect(self._on_wipe_db_clicked)
        toolbar_layout.addWidget(self.wipe_db_btn)

        self.bg_sync_label = QLabel("⏳ Fetching deep metadata in background...")
        self.bg_sync_label.setStyleSheet("color: #E65100; font-weight: bold; font-style: italic;")
        self.bg_sync_label.setVisible(False)
        toolbar_layout.addWidget(self.bg_sync_label)

        toolbar_layout.addStretch()
        main_layout.addLayout(toolbar_layout)

        # --- ROW 2: Primary Search & Quick Scope ---
        search_layout = QHBoxLayout()
        search_layout.setSpacing(8)

        search_layout.addWidget(QLabel("<b>🔍 Search:</b>"))

        self.spec_search_input = QLineEdit()
        self.spec_search_input.setStyleSheet(SPEC_FILTER_INPUT_STYLE)
        self.spec_search_input.setPlaceholderText("Search Spec # (e.g. 23.501) or title keywords...")
        self.spec_search_input.setText(self.saved_search)
        self.spec_search_input.setClearButtonEnabled(True)
        self.spec_search_input.setToolTip("Filter instantly by specification number or title keywords.")
        self.spec_search_input.textChanged.connect(self._on_search_changed)
        search_layout.addWidget(self.spec_search_input, stretch=3)

        ver_lbl = QLabel("Ver:")
        ver_lbl.setStyleSheet("color: #475569; font-weight: bold; font-size: 11px;")
        search_layout.addWidget(ver_lbl)

        self.version_search_input = QLineEdit()
        self.version_search_input.setStyleSheet(SPEC_FILTER_INPUT_STYLE)
        self.version_search_input.setPlaceholderText("e.g. 17 or 16.2")
        self.version_search_input.setText(self.saved_version)
        self.version_search_input.setClearButtonEnabled(True)
        self.version_search_input.setToolTip("Filter table by 3GPP release version (e.g. '15', '16.2', '17').")
        self.version_search_input.setFixedWidth(115)
        self.version_search_input.textChanged.connect(self._on_search_changed)
        search_layout.addWidget(self.version_search_input)

        # Downloaded Only Toggle Pill
        self.downloaded_only_btn = QPushButton("💾 Downloaded Only")
        self.downloaded_only_btn.setCheckable(True)
        self.downloaded_only_btn.setChecked(self.saved_downloaded_only)
        self.downloaded_only_btn.setCursor(Qt.PointingHandCursor)
        self.downloaded_only_btn.setStyleSheet(TOGGLE_PILL_STYLE)
        self.downloaded_only_btn.setToolTip("Show only specifications that currently have local cached files.")
        self.downloaded_only_btn.toggled.connect(self._toggle_downloaded_filter)
        self.downloaded_only_cb = self.downloaded_only_btn  # Backward compatibility alias
        search_layout.addWidget(self.downloaded_only_btn)

        # Favorites Toggle Pill
        self.fav_chip = QPushButton(f"⭐ Favorites ({len(self.favorites)})")
        self.fav_chip.setCheckable(True)
        self.fav_chip.setChecked(self.favorites_only)
        self.fav_chip.setCursor(Qt.PointingHandCursor)
        self.fav_chip.setStyleSheet(FAV_PILL_STYLE)
        self.fav_chip.setToolTip("Toggle view to show only your starred/favorite specifications.")
        self.fav_chip.toggled.connect(self._toggle_favorites_filter)
        search_layout.addWidget(self.fav_chip)

        main_layout.addLayout(search_layout)

        # --- ROW 3: Inline Faceted Filters (TSG, WG, Series, Tech, Type) ---
        filter_layout = QHBoxLayout()
        filter_layout.setSpacing(6)

        filters_lbl = QLabel("<b>🎯 Filters:</b>")
        filters_lbl.setStyleSheet("color: #334155; font-size: 11px;")
        filter_layout.addWidget(filters_lbl)

        # TSG Segmented Chips (All, RAN, SA, CT)
        tsg_lbl = QLabel("TSG:")
        tsg_lbl.setStyleSheet("color: #64748B; font-size: 11px; font-weight: bold;")
        filter_layout.addWidget(tsg_lbl)

        self.tsg_buttons = {}
        preset_tsgs = [
            ("All", ""),
            ("RAN", "RAN"),
            ("SA", "SA"),
            ("CT", "CT"),
        ]
        for label, tsg_val in preset_tsgs:
            btn = QPushButton(label)
            btn.setCheckable(True)
            btn.setCursor(Qt.PointingHandCursor)
            btn.setStyleSheet(SEGMENT_BTN_STYLE)
            btn.setToolTip(
                f"Filter specifications under TSG {label} (Plenary and Working Groups)"
                if tsg_val
                else "Show specifications across all Technical Specification Groups"
            )
            btn.clicked.connect(lambda _, t=tsg_val: self._on_tsg_button_clicked(t))
            self.tsg_buttons[tsg_val] = btn
            filter_layout.addWidget(btn)
        self.chip_buttons = self.tsg_buttons  # Backward compatibility alias

        # Working Group Dropdown
        wg_lbl = QLabel("WG:")
        wg_lbl.setStyleSheet("color: #64748B; font-size: 11px; font-weight: bold; margin-left: 4px;")
        filter_layout.addWidget(wg_lbl)

        self.wg_combo = QComboBox()
        self.wg_combo.setStyleSheet(SPEC_FILTER_COMBO_STYLE)
        self.wg_combo.setMinimumWidth(135)
        self.wg_combo.currentIndexChanged.connect(self._on_wg_combo_changed)
        filter_layout.addWidget(self.wg_combo)

        # Series Multi-Select CheckableComboBox
        self.series_combo = CheckableComboBox("Series")
        self.series_combo.setStyleSheet(SPEC_FILTER_COMBO_STYLE)
        self.series_combo.setMinimumWidth(130)
        self.series_combo.addItems(self.filter_options.get("series", []))
        self.series_combo.selectionChanged.connect(self._on_series_filter_changed)
        filter_layout.addWidget(self.series_combo)

        # Radio Technology Dropdown
        self.tech_combo = QComboBox()
        self.tech_combo.setStyleSheet(SPEC_FILTER_COMBO_STYLE)
        self.tech_combo.setMinimumWidth(100)
        self.tech_combo.addItem("All Tech", "")
        for tech in self.filter_options.get("techs", []):
            self.tech_combo.addItem(tech, tech)
        self.tech_combo.currentIndexChanged.connect(self._on_tech_combo_changed)
        filter_layout.addWidget(self.tech_combo)

        # Spec Type Segmented Buttons (All, TS, TR)
        type_lbl = QLabel("Type:")
        type_lbl.setStyleSheet("color: #64748B; font-size: 11px; font-weight: bold; margin-left: 4px;")
        filter_layout.addWidget(type_lbl)

        self.type_buttons = {}
        for t_label, t_val in [("All", "Any"), ("TS", "TS"), ("TR", "TR")]:
            btn = QPushButton(t_label)
            btn.setCheckable(True)
            btn.setCursor(Qt.PointingHandCursor)
            btn.setStyleSheet(SEGMENT_BTN_STYLE)
            btn.setToolTip(f"Filter by specification document type: {t_label}")
            btn.clicked.connect(lambda _, tv=t_val: self._on_type_button_clicked(tv))
            self.type_buttons[t_val] = btn
            filter_layout.addWidget(btn)

        # Tag Filter Dropdown
        tag_lbl = QLabel("Tag:")
        tag_lbl.setStyleSheet("color: #64748B; font-size: 11px; font-weight: bold; margin-left: 4px;")
        filter_layout.addWidget(tag_lbl)

        self.tag_filter_combo = QComboBox()
        self.tag_filter_combo.setStyleSheet(SPEC_FILTER_COMBO_STYLE)
        self.tag_filter_combo.setMinimumWidth(115)
        self.tag_filter_combo.currentIndexChanged.connect(self._on_tag_filter_changed)
        filter_layout.addWidget(self.tag_filter_combo)

        filter_layout.addStretch()

        # Reset All Button
        self.clear_all_btn = QPushButton("❌ Reset All")
        self.clear_all_btn.setCursor(Qt.PointingHandCursor)
        self.clear_all_btn.setStyleSheet(CLEAR_ALL_BTN_STYLE)
        self.clear_all_btn.setToolTip("Reset all search fields, active filters, and scopes.")
        self.clear_all_btn.clicked.connect(self._clear_all_filters)
        self.clear_all_btn.setVisible(False)
        filter_layout.addWidget(self.clear_all_btn)

        main_layout.addLayout(filter_layout)

        # --- ROW 4: Active Filter Tags & Count Banner ---
        self.active_tags_widget = QWidget()
        tags_bar_layout = QHBoxLayout(self.active_tags_widget)
        tags_bar_layout.setContentsMargins(0, 0, 0, 0)
        tags_bar_layout.setSpacing(6)

        self.active_tags_label = QLabel("Active:")
        self.active_tags_label.setStyleSheet("color: #64748B; font-size: 11px; font-weight: bold;")
        self.active_tags_label.setVisible(False)
        tags_bar_layout.addWidget(self.active_tags_label)

        self.active_tags_container = QWidget()
        self.active_tags_container_layout = QHBoxLayout(self.active_tags_container)
        self.active_tags_container_layout.setContentsMargins(0, 0, 0, 0)
        self.active_tags_container_layout.setSpacing(4)
        tags_bar_layout.addWidget(self.active_tags_container)

        tags_bar_layout.addStretch()

        self.count_badge = QLabel("0 specifications")
        self.count_badge.setStyleSheet("""
            QLabel {
                padding: 2px 10px;
                font-size: 11px;
                font-weight: bold;
                background-color: #F1F5F9;
                color: #475569;
                border: 1px solid #E2E8F0;
                border-radius: 10px;
            }
        """)
        tags_bar_layout.addWidget(self.count_badge)

        main_layout.addWidget(self.active_tags_widget)

        # --- ROW 5: Data Table ---
        self.table = QTableWidget()
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(["Specification", "Title", "Version / Documents"])
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Interactive)
        self.table.setColumnWidth(0, 195)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.setSectionResizeMode(2, QHeaderView.Interactive)
        self.table.setColumnWidth(2, 420)

        self.table.verticalHeader().setDefaultSectionSize(34)
        self.table.verticalHeader().setStyleSheet("QHeaderView::section { color: #94A3B8; font-size: 11px; }")

        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._show_context_menu)
        self.table.cellDoubleClicked.connect(self._on_table_cell_double_clicked)

        # Quick Copy keyboard shortcut (Ctrl+C)
        self.copy_shortcut = QShortcut(QKeySequence.Copy, self.table)
        self.copy_shortcut.activated.connect(self._copy_table_selection)

        main_layout.addWidget(self.table)
        self.setLayout(main_layout)

        self._sync_ui_to_filters()

    # ==========================================
    # --- FAVORITES MANAGEMENT ---
    # ==========================================
    def _toggle_favorite(self, spec_num: str):
        """Toggle favorite state for a single specification."""
        if spec_num in self.favorites:
            self.favorites.remove(spec_num)
        else:
            self.favorites.add(spec_num)

        self._save_settings()
        self._update_chip_styles()
        self._update_active_tags()
        self.refresh_table()

    def _batch_set_favorites(self, spec_nums: list, add: bool):
        """Batch add or remove specifications from favorites."""
        for num in spec_nums:
            if add:
                self.favorites.add(num)
            else:
                self.favorites.discard(num)

        self._save_settings()
        self._update_chip_styles()
        self._update_active_tags()
        self.refresh_table()

    def _toggle_favorites_filter(self, checked: bool = False):
        """Toggle the favorites-only filter mode."""
        self.favorites_only = self.fav_chip.isChecked()
        self.save_settings_timer.start()
        self._update_chip_styles()
        self._update_active_tags()
        self.filter_timer.start()

    def _toggle_downloaded_filter(self, checked: bool = False):
        """Toggle the downloaded-only filter mode."""
        self.save_settings_timer.start()
        self._update_active_tags()
        self.filter_timer.start()

    # ==========================================
    # --- EVENT HANDLERS & FILTERS ---
    # ==========================================
    def _open_settings_dialog(self):
        dialog = SpecsConfigDialog(self.download_dir, self)
        if dialog.exec_():
            new_path = dialog.get_download_path()
            if new_path and new_path != self.download_dir:
                self.download_dir = new_path
                self._save_settings()
                self.refresh_table()

    def _on_search_changed(self):
        self.search_timer.start()
        self.save_settings_timer.start()

    def _sync_ui_to_filters(self):
        """Synchronize UI controls with internal self.table_filters, favorites, and downloaded states."""
        # 1. Align TSG and WG
        grp = self.table_filters.get("group", "")
        if grp.startswith("SA"):
            self._current_tsg = "SA"
        elif grp.startswith("RAN"):
            self._current_tsg = "RAN"
        elif grp.startswith("CT") or grp.startswith("CN"):
            self._current_tsg = "CT"
        else:
            self._current_tsg = ""

        self._populate_wg_combo()
        if grp:
            idx = self.wg_combo.findData(grp)
            if idx >= 0:
                self.wg_combo.setCurrentIndex(idx)
            else:
                self.wg_combo.setCurrentIndex(0)
        else:
            self.wg_combo.setCurrentIndex(0)

        # 2. Align Series
        ser = self.table_filters.get("series", "")
        if ser:
            ser_list = [s.strip() for s in ser.split(",") if s.strip()]
            self.series_combo.setCheckedItems(ser_list)
        else:
            self.series_combo.setCheckedItems([])

        # 3. Align Tech
        tech = self.table_filters.get("tech", "")
        if tech:
            idx = self.tech_combo.findData(tech)
            if idx >= 0:
                self.tech_combo.setCurrentIndex(idx)
            else:
                self.tech_combo.setCurrentIndex(0)
        else:
            self.tech_combo.setCurrentIndex(0)

        # 4. Align Type
        stype = self.table_filters.get("spec_type", "Any")
        for t_val, btn in self.type_buttons.items():
            btn.setChecked(t_val == stype)

        # 5. Align Downloaded & Favorites
        self.downloaded_only_btn.setChecked(self.saved_downloaded_only)
        self.fav_chip.setChecked(self.favorites_only)

        self._populate_tag_filter_combo()
        self._update_chip_styles()
        self._update_active_tags()

    def _populate_wg_combo(self):
        self.wg_combo.blockSignals(True)
        current_data = self.wg_combo.currentData() or self.table_filters.get("group", "")
        self.wg_combo.clear()

        all_groups = self.filter_options.get("groups", [])
        clean_groups = [g for g in all_groups if g and g != "-"]

        if self._current_tsg == "SA":
            self.wg_combo.addItem("All SA WGs", "SA")
            sa_groups = [g for g in clean_groups if g.startswith("SA") and g != "SA"]
            for g in sorted(sa_groups):
                self.wg_combo.addItem(g, g)
        elif self._current_tsg == "RAN":
            self.wg_combo.addItem("All RAN WGs", "RAN")
            ran_groups = [g for g in clean_groups if g.startswith("RAN") and g != "RAN"]
            for g in sorted(ran_groups):
                self.wg_combo.addItem(g, g)
        elif self._current_tsg == "CT":
            self.wg_combo.addItem("All CT WGs", "CT")
            ct_groups = [g for g in clean_groups if (g.startswith("CT") or g.startswith("CN")) and g != "CT"]
            for g in sorted(ct_groups):
                self.wg_combo.addItem(g, g)
        else:
            self.wg_combo.addItem("All Working Groups", "")
            for g in sorted(clean_groups):
                self.wg_combo.addItem(g, g)

        if current_data:
            idx = self.wg_combo.findData(current_data)
            if idx >= 0:
                self.wg_combo.setCurrentIndex(idx)
            else:
                self.wg_combo.setCurrentIndex(0)
        else:
            self.wg_combo.setCurrentIndex(0)

        self.wg_combo.blockSignals(False)

    def _on_tsg_button_clicked(self, tsg_val: str):
        if self._current_tsg == tsg_val and tsg_val != "":
            self._current_tsg = ""
        else:
            self._current_tsg = tsg_val

        self._populate_wg_combo()
        self.table_filters["group"] = self.wg_combo.currentData() or ""
        self._update_chip_styles()
        self._update_active_tags()
        self.save_settings_timer.start()
        self.filter_timer.start()

    def _on_chip_clicked(self, tsg_val: str):
        """Backward compatibility wrapper for TSG chip clicks."""
        self._on_tsg_button_clicked(tsg_val)

    def _on_wg_combo_changed(self):
        data = self.wg_combo.currentData() or ""
        if self._current_tsg == "":
            if data.startswith("SA"):
                self._current_tsg = "SA"
            elif data.startswith("RAN"):
                self._current_tsg = "RAN"
            elif data.startswith("CT") or data.startswith("CN"):
                self._current_tsg = "CT"

        self.table_filters["group"] = data
        self._update_chip_styles()
        self._update_active_tags()
        self.save_settings_timer.start()
        self.filter_timer.start()

    def _on_wg_badge_clicked(self, wg: str):
        """Filters the specification list by the clicked Working Group badge."""
        if not wg or wg == "-":
            return
        tsg = ""
        wg_upper = wg.upper()
        if wg_upper.startswith("SA"):
            tsg = "SA"
        elif wg_upper.startswith("RAN") or wg_upper.startswith("GERAN"):
            tsg = "RAN"
        elif wg_upper.startswith("CT") or wg_upper.startswith("CN"):
            tsg = "CT"

        if self._current_tsg != tsg and tsg != "":
            self._current_tsg = tsg
            self._populate_wg_combo()
            self._update_chip_styles()

        idx = self.wg_combo.findData(wg)
        if idx >= 0:
            self.wg_combo.setCurrentIndex(idx)
        else:
            self._current_tsg = ""
            self._populate_wg_combo()
            self._update_chip_styles()
            fallback_idx = self.wg_combo.findData(wg)
            if fallback_idx >= 0:
                self.wg_combo.setCurrentIndex(fallback_idx)

    def _on_series_filter_changed(self, checked_items: list):
        if self.series_combo.isAllChecked() or not checked_items:
            self.table_filters["series"] = ""
        else:
            self.table_filters["series"] = ",".join(checked_items)

        self._update_active_tags()
        self.save_settings_timer.start()
        self.filter_timer.start()

    def _on_tech_combo_changed(self):
        data = self.tech_combo.currentData() or ""
        self.table_filters["tech"] = data
        self._update_active_tags()
        self.save_settings_timer.start()
        self.filter_timer.start()

    def _on_type_button_clicked(self, type_val: str):
        self.table_filters["spec_type"] = type_val
        self._update_chip_styles()
        self._update_active_tags()
        self.save_settings_timer.start()
        self.filter_timer.start()

    def _populate_tag_filter_combo(self):
        if not hasattr(self, "tag_filter_combo"):
            return
        current = self.tag_filter_combo.currentData() or ""
        self.tag_filter_combo.blockSignals(True)
        self.tag_filter_combo.clear()
        self.tag_filter_combo.addItem("All Tags", "")
        if self.db:
            unique_tags = self.db.get_all_unique_tags()
            for t in unique_tags:
                self.tag_filter_combo.addItem(f"🏷️ {t}", t)
        if current:
            idx = self.tag_filter_combo.findData(current)
            if idx >= 0:
                self.tag_filter_combo.setCurrentIndex(idx)
            else:
                self.tag_filter_combo.setCurrentIndex(0)
                self.tag_filter = None
        else:
            self.tag_filter_combo.setCurrentIndex(0)
        self.tag_filter_combo.blockSignals(False)

    def _on_tag_filter_changed(self):
        data = self.tag_filter_combo.currentData() or ""
        self.tag_filter = data if data else None
        self._update_active_tags()
        self.save_settings_timer.start()
        self.filter_timer.start()

    def _reset_tag_filter(self):
        self.tag_filter_combo.setCurrentIndex(0)
        self.tag_filter = None
        self._update_active_tags()
        self.save_settings_timer.start()
        self.filter_timer.start()

    def _update_chip_styles(self):
        fav_count = len(self.favorites)
        self.fav_chip.setText(f"⭐ Favorites ({fav_count})")
        self.fav_chip.setChecked(self.favorites_only)

        for tsg_val, btn in self.tsg_buttons.items():
            btn.setChecked(tsg_val == self._current_tsg)

        current_type = self.table_filters.get("spec_type", "Any")
        for t_val, btn in self.type_buttons.items():
            btn.setChecked(t_val == current_type)

    def _update_active_tags(self):
        """Re-render active filter tags dynamically with 1-click removal."""
        tags = []

        sq = self.spec_search_input.text().strip()
        if sq:
            disp_sq = sq if len(sq) <= 18 else f"{sq[:15]}..."
            tags.append((f'🔍 "{disp_sq}" ✕', lambda: (self.spec_search_input.clear(), self._on_search_changed())))

        vq = self.version_search_input.text().strip()
        if vq:
            tags.append((f'Ver: {vq} ✕', lambda: (self.version_search_input.clear(), self._on_search_changed())))

        if self.downloaded_only_btn.isChecked():
            tags.append(('💾 Cached ✕', lambda: self.downloaded_only_btn.setChecked(False)))

        if self.favorites_only:
            tags.append(('⭐ Favs ✕', lambda: self.fav_chip.setChecked(False)))

        if self.tag_filter:
            disp_tag = self.tag_filter if len(self.tag_filter) <= 15 else f"{self.tag_filter[:12]}..."
            tags.append((f'🏷️ {disp_tag} ✕', self._reset_tag_filter))

        grp = self.table_filters.get("group", "")
        if grp:
            tags.append((f'Group: {grp} ✕', self._reset_group_filter))

        ser = self.table_filters.get("series", "")
        if ser:
            disp_ser = ser if len(ser) <= 15 else f"{ser[:12]}..."
            tags.append((f'Series: {disp_ser} ✕', self._reset_series_filter))

        tech = self.table_filters.get("tech", "")
        if tech and tech != "Any":
            tags.append((f'Tech: {tech} ✕', self._reset_tech_filter))

        stype = self.table_filters.get("spec_type", "Any")
        if stype and stype != "Any":
            tags.append((f'Type: {stype} ✕', self._reset_type_filter))

        tag_labels = [label for label, _ in tags]
        if getattr(self, "_rendered_tag_labels", None) == tag_labels:
            return
        self._rendered_tag_labels = tag_labels

        while self.active_tags_container_layout.count():
            item = self.active_tags_container_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

        has_tags = len(tags) > 0
        self.active_tags_label.setVisible(has_tags)
        self.clear_all_btn.setVisible(has_tags)

        for label, callback in tags:
            tag_btn = QPushButton(label)
            tag_btn.setCursor(Qt.PointingHandCursor)
            tag_btn.setToolTip("Click to remove this filter constraint")
            tag_btn.setStyleSheet(DISMISSIBLE_TAG_STYLE)
            tag_btn.clicked.connect(callback)
            self.active_tags_container_layout.addWidget(tag_btn)

    def _reset_group_filter(self):
        self._current_tsg = ""
        self._populate_wg_combo()
        self.wg_combo.setCurrentIndex(0)
        self.table_filters["group"] = ""
        self._update_chip_styles()
        self._update_active_tags()
        self.save_settings_timer.start()
        self.filter_timer.start()

    def _reset_series_filter(self):
        self.series_combo.setCheckedItems([])
        self.table_filters["series"] = ""
        self._update_active_tags()
        self.save_settings_timer.start()
        self.filter_timer.start()

    def _reset_tech_filter(self):
        self.tech_combo.setCurrentIndex(0)
        self.table_filters["tech"] = ""
        self._update_active_tags()
        self.save_settings_timer.start()
        self.filter_timer.start()

    def _reset_type_filter(self):
        self.table_filters["spec_type"] = "Any"
        self._update_chip_styles()
        self._update_active_tags()
        self.save_settings_timer.start()
        self.filter_timer.start()

    def _clear_all_filters(self):
        self._rendered_tag_labels = None
        self.spec_search_input.blockSignals(True)
        self.version_search_input.blockSignals(True)
        self.downloaded_only_btn.blockSignals(True)
        self.fav_chip.blockSignals(True)
        self.wg_combo.blockSignals(True)
        self.tech_combo.blockSignals(True)
        if hasattr(self, "tag_filter_combo"):
            self.tag_filter_combo.blockSignals(True)

        self.spec_search_input.clear()
        self.version_search_input.clear()
        self.downloaded_only_btn.setChecked(False)
        self.favorites_only = False
        self.fav_chip.setChecked(False)
        self.tag_filter = None
        if hasattr(self, "tag_filter_combo"):
            self.tag_filter_combo.setCurrentIndex(0)

        self._current_tsg = ""
        self._populate_wg_combo()
        self.wg_combo.setCurrentIndex(0)

        self.series_combo.setCheckedItems([])
        self.tech_combo.setCurrentIndex(0)

        for t_val, btn in self.type_buttons.items():
            btn.setChecked(t_val == "Any")

        self.table_filters = {"series": "", "tech": "", "group": "", "spec_type": "Any"}

        self.spec_search_input.blockSignals(False)
        self.version_search_input.blockSignals(False)
        self.downloaded_only_btn.blockSignals(False)
        self.fav_chip.blockSignals(False)
        self.wg_combo.blockSignals(False)
        self.tech_combo.blockSignals(False)
        if hasattr(self, "tag_filter_combo"):
            self.tag_filter_combo.blockSignals(False)

        self._update_chip_styles()
        self._update_active_tags()
        self._save_settings()
        self.refresh_table()

    def set_bg_sync_active(self, is_active: bool):
        self.bg_sync_label.setVisible(is_active)
        if not is_active and self._is_loaded:
            self.refresh_table()

    def _open_download_dir(self):
        target_dir = Path(self.download_dir)
        if not target_dir.exists():
            try:
                target_dir.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                QMessageBox.warning(self, "Directory Error", f"Could not create directory:\n{e}")
                return
        try:
            os.startfile(str(target_dir))
        except Exception as e:
            QMessageBox.warning(self, "Explorer Error", f"Could not open directory:\n{e}")

    def _open_advanced_sync(self):
        dialog = AdvancedSyncDialog(self.db, self)
        if dialog.exec_():
            target_specs = dialog.matching_specs
            if target_specs:
                force_meta = self.force_meta_action.isChecked()
                self.update_specific_requested.emit(target_specs, force_meta)

    def _open_targeted_sync(self):
        dialog = TargetedSyncDialog(self)
        if dialog.exec_():
            targets = dialog.get_targets()
            if targets:
                force_meta = self.force_meta_action.isChecked()
                self.update_specific_requested.emit(targets, force_meta)

    def _open_table_filters(self):
        dialog = TableFilterDialog(self.db, self.table_filters, self)
        if dialog.exec_():
            self.table_filters = dialog.get_filters()
            self._sync_ui_to_filters()
            self.save_settings_timer.start()
            self.refresh_table()

    def _clear_table_filters(self):
        self._clear_all_filters()

    def _update_filter_ui(self):
        self._update_active_tags()

    def _show_context_menu(self, position):
        show_spec_context_menu(self, position)

    def _copy_table_selection(self):
        """Copies formatted citation and rich link for selected specifications on Ctrl+C."""
        selected_rows = self.table.selectionModel().selectedRows()
        if not selected_rows:
            return

        target_items = []
        for index in selected_rows:
            row = index.row()
            widget = self.table.cellWidget(row, 0)
            if widget:
                num_label = widget.findChild(QLabel, "specNumberLabel")
                if num_label:
                    canonical = num_label.property("specNumber")
                    spec_num = ""
                    if canonical:
                        spec_num = str(canonical).strip()
                    else:
                        display_text = num_label.text().strip()
                        match = re.match(r"^(\d{2}\.\d{2,3}(?:-[A-Za-z0-9]+)?)", display_text)
                        if match:
                            spec_num = match.group(1)

                    if spec_num:
                        cell_widget = self.table.cellWidget(row, 2)
                        v_info = None
                        if cell_widget:
                            combo = cell_widget.findChild(QComboBox)
                            if combo:
                                v_info = combo.currentData()
                        target_items.append((spec_num, v_info))

        if target_items:
            copy_specs_default(self, target_items, self.db)

    def _on_table_cell_double_clicked(self, row: int, column: int):
        if column in (0, 1):
            widget = self.table.cellWidget(row, 0)
            if widget:
                num_label = widget.findChild(QLabel, "specNumberLabel")
                if num_label:
                    canonical = num_label.property("specNumber") or num_label.text().strip().replace("⭐", "").strip()
                    self._open_edit_annotations(canonical)

    def _open_edit_annotations(self, spec_num: str):
        """Opens modal editor for specification tags and comments."""
        clean_num = str(spec_num).strip()
        if not clean_num or not self.db:
            return
        details = self.db.get_spec_details(clean_num)
        title = details.get("title", "")
        dlg = EditSpecAnnotationsDialog(clean_num, title, self.db, self)
        if dlg.exec_() == QDialog.Accepted:
            self._populate_tag_filter_combo()
            self.refresh_table()

    def _batch_add_tag(self, spec_nums: list):
        """Opens batch tag dialog for multiple selected specifications."""
        if not spec_nums or not self.db:
            return
        dlg = BatchTagDialog(spec_nums, self.db, self)
        if dlg.exec_() == QDialog.Accepted:
            self._populate_tag_filter_combo()
            self.refresh_table()

    def _show_spec_info(self, spec_num: str):
        details = self.db.get_spec_details(spec_num)
        if not details and "-" in spec_num:
            base_num = spec_num.split("-")[0]
            details = self.db.get_spec_details(base_num)
            if details:
                details["number"] = spec_num

        dialog = SpecInfoDialog(details or {}, self, db=self.db)
        dialog.exec_()
        self._populate_tag_filter_combo()
        self.refresh_table()

    def _open_spec_folder(self, spec_num: str):
        target_dir = Path(self.download_dir) / spec_num
        if target_dir.exists() and any(target_dir.iterdir()):
            try:
                os.startfile(str(target_dir))
            except Exception as e:
                QMessageBox.warning(self, "Explorer Error", f"Could not open directory:\n{e}")
        else:
            QMessageBox.information(
                self,
                "Folder Not Found",
                f"No downloaded files found for {spec_num}.\nDownload a document first to create the local folder.",
            )

    def _open_web_report(self, spec_num: str):
        url = build_dynareport_url(spec_num)
        if url:
            webbrowser.open(url)

    # ==========================================
    # --- TABLE REFRESH & RENDERING ---
    # ==========================================
    def _update_combo_item_badges(self, combo: QComboBox, index: int):
        """Update the display text, badges, and tooltips of a combo item based on disk state."""
        if index < 0 or index >= combo.count():
            return
        c_data = combo.itemData(index)
        if not c_data:
            return

        spec_target_dir = Path(self.download_dir) / c_data["spec_num"]
        stem = Path(c_data["fname"]).stem.lower()
        zip_exists = (spec_target_dir / c_data["fname"]).exists()

        local_files = []
        if spec_target_dir.exists():
            try:
                local_files = [f.name.lower() for f in spec_target_dir.iterdir() if f.is_file()]
            except Exception:
                local_files = []

        word_exists = any(f.startswith(stem) and (f.endswith(".docx") or f.endswith(".doc")) for f in local_files)
        pdf_exists = any(f.startswith(stem) and f.endswith(".pdf") for f in local_files)
        html_exists = any(f.startswith(stem) and f.endswith(".html") for f in local_files)
        md_exists = any(f.startswith(stem) and f.endswith(".md") for f in local_files)
        md_split_exists = (spec_target_dir / f"{stem}_clauses" / "index.md").exists()

        c_data["is_downloaded"] = zip_exists
        c_data["has_word"] = word_exists
        c_data["has_pdf"] = pdf_exists
        c_data["has_html"] = html_exists
        c_data["has_md"] = md_exists
        c_data["has_md_split"] = md_split_exists

        badges = []
        tooltips = []
        if word_exists:
            badges.append("📝")
            tooltips.append("Word (.docx) ready")
        if pdf_exists:
            badges.append("📕")
            tooltips.append("PDF (.pdf) ready")
        if html_exists:
            badges.append("🌐")
            tooltips.append("HTML (.html) ready")
        if md_exists:
            badges.append("📘")
            tooltips.append("Markdown (.md) ready")
        if md_split_exists:
            badges.append("📑")
            tooltips.append("Markdown Clauses ready")

        if not badges and zip_exists:
            badges.append("📦")
            tooltips.append("ZIP archive downloaded")
        elif not zip_exists:
            tooltips.append("Available on 3GPP FTP")

        status = "✅ " if zip_exists else ""
        u_date = c_data.get("upload_date")
        date_label = f" ({u_date})" if u_date else ""
        badge_str = f"  {' '.join(badges)}" if badges else ""
        ver = c_data.get("version", "")

        new_text = f"{status}v{ver}{date_label}{badge_str}"
        combo.setItemText(index, new_text)
        combo.setItemData(index, c_data)
        combo.setItemData(index, " • ".join(tooltips), Qt.ToolTipRole)

    def refresh_table(self):
        self.search_timer.stop()
        self.filter_timer.stop()
        self.table.setUpdatesEnabled(False)
        self.table.blockSignals(True)
        try:
            spec_query = self.spec_search_input.text().strip()
            version_query = self.version_search_input.text().strip()
            downloaded_only = self.downloaded_only_cb.isChecked()
            base_dl_dir = Path(self.download_dir)

            is_filtered = any(
                [
                    self.table_filters["series"],
                    self.table_filters["tech"],
                    self.table_filters["group"],
                    self.table_filters["spec_type"] != "Any",
                ]
            )

            specs = self.db.search_files(
                spec_number=spec_query if spec_query else None,
                release_version=version_query if version_query else None,
                tag=self.tag_filter,
                **self.table_filters,
            )

            self.table.setRowCount(0)

            grouped_specs = {}
            for row in specs:
                series, spec_num, title, spec_type, filename, version, url, upload_date, *rest = row
                row_tags = rest[0] if len(rest) > 0 else ""
                row_comments = rest[1] if len(rest) > 1 else ""
                row_wg = rest[2] if len(rest) > 2 else ""
                if filename and "-" not in spec_num:
                    part_match = PART_NUM_REGEX.search(filename)
                    if part_match:
                        spec_num = f"{spec_num}-{part_match.group(1)}"

                if spec_num not in grouped_specs:
                    grouped_specs[spec_num] = {
                        "title": title,
                        "type": spec_type if spec_type else "TS",
                        "wg": row_wg or "",
                        "versions": [],
                        "tags": [t.strip() for t in row_tags.split(",") if t.strip()] if row_tags else [],
                        "comments": row_comments or "",
                    }
                elif not grouped_specs[spec_num].get("wg") and row_wg:
                    grouped_specs[spec_num]["wg"] = row_wg
                grouped_specs[spec_num]["versions"].append((version, url, filename, upload_date))

            if downloaded_only and base_dl_dir.exists():
                try:
                    existing_dirs = {d.name for d in base_dl_dir.iterdir() if d.is_dir()}
                    grouped_specs = {s_num: s_data for s_num, s_data in grouped_specs.items() if s_num in existing_dirs}
                except Exception:
                    pass

            if self.favorites_only:
                filtered_grouped = {}
                for s_num, s_data in grouped_specs.items():
                    if s_num in self.favorites:
                        filtered_grouped[s_num] = s_data
                grouped_specs = filtered_grouped

            total_found = len(grouped_specs)
            rendered_specs = list(grouped_specs.items())[:100]

            if self.favorites_only:
                self.count_badge.setText(f"⭐ {total_found} Favorites")
                self.count_badge.setStyleSheet("""
                    padding: 2px 10px; font-size: 11px; font-weight: bold;
                    background-color: #FEF3C7; color: #92400E; border: 1px solid #F59E0B; border-radius: 10px;
                """)
            elif not spec_query and not version_query and not is_filtered and not downloaded_only:
                self.count_badge.setText(f"Top {len(rendered_specs)} specifications")
                self.count_badge.setStyleSheet("""
                    padding: 2px 10px; font-size: 11px; font-weight: bold;
                    background-color: #EBF8FF; color: #2B6CB0; border: 1px solid #BEE3F8; border-radius: 10px;
                """)
            elif total_found > 100:
                self.count_badge.setText(f"Showing 100 of {total_found} specs")
                self.count_badge.setStyleSheet("""
                    padding: 2px 10px; font-size: 11px; font-weight: bold;
                    background-color: #FFF3E0; color: #D83B01; border: 1px solid #FFB74D; border-radius: 10px;
                """)
            else:
                self.count_badge.setText(f"{total_found} specifications")
                self.count_badge.setStyleSheet("""
                    padding: 2px 10px; font-size: 11px; font-weight: bold;
                    background-color: #F1F5F9; color: #475569; border: 1px solid #E2E8F0; border-radius: 10px;
                """)

            self._update_active_tags()

            for row_idx, (spec_num, data) in enumerate(rendered_specs):
                self.table.insertRow(row_idx)
                spec_target_dir = base_dl_dir / spec_num
                is_fav = spec_num in self.favorites

                # --- COLUMN 0: Action Button + Badged Spec Number ---
                spec_widget = QWidget()
                spec_layout = QHBoxLayout(spec_widget)
                spec_layout.setContentsMargins(6, 0, 6, 0)
                spec_layout.setSpacing(6)

                action_btn = HoverMenuButton("⋮")
                action_btn.setFixedSize(22, 22)
                action_btn.setToolTip("Specification Actions")
                action_btn.setCursor(Qt.PointingHandCursor)
                action_btn.setStyleSheet("""
                    QPushButton { border: none; background: transparent; color: #718096; font-size: 18px; font-weight: bold; padding-bottom: 3px; }
                    QPushButton:hover { color: #0078D7; }
                    QPushButton::menu-indicator { image: none; width: 0px; }
                """)

                menu = QMenu(self)
                menu.setStyleSheet("""
                    QMenu { background-color: #FAFAFA; border: 1px solid #CCC; }
                    QMenu::item { padding: 5px 20px 5px 15px; color: #333333; }
                    QMenu::item:selected { background-color: #E1F0FF; color: #0078D7; }
                    QMenu::item:disabled { color: #AAAAAA; }
                """)

                fav_label = "⭐ Remove from Favorites" if is_fav else "⭐ Add to Favorites"
                menu.addAction(fav_label).triggered.connect(lambda _, s=spec_num: self._toggle_favorite(s))

                ann_action = menu.addAction("🏷️  Edit Tags & Notes...")
                ann_action.triggered.connect(lambda _, s=spec_num: self._open_edit_annotations(s))

                menu.addSeparator()

                info_action = menu.addAction("ℹ️  View Details")
                info_action.triggered.connect(lambda _, s=spec_num: self._show_spec_info(s))

                web_action = menu.addAction("🌐  Open 3GPP Web Report")
                web_action.triggered.connect(lambda _, s=spec_num: self._open_web_report(s))

                search_action = menu.addAction("🔎  Track Evolution in Spec Search")
                search_action.triggered.connect(lambda _, s=spec_num: self.search_in_spec_search_requested.emit(s, ""))

                menu.addSeparator()

                # --- Quick Copy Submenu ---
                copy_sub = menu.addMenu("📋  Quick Copy")
                copy_sub.setToolTipsVisible(True)
                copy_sub.setToolTip("Copy specification metadata in plain text and rich HTML form.")

                def _on_action_copy_about_to_show(sub=copy_sub, s=spec_num, r=row_idx):
                    sub.clear()
                    cell_widget = self.table.cellWidget(r, 2)
                    v_info = None
                    if cell_widget:
                        combo = cell_widget.findChild(QComboBox)
                        if combo:
                            v_info = combo.currentData()
                    populate_spec_copy_menu(sub, self, s, version_info=v_info, db=self.db)

                copy_sub.aboutToShow.connect(_on_action_copy_about_to_show)

                menu.addSeparator()

                folder_action = menu.addAction("📂  Open Local Folder")
                folder_action.triggered.connect(lambda _, s=spec_num: self._open_spec_folder(s))

                def _update_menu_state(act=folder_action, path=spec_target_dir):
                    if path.exists() and any(path.iterdir()):
                        act.setText("📂  Open Local Folder")
                        act.setEnabled(True)
                    else:
                        act.setText("📁  Folder Not Created")
                        act.setEnabled(False)

                menu.addSeparator()
                del_action = menu.addAction("🗑️  Delete Specification")
                del_action.triggered.connect(lambda _, s=spec_num: self._delete_single_spec(s))

                menu.aboutToShow.connect(_update_menu_state)
                _update_menu_state()
                action_btn.setMenu(menu)

                spec_type = (data["type"] or "TS").upper()
                type_badge = QLabel(spec_type)
                type_badge.setFixedWidth(24)
                type_badge.setAlignment(Qt.AlignCenter)
                if spec_type == "TR":
                    type_badge.setStyleSheet("""
                        background-color: #F3E8FF; color: #6B21A8; border: 1px solid #E9D5FF;
                        border-radius: 3px; padding: 1px 0px; font-size: 10px; font-weight: bold;
                    """)
                else:
                    type_badge.setStyleSheet("""
                        background-color: #EBF8FF; color: #2B6CB0; border: 1px solid #BEE3F8;
                        border-radius: 3px; padding: 1px 0px; font-size: 10px; font-weight: bold;
                    """)

                wg_name = (data.get("wg") or "").strip()
                wg_badge = QLabel(wg_name if wg_name and wg_name != "-" else "—")
                wg_badge.setFixedWidth(42)
                wg_badge.setAlignment(Qt.AlignCenter)
                if wg_name and wg_name != "-":
                    wg_badge.setCursor(Qt.PointingHandCursor)
                    wg_style, wg_desc = get_wg_badge_style_and_desc(wg_name)
                    wg_badge.setStyleSheet(wg_style)
                    wg_badge.setToolTip(f"Working Group: {wg_name} ({wg_desc})\nClick to filter by this Working Group")
                    wg_badge.mousePressEvent = lambda event, wg=wg_name: self._on_wg_badge_clicked(wg)
                else:
                    wg_badge.setStyleSheet("""
                        background-color: #F8FAFC; color: #94A3B8; border: 1px solid #E2E8F0;
                        border-radius: 3px; padding: 1px 0px; font-size: 10px; font-weight: bold;
                    """)
                    wg_badge.setToolTip("Working Group: Unassigned / Plenary")

                display_title = f"{spec_num} ⭐" if is_fav else spec_num
                spec_label = QLabel(display_title)
                spec_label.setObjectName("specNumberLabel")
                # Keep the canonical identifier separate from decorated display text.
                spec_label.setProperty("specNumber", spec_num)
                spec_label.setStyleSheet("font-size: 12px; font-weight: bold; color: #1A202C;")
                if is_fav:
                    spec_label.setToolTip("Marked as Favorite")

                spec_layout.addWidget(action_btn)
                spec_layout.addWidget(type_badge)
                spec_layout.addWidget(wg_badge)
                spec_layout.addWidget(spec_label)
                spec_layout.addStretch()
                self.table.setCellWidget(row_idx, 0, spec_widget)

                # --- COLUMN 1: Title + Tags & Comments ---
                title_text = data["title"] if data["title"] else "Unknown Title"
                tags = data.get("tags", [])
                comments = data.get("comments", "").strip()

                if tags or comments:
                    title_widget = QWidget()
                    title_layout = QVBoxLayout(title_widget)
                    title_layout.setContentsMargins(4, 3, 4, 3)
                    title_layout.setSpacing(2)

                    title_lbl = QLabel(title_text)
                    title_lbl.setStyleSheet("font-size: 12px; color: #1E293B;")
                    title_lbl.setWordWrap(True)
                    title_layout.addWidget(title_lbl)

                    meta_row = QHBoxLayout()
                    meta_row.setContentsMargins(0, 0, 0, 0)
                    meta_row.setSpacing(4)

                    for tag_str in tags:
                        tag_pill = QLabel(f"🏷️ {tag_str}")
                        tag_pill.setStyleSheet("""
                            background-color: #EBF3FC; color: #1E5C99; border: 1px solid #BDD7EE;
                            border-radius: 9px; padding: 1px 6px; font-size: 10px; font-weight: bold;
                        """)
                        tag_pill.setToolTip(f"User Tag: {tag_str}")
                        meta_row.addWidget(tag_pill)

                    if comments:
                        note_pill = QLabel("📝 Note")
                        note_pill.setStyleSheet("""
                            background-color: #FEF3C7; color: #92400E; border: 1px solid #FCD34D;
                            border-radius: 9px; padding: 1px 6px; font-size: 10px; font-weight: bold;
                        """)
                        preview = comments if len(comments) <= 250 else f"{comments[:247]}..."
                        note_pill.setToolTip(f"Personal Note:\n{preview}")
                        meta_row.addWidget(note_pill)

                    meta_row.addStretch()
                    title_layout.addLayout(meta_row)
                    self.table.setCellWidget(row_idx, 1, title_widget)
                    self.table.setRowHeight(row_idx, 48)
                else:
                    self.table.setItem(row_idx, 1, QTableWidgetItem(title_text))
                    self.table.setRowHeight(row_idx, 34)

                # --- COLUMN 2: Documents Action Bar ---
                version_combo = QComboBox()
                version_combo.setFixedWidth(230)
                version_combo.setFixedHeight(28)
                version_combo.view().setMinimumWidth(255)

                # Snapshot local files in memory once per specification
                local_files = []
                if spec_target_dir.exists():
                    try:
                        local_files = [f.name.lower() for f in spec_target_dir.iterdir() if f.is_file()]
                    except Exception:
                        local_files = []

                def parse_ver(v_str):
                    return [(0, int(x)) if x.isdigit() else (1, str(x)) for x in str(v_str).split(".")]

                sorted_versions = sorted(data["versions"], key=lambda x: parse_ver(x[0]), reverse=True)

                for ver, url, fname, u_date in sorted_versions:
                    if not fname:
                        continue

                    stem = Path(fname).stem.lower()
                    fname_lower = fname.lower()
                    zip_exists = fname_lower in local_files
                    word_exists = any(
                        f.startswith(stem) and (f.endswith(".docx") or f.endswith(".doc")) for f in local_files
                    )
                    pdf_exists = any(f.startswith(stem) and f.endswith(".pdf") for f in local_files)
                    html_exists = any(f.startswith(stem) and f.endswith(".html") for f in local_files)

                    badges = []
                    tooltips = []
                    if word_exists:
                        badges.append("📝")
                        tooltips.append("Word (.docx) ready")
                    if pdf_exists:
                        badges.append("📕")
                        tooltips.append("PDF (.pdf) ready")
                    if html_exists:
                        badges.append("🌐")
                        tooltips.append("HTML (.html) ready")

                    if not badges and zip_exists:
                        badges.append("📦")
                        tooltips.append("ZIP archive downloaded")
                    elif not zip_exists:
                        tooltips.append("Available on 3GPP FTP")

                    status = "✅ " if zip_exists else ""
                    date_label = f" ({u_date})" if u_date else ""
                    badge_str = f"  {' '.join(badges)}" if badges else ""
                    display_text = f"{status}v{ver}{date_label}{badge_str}"

                    item_idx = version_combo.count()
                    version_combo.addItem(
                        display_text,
                        userData={
                            "url": url,
                            "fname": fname,
                            "spec_num": spec_num,
                            "version": ver,
                            "is_downloaded": zip_exists,
                            "upload_date": u_date,
                            "has_word": word_exists,
                            "has_pdf": pdf_exists,
                            "has_html": html_exists,
                        },
                    )
                    version_combo.setItemData(item_idx, " • ".join(tooltips), Qt.ToolTipRole)

                doc_action_btn = QPushButton()
                doc_action_btn.setFixedWidth(150)
                doc_action_btn.setFixedHeight(28)
                doc_action_btn.setCursor(Qt.PointingHandCursor)

                doc_menu = QMenu(self)
                doc_menu.setStyleSheet("""
                    QMenu { background-color: #FAFAFA; border: 1px solid #CCC; }
                    QMenu::item { padding: 6px 22px 6px 14px; color: #333333; font-size: 12px; }
                    QMenu::item:selected { background-color: #E1F0FF; color: #0078D7; }
                    QMenu::item:disabled { color: #AAAAAA; }
                    QMenu::separator { height: 1px; background-color: #E2E8F0; margin: 4px 0; }
                """)
                doc_action_btn.setMenu(doc_menu)

                def _update_btn_state(index_ignore=0, c=version_combo, btn=doc_action_btn, menu=doc_menu,
                                      files_cache=local_files):
                    c_data = c.currentData()
                    if not c_data:
                        return

                    current_dir = Path(self.download_dir) / c_data["spec_num"]
                    stem = Path(c_data["fname"]).stem.lower()

                    # Re-scan the directory whenever button state is refreshed.
                    # The old code used the refresh_table() snapshot (files_cache),
                    # which became stale immediately after extracting/converting a file.
                    current_files = []
                    if current_dir.exists():
                        try:
                            current_files = [
                                f.name.lower()
                                for f in current_dir.iterdir()
                                if f.is_file()
                            ]
                        except OSError:
                            current_files = []

                    zip_exists = c_data["fname"].lower() in current_files
                    word_exists = any(
                        f.startswith(stem) and (f.endswith(".docx") or f.endswith(".doc"))
                        for f in current_files
                    )
                    pdf_exists = any(f.startswith(stem) and f.endswith(".pdf") for f in current_files)
                    html_exists = any(f.startswith(stem) and f.endswith(".html") for f in current_files)
                    txt_exists = any(f.startswith(stem) and f.endswith(".txt") for f in current_files)
                    md_exists = any(f.startswith(stem) and f.endswith(".md") for f in current_files)
                    md_split_exists = (spec_target_dir / f"{stem}_clauses" / "index.md").exists()
                    dir_ready = bool(current_files)

                    base_style = """
                        QPushButton {
                            font-size: 11px;
                            font-weight: bold;
                            padding: 2px 8px;
                            border-radius: 4px;
                            text-align: center;
                        }
                        QPushButton::menu-indicator { image: none; width: 0px; }
                    """

                    if word_exists:
                        btn.setText("📝 Open Word ▾")
                        btn.setStyleSheet(base_style + """
                            QPushButton {
                                background-color: #E8F5E9;
                                color: #2E7D32;
                                border: 1px solid #2E7D32;
                            }
                            QPushButton:hover { background-color: #C8E6C9; }
                        """)
                    elif zip_exists:
                        btn.setText("⚙️ Extract Word ▾")
                        btn.setStyleSheet(base_style + """
                            QPushButton {
                                background-color: #FFF3E0;
                                color: #E65100;
                                border: 1px solid #FFB74D;
                            }
                            QPushButton:hover { background-color: #FFE0B2; }
                        """)
                    else:
                        btn.setText("⬇️ Get Word ▾")
                        btn.setStyleSheet(base_style + """
                            QPushButton {
                                background-color: #EBF3FB;
                                color: #0066CC;
                                border: 1px solid #B0D0F0;
                            }
                            QPushButton:hover { background-color: #D6E8FA; border-color: #0066CC; }
                        """)

                    menu.clear()

                    word_label = "📝 Open Word Document" if word_exists else (
                        "⚙️ Extract Word Document" if zip_exists else "⬇️ Download && Open Word")
                    act_word = menu.addAction(f"{word_label} {'✅' if word_exists else ''}".strip())
                    act_word.triggered.connect(lambda: self._handle_document_action(c, "word", btn))

                    pdf_label = "📕 Open PDF" if pdf_exists else (
                        "⚙️ Convert to PDF" if (word_exists or zip_exists) else "⬇️ Get && Convert to PDF")
                    act_pdf = menu.addAction(f"{pdf_label} {'✅' if pdf_exists else ''}".strip())
                    act_pdf.triggered.connect(lambda: self._handle_document_action(c, "pdf", btn))

                    html_label = "🌐 Open HTML" if html_exists else (
                        "⚙️ Convert to HTML" if (word_exists or zip_exists) else "⬇️ Get && Convert to HTML")
                    act_html = menu.addAction(f"{html_label} {'✅' if html_exists else ''}".strip())
                    act_html.triggered.connect(lambda: self._handle_document_action(c, "html", btn))

                    txt_label = "📄 Open TXT" if txt_exists else (
                        "⚙️ Convert to TXT" if (word_exists or zip_exists) else "⬇️ Get && Convert to TXT")
                    act_txt = menu.addAction(f"{txt_label} {'✅' if txt_exists else ''}".strip())
                    act_txt.triggered.connect(lambda: self._handle_document_action(c, "txt", btn))

                    menu.addSeparator()

                    md_label = "📘 Open Markdown (.md)" if md_exists else (
                        "⚙️ Convert to Markdown (.md)" if (word_exists or zip_exists) else "⬇️ Get && Convert to Markdown (.md)")
                    act_md = menu.addAction(f"{md_label} {'✅' if md_exists else ''}".strip())
                    act_md.triggered.connect(lambda: self._handle_document_action(c, "md", btn))

                    md_split_label = "📑 Open Markdown Clauses" if md_split_exists else (
                        "⚙️ Convert to Markdown Clauses" if (word_exists or zip_exists) else "⬇️ Get && Convert to Markdown Clauses")
                    act_md_split = menu.addAction(f"{md_split_label} {'✅' if md_split_exists else ''}".strip())
                    act_md_split.triggered.connect(lambda: self._handle_document_action(c, "md_split", btn))

                    menu.addSeparator()

                    # Add the new Clause Excerpt action here:
                    act_excerpt = menu.addAction("✂️ Extract Clauses (Excerpt)...")
                    act_excerpt.triggered.connect(lambda: self._open_clause_excerpt_dialog(c))

                    act_cr_pcr = menu.addAction("📝 Draft Contribution (CR / pCR)...")
                    act_cr_pcr.triggered.connect(lambda: self._open_cr_pcr_dialog(c))

                    menu.addSeparator()

                    act_search_tab = menu.addAction("🔎 Track Evolution / Clauses in Spec Search")
                    act_search_tab.triggered.connect(lambda: self.search_in_spec_search_requested.emit(c_data["spec_num"], c_data.get("version", "")))

                    act_index_search = menu.addAction("📥 Index into Spec Search DB")
                    act_index_search.triggered.connect(lambda: self.index_in_spec_search_requested.emit(c_data["spec_num"], c_data.get("version", "")))

                    menu.addSeparator()

                    zip_label = "📦 Show ZIP Archive ✅" if zip_exists else "📥 Download ZIP Archive"
                    act_zip = menu.addAction(zip_label)
                    act_zip.triggered.connect(lambda: self._handle_zip_action(c, btn))

                    menu.addSeparator()
                    ver_title = f"v{c_data['version']}" if not str(c_data.get('version', '')).startswith('v') else str(c_data.get('version', ''))
                    doc_copy_sub = menu.addMenu(f"📋 Quick Copy ({ver_title})")
                    doc_copy_sub.setToolTipsVisible(True)
                    doc_copy_sub.setToolTip(f"Copy specification metadata and download links for version {ver_title}.")

                    def _on_doc_copy_about_to_show(sub=doc_copy_sub, s=c_data["spec_num"], v=c_data):
                        sub.clear()
                        populate_spec_copy_menu(sub, self, s, version_info=v, db=self.db)

                    doc_copy_sub.aboutToShow.connect(_on_doc_copy_about_to_show)

                    if dir_ready:
                        act_dir = menu.addAction("📂 Open Local Folder")
                        act_dir.setEnabled(True)
                        act_dir.triggered.connect(lambda _, s=c_data["spec_num"]: self._open_spec_folder(s))
                    else:
                        act_dir = menu.addAction("📁 Folder Not Created")
                        act_dir.setEnabled(False)

                version_combo.currentIndexChanged.connect(_update_btn_state)
                _update_btn_state()

                cell_widget = QWidget()
                layout = QHBoxLayout(cell_widget)
                layout.setContentsMargins(4, 0, 4, 0)
                layout.setSpacing(6)

                layout.addWidget(version_combo)
                layout.addWidget(doc_action_btn)
                layout.addStretch()

                self.table.setCellWidget(row_idx, 2, cell_widget)

        except Exception as e:
            logger.error(f"Error during refresh_table: {e}", exc_info=True)
        finally:
            self.table.blockSignals(False)
            self.table.setUpdatesEnabled(True)

    # ==========================================
    # --- DOCUMENT FETCH & EXTRACTION ---
    # ==========================================
    def _open_clause_excerpt_dialog(self, combo: QComboBox):
        if sip.isdeleted(combo):
            return
        c_data = combo.currentData()
        if not c_data:
            return

        spec_num = c_data["spec_num"]
        spec_dl_dir = Path(self.download_dir) / spec_num
        spec_type = "TS"
        details = self.db.get_spec_details(spec_num)
        if not details and "-" in spec_num:
            base_num = spec_num.split("-")[0]
            details = self.db.get_spec_details(base_num)
        if details and details.get("type"):
            spec_type = details["type"]

        dialog = SpecClauseExcerptDialog(
            spec_num=spec_num,
            spec_type=spec_type,
            version=c_data.get("version", ""),
            fname=c_data.get("fname", ""),
            url=c_data.get("url", ""),
            spec_dl_dir=spec_dl_dir,
            parent=self,
        )

        def _on_extraction_done(p: Path, c=combo):
            try:
                if not sip.isdeleted(c):
                    idx = c.currentIndex()
                    self._update_combo_item_badges(c, idx)
                    c.currentIndexChanged.emit(idx)
            except RuntimeError:
                pass

        dialog.extraction_completed.connect(_on_extraction_done)
        dialog.show()

    def _open_cr_pcr_dialog(self, combo: QComboBox):
        if sip.isdeleted(combo):
            return
        c_data = combo.currentData()
        if not c_data:
            return

        from modules.word_tools.ui.cr_pcr_dialog import CrPcrDialog

        spec_num = c_data["spec_num"]
        spec_dl_dir = Path(self.download_dir) / spec_num
        spec_type = "TS"
        release = c_data.get("release", "")
        details = self.db.get_spec_details(spec_num)
        if not details and "-" in spec_num:
            base_num = spec_num.split("-")[0]
            details = self.db.get_spec_details(base_num)
        if details and details.get("type"):
            spec_type = details["type"]

        cr_dlg = CrPcrDialog(
            spec_num=spec_num,
            spec_type=spec_type,
            version=c_data.get("version", ""),
            release=release,
            fname=c_data.get("fname", ""),
            url=c_data.get("url", ""),
            spec_dl_dir=spec_dl_dir,
            parent=self,
        )
        cr_dlg.show()

    def _handle_document_action(self, combo: QComboBox, doc_type: str, btn: QPushButton):
        handle_document_action(self, combo, doc_type, btn)

    def _handle_zip_action(self, combo: QComboBox, btn: QPushButton):
        handle_zip_action(self, combo, btn)

    def _open_add_spec_dialog(self):
        dialog = AddSpecDialog(self.db, self)
        if dialog.exec_() == QDialog.Accepted:
            self.refresh_table()
            if dialog.sync_requested and dialog.saved_spec_number:
                force_meta = self.force_meta_action.isChecked()
                self.update_specific_requested.emit([dialog.saved_spec_number], force_meta)

    def _delete_single_spec(self, spec_num: str):
        reply = QMessageBox.question(
            self,
            "Confirm Delete",
            f"Are you sure you want to delete specification TS {spec_num} from the database?",
            QMessageBox.Yes | QMessageBox.No,
        )
        if reply == QMessageBox.Yes:
            deleted = self.db.delete_specification(spec_num)
            if not deleted and "-" in spec_num:
                base_num = spec_num.split("-")[0]
                deleted = self.db.delete_specification(base_num)
                if deleted:
                    self.favorites.discard(base_num)
            if deleted:
                self.favorites.discard(spec_num)
                self._save_settings()
                self.refresh_table()
                logger.log(logging.INFO, f"🗑️ Deleted TS {spec_num} from specifications database.")
            else:
                QMessageBox.warning(self, "Delete Error", f"Could not delete TS {spec_num}.")

    def _delete_batch_specs(self, spec_nums: list):
        reply = QMessageBox.question(
            self,
            "Confirm Delete",
            f"Are you sure you want to delete {len(spec_nums)} selected specifications from the database?",
            QMessageBox.Yes | QMessageBox.No,
        )
        if reply == QMessageBox.Yes:
            deleted_count = 0
            for num in spec_nums:
                deleted = self.db.delete_specification(num)
                if not deleted and "-" in num:
                    base_num = num.split("-")[0]
                    deleted = self.db.delete_specification(base_num)
                    if deleted:
                        self.favorites.discard(base_num)
                if deleted:
                    self.favorites.discard(num)
                    deleted_count += 1
            self._save_settings()
            self.refresh_table()
            logger.log(logging.INFO, f"🗑️ Deleted {deleted_count} specification(s) from database.")

    def _on_wipe_db_clicked(self):
        reply = QMessageBox.critical(
            self,
            "Confirm Wipe Database",
            "This will delete ALL specification metadata, file versions, and related mappings from 3gpp_data.db.\n\nContinue?",
            QMessageBox.Yes | QMessageBox.No,
        )
        if reply == QMessageBox.Yes:
            if self.db.wipe_database():
                self.favorites.clear()
                self._save_settings()
                self.refresh_table()
                logger.log(logging.INFO, "🧹 Specifications database wiped successfully.")
            else:
                QMessageBox.critical(self, "Wipe Failed", "Failed to wipe specifications database.")