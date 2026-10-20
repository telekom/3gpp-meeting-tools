# --- File: src/modules/emails/ui/email_toolbar.py ---
import logging
from typing import Optional

from PyQt5.QtCore import Qt, QDate, pyqtSignal
from PyQt5.QtWidgets import (
    QDateEdit,
    QFrame,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from modules.meetings.ui.tdocs_components import CheckableComboBox

logger = logging.getLogger(__name__)


def _get_btn_style(primary: bool = False) -> str:
    if primary:
        return (
            "QPushButton { font-weight: bold; background-color: #0078D7; color: white; "
            "padding: 5px 12px; border-radius: 4px; border: 1px solid #005A9E; } "
            "QPushButton:hover { background-color: #005A9E; }"
        )
    return (
        "QPushButton { font-weight: bold; background-color: #FFFFFF; color: #333333; "
        "padding: 5px 12px; border-radius: 4px; border: 1px solid #CCCCCC; } "
        "QPushButton:hover { background-color: #F0F4F8; border-color: #005A9E; color: #005A9E; }"
    )


class EmailControlToolbar(QFrame):
    """
    Unified top control panel for EmailManagerWindow:
    - Action buttons: Sync, Move, Move All, Scan Target, Statistics, Folders
    - Filter controls: Date range, Starred, Followed AIs, AI/Company/Sender combo boxes
    - Status & thread count labels
    """

    sync_requested = pyqtSignal()
    move_requested = pyqtSignal()
    move_all_requested = pyqtSignal()
    rescan_requested = pyqtSignal()
    stats_requested = pyqtSignal()
    config_requested = pyqtSignal()
    filters_changed = pyqtSignal()
    date_range_changed = pyqtSignal(str, str)

    def __init__(self, start_date: str = "", end_date: str = "", parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.start_date = start_date
        self.end_date = end_date
        self._setup_ui()

    def _setup_ui(self):
        self.setStyleSheet(
            "QFrame { background-color: #FFFFFF; border: 1px solid #E0E0E0; border-radius: 8px; } "
            "QLabel { font-weight: bold; color: #555; border: none; } "
            "QLineEdit, QComboBox, QDateEdit { padding: 4px 6px; border: 1px solid #CCC; border-radius: 4px; background: #FFF; color: #333; height: 22px; }"
        )
        top_layout = QVBoxLayout(self)
        top_layout.setContentsMargins(10, 8, 10, 8)
        top_layout.setSpacing(6)

        # --- ROW 1: Buttons & Status ---
        row1_layout = QHBoxLayout()
        row1_layout.setSpacing(6)

        self.btn_sync = QPushButton("🔄 Sync Source")
        self.btn_sync.setStyleSheet(_get_btn_style(primary=True))
        self.btn_sync.setToolTip("Scan your configured Outlook Source folder to download and index new emails.")
        self.btn_sync.clicked.connect(self.sync_requested.emit)

        self.btn_move = QPushButton("➡ Move Selected")
        self.btn_move.setStyleSheet(_get_btn_style())
        self.btn_move.setToolTip("Move the currently selected emails from the Source folder to the Target folder.")
        self.btn_move.clicked.connect(self.move_requested.emit)

        self.btn_move_all = QPushButton("⏭ Move All")
        self.btn_move_all.setStyleSheet(_get_btn_style())
        self.btn_move_all.setToolTip("Move ALL visible source emails to the Target folder, auto-organizing them by Agenda Item.")
        self.btn_move_all.clicked.connect(self.move_all_requested.emit)

        self.btn_rescan = QPushButton("🔁 Scan Target")
        self.btn_rescan.setStyleSheet(_get_btn_style())
        self.btn_rescan.setToolTip("Re-scan your configured Target archive folder to index previously moved emails.")
        self.btn_rescan.clicked.connect(self.rescan_requested.emit)

        self.btn_stats = QPushButton("📊 Statistics")
        self.btn_stats.setStyleSheet(_get_btn_style())
        self.btn_stats.setToolTip("Generate an interactive HTML analytics dashboard for email traffic.")
        self.btn_stats.clicked.connect(self.stats_requested.emit)

        self.btn_config = QPushButton("⚙ Folders")
        self.btn_config.setStyleSheet(_get_btn_style())
        self.btn_config.setToolTip("Configure your Outlook Source/Target folders and Statistics preferences.")
        self.btn_config.clicked.connect(self.config_requested.emit)

        self.lbl_status = QLabel("Ready.")
        self.lbl_status.setStyleSheet("color: #666; font-style: italic; margin-left: 10px;")

        self.lbl_count = QLabel("Showing 0 Threads | 0 Emails")
        self.lbl_count.setStyleSheet(
            "font-size: 13px; color: #0078D7; font-weight: bold; padding-left: 10px; border: none;"
        )

        for btn in [self.btn_sync, self.btn_move, self.btn_move_all, self.btn_rescan, self.btn_stats, self.btn_config]:
            row1_layout.addWidget(btn)
        row1_layout.addWidget(self.lbl_status)
        row1_layout.addStretch()
        row1_layout.addWidget(self.lbl_count)
        top_layout.addLayout(row1_layout)

        # --- ROW 2: Filters ---
        row2_layout = QHBoxLayout()
        row2_layout.setSpacing(8)

        self.dt_start = QDateEdit()
        self.dt_start.setCalendarPopup(True)
        self.dt_start.setToolTip("Limit loaded emails to those received ON or AFTER this date.")
        if self.start_date:
            self.dt_start.setDate(QDate.fromString(self.start_date, Qt.ISODate))
        else:
            self.dt_start.setDate(QDate.currentDate().addDays(-7))

        self.dt_end = QDateEdit()
        self.dt_end.setCalendarPopup(True)
        self.dt_end.setToolTip("Limit loaded emails to those received ON or BEFORE this date.")
        if self.end_date:
            self.dt_end.setDate(QDate.fromString(self.end_date, Qt.ISODate))
        else:
            self.dt_end.setDate(QDate.currentDate().addDays(7))

        self.dt_start.dateChanged.connect(self._on_date_changed)
        self.dt_end.dateChanged.connect(self._on_date_changed)

        self.btn_filter_star = QPushButton("⭐ Starred")
        self.btn_filter_star.setCheckable(True)
        self.btn_filter_star.setToolTip("Filter left panel to only show threads containing a Starred TDoc.")
        self.btn_filter_star.setStyleSheet(
            "QPushButton { padding: 4px 8px; background: white; border-radius: 3px; border: 1px solid #CCC; color: #333; } "
            "QPushButton:checked { background-color: #FFF4CE; font-weight: bold; border-color: #E2C08D; }"
        )
        self.btn_filter_star.toggled.connect(self.filters_changed.emit)

        self.btn_filter_follow = QPushButton("👀 Followed AIs")
        self.btn_filter_follow.setCheckable(True)
        self.btn_filter_follow.setToolTip("Filter left panel to only show threads belonging to a Followed Agenda Item.")
        self.btn_filter_follow.setStyleSheet(
            "QPushButton { padding: 4px 8px; background: white; border-radius: 3px; border: 1px solid #CCC; color: #333; } "
            "QPushButton:checked { background-color: #E6F4E6; font-weight: bold; color: #0C6B0C; border-color: #0C6B0C; }"
        )
        self.btn_filter_follow.toggled.connect(self.filters_changed.emit)

        self.cb_ai = CheckableComboBox("Filter by AI")
        self.cb_ai.setToolTip("Filter threads by specific Agenda Items.")
        self.cb_ai.selectionChanged.connect(self.filters_changed.emit)

        self.cb_company = CheckableComboBox("Company")
        self.cb_company.setMinimumWidth(110)
        self.cb_company.setToolTip("Filter the right panel emails by contributing Company.")
        self.cb_company.selectionChanged.connect(self.filters_changed.emit)

        self.cb_sender = CheckableComboBox("Sender")
        self.cb_sender.setMinimumWidth(110)
        self.cb_sender.setToolTip("Filter the right panel emails by specific Senders.")
        self.cb_sender.selectionChanged.connect(self.filters_changed.emit)

        row2_layout.addWidget(QLabel("📅 Limit Data:"))
        row2_layout.addWidget(self.dt_start)
        row2_layout.addWidget(QLabel("-"))
        row2_layout.addWidget(self.dt_end)
        row2_layout.addSpacing(10)
        row2_layout.addWidget(self.btn_filter_star)
        row2_layout.addWidget(self.btn_filter_follow)
        row2_layout.addWidget(self.cb_ai)
        row2_layout.addWidget(self.cb_company)
        row2_layout.addWidget(self.cb_sender)
        row2_layout.addStretch()

        top_layout.addLayout(row2_layout)

    def _on_date_changed(self):
        s_date = self.dt_start.date().toString(Qt.ISODate)
        e_date = self.dt_end.date().toString(Qt.ISODate)
        self.start_date = s_date
        self.end_date = e_date
        self.date_range_changed.emit(s_date, e_date)
        self.filters_changed.emit()

    def set_buttons_enabled(self, enabled: bool):
        for btn in [self.btn_sync, self.btn_move, self.btn_move_all, self.btn_rescan, self.btn_stats, self.btn_config]:
            btn.setEnabled(enabled)

    def set_status(self, text: str):
        self.lbl_status.setText(text)

    def set_count_text(self, text: str):
        self.lbl_count.setText(text)
