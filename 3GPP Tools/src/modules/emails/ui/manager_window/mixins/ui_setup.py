"""UI setup and folder/stats config persistence mixin for EmailManagerWindow."""

from __future__ import annotations

import json

from PyQt5.QtCore import QSettings, Qt
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QPushButton,
    QSplitter,
    QTableView,
    QTextBrowser,
    QVBoxLayout,
    QWidget,
)

from core.ui.font_scale import bind_reader_zoom_settings, create_help_tips_pill
from core.ui.theme import C
from modules.emails.core.thread_markdown_exporter import apply_email_reader_document_style
from modules.emails.ui.email_models import (
    EmailProxyModel,
    EmailTableModel,
    TDocProxyModel,
    TDocSummaryModel,
)
from modules.emails.ui.email_toolbar import EmailControlToolbar


class EmailManagerUiSetupMixin:
    """Config persistence and split-pane Qt layout setup for EmailManagerWindow."""

    def _load_config(self) -> dict:
        if self.config_path.exists():
            try:
                with open(self.config_path, "r") as f:
                    return json.load(f)
            except Exception:
                pass
        return {
            "source_folder": "",
            "target_folder": "",
            "stats_config": {
                "email_top_companies": 25,
                "email_top_delegates": 25,
                "email_heatmap_top_comps": 25,
                "email_heatmap_top_ais": 25,
            },
        }

    def _save_config(
        self,
        source_folder: str,
        target_folder: str,
        sd: str,
        ed: str,
        stats_cfg: dict = None,
    ):
        self.source_folder = source_folder
        self.target_folder = target_folder
        self.start_date = sd
        self.end_date = ed

        if stats_cfg is not None:
            self.stats_config = stats_cfg
        elif not hasattr(self, "stats_config"):
            self.stats_config = self._load_config().get("stats_config", {})

        with open(self.config_path, "w") as f:
            json.dump(
                {
                    "source_folder": source_folder,
                    "target_folder": target_folder,
                    "start_date": sd,
                    "end_date": ed,
                    "stats_config": self.stats_config,
                },
                f,
                indent=4,
            )

    def _configure_folders(self):
        from modules.emails.ui.config_dialog import EmailConfigDialog

        current_stats = getattr(self, "stats_config", self._load_config().get("stats_config", {}))

        dialog = EmailConfigDialog(self.source_folder, self.target_folder, current_stats, self)
        if dialog.exec_() == QDialog.Accepted:
            config_data = dialog.get_config_data()
            src = config_data["source_folder"]
            tgt = config_data["target_folder"]
            new_stats = config_data["stats_config"]

            self.source_folder = src
            self.target_folder = tgt
            self.stats_config = new_stats
            self._save_config(src, tgt, self.start_date, self.end_date, new_stats)

    def _setup_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(10)

        self.toolbar = EmailControlToolbar(
            start_date=getattr(self, "start_date", ""),
            end_date=getattr(self, "end_date", ""),
            parent=self,
        )
        self.toolbar.sync_requested.connect(self._run_sync)
        self.toolbar.move_requested.connect(self._run_move)
        self.toolbar.move_all_requested.connect(self._run_move_all)
        self.toolbar.rescan_requested.connect(self._run_target_rescan)
        self.toolbar.stats_requested.connect(self._generate_statistics)
        self.toolbar.config_requested.connect(self._configure_folders)
        self.toolbar.filters_changed.connect(self._apply_filters)
        self.toolbar.date_range_changed.connect(self._on_date_range_changed)

        self.btn_sync = self.toolbar.btn_sync
        self.btn_move = self.toolbar.btn_move
        self.btn_move_all = self.toolbar.btn_move_all
        self.btn_rescan = self.toolbar.btn_rescan
        self.btn_stats = self.toolbar.btn_stats
        self.btn_config = self.toolbar.btn_config
        self.lbl_status = self.toolbar.lbl_status
        self.lbl_count = self.toolbar.lbl_count
        self.dt_start = self.toolbar.dt_start
        self.dt_end = self.toolbar.dt_end
        self.btn_filter_star = self.toolbar.btn_filter_star
        self.btn_filter_follow = self.toolbar.btn_filter_follow
        self.cb_ai = self.toolbar.cb_ai
        self.cb_company = self.toolbar.cb_company
        self.cb_sender = self.toolbar.cb_sender

        main_layout.addWidget(self.toolbar, stretch=0)

        self.main_splitter = QSplitter(Qt.Horizontal)

        # Left panel: TDoc threads
        self.left_panel = QWidget()
        left_layout = QVBoxLayout(self.left_panel)
        left_layout.setContentsMargins(0, 0, 0, 0)

        left_header_layout = QHBoxLayout()
        self.lbl_left_title = QLabel("📄 Active Threads")
        self.lbl_left_title.setStyleSheet("font-weight: bold; color: #555;")

        self.tdoc_search_input = QLineEdit()
        self.tdoc_search_input.setPlaceholderText("🔍 Search threads...")
        self.tdoc_search_input.setToolTip("Filter the thread list by TDoc number or Subject keywords.")
        self.tdoc_search_input.setStyleSheet(
            "padding: 2px 4px; border: 1px solid #CCC; border-radius: 4px; background: #FFF; color: #333;"
        )
        self.tdoc_search_input.textChanged.connect(lambda _: self.search_timer.start())

        left_header_layout.addWidget(self.lbl_left_title)
        left_header_layout.addStretch()
        left_header_layout.addWidget(self.tdoc_search_input)
        left_layout.addLayout(left_header_layout)

        self.tdoc_view = QTableView()
        self.tdoc_view.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.tdoc_view.setSelectionMode(QAbstractItemView.SingleSelection)
        self.tdoc_view.verticalHeader().setVisible(False)
        self.tdoc_view.setAlternatingRowColors(True)
        self.tdoc_view.doubleClicked.connect(self._on_tdoc_double_clicked)
        self.tdoc_view.setStyleSheet(
            f"QTableView {{ gridline-color: {C.BORDER_GRAY}; border: 1px solid {C.BORDER_GRAY}; background-color: {C.BG_CARD}; }} "
            f"QHeaderView::section {{ background-color: {C.BG_SURFACE}; padding: 4px; font-weight: bold; border: 1px solid {C.BORDER_GRAY}; }}"
        )
        left_layout.addWidget(self.tdoc_view)

        # Right panel: Emails & reader
        self.right_panel = QWidget()
        right_layout = QVBoxLayout(self.right_panel)
        right_layout.setContentsMargins(0, 0, 0, 0)

        right_header_layout = QHBoxLayout()
        self.lbl_right_title = QLabel("📧 Thread History")
        self.lbl_right_title.setStyleSheet(f"font-weight: bold; color: {C.ACCENT_BRIGHT};")

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("🔍 Search this thread...")
        self.search_input.setToolTip(
            "Search for specific text, sender, or company within the currently selected thread."
        )
        self.search_input.setMaximumWidth(200)
        self.search_input.setStyleSheet(
            "padding: 2px 4px; border: 1px solid #CCC; border-radius: 4px; background: #FFF; color: #333;"
        )
        self.search_input.textChanged.connect(lambda _: self.search_timer.start())

        right_header_layout.addWidget(self.lbl_right_title)
        right_header_layout.addStretch()
        right_header_layout.addWidget(self.search_input)
        right_layout.addLayout(right_header_layout)

        right_splitter = QSplitter(Qt.Vertical)

        self.email_view = QTableView()
        self.email_view.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.email_view.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.email_view.verticalHeader().setVisible(False)
        self.email_view.setAlternatingRowColors(True)
        self.email_view.setCursor(Qt.PointingHandCursor)
        self.email_view.clicked.connect(self._on_table_clicked)
        self.email_view.doubleClicked.connect(self._on_email_double_clicked)
        self.email_view.setContextMenuPolicy(Qt.CustomContextMenu)
        self.email_view.customContextMenuRequested.connect(self._show_context_menu)
        self.email_view.setStyleSheet(
            f"QTableView {{ gridline-color: {C.BORDER_GRAY}; border: 1px solid {C.BORDER_GRAY}; background-color: {C.BG_CARD}; }} "
            f"QHeaderView::section {{ background-color: {C.BG_SURFACE}; padding: 4px; font-weight: bold; border: 1px solid {C.BORDER_GRAY}; }}"
        )
        right_splitter.addWidget(self.email_view)

        pane_widget = QWidget()
        pane_layout = QVBoxLayout(pane_widget)
        pane_layout.setContentsMargins(0, 10, 0, 0)

        self.btn_open_msg = QPushButton("📄 Open .msg File")
        self.btn_open_msg.setToolTip("Open this email directly in native Microsoft Outlook.")
        self.btn_open_msg.setEnabled(False)
        self.btn_open_msg.clicked.connect(self._open_current_msg)

        self.btn_open_folder = QPushButton("📁 Open Folder")
        self.btn_open_folder.setToolTip("Open the folder containing this .msg file.")
        self.btn_open_folder.setEnabled(False)
        self.btn_open_folder.clicked.connect(self._open_current_msg_folder)

        self.btn_toggle_star = QPushButton("⭐ Star TDoc")
        self.btn_toggle_star.setToolTip("Toggle the Star status for this TDoc globally.")
        self.btn_toggle_star.setEnabled(False)
        self.btn_toggle_star.clicked.connect(self._toggle_current_star)

        self.btn_toggle_follow = QPushButton("★ Follow AI")
        self.btn_toggle_follow.setToolTip("Toggle the Follow status for this Agenda Item globally.")
        self.btn_toggle_follow.setEnabled(False)
        self.btn_toggle_follow.clicked.connect(self._toggle_current_ai_follow)

        btn_layout = QHBoxLayout()
        btn_layout.addWidget(self.btn_open_msg)
        btn_layout.addWidget(self.btn_open_folder)
        btn_layout.addWidget(self.btn_toggle_star)
        btn_layout.addWidget(self.btn_toggle_follow)
        btn_layout.addStretch()
        self.help_pill_btn = create_help_tips_pill(
            ["Ctrl + Mouse Wheel in reading pane: zoom text larger/smaller (Ctrl+0 to reset; saved automatically)"],
            parent=pane_widget,
            title="Email Reader Tips",
        )
        btn_layout.addWidget(self.help_pill_btn)

        self.reading_pane = QTextBrowser()
        self._reader_settings = QSettings("3GPP_Tools", "EmailManagerWindow")
        bind_reader_zoom_settings(self.reading_pane, self._reader_settings, key="zoom_steps")
        apply_email_reader_document_style(self.reading_pane)
        self.reading_pane.setStyleSheet(
            "background: white; border: 1px solid #CCC; border-radius: 4px; padding: 10px; color: #333;"
        )
        pane_layout.addLayout(btn_layout)
        pane_layout.addWidget(self.reading_pane)

        right_splitter.addWidget(pane_widget)
        right_splitter.setSizes([500, 250])

        right_layout.addWidget(right_splitter)

        self.main_splitter.addWidget(self.left_panel)
        self.main_splitter.addWidget(self.right_panel)
        self.main_splitter.setSizes([350, 1000])
        main_layout.addWidget(self.main_splitter, stretch=1)

        # Model setup
        self.tdoc_model = TDocSummaryModel()
        self.tdoc_proxy = TDocProxyModel()
        self.tdoc_proxy.setSourceModel(self.tdoc_model)
        self.tdoc_view.setModel(self.tdoc_proxy)

        self.email_model = EmailTableModel()
        self.email_proxy = EmailProxyModel()
        self.email_proxy.setSourceModel(self.email_model)
        self.email_view.setModel(self.email_proxy)

        self.tdoc_view.selectionModel().selectionChanged.connect(self._on_tdoc_selected)
        self.email_view.selectionModel().selectionChanged.connect(self._on_email_selected)

        self.tdoc_proxy.layoutChanged.connect(self._update_count_label)
        self.email_proxy.layoutChanged.connect(self._update_count_label)

        tdoc_header = self.tdoc_view.horizontalHeader()
        tdoc_header.setSectionResizeMode(QHeaderView.Interactive)
        tdoc_header.resizeSection(0, 25)
        tdoc_header.setSectionResizeMode(1, QHeaderView.Stretch)
        tdoc_header.resizeSection(2, 60)
        tdoc_header.resizeSection(3, 45)

        email_header = self.email_view.horizontalHeader()
        email_header.setSectionResizeMode(QHeaderView.Interactive)
        email_header.resizeSection(0, 30)
        email_header.resizeSection(1, 85)
        email_header.resizeSection(2, 75)
        email_header.resizeSection(3, 120)
        email_header.resizeSection(4, 90)
        email_header.resizeSection(5, 60)
        email_header.resizeSection(6, 70)
        email_header.resizeSection(7, 110)
        email_header.resizeSection(8, 140)
        email_header.setSectionResizeMode(9, QHeaderView.Stretch)
