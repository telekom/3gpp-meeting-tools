# --- File: src/modules/meetings/ui/tdocs/window/mixins/filter_setup.py ---
"""
Filter bar construction and simple filter-change handlers / toggle buttons.

Mixin of TDocsWindow: relies on attributes set in TDocsWindow.__init__ / _setup_*.
"""
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtWidgets import (
    QWidget,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QFrame,
    QPushButton,
    QCheckBox,
    QComboBox,
    QSizePolicy,
)
from core.ui.flow_layout import FlowLayout
from core.ui.font_scale import scale_px
from modules.meetings.ui.tdocs_components import CheckableComboBox, make_rich_tooltip
from modules.emails.ui.sync_progress_bar import AsyncSyncProgressBar
from core.ui.theme import C


class FilterSetupMixin:
    """Filter bar construction and simple filter-change handlers / toggle buttons."""

    def _setup_filters(self, layout):
        self.search_timer = QTimer(self)
        self.search_timer.setSingleShot(True)
        self.search_timer.setInterval(300)
        self.search_timer.timeout.connect(self._apply_search_filter)

        filter_frame = QFrame()
        filter_frame.setObjectName("tdocFilterFrame")
        filter_sp = QSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)
        filter_sp.setHeightForWidth(True)
        filter_frame.setSizePolicy(filter_sp)
        filter_frame.setStyleSheet(f"""
            QFrame#tdocFilterFrame {{
                background-color: {C.BG_CARD};
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 6px;
            }}
            QLabel {{
                font-weight: bold;
                color: {C.TEXT_SECONDARY};
                border: none;
            }}
            QLineEdit, QComboBox {{
                padding: 4px 8px;
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 4px;
                background: {C.BG_CARD};
            }}
            QComboBox[activeFilter="true"] {{
                background-color: {C.INFO_SOFT};
                border: 1.5px solid {C.INFO_ACCENT};
                color: {C.INFO};
                font-weight: bold;
            }}
        """)
        filter_layout = FlowLayout(filter_frame, margin=0, h_spacing=5, v_spacing=4)
        filter_layout.setContentsMargins(6, 4, 6, 4)
        self._filter_frame = filter_frame

        search_group = QWidget(filter_frame)
        search_group.setStyleSheet(".QWidget { background: transparent; border: none; }")
        search_layout = QHBoxLayout(search_group)
        search_layout.setContentsMargins(0, 0, 0, 0)
        search_layout.setSpacing(5)

        search_lbl = QLabel("🔍")
        search_lbl.setToolTip("Full-text search across TDoc number, title, source, remarks, and related metadata.")
        search_layout.addWidget(search_lbl)

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText('Search TDocs (e.g. "CR", -draft)...')
        self.search_input.setToolTip(
            'Search across TDoc metadata.\n'
            '• Exact phrases: "draft baseline"\n'
            '• Exclude words: -draft or !draft\n'
            '• Substrings & numbers: S2-260855, TS 23.501'
        )
        self.search_input.setClearButtonEnabled(True)
        self.search_input.setMinimumWidth(scale_px(180))
        self.search_input.setMaximumWidth(scale_px(240))
        self.search_input.textChanged.connect(lambda _: self.search_timer.start())
        search_layout.addWidget(self.search_input, stretch=1)

        self.btn_followed_wis = QPushButton("🎯")
        self.btn_followed_wis.setCheckable(True)
        self.btn_followed_wis.setCursor(Qt.PointingHandCursor)
        self.btn_followed_wis.setFixedSize(scale_px(28), scale_px(28))
        self.btn_followed_wis.setContextMenuPolicy(Qt.CustomContextMenu)
        self.btn_followed_wis.customContextMenuRequested.connect(self._show_tracked_wi_context_menu)
        self.btn_followed_wis.setToolTip(make_rich_tooltip(
            title="🎯 Followed Work Items Filter",
            description="Filter to TDocs associated with Work Items you have selected to follow.",
            tip="Right-click to manage your Followed Work Items watchlist."
        ))
        self.btn_followed_wis.clicked.connect(self._toggle_followed_wis_filter)
        search_layout.addWidget(self.btn_followed_wis)

        self.btn_starred_wis = QPushButton("⭐")
        self.btn_starred_wis.setCheckable(True)
        self.btn_starred_wis.setCursor(Qt.PointingHandCursor)
        self.btn_starred_wis.setFixedSize(scale_px(28), scale_px(28))
        self.btn_starred_wis.setContextMenuPolicy(Qt.CustomContextMenu)
        self.btn_starred_wis.customContextMenuRequested.connect(self._show_tracked_wi_context_menu)
        self.btn_starred_wis.setToolTip(make_rich_tooltip(
            title="⭐ Starred Work Items Filter",
            description="Filter to TDocs associated with Work Items you have starred (personal priority watchlist).",
            tip="Right-click to toggle including unassigned incoming Liaison Statements (LS)."
        ))
        self.btn_starred_wis.clicked.connect(self._toggle_starred_wis_filter)
        search_layout.addWidget(self.btn_starred_wis)

        self.btn_agreements_filter = QPushButton("⚖️")
        self.btn_agreements_filter.setCheckable(True)
        self.btn_agreements_filter.setCursor(Qt.PointingHandCursor)
        self.btn_agreements_filter.setFixedSize(scale_px(28), scale_px(28))
        self.btn_agreements_filter.setToolTip(make_rich_tooltip(
            title="⚖️ Chairman's Notes Agreements Filter (RAN)",
            description="Filter table to only TDocs with Chairman's Notes text agreements or working assumptions.",
            tip="Agreements are highlighted in soft green and working assumptions in warm amber."
        ))
        self.btn_agreements_filter.clicked.connect(self._toggle_agreements_filter)
        if self.is_ran:
            search_layout.addWidget(self.btn_agreements_filter)
        else:
            self.btn_agreements_filter.setVisible(False)

        self.btn_unread_emails = QPushButton("🔵")
        self.btn_unread_emails.setCheckable(True)
        self.btn_unread_emails.setCursor(Qt.PointingHandCursor)
        self.btn_unread_emails.setFixedSize(scale_px(28), scale_px(28))
        self.btn_unread_emails.setToolTip(make_rich_tooltip(
            title="🔵 Unread Emails Filter",
            description="Show only TDocs and Tracked Email Threads that have unread emails.",
            tip="Click again to clear, or click 🚩 to switch to flagged emails."
        ))
        self.btn_unread_emails.clicked.connect(self._toggle_unread_emails_filter)
        search_layout.addWidget(self.btn_unread_emails)

        self.btn_flagged_emails = QPushButton("🚩")
        self.btn_flagged_emails.setCheckable(True)
        self.btn_flagged_emails.setCursor(Qt.PointingHandCursor)
        self.btn_flagged_emails.setFixedSize(scale_px(28), scale_px(28))
        self.btn_flagged_emails.setToolTip(make_rich_tooltip(
            title="🚩 Flagged Emails Filter",
            description="Show only TDocs and Tracked Email Threads that have Outlook-flagged emails.",
            tip="Click again to clear, or click 🔵 to switch to unread emails."
        ))
        self.btn_flagged_emails.clicked.connect(self._toggle_flagged_emails_filter)
        search_layout.addWidget(self.btn_flagged_emails)

        self._update_tracked_wi_button_styles()
        filter_layout.addWidget(search_group)

        self.company_combo = CheckableComboBox("Company")
        self.company_combo.setMaximumWidth(scale_px(165))
        self.company_combo.setContextMenuPolicy(Qt.CustomContextMenu)
        self.company_combo.customContextMenuRequested.connect(self._show_company_filter_context_menu)
        self.company_combo.setToolTip(
            'Filter by contributing company (excludes unconfirmed `?` and `[...]` sources by default).\n'
            'Right-click to toggle including unconfirmed (?, [...]) carry-over sources.'
        )
        self.company_combo.selectionChanged.connect(self._on_company_changed)
        filter_layout.addWidget(self.company_combo)

        self.type_combo = CheckableComboBox("Type")
        self.type_combo.setMaximumWidth(scale_px(125))
        self.type_combo.setToolTip('Filter by TDoc type. Grey entries currently have no intersection with the other active filters but remain selectable.')
        self.type_combo.selectionChanged.connect(self._on_type_changed)
        filter_layout.addWidget(self.type_combo)

        self.ai_combo = CheckableComboBox("AI")
        self.ai_combo.setMaximumWidth(scale_px(130))
        self.ai_combo.setToolTip('Filter by Agenda Item. When agenda.csv is available, associated acronyms are shown in parentheses.')
        self.ai_combo.selectionChanged.connect(self._on_ai_changed)
        filter_layout.addWidget(self.ai_combo)

        self.sub_item_combo = CheckableComboBox("Sub-Item")
        self.sub_item_combo.setMaximumWidth(scale_px(150))
        self.sub_item_combo.setToolTip('Filter by Agenda Sub-Item / topic separator from TdocsByAgenda.htm or Chairman notes.')
        self.sub_item_combo.selectionChanged.connect(self._on_sub_item_changed)
        self.sub_item_combo.setVisible(False)
        filter_layout.addWidget(self.sub_item_combo)

        self.wi_combo = CheckableComboBox("WI")
        self.wi_combo.setMaximumWidth(scale_px(135))
        self.wi_combo.setToolTip('Filter by the Related WIs metadata from the official TDoc list. Grey entries currently have no intersection with the other active filters.')
        self.wi_combo.selectionChanged.connect(self._on_wi_changed)
        filter_layout.addWidget(self.wi_combo)

        self.status_combo = CheckableComboBox("TDoc Status")
        self.status_combo.setMaximumWidth(scale_px(155))
        self.status_combo.setToolTip('Filter by the meeting/secretary TDoc status.')
        self.status_combo.selectionChanged.connect(self._on_status_changed)
        filter_layout.addWidget(self.status_combo)

        self.my_status_combo = CheckableComboBox("My Status")
        self.my_status_combo.setMaximumWidth(scale_px(145))
        self.my_status_combo.setToolTip('Filter by your personal TDoc status.')
        self.my_status_combo.selectionChanged.connect(self._on_my_status_changed)
        filter_layout.addWidget(self.my_status_combo)

        self.tags_combo = CheckableComboBox("Tags")
        self.tags_combo.setMaximumWidth(scale_px(145))
        self.tags_combo.setToolTip('Filter TDocs and tracked email threads by your personal tags.')
        self.tags_combo.selectionChanged.connect(self._on_tags_changed)
        self.tags_combo.setVisible(False)
        filter_layout.addWidget(self.tags_combo)

        tail_group = QWidget(filter_frame)
        tail_group.setStyleSheet(".QWidget { background: transparent; border: none; }")
        tail_layout = QHBoxLayout(tail_group)
        tail_layout.setContentsMargins(0, 0, 0, 0)
        tail_layout.setSpacing(5)

        if self.is_sa2:
            self.chk_no_comments = QCheckBox("No Comments")
            self.chk_no_comments.setToolTip('Show only TDocs without secretary/chair comments.')
            self.chk_no_comments.toggled.connect(self._on_no_comments_toggled)
            tail_layout.addWidget(self.chk_no_comments)

        order_lbl = QLabel("Order:")
        order_lbl.setStyleSheet(f"font-size: 11px; color: {C.TEXT_MUTED}; font-weight: bold;")
        tail_layout.addWidget(order_lbl)

        self.order_combo = QComboBox()
        self.order_combo.setSizeAdjustPolicy(QComboBox.AdjustToContents)
        self.order_combo.setMaximumWidth(scale_px(165))
        self.order_combo.setToolTip(
            "Change the display ordering of TDocs.\n"
            "• 📋 Agenda Order: Sort by Agenda Item and Sub-Item discussion order (superseded by TdocsByAgenda.htm).\n"
            "• 🔢 TDoc Number: Sort alphanumerically by TDoc identifier.\n"
            "• 🏢 Company: Group by contributing company.\n"
            "• ✉️ Most Emails: Show most discussed TDoc families first.\n"
            "• 🗂️ Default / File Order: Original document arrival order."
        )
        self.order_combo.addItem("📋 Agenda Order", "AGENDA")
        self.order_combo.addItem("🔢 TDoc Number", "TDOC")
        self.order_combo.addItem("🏢 Company", "COMPANY")
        self.order_combo.addItem("✉️ Most Emails", "EMAILS")
        self.order_combo.addItem("🗂️ Default / File Order", "DEFAULT")
        self.order_combo.currentIndexChanged.connect(self._on_order_changed)
        tail_layout.addWidget(self.order_combo)

        self.row_mode_combo = QComboBox()
        self.row_mode_combo.setSizeAdjustPolicy(QComboBox.AdjustToContents)
        self.row_mode_combo.setMaximumWidth(scale_px(135))
        self.row_mode_combo.setToolTip("Filter between all rows, official TDocs only, or pinned 📧 Email Threads only.")
        self.row_mode_combo.addItem("📋 All Rows", "ALL")
        self.row_mode_combo.addItem("📄 TDocs Only", "TDOCS")
        self.row_mode_combo.addItem("📧 Threads Only", "THREADS")
        self.row_mode_combo.currentIndexChanged.connect(
            lambda _: self.proxy.setRowModeFilter(self.row_mode_combo.currentData()) if hasattr(self, "proxy") else None
        )
        tail_layout.addWidget(self.row_mode_combo)

        self.btn_reset_filters = QPushButton("✕ Clear")
        self.btn_reset_filters.setToolTip("Reset all search text and active filters to show all TDocs.")
        self.btn_reset_filters.setCursor(Qt.PointingHandCursor)
        self.btn_reset_filters.setStyleSheet(f"""
            QPushButton {{
                background-color: {C.BG_SURFACE};
                color: {C.TEXT_MUTED};
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 4px;
                padding: 4px 8px;
                font-size: 11px;
                font-weight: bold;
            }}
            QPushButton:hover {{
                background-color: {C.DANGER_SOFT_HOVER};
                border-color: {C.DANGER_BORDER_ALT};
                color: {C.DANGER_ALT};
            }}
        """)
        self.btn_reset_filters.clicked.connect(self._clear_all_filters)
        tail_layout.addWidget(self.btn_reset_filters)
        filter_layout.addWidget(tail_group)

        layout.addWidget(filter_frame)

        self.thread_sync_bar = AsyncSyncProgressBar(self)
        self.thread_sync_bar.cancel_requested.connect(self._on_cancel_tracked_threads_sync)
        layout.addWidget(self.thread_sync_bar)
        self._tracked_threads_sync_worker = None
        self._sync_silent_mode = False

    def _apply_search_filter(self):
        self.proxy.setGlobalFilter(self.search_input.text())
        self._update_filter_facets()
        QTimer.singleShot(0, self._update_count_label)

    def _on_type_changed(self, types):
        self.proxy.setTypeFilters(types)
        self._update_filter_facets()
        QTimer.singleShot(0, self._update_count_label)

    def _on_ai_changed(self, ais):
        self.proxy.setAIFilters(ais)
        self._update_filter_facets()
        QTimer.singleShot(0, self._update_count_label)

    def _on_wi_changed(self, wis):
        self.proxy.setWIFilters(wis)
        self._update_filter_facets()
        QTimer.singleShot(0, self._update_count_label)

    def _on_status_changed(self, statuses):
        self.proxy.setStatusFilters(statuses)
        self._update_filter_facets()
        QTimer.singleShot(0, self._update_count_label)

    def _on_my_status_changed(self, statuses):
        self.proxy.setMyStatusFilters(statuses)
        self._update_filter_facets()
        QTimer.singleShot(0, self._update_count_label)

    def _on_tags_changed(self, tags):
        all_checked = (
            self.tags_combo.model().rowCount() > 0
            and self.tags_combo.model().item(0).checkState() == Qt.Checked
        )
        self.proxy.setTagFilters(tags, active=not all_checked)
        self._update_filter_facets()
        QTimer.singleShot(0, self._update_count_label)

    def _on_company_changed(self, companies):
        self.proxy.setCompanyFilters(companies)
        self._update_filter_facets()
        QTimer.singleShot(0, self._update_count_label)

    def _on_sub_item_changed(self, sub_items):
        self.proxy.setSubItemFilters(sub_items)
        self._update_filter_facets()
        QTimer.singleShot(0, self._update_count_label)

    def _on_order_changed(self, index):
        mode = self.order_combo.currentData() or "AGENDA"
        self.proxy.setSortMode(mode)
        if mode == "AGENDA":
            self.table.sortByColumn(10, Qt.AscendingOrder)
        elif mode == "TDOC":
            self.table.sortByColumn(1, Qt.AscendingOrder)
        elif mode == "COMPANY":
            self.table.sortByColumn(3, Qt.AscendingOrder)
        elif mode == "EMAILS":
            self.table.sortByColumn(13, Qt.DescendingOrder)
        elif mode == "DEFAULT":
            self.proxy.sort(-1)
        QTimer.singleShot(0, self._update_count_label)

    def _toggle_followed_wis_filter(self):
        self._load_tracked_wi_tokens()
        self.proxy.setTrackedWIFilters(
            self.btn_followed_wis.isChecked(),
            self.btn_starred_wis.isChecked()
        )
        self._update_tracked_wi_button_styles()
        self._update_filter_facets()
        QTimer.singleShot(0, self._update_count_label)

    def _toggle_agreements_filter(self):
        """Toggles table filtering to only TDocs with agreements or working assumptions."""
        checked = self.btn_agreements_filter.isChecked()
        self.proxy.setAgreementsFilter(checked)
        self._update_tracked_wi_button_styles()

    def _toggle_unread_emails_filter(self):
        """Toggles 1-click filtering to rows with unread emails."""
        checked = self.btn_unread_emails.isChecked()
        if checked and hasattr(self, "btn_flagged_emails"):
            self.btn_flagged_emails.blockSignals(True)
            self.btn_flagged_emails.setChecked(False)
            self.btn_flagged_emails.blockSignals(False)
        self.proxy.setEmailStateFilter("UNREAD" if checked else "ALL")
        self._update_tracked_wi_button_styles()
        self._update_filter_facets()
        QTimer.singleShot(0, self._update_count_label)

    def _toggle_flagged_emails_filter(self):
        """Toggles 1-click filtering to rows with Outlook-flagged emails."""
        checked = self.btn_flagged_emails.isChecked()
        if checked and hasattr(self, "btn_unread_emails"):
            self.btn_unread_emails.blockSignals(True)
            self.btn_unread_emails.setChecked(False)
            self.btn_unread_emails.blockSignals(False)
        self.proxy.setEmailStateFilter("FLAGGED" if checked else "ALL")
        self._update_tracked_wi_button_styles()
        self._update_filter_facets()
        QTimer.singleShot(0, self._update_count_label)

    def _on_no_comments_toggled(self, checked):
        self.proxy.setNoCommentsFilter(checked)
        self._update_filter_facets()
        QTimer.singleShot(0, self._update_count_label)
