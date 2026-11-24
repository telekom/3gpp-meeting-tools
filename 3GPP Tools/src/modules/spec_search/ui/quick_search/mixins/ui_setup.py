"""UI setup and styling mixin for QuickSpecSearchPanel."""

from __future__ import annotations

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QPushButton,
    QSplitter,
    QTextBrowser,
    QVBoxLayout,
)

from core.ui.icons import lucide_icon
from core.ui.theme import C, S
from ..delegate import _QuickSearchResultDelegate
from ..styles import apply_quick_search_styles


class QuickSearchUiSetupMixin:
    """Builds filter inputs, result list, collapsible clause preview, and input helpers."""

    def _setup_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(S.SM, S.SM, S.SM, S.SM)
        root.setSpacing(S.XS)

        # --- Search & Filter Header Card ---
        filter_card = QFrame(self)
        filter_card.setObjectName("quickSearchFilterCard")
        fc_layout = QVBoxLayout(filter_card)
        fc_layout.setContentsMargins(S.SM, S.SM, S.SM, S.SM)
        fc_layout.setSpacing(S.XS)

        self.search_input = QLineEdit(filter_card)
        self.search_input.setObjectName("quickSearchInput")
        self.search_input.setPlaceholderText("Search clause text (e.g. registration, DNN, PEI)…")
        self.search_input.setClearButtonEnabled(True)
        self.search_input.textChanged.connect(self._schedule_search)
        self.search_input.returnPressed.connect(self._execute_search)
        fc_layout.addWidget(self.search_input)

        row2 = QHBoxLayout()
        row2.setContentsMargins(0, 0, 0, 0)
        row2.setSpacing(S.XS)

        self.spec_filter_input = QLineEdit(filter_card)
        self.spec_filter_input.setObjectName("quickSearchSubInput")
        self.spec_filter_input.setPlaceholderText("Spec: e.g. 23.501")
        self.spec_filter_input.setClearButtonEnabled(True)
        self.spec_filter_input.textChanged.connect(self._schedule_search)
        row2.addWidget(self.spec_filter_input, 1)

        self.clause_filter_input = QLineEdit(filter_card)
        self.clause_filter_input.setObjectName("quickSearchSubInput")
        self.clause_filter_input.setPlaceholderText("Clause: e.g. 5.5")
        self.clause_filter_input.setClearButtonEnabled(True)
        self.clause_filter_input.textChanged.connect(self._schedule_search)
        row2.addWidget(self.clause_filter_input, 1)

        self.latest_only_btn = QPushButton("Latest per Spec", filter_card)
        self.latest_only_btn.setObjectName("quickSearchTogglePill")
        self.latest_only_btn.setCheckable(True)
        self.latest_only_btn.setChecked(True)
        self.latest_only_btn.setCursor(Qt.PointingHandCursor)
        self.latest_only_btn.setToolTip(
            "When checked, limits results to the newest indexed version per specification."
        )
        self.latest_only_btn.toggled.connect(self._schedule_search)
        row2.addWidget(self.latest_only_btn)

        fc_layout.addLayout(row2)
        root.addWidget(filter_card)

        # --- Status Strip ---
        self.status_label = QLabel("Type keywords to search indexed clauses.", self)
        self.status_label.setObjectName("quickSearchStatusLabel")
        self.status_label.setWordWrap(True)
        root.addWidget(self.status_label)

        # --- Results + Collapsible Preview Splitter ---
        self.results_splitter = QSplitter(Qt.Vertical, self)
        self.results_splitter.setChildrenCollapsible(False)

        self.results_list = QListWidget(self.results_splitter)
        self.results_list.setObjectName("quickSearchResultsList")
        self.results_list.setUniformItemSizes(False)
        self.results_list.setVerticalScrollMode(QListWidget.ScrollPerPixel)
        self.results_list.setItemDelegate(_QuickSearchResultDelegate(self.results_list))
        self.results_list.itemClicked.connect(self._on_result_selected)
        self.results_list.currentItemChanged.connect(self._on_current_item_changed)
        self.results_splitter.addWidget(self.results_list)

        self.preview_frame = QFrame(self.results_splitter)
        self.preview_frame.setObjectName("quickSearchPreviewFrame")
        prev_layout = QVBoxLayout(self.preview_frame)
        prev_layout.setContentsMargins(S.SM, S.SM, S.SM, S.SM)
        prev_layout.setSpacing(S.XS)

        prev_hdr = QHBoxLayout()
        prev_hdr.setContentsMargins(0, 0, 0, 0)
        prev_hdr.setSpacing(S.XS)

        self.preview_title_label = QLabel("Clause Preview", self.preview_frame)
        self.preview_title_label.setObjectName("quickSearchPreviewTitle")
        self.preview_title_label.setWordWrap(True)
        prev_hdr.addWidget(self.preview_title_label, 1)

        self.close_preview_btn = QPushButton("", self.preview_frame)
        self.close_preview_btn.setObjectName("quickSearchIconBtn")
        self.close_preview_btn.setIcon(
            lucide_icon("x", color=C.TEXT_MUTED, active_color=C.TEXT_PRIMARY, size=13)
        )
        self.close_preview_btn.setToolTip("Collapse clause preview")
        self.close_preview_btn.setCursor(Qt.PointingHandCursor)
        self.close_preview_btn.clicked.connect(lambda: self.preview_frame.setVisible(False))
        prev_hdr.addWidget(self.close_preview_btn)
        prev_layout.addLayout(prev_hdr)

        self.preview_browser = QTextBrowser(self.preview_frame)
        self.preview_browser.setObjectName("quickSearchPreviewBrowser")
        self.preview_browser.setOpenExternalLinks(False)
        prev_layout.addWidget(self.preview_browser, 1)

        action_row = QHBoxLayout()
        action_row.setContentsMargins(0, 0, 0, 0)
        action_row.setSpacing(S.XS)

        self.open_search_btn = QPushButton("Spec Search", self.preview_frame)
        self.open_search_btn.setObjectName("quickSearchActionBtn")
        self.open_search_btn.setIcon(lucide_icon("search", color=C.ACCENT_PRIMARY, size=13))
        self.open_search_btn.setCursor(Qt.PointingHandCursor)
        self.open_search_btn.setToolTip("Open this spec and clause in full Spec Search tab")
        self.open_search_btn.clicked.connect(self._on_open_in_spec_search)
        action_row.addWidget(self.open_search_btn)

        self.open_specs_btn = QPushButton("Specs", self.preview_frame)
        self.open_specs_btn.setObjectName("quickSearchActionBtn")
        self.open_specs_btn.setIcon(lucide_icon("book-open", color=C.TEXT_SECONDARY, size=13))
        self.open_specs_btn.setCursor(Qt.PointingHandCursor)
        self.open_specs_btn.setToolTip("Open this specification in the Specifications tab")
        self.open_specs_btn.clicked.connect(self._on_open_in_specs)
        action_row.addWidget(self.open_specs_btn)

        action_row.addStretch(1)

        self.copy_citation_btn = QPushButton("Copy", self.preview_frame)
        self.copy_citation_btn.setObjectName("quickSearchActionBtn")
        self.copy_citation_btn.setIcon(lucide_icon("copy", color=C.TEXT_SECONDARY, size=13))
        self.copy_citation_btn.setCursor(Qt.PointingHandCursor)
        self.copy_citation_btn.setToolTip("Copy clause citation and excerpt to clipboard")
        self.copy_citation_btn.clicked.connect(self._on_copy_citation)
        action_row.addWidget(self.copy_citation_btn)

        prev_layout.addLayout(action_row)
        self.preview_frame.setVisible(False)
        self.results_splitter.addWidget(self.preview_frame)
        self.results_splitter.setStretchFactor(0, 1)
        self.results_splitter.setStretchFactor(1, 0)

        root.addWidget(self.results_splitter, 1)

    def _apply_styles(self) -> None:
        apply_quick_search_styles(self)

    def current_query(self) -> str:
        return self.search_input.text().strip()

    def focus_input(self) -> None:
        self.search_input.setFocus()
        self.search_input.selectAll()
