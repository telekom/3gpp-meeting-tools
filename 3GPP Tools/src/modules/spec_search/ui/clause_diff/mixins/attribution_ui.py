"""Modifying TDoc Attribution tab UI and background worker lifecycle mixin."""

import html
from typing import Any, Dict, List

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QAbstractItemView, QFrame, QHBoxLayout, QHeaderView, QLabel,
    QProgressBar, QPushButton, QSplitter, QTableWidget, QTextBrowser,
    QVBoxLayout, QWidget,
)

from core.ui.theme import C


class _AttributionUiMixin:
    """Builds Tab 3 (Modifying TDoc Attribution) and runs Stage 1 background attribution."""

    def _setup_attribution_tab(self):
        """Sets up Tab 3: Modifying TDoc Attribution (Stage 1 Metadata & Stage 2 Document Fingerprinting)."""
        self.tab_attribution = QWidget()
        layout = QVBoxLayout(self.tab_attribution)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(8)

        # 1. Header Control Bar
        header_bar = QHBoxLayout()
        header_bar.setSpacing(8)

        self.attr_status_lbl = QLabel(f"<b>Clause Attribution:</b> Identifying originating CR/p-CR for Clause {self.clause_number}...")
        self.attr_status_lbl.setStyleSheet(f"font-size: 12px; color: {C.TEXT_STRONG};")
        header_bar.addWidget(self.attr_status_lbl, stretch=1)

        self.btn_attr_run = QPushButton("🔄 Re-Analyze")
        self.btn_attr_run.setToolTip("Re-run multi-stage clause attribution for current releases.")
        self.btn_attr_run.setStyleSheet(f"padding: 4px 12px; font-weight: bold; background-color: {C.BG_CANVAS}; border: 1px solid {C.BORDER_STRONG}; border-radius: 4px;")
        self.btn_attr_run.setCursor(Qt.PointingHandCursor)
        self.btn_attr_run.clicked.connect(lambda: self._run_attribution(force=True))
        header_bar.addWidget(self.btn_attr_run)

        self.btn_attr_deep_verify = QPushButton("🔬 Deep Verify Top 3 (DOCX)")
        self.btn_attr_deep_verify.setToolTip("Download and inspect Word OpenXML track changes in top 3 candidate documents for 100% deterministic text match.")
        self.btn_attr_deep_verify.setStyleSheet(f"padding: 4px 12px; font-weight: bold; background-color: {C.INFO_SOFT}; color: {C.INFO}; border: 1px solid {C.INFO_BORDER_MEDIUM}; border-radius: 4px;")
        self.btn_attr_deep_verify.setCursor(Qt.PointingHandCursor)
        self.btn_attr_deep_verify.setEnabled(False)
        self.btn_attr_deep_verify.clicked.connect(self._run_deep_verify)
        header_bar.addWidget(self.btn_attr_deep_verify)

        layout.addLayout(header_bar)

        # Progress Bar
        self.attr_progress_bar = QProgressBar()
        self.attr_progress_bar.setRange(0, 100)
        self.attr_progress_bar.setFixedHeight(8)
        self.attr_progress_bar.setTextVisible(False)
        self.attr_progress_bar.setStyleSheet(f"""
            QProgressBar {{
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 3px;
                background-color: {C.BG_CANVAS};
            }}
            QProgressBar::chunk {{
                background-color: {C.ACCENT_SKY};
                border-radius: 2px;
            }}
        """)
        self.attr_progress_bar.setVisible(False)
        layout.addWidget(self.attr_progress_bar)

        # 2. Hero Banner Card (Top Match)
        self.attr_hero_banner = QFrame()
        self.attr_hero_banner.setStyleSheet(f"""
            QFrame {{
                background-color: {C.SUCCESS_SOFT};
                border: 1px solid {C.SUCCESS_BORDER};
                border-radius: 6px;
                padding: 6px;
            }}
        """)
        self.attr_hero_banner.setVisible(False)
        hero_layout = QVBoxLayout(self.attr_hero_banner)
        hero_layout.setContentsMargins(8, 8, 8, 8)
        hero_layout.setSpacing(6)

        # Hero Top Row: Title + Action Buttons
        hero_row1 = QHBoxLayout()
        self.hero_match_title_lbl = QLabel("🎯 <b>Top Candidate:</b> Loading...")
        self.hero_match_title_lbl.setStyleSheet(f"font-size: 13px; color: {C.SUCCESS};")
        hero_row1.addWidget(self.hero_match_title_lbl, stretch=1)

        self.btn_hero_open = QPushButton("📄 Open TDoc")
        self.btn_hero_open.setToolTip("Open contribution document (.docx) in Microsoft Word.")
        self.btn_hero_open.setStyleSheet(f"padding: 3px 10px; font-weight: bold; background-color: {C.SUCCESS_DARK}; color: white; border-radius: 4px;")
        self.btn_hero_open.setCursor(Qt.PointingHandCursor)
        hero_row1.addWidget(self.btn_hero_open)

        self.btn_hero_jump = QPushButton("🚀 Jump to Meetings")
        self.btn_hero_jump.setToolTip("Navigate to Meetings tab with this TDoc focused.")
        self.btn_hero_jump.setStyleSheet(f"padding: 3px 10px; font-weight: bold; background-color: {C.BG_CANVAS}; color: {C.TEXT_BODY}; border: 1px solid {C.BORDER_STRONG}; border-radius: 4px;")
        self.btn_hero_jump.setCursor(Qt.PointingHandCursor)
        hero_row1.addWidget(self.btn_hero_jump)

        self.btn_hero_copy = QPushButton("📋 Copy Citation")
        self.btn_hero_copy.setToolTip("Copy plain text and rich HTML citation.")
        self.btn_hero_copy.setStyleSheet(f"padding: 3px 10px; font-weight: bold; background-color: {C.BG_CANVAS}; color: {C.TEXT_BODY}; border: 1px solid {C.BORDER_STRONG}; border-radius: 4px;")
        self.btn_hero_copy.setCursor(Qt.PointingHandCursor)
        hero_row1.addWidget(self.btn_hero_copy)

        hero_layout.addLayout(hero_row1)

        self.hero_doc_title_lbl = QLabel("")
        self.hero_doc_title_lbl.setStyleSheet(f"font-size: 11px; color: {C.TEXT_BODY}; font-style: italic;")
        self.hero_doc_title_lbl.setWordWrap(True)
        hero_layout.addWidget(self.hero_doc_title_lbl)

        self.hero_badges_layout = QHBoxLayout()
        self.hero_badges_layout.setSpacing(6)
        hero_layout.addLayout(self.hero_badges_layout)

        layout.addWidget(self.attr_hero_banner)

        # 3. Middle & Bottom Splitter: Candidates Table & Details Browser
        splitter = QSplitter(Qt.Vertical)

        # Table of Candidates
        self.attr_table = QTableWidget(0, 6)
        self.attr_table.setHorizontalHeaderLabels(["Confidence", "TDoc / CR #", "Source", "Title / Subject", "Evidence Signals", "Actions"])
        self.attr_table.horizontalHeader().setStretchLastSection(False)
        self.attr_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.attr_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.attr_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.attr_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        self.attr_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
        self.attr_table.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeToContents)
        self.attr_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.attr_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.attr_table.verticalHeader().setVisible(False)
        self.attr_table.itemSelectionChanged.connect(self._on_candidate_table_selection_changed)
        self.attr_table.setStyleSheet(f"""
            QTableWidget {{
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 4px;
                background-color: {C.BG_CARD};
                font-size: 11px;
                selection-background-color: {C.ACCENT_SOFT};
                selection-color: {C.TEXT_PRIMARY};
            }}
            QHeaderView::section {{
                background-color: {C.BG_SURFACE};
                font-weight: bold;
                color: {C.TEXT_SECONDARY};
                padding: 4px 6px;
                border: none;
                border-bottom: 1px solid {C.BORDER_STRONG};
            }}
        """)
        splitter.addWidget(self.attr_table)

        # Details Browser
        self.attr_details_browser = QTextBrowser()
        self.attr_details_browser.setOpenExternalLinks(True)
        self.attr_details_browser.setPlaceholderText("Select a candidate contribution above to inspect full metadata, meeting agenda context, and verified text excerpts...")
        self.attr_details_browser.setStyleSheet(f"""
            QTextBrowser {{
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 4px;
                background-color: {C.BG_SURFACE};
                padding: 6px;
                font-size: 11px;
                line-height: 1.4;
            }}
        """)
        splitter.addWidget(self.attr_details_browser)
        splitter.setSizes([300, 160])

        layout.addWidget(splitter)
        self.tab_widget.addTab(self.tab_attribution, "🎯 Modifying TDoc")

    def _focus_and_run_attribution(self):
        """Switches to the Modifying TDoc tab and triggers attribution."""
        self.tab_widget.setCurrentIndex(2)
        self._run_attribution()

    def _on_tab_changed(self, index: int):
        """Trigger attribution lazily when user navigates to the Attribution tab."""
        if index == 2:
            query_key = (self.spec_number, self.clause_number, self.version_b, self.version_a)
            if self._attribution_query_key != query_key or self._cached_attribution_results is None:
                self._run_attribution()

    def _run_attribution(self, force: bool = False):
        """Executes multi-stage clause attribution via background thread."""
        query_key = (self.spec_number, self.clause_number, self.version_b, self.version_a)
        if not force and self._attribution_query_key == query_key and self._cached_attribution_results is not None:
            return

        if not self.version_b:
            self.attr_status_lbl.setText("Select target specification release first.")
            return

        self._attribution_query_key = query_key
        self.attr_progress_bar.setVisible(True)
        self.attr_progress_bar.setValue(10)
        self.attr_status_lbl.setText(f"Analyzing candidates for TS {self.spec_number} Clause {self.clause_number} in v{self.version_b}...")
        self.btn_attr_run.setEnabled(False)
        self.btn_attr_deep_verify.setEnabled(False)

        target_clause = self.db.get_clause_content_by_spec_ver(self.spec_number, self.version_b, self.clause_number)
        target_content = target_clause.get("content", "") if target_clause else ""

        if not hasattr(self, "attr_engine"):
            from modules.specifications.core.spec_clause_attribution import SpecClauseAttributionEngine
            self.attr_engine = SpecClauseAttributionEngine(search_db=self.db)

        from modules.specifications.core.spec_clause_attribution import SpecClauseAttributionThread
        self._attr_worker = SpecClauseAttributionThread(
            engine=self.attr_engine,
            spec_number=self.spec_number,
            clause_number=self.clause_number,
            target_version=self.version_b,
            base_version=self.version_a,
            clause_title=self.clause_title,
            diff_text=self._raw_diff_text,
            clause_content=target_content,
            parent=self,
        )
        self._attr_worker.progress.connect(self._on_attribution_progress)
        self._attr_worker.finished.connect(self._on_attribution_finished)
        self._attr_worker.error.connect(self._on_attribution_error)
        self._attr_worker.start()

    def _on_attribution_progress(self, val: int, msg: str):
        self.attr_progress_bar.setValue(val)
        self.attr_status_lbl.setText(f"<b>Attribution:</b> {html.escape(msg)}")

    def _on_attribution_finished(self, candidates: List[Dict[str, Any]]):
        self.attr_progress_bar.setVisible(False)
        self.btn_attr_run.setEnabled(True)
        self._cached_attribution_results = candidates
        self._populate_attribution_view(candidates)

    def _on_attribution_error(self, err_msg: str):
        self.attr_progress_bar.setVisible(False)
        self.btn_attr_run.setEnabled(True)
        self.attr_status_lbl.setText(f"<span style='color: {C.DANGER};'><b>Attribution Error:</b> {html.escape(err_msg)}</span>")
