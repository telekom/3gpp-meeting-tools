"""UI setup mixin for the modeless EmailBriefingsDialog (Per-AI & Saved Selections)."""

from __future__ import annotations

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMenu,
    QPushButton,
    QSplitter,
    QTextBrowser,
    QTreeWidget,
    QVBoxLayout,
    QWidget,
)

from core.ui.font_scale import bind_reader_zoom_settings, scale_px
from core.ui.theme import C
from core.ui.ui_components import BUTTON_STYLE_TOOLBAR_SECONDARY
from modules.emails.core.markdown_export import apply_email_reader_document_style


class BriefingsUiSetupMixin:
    """Builds the master-detail Per-AI & Saved Selections briefing cockpit."""

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(10, 8, 10, 8)
        root.setSpacing(6)

        self.main_splitter = QSplitter(Qt.Horizontal, self)
        self.main_splitter.setChildrenCollapsible(False)

        # --- Left Navigation Tree ---
        left_panel = QWidget(self.main_splitter)
        left_panel.setMinimumWidth(scale_px(250))
        left_lay = QVBoxLayout(left_panel)
        left_lay.setContentsMargins(0, 0, 4, 0)
        left_lay.setSpacing(6)

        nav_hdr = QHBoxLayout()
        nav_lbl = QLabel("<b>🧭 Agenda Items & Selections</b>", left_panel)
        nav_lbl.setStyleSheet(f"color: {C.TEXT_PRIMARY}; font-size: 13px;")
        nav_hdr.addWidget(nav_lbl, 1)

        self.btn_refresh_tree = QPushButton("🔄", left_panel)
        self.btn_refresh_tree.setToolTip("Reload AI topics, email counts, and freshness status.")
        self.btn_refresh_tree.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_refresh_tree.clicked.connect(self.refresh_navigation)
        nav_hdr.addWidget(self.btn_refresh_tree)
        left_lay.addLayout(nav_hdr)

        self.nav_filter_edit = QLineEdit(left_panel)
        self.nav_filter_edit.setClearButtonEnabled(True)
        self.nav_filter_edit.setPlaceholderText("🔍 Filter AI or selection...")
        self.nav_filter_edit.setMinimumHeight(scale_px(26))
        self.nav_filter_edit.textChanged.connect(self._filter_nav_tree)
        left_lay.addWidget(self.nav_filter_edit)

        self.nav_tree = QTreeWidget(left_panel)
        self.nav_tree.setHeaderHidden(True)
        self.nav_tree.setIndentation(16)
        self.nav_tree.setStyleSheet(
            f"QTreeWidget {{ background-color: {C.BG_CARD}; border: 1px solid {C.BORDER_STRONG}; border-radius: 5px; }}"
        )
        self.nav_tree.currentItemChanged.connect(self._on_nav_selection_changed)
        self.nav_tree.setContextMenuPolicy(Qt.CustomContextMenu)
        self.nav_tree.customContextMenuRequested.connect(self._on_nav_context_menu)
        left_lay.addWidget(self.nav_tree, 1)

        # --- Right Briefing Detail Pane ---
        right_panel = QWidget(self.main_splitter)
        right_panel.setMinimumWidth(scale_px(420))
        right_lay = QVBoxLayout(right_panel)
        right_lay.setContentsMargins(4, 0, 0, 0)
        right_lay.setSpacing(6)

        hdr_card = QFrame(right_panel)
        hdr_card.setObjectName("BriefingHeaderCard")
        hdr_card.setStyleSheet(
            f"#BriefingHeaderCard {{ background-color: {C.BG_SURFACE}; border: 1px solid {C.BORDER_SUBTLE}; border-radius: 6px; }}"
        )
        hdr_lay = QVBoxLayout(hdr_card)
        hdr_lay.setContentsMargins(10, 8, 10, 8)
        hdr_lay.setSpacing(6)

        title_row = QHBoxLayout()
        title_row.setSpacing(8)
        self.scope_title_lbl = QLabel("Select an Agenda Item or Selection", hdr_card)
        self.scope_title_lbl.setWordWrap(True)
        self.scope_title_lbl.setStyleSheet(f"font-size: 14px; font-weight: bold; color: {C.TEXT_PRIMARY};")
        title_row.addWidget(self.scope_title_lbl, 1)

        self.freshness_lbl = QLabel("", hdr_card)
        self.freshness_lbl.setStyleSheet(
            f"background-color: {C.BG_CARD}; color: {C.TEXT_SECONDARY}; border: 1px solid {C.BORDER_SUBTLE}; "
            f"border-radius: 4px; padding: 2px 8px; font-size: 11px; font-weight: bold;"
        )
        title_row.addWidget(self.freshness_lbl)
        hdr_lay.addLayout(title_row)

        ctrl_row = QHBoxLayout()
        ctrl_row.setSpacing(6)
        ctrl_h = scale_px(26)

        self.btn_update = QPushButton("⚡ Generate / Update", hdr_card)
        self.btn_update.setMinimumHeight(ctrl_h)
        self.btn_update.setCursor(Qt.PointingHandCursor)
        self.btn_update.setStyleSheet(
            f"QPushButton {{ background-color: {C.ACCENT_PRIMARY}; color: {C.TEXT_ON_ACCENT}; "
            f"border: none; border-radius: 4px; padding: 4px 12px; font-weight: bold; font-size: 12px; }} "
            f"QPushButton:hover {{ background-color: {C.ACCENT_HOVER}; }} "
            f"QPushButton:disabled {{ background-color: {C.BORDER_STRONG}; color: {C.TEXT_MUTED}; }}"
        )
        self.btn_update.clicked.connect(lambda: self._run_active_briefing(mode="auto"))
        ctrl_row.addWidget(self.btn_update)

        self.btn_rebuild = QPushButton("🔄 Full Rebuild", hdr_card)
        self.btn_rebuild.setMinimumHeight(ctrl_h)
        self.btn_rebuild.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_rebuild.setToolTip("Re-summarize all stale/all leaf topics and rebuild the cross-thinking briefing.")
        self.btn_rebuild.clicked.connect(lambda: self._run_active_briefing(mode="full"))
        ctrl_row.addWidget(self.btn_rebuild)

        self.btn_model = QPushButton("⚙️ Model ▾", hdr_card)
        self.btn_model.setMinimumHeight(ctrl_h)
        self.btn_model.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.model_menu = QMenu(self)
        self.model_menu.aboutToShow.connect(self._populate_model_menu)
        self.btn_model.setMenu(self.model_menu)
        ctrl_row.addWidget(self.btn_model)

        self.btn_m365_copilot = QPushButton("💠 M365 Copilot ▾", hdr_card)
        self.btn_m365_copilot.setMinimumHeight(ctrl_h)
        self.btn_m365_copilot.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_m365_copilot.setToolTip(
            "Prepare Per-AI or Selection briefing pack for Microsoft 365 Copilot and paste the response back."
        )
        m365_menu = QMenu(self)
        act_m365_send = m365_menu.addAction("🚀 Send Briefing Pack to M365 Copilot...")
        act_m365_send.triggered.connect(lambda _=False: self._open_m365_copilot_handoff(auto_launch=True))
        act_m365_paste = m365_menu.addAction("📋 Paste M365 / External Briefing...")
        act_m365_paste.triggered.connect(lambda _=False: self._open_m365_copilot_handoff(auto_launch=False))
        self.btn_m365_copilot.setMenu(m365_menu)
        ctrl_row.addWidget(self.btn_m365_copilot)

        self.btn_focus_toggle = QPushButton("🎯 Focus", hdr_card)
        self.btn_focus_toggle.setCheckable(True)
        self.btn_focus_toggle.setMinimumHeight(ctrl_h)
        self.btn_focus_toggle.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_focus_toggle.setToolTip("Show or hide optional strategic focus prompt for this briefing.")
        self.btn_focus_toggle.toggled.connect(lambda v: self.focus_row_widget.setVisible(bool(v)))
        ctrl_row.addWidget(self.btn_focus_toggle)

        ctrl_row.addStretch()

        self.btn_copy = QPushButton("📋 Copy", hdr_card)
        self.btn_copy.setMinimumHeight(ctrl_h)
        self.btn_copy.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_copy.clicked.connect(self._on_copy_briefing)
        ctrl_row.addWidget(self.btn_copy)

        self.btn_export_md = QPushButton("💾 Export .md", hdr_card)
        self.btn_export_md.setMinimumHeight(ctrl_h)
        self.btn_export_md.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_export_md.clicked.connect(self._on_export_briefing_md)
        ctrl_row.addWidget(self.btn_export_md)

        self.btn_stop = QPushButton("⏹ Stop", hdr_card)
        self.btn_stop.setMinimumHeight(ctrl_h)
        self.btn_stop.setStyleSheet(
            f"QPushButton {{ background-color: {C.DANGER_SOFT}; color: {C.DANGER}; "
            f"border: 1px solid {C.DANGER_BORDER_HOVER}; border-radius: 4px; padding: 3px 10px; font-weight: bold; }}"
        )
        self.btn_stop.setVisible(False)
        self.btn_stop.clicked.connect(self._cancel_active_worker)
        ctrl_row.addWidget(self.btn_stop)

        hdr_lay.addLayout(ctrl_row)

        self.focus_row_widget = QWidget(hdr_card)
        focus_lay = QHBoxLayout(self.focus_row_widget)
        focus_lay.setContentsMargins(0, 2, 0, 0)
        focus_lay.setSpacing(6)
        focus_lay.addWidget(QLabel("Focus:", self.focus_row_widget))
        self.focus_prompt_edit = QLineEdit(self.focus_row_widget)
        self.focus_prompt_edit.setMinimumHeight(scale_px(24))
        self.focus_prompt_edit.setPlaceholderText("Optional lens (e.g., 'Compare Option 1 vs Option 2' or 'Highlight roaming/signaling impacts')...")
        focus_lay.addWidget(self.focus_prompt_edit, 1)
        self.focus_row_widget.setVisible(False)
        hdr_lay.addWidget(self.focus_row_widget)

        self.progress_lbl = QLabel("", hdr_card)
        self.progress_lbl.setStyleSheet(f"color: {C.INFO}; font-size: 11px; font-weight: 500;")
        self.progress_lbl.setVisible(False)
        hdr_lay.addWidget(self.progress_lbl)

        right_lay.addWidget(hdr_card)

        self.right_vsplit = QSplitter(Qt.Vertical, right_panel)
        self.right_vsplit.setChildrenCollapsible(False)

        self.briefing_browser = QTextBrowser(self.right_vsplit)
        self.briefing_browser.setOpenLinks(False)
        apply_email_reader_document_style(self.briefing_browser)
        self.briefing_browser.anchorClicked.connect(self._on_anchor_clicked)
        bind_reader_zoom_settings(self.briefing_browser, self._settings, key="zoom_steps")

        # Collapsible Underlying Leaf Summaries Drawer
        self.leaves_frame = QFrame(self.right_vsplit)
        self.leaves_frame.setObjectName("BriefingLeavesDrawer")
        self.leaves_frame.setStyleSheet(
            f"#BriefingLeavesDrawer {{ background-color: {C.BG_SURFACE}; border: 1px solid {C.BORDER_SUBTLE}; border-radius: 6px; }}"
        )
        leaves_lay = QVBoxLayout(self.leaves_frame)
        leaves_lay.setContentsMargins(8, 4, 8, 6)
        leaves_lay.setSpacing(4)

        self.btn_toggle_leaves = QPushButton("▸ Underlying Thread / TDoc Summaries (0)", self.leaves_frame)
        self.btn_toggle_leaves.setCursor(Qt.PointingHandCursor)
        self.btn_toggle_leaves.setStyleSheet(
            f"QPushButton {{ background: transparent; border: none; color: {C.TEXT_STRONG}; "
            f"font-weight: bold; font-size: 11px; text-align: left; padding: 2px 0; }}"
        )
        self.btn_toggle_leaves.clicked.connect(self._on_toggle_leaves_drawer)
        leaves_lay.addWidget(self.btn_toggle_leaves)

        self.leaves_browser = QTextBrowser(self.leaves_frame)
        self.leaves_browser.setOpenLinks(False)
        apply_email_reader_document_style(self.leaves_browser)
        self.leaves_browser.anchorClicked.connect(self._on_anchor_clicked)
        self.leaves_browser.setVisible(False)
        leaves_lay.addWidget(self.leaves_browser, 1)

        self.right_vsplit.setSizes([520, 120])
        right_lay.addWidget(self.right_vsplit, 1)

        self.main_splitter.setSizes([290, 810])
        root.addWidget(self.main_splitter, 1)
