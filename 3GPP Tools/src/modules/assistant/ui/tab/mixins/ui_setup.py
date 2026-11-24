"""Widget construction and layout setup mixin for AssistantTab."""

from __future__ import annotations

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QSplitter,
    QTabWidget,
    QTextBrowser,
    QTreeWidget,
    QVBoxLayout,
    QWidget,
)

from core.ui.page_ribbon import PageRibbonCard, apply_lucide_icon
from core.ui.theme import C, R, S, T
from core.ui.ui_components import (
    BUTTON_STYLE_TOOLBAR_DANGER,
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    TREE_STYLE_CLEAN,
)
from modules.assistant.ui.assistant_components import (
    ActivityTimelineWidget,
    CARD_FRAME_STYLE,
    DiagnosticsTraceWidget,
    EvidenceInspectorWidget,
    MultilinePromptEdit,
    StarterChipsWidget,
    STATUS_PILL_NEUTRAL,
)


class AssistantUiSetupMixin:
    """Builds the main workspace, dock strip, chat card, and cockpit tabs."""

    def _setup_ui(self) -> None:
        self._host_layout = QVBoxLayout(self)
        self._host_layout.setContentsMargins(10, 8, 10, 8)
        self._host_layout.setSpacing(6)

        # Single-instance reparentable workspace body (tab <-> RightSidePanel)
        self.body_container = QWidget(self)
        self.body_container.setObjectName("assistantBodyContainer")
        self._body_layout = QVBoxLayout(self.body_container)
        self._body_layout.setContentsMargins(0, 0, 0, 0)
        self._body_layout.setSpacing(6)

        # ---------------------------------------------------------------------
        # Top Header Ribbon (Model Status, Session Actions, Layout Switcher)
        # ---------------------------------------------------------------------
        self.ribbon_card = PageRibbonCard("3GPP Assistant", parent=self.body_container)

        # Model status pill
        self.model_status_btn = QPushButton("⚪ Ollama: Checking...")
        self.model_status_btn.setStyleSheet(STATUS_PILL_NEUTRAL)
        self.model_status_btn.setCursor(Qt.PointingHandCursor)
        self.model_status_btn.setToolTip("Click to configure Ollama settings.")
        self.model_status_btn.clicked.connect(self._on_open_ollama_settings)
        self.ribbon_card.add_primary_widget(self.model_status_btn)

        self.status_label = QLabel("Ready")
        self.status_label.setStyleSheet(f"color: {C.TEXT_MUTED}; font-weight: 500; font-size: 12px;")
        self.ribbon_card.add_primary_widget(self.status_label)

        # Session Action Buttons (Utility strip)
        self.new_chat_btn = QPushButton("New Chat")
        self.new_chat_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        apply_lucide_icon(self.new_chat_btn, "plus")
        self.new_chat_btn.setToolTip("Clear the conversation and begin a new research session.")
        self.new_chat_btn.clicked.connect(self._on_new_chat)
        self.ribbon_card.add_utility_widget(self.new_chat_btn)

        self.copy_chat_btn = QPushButton("Copy Chat")
        self.copy_chat_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        apply_lucide_icon(self.copy_chat_btn, "copy")
        self.copy_chat_btn.setToolTip("Copy the full conversation transcript and citations to the clipboard.")
        self.copy_chat_btn.clicked.connect(self._on_copy_conversation)
        self.ribbon_card.add_utility_widget(self.copy_chat_btn)

        self.export_chat_btn = QPushButton("Export...")
        self.export_chat_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        apply_lucide_icon(self.export_chat_btn, "download")
        self.export_chat_btn.setToolTip("Export the research conversation to a Markdown (.md) file.")
        self.export_chat_btn.clicked.connect(self._on_export_markdown)
        self.ribbon_card.add_utility_widget(self.export_chat_btn)

        # Layout Switcher Button
        self.layout_toggle_btn = QPushButton("⬒ Bottom Split")
        self.layout_toggle_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        apply_lucide_icon(self.layout_toggle_btn, "columns")
        self.layout_toggle_btn.setCheckable(True)
        self.layout_toggle_btn.setChecked(True)  # True = Side-by-Side (Horizontal)
        self.layout_toggle_btn.setToolTip("Toggle between Side-by-Side and Bottom Split layouts.")
        self.layout_toggle_btn.clicked.connect(self._toggle_layout)
        self.ribbon_card.add_utility_widget(self.layout_toggle_btn)

        self._body_layout.addWidget(self.ribbon_card)

        # Slim docked status strip (shown instead of ribbon_card when docked)
        self.dock_compact_strip = QFrame(self.body_container)
        self.dock_compact_strip.setObjectName("assistantDockCompactStrip")
        self.dock_compact_strip.setFixedHeight(32)
        self.dock_compact_strip.setStyleSheet(
            f"QFrame#assistantDockCompactStrip {{"
            f" background-color: {C.BG_SURFACE}; border: 1px solid {C.BORDER_SUBTLE};"
            f" border-radius: {R.MD}px; }}"
        )
        strip_layout = QHBoxLayout(self.dock_compact_strip)
        strip_layout.setContentsMargins(S.XS, 2, S.XS, 2)
        strip_layout.setSpacing(S.XS)

        self.dock_model_status_btn = QPushButton("⚪ Ollama", self.dock_compact_strip)
        self.dock_model_status_btn.setStyleSheet(STATUS_PILL_NEUTRAL)
        self.dock_model_status_btn.setCursor(Qt.PointingHandCursor)
        self.dock_model_status_btn.setToolTip("Click to configure Ollama settings.")
        self.dock_model_status_btn.clicked.connect(self._on_open_ollama_settings)
        strip_layout.addWidget(self.dock_model_status_btn)

        self.dock_status_label = QLabel("Ready", self.dock_compact_strip)
        self.dock_status_label.setStyleSheet(
            f"color: {C.TEXT_MUTED}; font-weight: 500; font-size: {T.SIZE_CAPTION}px;"
        )
        strip_layout.addWidget(self.dock_status_label)
        strip_layout.addStretch(1)

        self.dock_sources_toggle_btn = QPushButton("📚 Sources (0)", self.dock_compact_strip)
        self.dock_sources_toggle_btn.setObjectName("dockSourcesToggleBtn")
        self.dock_sources_toggle_btn.setCheckable(True)
        self.dock_sources_toggle_btn.setCursor(Qt.PointingHandCursor)
        self.dock_sources_toggle_btn.setToolTip("Show or hide retrieved sources and activity drawer.")
        self.dock_sources_toggle_btn.setStyleSheet(
            f"QPushButton#dockSourcesToggleBtn {{"
            f" background-color: {C.BG_CARD}; color: {C.TEXT_PRIMARY};"
            f" border: 1px solid {C.BORDER_SUBTLE}; border-radius: {R.SM}px;"
            f" padding: 2px 8px; font-size: {T.SIZE_CAPTION}px; font-weight: 600; }}"
            f"QPushButton#dockSourcesToggleBtn:checked {{"
            f" background-color: {C.ACCENT_SOFT}; color: {C.ACCENT_PRIMARY};"
            f" border-color: {C.ACCENT_PRIMARY}; }}"
        )
        self.dock_sources_toggle_btn.toggled.connect(self.set_cockpit_visible_in_dock)
        strip_layout.addWidget(self.dock_sources_toggle_btn)

        self.dock_new_chat_btn = QPushButton("New", self.dock_compact_strip)
        self.dock_new_chat_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        apply_lucide_icon(self.dock_new_chat_btn, "plus")
        self.dock_new_chat_btn.setToolTip("Clear conversation and start a new research session.")
        self.dock_new_chat_btn.clicked.connect(self._on_new_chat)
        strip_layout.addWidget(self.dock_new_chat_btn)

        self.dock_compact_strip.setVisible(False)
        self._body_layout.addWidget(self.dock_compact_strip)

        # ---------------------------------------------------------------------
        # Main Splitter (Left: Chat & Input, Right: Research Cockpit Drawer)
        # ---------------------------------------------------------------------
        self.main_splitter = QSplitter(Qt.Horizontal)
        self.main_splitter.setChildrenCollapsible(False)

        # Left / Chat Widget
        chat_widget = QWidget()
        chat_layout = QVBoxLayout(chat_widget)
        chat_layout.setContentsMargins(0, 0, 0, 0)
        chat_layout.setSpacing(6)

        # Conversation Card Frame
        conv_frame = QFrame()
        conv_frame.setObjectName("assistantCardFrame")
        conv_frame.setStyleSheet(CARD_FRAME_STYLE)
        conv_frame_layout = QVBoxLayout(conv_frame)
        conv_frame_layout.setContentsMargins(8, 8, 8, 8)
        conv_frame_layout.setSpacing(6)

        # Starter Chips for empty state
        self.starter_chips = StarterChipsWidget()
        self.starter_chips.prompt_selected.connect(self._on_starter_chip_selected)
        conv_frame_layout.addWidget(self.starter_chips)

        self.conversation_browser = QTextBrowser()
        self.conversation_browser.setOpenExternalLinks(False)
        self.conversation_browser.anchorClicked.connect(self._on_anchor_clicked)
        self.conversation_browser.setStyleSheet(f"""
            QTextBrowser {{
                background-color: {C.BG_CARD};
                border: none;
                font-family: "Segoe UI", Arial, sans-serif;
                font-size: 13px;
                color: {C.TEXT_STRONG};
            }}
        """)
        self.conversation_browser.setPlaceholderText(
            "Ask a 3GPP question. The Assistant will research local specification "
            "and protocol knowledge before answering."
        )
        conv_frame_layout.addWidget(self.conversation_browser, stretch=1)
        chat_layout.addWidget(conv_frame, stretch=1)

        # Input Card Frame
        input_frame = QFrame()
        input_frame.setObjectName("assistantCardFrame")
        input_frame.setStyleSheet(CARD_FRAME_STYLE)
        input_frame_layout = QVBoxLayout(input_frame)
        input_frame_layout.setContentsMargins(8, 8, 8, 8)
        input_frame_layout.setSpacing(6)

        self.question_input = MultilinePromptEdit()
        self.question_input.submit_requested.connect(self._on_send)
        self.question_input.textChanged.connect(self._on_prompt_changed)
        input_frame_layout.addWidget(self.question_input)

        # Action / Control bar under input
        input_action_bar = QHBoxLayout()
        input_action_bar.setContentsMargins(0, 0, 0, 0)
        input_action_bar.setSpacing(6)

        self.char_count_label = QLabel("0 chars")
        self.char_count_label.setStyleSheet(f"color: {C.TEXT_FAINT}; font-size: 11px;")
        input_action_bar.addWidget(self.char_count_label)

        self.clear_input_btn = QPushButton("Clear")
        self.clear_input_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        apply_lucide_icon(self.clear_input_btn, "rotate-ccw")
        self.clear_input_btn.setEnabled(False)
        self.clear_input_btn.clicked.connect(self.question_input.clear)
        input_action_bar.addWidget(self.clear_input_btn)

        input_action_bar.addStretch()

        self.send_btn = QPushButton("Send")
        self.send_btn.setObjectName("primaryBtn")
        apply_lucide_icon(self.send_btn, "send", color=C.TEXT_ON_ACCENT)
        self.send_btn.clicked.connect(self._on_send)
        input_action_bar.addWidget(self.send_btn)

        self.stop_btn = QPushButton("■ Stop")
        self.stop_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_DANGER)
        self.stop_btn.setEnabled(False)
        self.stop_btn.clicked.connect(self._on_stop)
        input_action_bar.addWidget(self.stop_btn)

        input_frame_layout.addLayout(input_action_bar)
        chat_layout.addWidget(input_frame)

        self.main_splitter.addWidget(chat_widget)

        # ---------------------------------------------------------------------
        # Right / Research Cockpit Drawer (Tabbed Sources, Activity, Trace)
        # ---------------------------------------------------------------------
        self.cockpit_tabs = QTabWidget()
        self.cockpit_tabs.setStyleSheet(f"""
            QTabWidget::pane {{
                border: 1px solid {C.BORDER_SUBTLE};
                border-radius: 6px;
                background-color: {C.BG_CARD};
                top: -1px;
            }}
            QTabBar::tab {{
                background-color: {C.BG_SURFACE};
                border: 1px solid {C.BORDER_SUBTLE};
                padding: 6px 12px;
                margin-right: 2px;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
                font-weight: 600;
                font-size: 12px;
                color: {C.TEXT_MUTED};
            }}
            QTabBar::tab:selected {{
                background-color: {C.BG_CARD};
                color: {C.ACCENT_PRIMARY};
                border-bottom: 1px solid {C.BG_CARD};
            }}
        """)

        # Tab 1: Sources & Evidence Inspector
        sources_widget = QWidget()
        sources_layout = QVBoxLayout(sources_widget)
        sources_layout.setContentsMargins(6, 6, 6, 6)
        sources_layout.setSpacing(6)

        sources_splitter = QSplitter(Qt.Vertical)
        sources_splitter.setChildrenCollapsible(False)

        # Sources Tree Card
        self.sources_tree = QTreeWidget()
        self.sources_tree.setStyleSheet(TREE_STYLE_CLEAN)
        self.sources_tree.setHeaderLabels(["ID", "Source", "Version / Clause"])
        self.sources_tree.setAlternatingRowColors(True)
        self.sources_tree.setRootIsDecorated(False)
        self.sources_tree.itemClicked.connect(self._on_source_tree_item_clicked)
        self.sources_tree.itemDoubleClicked.connect(self._on_source_tree_item_clicked)
        sources_splitter.addWidget(self.sources_tree)

        # Embedded Evidence Inspector
        self.evidence_inspector = EvidenceInspectorWidget()
        self.evidence_inspector.open_spec_search_requested.connect(self.open_spec_search_requested)
        self.evidence_inspector.open_specs_requested.connect(self.open_specs_requested)
        self.evidence_inspector.open_protocol_requested.connect(self.open_protocol_requested)
        sources_splitter.addWidget(self.evidence_inspector)

        sources_splitter.setSizes([180, 240])
        sources_layout.addWidget(sources_splitter)
        self.cockpit_tabs.addTab(sources_widget, "📚 Sources (0)")

        # Tab 2: Research Activity Timeline
        self.activity_widget = ActivityTimelineWidget()
        self.cockpit_tabs.addTab(self.activity_widget, "⚡ Activity (0)")

        # Tab 3: Diagnostics & Ollama Trace
        self.diagnostics_widget = DiagnosticsTraceWidget()
        self.cockpit_tabs.addTab(self.diagnostics_widget, "🔬 Diagnostics")

        self.main_splitter.addWidget(self.cockpit_tabs)

        # Default side-by-side sizes (65% Chat, 35% Drawer)
        self.main_splitter.setSizes([680, 420])
        self._body_layout.addWidget(self.main_splitter, stretch=1)

        self._host_layout.addWidget(self.body_container, stretch=1)

        # Placeholder shown in tab 8 only while Assistant body is docked on the right
        self.docked_placeholder_card = QFrame(self)
        self.docked_placeholder_card.setObjectName("assistantDockedPlaceholder")
        self.docked_placeholder_card.setStyleSheet(CARD_FRAME_STYLE)
        ph_layout = QVBoxLayout(self.docked_placeholder_card)
        ph_layout.setContentsMargins(S.LG, S.LG, S.LG, S.LG)
        ph_layout.setSpacing(S.SM)
        ph_layout.addStretch(1)

        ph_title = QLabel("Assistant is currently docked in the right side panel")
        ph_title.setAlignment(Qt.AlignCenter)
        ph_title.setStyleSheet(
            f"color: {C.TEXT_STRONG}; font-family: {T.FONT_FAMILY_UI};"
            f" font-size: {T.SIZE_SECTION_TITLE}px; font-weight: 600;"
        )
        ph_layout.addWidget(ph_title)

        ph_sub = QLabel(
            "Your active conversation, streaming progress, and evidence citations remain live in the side panel."
        )
        ph_sub.setWordWrap(True)
        ph_sub.setAlignment(Qt.AlignCenter)
        ph_sub.setStyleSheet(
            f"color: {C.TEXT_MUTED}; font-family: {T.FONT_FAMILY_UI}; font-size: {T.SIZE_BODY}px;"
        )
        ph_layout.addWidget(ph_sub)

        btn_row = QHBoxLayout()
        btn_row.addStretch(1)
        self.return_to_tab_btn = QPushButton("Return Assistant to Full Tab")
        self.return_to_tab_btn.setObjectName("primaryBtn")
        apply_lucide_icon(self.return_to_tab_btn, "maximize-2", color=C.TEXT_ON_ACCENT)
        self.return_to_tab_btn.clicked.connect(self.undock_requested.emit)
        btn_row.addWidget(self.return_to_tab_btn)
        btn_row.addStretch(1)
        ph_layout.addLayout(btn_row)
        ph_layout.addStretch(1)

        self.docked_placeholder_card.setVisible(False)
        self._host_layout.addWidget(self.docked_placeholder_card, stretch=1)
