"""Research-oriented Qt UI for the local 3GPP Assistant."""

from __future__ import annotations

import logging
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Set

from PyQt5.QtCore import Qt, QTimer, QUrl, pyqtSignal
from PyQt5.QtWidgets import (
    QApplication, QFileDialog, QFrame, QHBoxLayout, QLabel, QMessageBox,
    QPushButton, QSplitter, QTabWidget, QTextBrowser, QTreeWidget, QTreeWidgetItem,
    QVBoxLayout, QWidget,
)

from core.ui.page_ribbon import PageRibbonCard, apply_lucide_icon
from core.ui.toast import notify_info, notify_success
from core.ui.ui_components import (
    BUTTON_STYLE_TOOLBAR_DANGER,
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    TREE_STYLE_CLEAN,
)
from modules.assistant.core.agent_controller import AgentControllerConfig
from modules.assistant.core.agent_models import (
    AgentResult, Evidence, EvidenceSourceType, ResearchActivity,
)
from modules.assistant.core.ollama_agent_adapter import OllamaAgentAdapter
from modules.assistant.core.tools import build_read_only_tool_registry
from modules.assistant.ui.assistant_components import (
    ActivityTimelineWidget,
    CARD_FRAME_STYLE,
    DiagnosticsTraceWidget,
    EvidenceInspectorWidget,
    MultilinePromptEdit,
    StarterChipsWidget,
    STATUS_PILL_OFFLINE,
    STATUS_PILL_ONLINE,
    STATUS_PILL_NEUTRAL,
    render_assistant_message_card,
    render_markdown_to_html,
    render_system_message_card,
    render_user_message_card,
)
from modules.assistant.ui.assistant_worker import AgentResearchWorker
from core.ui.theme import C, R, S, T

logger = logging.getLogger(__name__)


class AssistantTab(QWidget):
    """Conversation + objective research activity + application-owned sources."""

    open_ollama_settings_requested = pyqtSignal()
    open_spec_search_requested = pyqtSignal(str, str, str)  # (spec, version, clause)
    open_specs_requested = pyqtSignal(str, str)             # (spec, version)
    open_protocol_requested = pyqtSignal(str, str)          # (protocol, message)
    undock_requested = pyqtSignal()
    docked_changed = pyqtSignal(bool)

    def __init__(
        self,
        *,
        specs_db_path: Path,
        spec_search_db_path: Path,
        protocol_db_path: Path,
        parent=None,
    ):
        super().__init__(parent)
        self.specs_db_path = Path(specs_db_path)
        self.spec_search_db_path = Path(spec_search_db_path)
        self.protocol_db_path = Path(protocol_db_path)

        self.tool_registry = build_read_only_tool_registry(
            specs_db_path=self.specs_db_path,
            spec_search_db_path=self.spec_search_db_path,
            protocol_db_path=self.protocol_db_path,
        )
        self._worker: Optional[AgentResearchWorker] = None
        self._conversation: List[Dict[str, str]] = []
        self._activities: List[ResearchActivity] = []
        self._evidence: Dict[str, Evidence] = {}
        self._last_result: Optional[AgentResult] = None
        self._exchange_events: List[dict] = []
        self._is_docked: bool = False
        self._setup_ui()
        # Explicit idle state; independent of Ollama monitor state.
        self._set_running(False)
        self._set_status_text("Ready")
        self._refresh_ollama_status()

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

    @property
    def is_docked(self) -> bool:
        return self._is_docked

    def set_cockpit_visible_in_dock(self, visible: bool) -> None:
        """Show or collapse the lower Sources/Activity drawer while in compact dock mode."""
        visible = bool(visible)
        if hasattr(self, "dock_sources_toggle_btn"):
            self.dock_sources_toggle_btn.blockSignals(True)
            self.dock_sources_toggle_btn.setChecked(visible)
            self.dock_sources_toggle_btn.blockSignals(False)
        self.cockpit_tabs.setVisible(visible)
        if self._is_docked:
            self.main_splitter.setOrientation(Qt.Vertical)
            self.main_splitter.setSizes([380, 240] if visible else [600, 0])

    def set_docked_mode(self, docked: bool) -> None:
        """Toggle compact layout when Assistant body is hosted inside RightSidePanel."""
        docked = bool(docked)
        changed = docked != self._is_docked
        self._is_docked = docked

        self.ribbon_card.setVisible(not docked)
        self.dock_compact_strip.setVisible(docked)
        self.docked_placeholder_card.setVisible(docked)

        if docked:
            self._body_layout.setContentsMargins(S.XS, S.XS, S.XS, S.XS)
            self.main_splitter.setOrientation(Qt.Vertical)
            self.set_cockpit_visible_in_dock(False)
        else:
            self._body_layout.setContentsMargins(0, 0, 0, 0)
            self.cockpit_tabs.setVisible(True)
            self._toggle_layout(self.layout_toggle_btn.isChecked())

        if changed:
            self.docked_changed.emit(docked)

    def restore_body_into_tab(self) -> None:
        """Re-attach ``body_container`` into this tab and exit compact dock mode."""
        if self._host_layout.indexOf(self.body_container) < 0:
            self._host_layout.insertWidget(0, self.body_container, 1)
        self.body_container.setVisible(True)
        self.set_docked_mode(False)

    def _set_status_text(self, text: str) -> None:
        self.status_label.setText(text)
        if hasattr(self, "dock_status_label"):
            self.dock_status_label.setText(text)

    def _update_source_counts(self) -> None:
        count = len(self._evidence)
        label = f"📚 Sources ({count})"
        self.cockpit_tabs.setTabText(0, label)
        if hasattr(self, "dock_sources_toggle_btn"):
            self.dock_sources_toggle_btn.setText(label)

    def _toggle_layout(self, checked: bool) -> None:
        """Toggle between Side-by-Side (Horizontal) and Bottom Split (Vertical)."""
        if checked:
            self.main_splitter.setOrientation(Qt.Horizontal)
            self.main_splitter.setSizes([680, 420])
            self.layout_toggle_btn.setText("⬒ Bottom Split")
            self.layout_toggle_btn.setToolTip("Switch to stacked top/bottom layout")
        else:
            self.main_splitter.setOrientation(Qt.Vertical)
            self.main_splitter.setSizes([450, 280])
            self.layout_toggle_btn.setText("◫ Side Drawer")
            self.layout_toggle_btn.setToolTip("Switch to side-by-side drawer layout")

    def _on_prompt_changed(self) -> None:
        text = self.question_input.toPlainText()
        count = len(text)
        self.char_count_label.setText(f"{count} char" + ("s" if count != 1 else ""))
        self.clear_input_btn.setEnabled(count > 0)

    def _on_starter_chip_selected(self, prompt: str) -> None:
        self.question_input.setPlainText(prompt)
        self._on_send()

    def _on_anchor_clicked(self, url: QUrl) -> None:
        raw_url = url.toString() if hasattr(url, "toString") else str(url)
        if raw_url.startswith("evidence:"):
            eid = raw_url.split(":", 1)[1].strip()
            self._select_evidence_by_id(eid)
        elif raw_url.startswith("copy_msg:"):
            try:
                idx = int(raw_url.split(":", 1)[1].strip())
                if 0 <= idx < len(self._conversation):
                    text = self._conversation[idx].get("content", "")
                    QApplication.clipboard().setText(text)
                    self.status_label.setText("Answer copied to clipboard!")
                    QTimer.singleShot(
                        2000,
                        lambda: self.status_label.setText(
                            "Ready" if not (self._worker and self._worker.isRunning()) else "Researching…"
                        ),
                    )
            except Exception as exc:
                logger.error("Failed to copy message: %s", exc)

    def _select_evidence_by_id(self, eid: str) -> None:
        """Highlights the evidence row in sources tree, displays it in inspector, and switches to Sources tab."""
        if self._is_docked and not self.cockpit_tabs.isVisible():
            self.set_cockpit_visible_in_dock(True)
        self.cockpit_tabs.setCurrentIndex(0)  # Tab 0 is Sources
        for i in range(self.sources_tree.topLevelItemCount()):
            item = self.sources_tree.topLevelItem(i)
            if item.data(0, Qt.UserRole) == eid:
                self.sources_tree.setCurrentItem(item)
                self.sources_tree.scrollToItem(item)
                ev = self._evidence.get(eid)
                self.evidence_inspector.set_evidence(ev)
                break

    def _on_source_tree_item_clicked(self, item: QTreeWidgetItem, column: int = 0) -> None:
        evidence_id = str(item.data(0, Qt.UserRole) or "")
        evidence = self._evidence.get(evidence_id)
        self.evidence_inspector.set_evidence(evidence)

    def _refresh_ollama_status(self) -> None:
        try:
            from core.ai.ollama_client import load_ollama_config, is_ollama_port_open
            cfg = load_ollama_config()
            host = cfg.get("host", "http://127.0.0.1:11434")
            model = cfg.get("selected_model", "")
            is_open = is_ollama_port_open(host)
            if is_open and model:
                self.model_status_btn.setText(f"🟢 Ollama: {model}")
                self.model_status_btn.setStyleSheet(STATUS_PILL_ONLINE)
                self.model_status_btn.setToolTip(f"Host: {host}\nModel: {model}\nClick to configure Ollama settings.")
            elif is_open:
                self.model_status_btn.setText("🟡 Ollama: No Model Selected")
                self.model_status_btn.setStyleSheet(STATUS_PILL_NEUTRAL)
                self.model_status_btn.setToolTip(f"Host: {host}\nNo model selected.\nClick to configure.")
            else:
                self.model_status_btn.setText("🔴 Ollama: Offline")
                self.model_status_btn.setStyleSheet(STATUS_PILL_OFFLINE)
                self.model_status_btn.setToolTip(f"Host: {host}\nDaemon not responding.\nClick to configure or start service.")
        except Exception as exc:
            logger.debug("Failed to query Ollama status: %s", exc)
            self.model_status_btn.setText("⚪ Ollama: Unknown")
            self.model_status_btn.setStyleSheet(STATUS_PILL_NEUTRAL)
        if hasattr(self, "dock_model_status_btn"):
            self.dock_model_status_btn.setText(self.model_status_btn.text())
            self.dock_model_status_btn.setStyleSheet(self.model_status_btn.styleSheet())
            self.dock_model_status_btn.setToolTip(self.model_status_btn.toolTip())

    def _on_open_ollama_settings(self) -> None:
        if self.receivers(self.open_ollama_settings_requested) > 0:
            self.open_ollama_settings_requested.emit()
        else:
            try:
                from core.ai.ollama_dialog import OllamaConfigDialog
                dlg = OllamaConfigDialog(self)
                dlg.exec_()
            except Exception as exc:
                logger.error("Could not open Ollama settings dialog: %s", exc)
        QTimer.singleShot(1500, self._refresh_ollama_status)

    def _on_new_chat(self) -> None:
        if self._worker is not None and self._worker.isRunning():
            reply = QMessageBox.question(
                self,
                "Active Research in Progress",
                "A research query is currently running. Do you want to stop it and begin a new session?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if reply == QMessageBox.Yes:
                self._on_stop()
            else:
                return

        self._conversation.clear()
        self._activities.clear()
        self._evidence.clear()
        self._last_result = None
        self._exchange_events.clear()
        self.conversation_browser.clear()
        self.activity_widget.clear()
        self.sources_tree.clear()
        self.evidence_inspector.set_evidence(None)
        self.diagnostics_widget.clear()
        self._update_activity_count()
        self._update_source_counts()
        self.cockpit_tabs.setTabText(1, "⚡ Activity (0)")
        self._set_status_text("Ready")
        self.starter_chips.setVisible(True)
        self.question_input.clear()
        self.question_input.setFocus()

    def _on_copy_conversation(self) -> None:
        if not self._conversation:
            notify_info(self, "Copy Chat", "Conversation transcript is currently empty.")
            return
        lines = ["# 3GPP Research Assistant Transcript\n"]
        for msg in self._conversation:
            role = "### User" if msg.get("role") == "user" else "### Assistant"
            content = msg.get("content", "")
            lines.append(f"{role}\n\n{content}\n")
        if self._evidence:
            lines.append("## Referenced Sources & Evidence\n")
            for eid, ev in sorted(self._evidence.items()):
                label = self._source_label(ev)
                lines.append(f"- **[{eid}]** {label}")
        full_text = "\n".join(lines)
        QApplication.clipboard().setText(full_text)
        self.status_label.setText("Transcript copied!")
        notify_success(self, "Transcript Copied", "Conversation transcript copied to clipboard.")
        QTimer.singleShot(
            2500,
            lambda: self.status_label.setText(
                "Ready" if not (self._worker and self._worker.isRunning()) else "Researching…"
            ),
        )

    def _on_export_markdown(self) -> None:
        if not self._conversation:
            notify_info(self, "Export Transcript", "Conversation transcript is currently empty.")
            return
        filename, _ = QFileDialog.getSaveFileName(
            self,
            "Export Research Transcript",
            f"3gpp_assistant_session_{datetime.now().strftime('%Y%m%d_%H%M%S')}.md",
            "Markdown Files (*.md);;All Files (*)",
        )
        if not filename:
            return
        lines = [
            "# 3GPP Research Assistant Transcript\n",
            f"*Exported: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}*\n",
        ]
        for msg in self._conversation:
            role = "### User" if msg.get("role") == "user" else "### Assistant"
            content = msg.get("content", "")
            lines.append(f"{role}\n\n{content}\n")
        if self._evidence:
            lines.append("## Referenced Sources & Evidence\n")
            for eid, ev in sorted(self._evidence.items()):
                label = self._source_label(ev)
                lines.append(f"- **[{eid}]** {label}")
        try:
            Path(filename).write_text("\n".join(lines), encoding="utf-8")
            self.status_label.setText("Exported successfully!")
            notify_success(
                self,
                "Transcript Exported",
                f"Saved Markdown transcript to {Path(filename).name}",
            )
            QTimer.singleShot(
                2500,
                lambda: self.status_label.setText(
                    "Ready" if not (self._worker and self._worker.isRunning()) else "Researching…"
                ),
            )
        except Exception as exc:
            QMessageBox.critical(self, "Export Failed", f"Failed to save transcript: {exc}")

    def _on_send(self) -> None:
        if self._worker is not None and self._worker.isRunning():
            return
        question = self.question_input.toPlainText().strip()
        if not question:
            return

        self.question_input.clear()
        self.starter_chips.setVisible(False)
        self._append_conversation("user", question)
        self._activities.clear()
        self._evidence.clear()
        self._last_result = None
        self.activity_widget.clear()
        self.sources_tree.clear()
        self.evidence_inspector.set_evidence(None)
        self.diagnostics_widget.clear()
        self._exchange_events.clear()
        self._update_activity_count()
        self._update_source_counts()
        self.cockpit_tabs.setTabText(1, "⚡ Activity (0)")

        try:
            adapter = OllamaAgentAdapter()
        except Exception as exc:
            self._append_system_message(f"Could not initialize the local model: {exc}", error=True)
            return

        worker = AgentResearchWorker(
            question=question,
            tool_registry=self.tool_registry,
            adapter=adapter,
            conversation_context=self._conversation[:-1],
            controller_config=AgentControllerConfig(),
            parent=self,
        )
        self._worker = worker
        worker.research_started.connect(self._on_research_started)
        worker.activity_updated.connect(self._on_activity)
        worker.evidence_added.connect(self._on_evidence)
        worker.exchange_observed.connect(self._on_exchange)
        worker.result_ready.connect(self._on_result)
        worker.research_cancelled.connect(self._on_cancelled)
        worker.research_failed.connect(self._on_failed)
        worker.finished.connect(self._on_worker_finished)
        self._set_running(True)
        self.activity_widget.set_running(True)
        # Switch to Research Activity tab during research so user sees live execution
        self.cockpit_tabs.setCurrentIndex(1)
        worker.start()

    def _on_stop(self) -> None:
        worker = self._worker
        if worker is None or not worker.isRunning():
            return
        self._set_status_text("Cancelling…")
        self.stop_btn.setEnabled(False)
        worker.request_cancel()

    def _on_research_started(self) -> None:
        self._set_status_text("Researching…")

    def _on_activity(self, activity: ResearchActivity) -> None:
        self._activities.append(activity)
        self._update_activity_count()
        self.activity_widget.add_activity(activity)

    def _on_evidence(self, evidence: Evidence) -> None:
        self._evidence[evidence.id] = evidence
        self._add_source_row(evidence)
        self._update_source_counts()
        # Automatically display the first retrieved evidence in the inspector
        if len(self._evidence) == 1:
            self.evidence_inspector.set_evidence(evidence)

    def _on_exchange(self, event: object) -> None:
        if not isinstance(event, dict):
            return
        event = dict(event)
        event["timestamp"] = datetime.now().strftime("%H:%M:%S.%f")[:-3]
        self._exchange_events.append(event)
        self.diagnostics_widget.add_exchange_event(event)

    def _on_result(self, result: AgentResult) -> None:
        self._last_result = result
        # Final result is authoritative in case an evidence signal was missed.
        for evidence in result.evidence:
            if evidence.id not in self._evidence:
                self._evidence[evidence.id] = evidence
                self._add_source_row(evidence)

        self._update_source_counts()
        self.activity_widget.set_running(False)

        # Switch to Sources tab so user can explore evidence citations immediately
        if self._evidence:
            self.cockpit_tabs.setCurrentIndex(0)
            if self.sources_tree.topLevelItemCount() > 0:
                first_item = self.sources_tree.topLevelItem(0)
                self.sources_tree.setCurrentItem(first_item)
                self._on_source_tree_item_clicked(first_item)

        # Compute KPI badge and update Diagnostics metrics
        stamp = datetime.now().strftime("%H:%M")
        msg_idx = len(self._conversation)

        total_duration = 0.0
        prompt_tokens = 0
        eval_tokens = 0
        eval_duration = 0.0

        if result.trace and result.trace.events:
            for ev in result.trace.events:
                if ev.get("type") == "model_call_finished":
                    total_duration += float(ev.get("duration_seconds", 0.0))
                    metrics = ev.get("metrics") or {}
                    prompt_tokens += int(metrics.get("prompt_eval_count") or 0)
                    eval_tokens += int(metrics.get("eval_count") or 0)
                    eval_duration += float(metrics.get("eval_duration") or 0) / 1e9

        self.diagnostics_widget.update_metrics(
            total_seconds=total_duration,
            prompt_tokens=prompt_tokens,
            eval_tokens=eval_tokens,
            eval_seconds=eval_duration,
        )

        kpi_parts = []
        if total_duration > 0:
            kpi_parts.append(f"⏱️ {total_duration:.1f}s")
        if self._activities:
            kpi_parts.append(f"{len(self._activities)} steps")
        if self._evidence:
            kpi_parts.append(f"{len(self._evidence)} sources")
        kpi_str = " · ".join(kpi_parts)

        card_html = render_assistant_message_card(
            result.answer,
            timestamp=stamp,
            valid_evidence_ids=set(self._evidence),
            msg_index=msg_idx,
            kpi_badge=kpi_str,
        )
        self.conversation_browser.append(card_html)
        self._conversation.append({
            "role": "assistant",
            "content": result.answer,
            "timestamp": stamp,
            "kpi": kpi_str,
        })

        if result.warnings:
            for w in result.warnings:
                self.conversation_browser.append(render_system_message_card(w, error=False))

        self._set_status_text("Ready")

    def _on_cancelled(self, result) -> None:
        self.activity_widget.set_running(False)
        self._append_system_message("Research stopped.")
        self._set_status_text("Stopped")

    def _on_failed(self, message: str, result) -> None:
        self.activity_widget.set_running(False)
        logger.error("Assistant research failed: %s", message)
        self._append_system_message(message, error=True)
        self._set_status_text("Failed")

    def _on_worker_finished(self) -> None:
        sender = self.sender()
        if sender is not None and sender is not self._worker:
            return
        worker = self._worker
        if worker is None:
            return
        self._set_running(False)
        self.activity_widget.set_running(False)
        worker.deleteLater()
        self._worker = None
        if self.status_label.text() == "Cancelling…":
            self._set_status_text("Stopped")

    def shutdown(self, wait_ms: int = 500) -> None:
        """Request cooperative cancellation during application shutdown."""
        worker = self._worker
        if worker is None or not worker.isRunning():
            return
        worker.request_cancel()
        if not worker.wait(wait_ms):
            logger.warning(
                "Assistant worker is still waiting for an active operation during shutdown; "
                "it was not force-terminated."
            )

    def _set_running(self, running: bool) -> None:
        self.send_btn.setEnabled(not running)
        self.question_input.setEnabled(not running)
        self.stop_btn.setEnabled(running)
        if not running:
            self.question_input.setFocus()

    def _append_conversation(self, role: str, content: str) -> None:
        stamp = datetime.now().strftime("%H:%M")
        if role == "user":
            card_html = render_user_message_card(content, timestamp=stamp)
            self.conversation_browser.append(card_html)
            self._conversation.append({"role": role, "content": content, "timestamp": stamp})

    def _append_system_message(self, message: str, error: bool = False) -> None:
        card_html = render_system_message_card(message, error=error)
        self.conversation_browser.append(card_html)

    def _render_answer(self, answer: str) -> str:
        return render_markdown_to_html(answer, set(self._evidence))

    def _add_source_row(self, evidence: Evidence) -> None:
        source = self._source_label(evidence)
        detail_parts = []
        if evidence.version:
            detail_parts.append(f"v{evidence.version}")
        if evidence.clause:
            detail_parts.append(f"Clause {evidence.clause}")
        item = QTreeWidgetItem(
            [evidence.id, source, " — ".join(detail_parts)]
        )
        item.setData(0, Qt.UserRole, evidence.id)
        tooltip = evidence.title or source
        item.setToolTip(1, tooltip)
        self.sources_tree.addTopLevelItem(item)
        for column in range(3):
            self.sources_tree.resizeColumnToContents(column)

    @staticmethod
    def _source_label(evidence: Evidence) -> str:
        if evidence.source_type == EvidenceSourceType.SPECIFICATION_TEXT:
            prefix = f"TS {evidence.specification}" if evidence.specification else evidence.source_name
            return f"{prefix}: {evidence.title or 'indexed clause'}"
        if evidence.source_type == EvidenceSourceType.SPECIFICATION_CATALOGUE:
            prefix = f"TS {evidence.specification}" if evidence.specification else evidence.source_name
            return f"{prefix}: catalogue"
        if evidence.source_type == EvidenceSourceType.PROTOCOL_KNOWLEDGE:
            return evidence.title or "Structured protocol knowledge"
        return evidence.source_name

    def _update_activity_count(self) -> None:
        count = len(self._activities)
        self.cockpit_tabs.setTabText(1, f"⚡ Activity ({count})")
        self.activity_widget.update_count(count)
