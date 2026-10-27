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

logger = logging.getLogger(__name__)


class AssistantTab(QWidget):
    """Conversation + objective research activity + application-owned sources."""

    open_ollama_settings_requested = pyqtSignal()
    open_spec_search_requested = pyqtSignal(str, str, str)  # (spec, version, clause)
    open_specs_requested = pyqtSignal(str, str)             # (spec, version)
    open_protocol_requested = pyqtSignal(str, str)          # (protocol, message)

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
        self._setup_ui()
        # Explicit idle state; independent of Ollama monitor state.
        self._set_running(False)
        self.status_label.setText("Ready")
        self._refresh_ollama_status()

    def _setup_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 8, 10, 8)
        layout.setSpacing(6)

        # ---------------------------------------------------------------------
        # Top Header Ribbon (Model Status, Session Actions, Layout Switcher)
        # ---------------------------------------------------------------------
        header_card = QFrame()
        header_card.setObjectName("assistantCardHeader")
        header_card.setStyleSheet("""
            QFrame#assistantCardHeader {
                background-color: #F8FAFC;
                border: 1px solid #E2E8F0;
                border-radius: 6px;
                padding: 4px;
            }
        """)
        header_layout = QHBoxLayout(header_card)
        header_layout.setContentsMargins(8, 4, 8, 4)
        header_layout.setSpacing(8)

        title = QLabel("🤖 3GPP Assistant")
        title.setStyleSheet("font-weight: bold; font-size: 13px; color: #1E293B;")
        header_layout.addWidget(title)

        # Model status pill
        self.model_status_btn = QPushButton("⚪ Ollama: Checking...")
        self.model_status_btn.setStyleSheet(STATUS_PILL_NEUTRAL)
        self.model_status_btn.setCursor(Qt.PointingHandCursor)
        self.model_status_btn.setToolTip("Click to configure Ollama settings.")
        self.model_status_btn.clicked.connect(self._on_open_ollama_settings)
        header_layout.addWidget(self.model_status_btn)

        self.status_label = QLabel("Ready")
        self.status_label.setStyleSheet("color: #64748B; font-weight: 500; font-size: 12px;")
        header_layout.addWidget(self.status_label)

        header_layout.addStretch()

        # Session Action Buttons
        self.new_chat_btn = QPushButton("➕ New Chat")
        self.new_chat_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.new_chat_btn.setToolTip("Clear the conversation and begin a new research session.")
        self.new_chat_btn.clicked.connect(self._on_new_chat)
        header_layout.addWidget(self.new_chat_btn)

        self.copy_chat_btn = QPushButton("📋 Copy Chat")
        self.copy_chat_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.copy_chat_btn.setToolTip("Copy the full conversation transcript and citations to the clipboard.")
        self.copy_chat_btn.clicked.connect(self._on_copy_conversation)
        header_layout.addWidget(self.copy_chat_btn)

        self.export_chat_btn = QPushButton("💾 Export...")
        self.export_chat_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.export_chat_btn.setToolTip("Export the research conversation to a Markdown (.md) file.")
        self.export_chat_btn.clicked.connect(self._on_export_markdown)
        header_layout.addWidget(self.export_chat_btn)

        # Layout Switcher Button
        self.layout_toggle_btn = QPushButton("⬒ Bottom Split")
        self.layout_toggle_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.layout_toggle_btn.setCheckable(True)
        self.layout_toggle_btn.setChecked(True)  # True = Side-by-Side (Horizontal)
        self.layout_toggle_btn.setToolTip("Toggle between Side-by-Side and Bottom Split layouts.")
        self.layout_toggle_btn.clicked.connect(self._toggle_layout)
        header_layout.addWidget(self.layout_toggle_btn)

        layout.addWidget(header_card)

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
        self.conversation_browser.setStyleSheet("""
            QTextBrowser {
                background-color: #FFFFFF;
                border: none;
                font-family: "Segoe UI", Arial, sans-serif;
                font-size: 13px;
                color: #1E293B;
            }
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
        self.char_count_label.setStyleSheet("color: #94A3B8; font-size: 11px;")
        input_action_bar.addWidget(self.char_count_label)

        self.clear_input_btn = QPushButton("✕ Clear")
        self.clear_input_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.clear_input_btn.setEnabled(False)
        self.clear_input_btn.clicked.connect(self.question_input.clear)
        input_action_bar.addWidget(self.clear_input_btn)

        input_action_bar.addStretch()

        self.send_btn = QPushButton("Send ↵")
        self.send_btn.setObjectName("primaryBtn")
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
        self.cockpit_tabs.setStyleSheet("""
            QTabWidget::pane {
                border: 1px solid #E2E8F0;
                border-radius: 6px;
                background-color: #FFFFFF;
                top: -1px;
            }
            QTabBar::tab {
                background-color: #F8FAFC;
                border: 1px solid #E2E8F0;
                padding: 6px 12px;
                margin-right: 2px;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
                font-weight: 600;
                font-size: 12px;
                color: #64748B;
            }
            QTabBar::tab:selected {
                background-color: #FFFFFF;
                color: #1E5C99;
                border-bottom: 1px solid #FFFFFF;
            }
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
        layout.addWidget(self.main_splitter, stretch=1)

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
        self.cockpit_tabs.setTabText(0, "📚 Sources (0)")
        self.cockpit_tabs.setTabText(1, "⚡ Activity (0)")
        self.status_label.setText("Ready")
        self.starter_chips.setVisible(True)
        self.question_input.clear()
        self.question_input.setFocus()

    def _on_copy_conversation(self) -> None:
        if not self._conversation:
            QMessageBox.information(self, "Copy Chat", "Conversation transcript is currently empty.")
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
        QTimer.singleShot(
            2500,
            lambda: self.status_label.setText(
                "Ready" if not (self._worker and self._worker.isRunning()) else "Researching…"
            ),
        )

    def _on_export_markdown(self) -> None:
        if not self._conversation:
            QMessageBox.information(self, "Export Transcript", "Conversation transcript is currently empty.")
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
        self.cockpit_tabs.setTabText(0, "📚 Sources (0)")
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
        self.status_label.setText("Cancelling…")
        self.stop_btn.setEnabled(False)
        worker.request_cancel()

    def _on_research_started(self) -> None:
        self.status_label.setText("Researching…")

    def _on_activity(self, activity: ResearchActivity) -> None:
        self._activities.append(activity)
        self._update_activity_count()
        self.activity_widget.add_activity(activity)

    def _on_evidence(self, evidence: Evidence) -> None:
        self._evidence[evidence.id] = evidence
        self._add_source_row(evidence)
        self.cockpit_tabs.setTabText(0, f"📚 Sources ({len(self._evidence)})")
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

        self.cockpit_tabs.setTabText(0, f"📚 Sources ({len(self._evidence)})")
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

        self.status_label.setText("Ready")

    def _on_cancelled(self, result) -> None:
        self.activity_widget.set_running(False)
        self._append_system_message("Research stopped.")
        self.status_label.setText("Stopped")

    def _on_failed(self, message: str, result) -> None:
        self.activity_widget.set_running(False)
        logger.error("Assistant research failed: %s", message)
        self._append_system_message(message, error=True)
        self.status_label.setText("Failed")

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
            self.status_label.setText("Stopped")

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
