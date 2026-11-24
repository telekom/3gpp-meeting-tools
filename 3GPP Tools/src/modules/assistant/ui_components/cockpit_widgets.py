"""Evidence inspector, activity timeline, and raw Ollama trace cockpit widgets."""

from __future__ import annotations

import html
import json
from typing import Optional

from PyQt5.QtCore import QTimer, pyqtSignal
from PyQt5.QtWidgets import (
    QApplication,
    QFrame,
    QHBoxLayout,
    QLabel,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QTextBrowser,
    QVBoxLayout,
    QWidget,
)

from core.ui.theme import C
from core.ui.ui_components import BUTTON_STYLE_TOOLBAR_SECONDARY
from modules.assistant.core.agent_models import Evidence, EvidenceSourceType, ResearchActivity
from modules.assistant.ui_components.styles import (
    BADGE_CATALOGUE,
    BADGE_PROTOCOL,
    BADGE_SPEC_TEXT,
    CARD_FRAME_STYLE,
    TOOL_ICONS,
)


class EvidenceInspectorWidget(QFrame):
    """Embedded inspector previewing retrieved clause/protocol evidence with deep links."""

    open_spec_search_requested = pyqtSignal(str, str, str)  # (spec, version, clause)
    open_specs_requested = pyqtSignal(str, str)             # (spec, version)
    open_protocol_requested = pyqtSignal(str, str)          # (protocol, message)

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setObjectName("assistantCardFrame")
        self.setStyleSheet(CARD_FRAME_STYLE)
        self._current_evidence: Optional[Evidence] = None

        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(6)

        hdr = QHBoxLayout()
        hdr.setContentsMargins(0, 0, 0, 0)
        hdr.setSpacing(6)

        self.badge = QLabel("Evidence")
        self.badge.setStyleSheet(BADGE_CATALOGUE)
        hdr.addWidget(self.badge)

        self.title_label = QLabel("No evidence selected")
        self.title_label.setStyleSheet(f"font-weight: bold; color: {C.TEXT_STRONG}; font-size: 12px;")
        hdr.addWidget(self.title_label, stretch=1)

        self.spec_search_btn = QPushButton("🔎 Spec Search")
        self.spec_search_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.spec_search_btn.setToolTip("Open and navigate to this clause in Spec Search tab")
        self.spec_search_btn.clicked.connect(self._on_open_spec_search)
        self.spec_search_btn.setVisible(False)
        hdr.addWidget(self.spec_search_btn)

        self.protocol_btn = QPushButton("🔬 In Protocols")
        self.protocol_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.protocol_btn.setToolTip("Inspect this message in Protocols tab")
        self.protocol_btn.clicked.connect(self._on_open_protocol)
        self.protocol_btn.setVisible(False)
        hdr.addWidget(self.protocol_btn)

        self.specs_tab_btn = QPushButton("📚 In Specs")
        self.specs_tab_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.specs_tab_btn.setToolTip("View specification metadata in Specifications tab")
        self.specs_tab_btn.clicked.connect(self._on_open_specs)
        self.specs_tab_btn.setVisible(False)
        hdr.addWidget(self.specs_tab_btn)

        self.copy_btn = QPushButton("📋 Copy Excerpt")
        self.copy_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.copy_btn.clicked.connect(self._on_copy_excerpt)
        self.copy_btn.setEnabled(False)
        hdr.addWidget(self.copy_btn)

        layout.addLayout(hdr)

        self.content_browser = QTextBrowser()
        self.content_browser.setStyleSheet(
            f"""
            QTextBrowser {{
                background-color: {C.BG_SURFACE};
                border: 1px solid {C.BORDER_SUBTLE};
                border-radius: 4px;
                font-family: Consolas, 'Courier New', monospace;
                font-size: 12px;
                color: {C.TEXT_STRONG};
                padding: 6px;
            }}
        """
        )
        self.content_browser.setPlaceholderText(
            "Select an evidence item from the list above or click an [E#] citation in the chat to view its retrieved content."
        )
        layout.addWidget(self.content_browser, stretch=1)

    def set_evidence(self, ev: Optional[Evidence]) -> None:
        self._current_evidence = ev
        if ev is None:
            self.badge.setText("Evidence")
            self.badge.setStyleSheet(BADGE_CATALOGUE)
            self.title_label.setText("No evidence selected")
            self.content_browser.clear()
            self.spec_search_btn.setVisible(False)
            self.protocol_btn.setVisible(False)
            self.specs_tab_btn.setVisible(False)
            self.copy_btn.setEnabled(False)
            return

        if ev.source_type == EvidenceSourceType.SPECIFICATION_TEXT:
            self.badge.setText("Spec Text")
            self.badge.setStyleSheet(BADGE_SPEC_TEXT)
        elif ev.source_type == EvidenceSourceType.PROTOCOL_KNOWLEDGE:
            self.badge.setText("Protocol")
            self.badge.setStyleSheet(BADGE_PROTOCOL)
        else:
            self.badge.setText("Catalogue")
            self.badge.setStyleSheet(BADGE_CATALOGUE)

        parts = [ev.id]
        if ev.specification:
            parts.append(f"TS {ev.specification}")
        if ev.version:
            parts.append(f"v{ev.version}")
        if ev.clause:
            parts.append(f"Clause {ev.clause}")
        elif ev.title:
            parts.append(ev.title)
        self.title_label.setText(" — ".join(parts))

        self.spec_search_btn.setVisible(bool(ev.specification and ev.clause))
        self.specs_tab_btn.setVisible(bool(ev.specification))
        self.protocol_btn.setVisible(bool(ev.source_type == EvidenceSourceType.PROTOCOL_KNOWLEDGE))
        self.copy_btn.setEnabled(True)

        content = ev.content
        if isinstance(content, dict):
            text = json.dumps(content, indent=2, ensure_ascii=False, default=str)
        else:
            text = str(content or "")
        self.content_browser.setPlainText(text)

    def _on_copy_excerpt(self) -> None:
        text = self.content_browser.toPlainText()
        if text:
            QApplication.clipboard().setText(text)
            orig = self.copy_btn.text()
            self.copy_btn.setText("✓ Copied!")
            QTimer.singleShot(1500, lambda: self.copy_btn.setText(orig))

    def _on_open_spec_search(self) -> None:
        if self._current_evidence and self._current_evidence.specification:
            spec = self._current_evidence.specification
            ver = self._current_evidence.version or ""
            clause = self._current_evidence.clause or ""
            self.open_spec_search_requested.emit(spec, ver, clause)

    def _on_open_specs(self) -> None:
        if self._current_evidence and self._current_evidence.specification:
            spec = self._current_evidence.specification
            ver = self._current_evidence.version or ""
            self.open_specs_requested.emit(spec, ver)

    def _on_open_protocol(self) -> None:
        if self._current_evidence and self._current_evidence.content:
            proto = ""
            msg = ""
            if isinstance(self._current_evidence.content, dict):
                proto = str(self._current_evidence.content.get("protocol") or "")
                msg = str(self._current_evidence.content.get("message") or "")
            self.open_protocol_requested.emit(proto, msg)


class ActivityTimelineWidget(QWidget):
    """Visual timeline displaying real-time research steps with semantic icons."""

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(6, 6, 6, 6)
        layout.setSpacing(6)

        header = QHBoxLayout()
        header.setContentsMargins(2, 0, 2, 0)
        self.step_label = QLabel("0 steps executed")
        self.step_label.setStyleSheet(f"font-weight: 600; color: {C.TEXT_SECONDARY}; font-size: 11px;")
        header.addWidget(self.step_label)
        header.addStretch()

        self.clear_btn = QPushButton("Clear")
        self.clear_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.clear_btn.clicked.connect(self.clear)
        header.addWidget(self.clear_btn)
        layout.addLayout(header)

        self.progress_bar = QProgressBar()
        self.progress_bar.setFixedHeight(4)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setRange(0, 0)
        self.progress_bar.setStyleSheet(
            f"""
            QProgressBar {{
                border: none;
                background-color: {C.BORDER_SUBTLE};
                border-radius: 2px;
            }}
            QProgressBar::chunk {{
                background-color: {C.ACCENT_PRIMARY};
                border-radius: 2px;
            }}
        """
        )
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        self.browser = QTextBrowser()
        self.browser.setStyleSheet(
            f"""
            QTextBrowser {{
                background-color: {C.BG_CARD};
                border: 1px solid {C.BORDER_SUBTLE};
                border-radius: 6px;
                font-family: "Segoe UI", Arial, sans-serif;
                font-size: 12px;
                color: {C.TEXT_STRONG};
                padding: 6px;
            }}
        """
        )
        layout.addWidget(self.browser, stretch=1)

    def set_running(self, running: bool) -> None:
        self.progress_bar.setVisible(running)

    def add_activity(self, activity: ResearchActivity) -> None:
        icon = TOOL_ICONS.get(activity.kind, "•")
        tool = activity.metadata.get("tool", "")
        if tool in TOOL_ICONS:
            icon = TOOL_ICONS[tool]

        msg = html.escape(activity.message)
        eids = activity.metadata.get("evidence_ids") or []
        eid_html = ""
        if eids:
            eid_chips = " ".join(
                f'<span style="background:{C.SKY_SOFT}; color:{C.ACCENT_SKY_DARK}; font-weight:bold; padding:1px 4px; border-radius:3px; font-size:10px;">{eid}</span>'
                for eid in eids
            )
            eid_html = f" &nbsp;{eid_chips}"

        item_html = f"""
        <div style="margin: 3px 0; padding: 4px 6px; border-left: 2px solid {C.BORDER_STRONG}; background-color: {C.BG_SURFACE};">
          <span style="font-size: 13px;">{icon}</span> <b>{msg}</b>{eid_html}
        </div>
        """
        self.browser.append(item_html)

    def update_count(self, count: int) -> None:
        self.step_label.setText(f"{count} step" + ("s" if count != 1 else "") + " executed")

    def clear(self) -> None:
        self.browser.clear()
        self.update_count(0)
        self.set_running(False)


class DiagnosticsTraceWidget(QWidget):
    """Detailed diagnostics panel for raw Ollama exchange trace with KPI metrics."""

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(6, 6, 6, 6)
        layout.setSpacing(6)

        kpi_card = QFrame()
        kpi_card.setStyleSheet(
            f"""
            QFrame {{
                background-color: {C.BG_SURFACE};
                border: 1px solid {C.BORDER_SUBTLE};
                border-radius: 6px;
                padding: 4px 8px;
            }}
        """
        )
        kpi_layout = QHBoxLayout(kpi_card)
        kpi_layout.setContentsMargins(4, 2, 4, 2)
        kpi_layout.setSpacing(12)

        self.kpi_latency = QLabel("⏱️ Latency: 0.0s")
        self.kpi_latency.setStyleSheet(f"color: {C.TEXT_SECONDARY}; font-size: 11px; font-weight: 600;")
        kpi_layout.addWidget(self.kpi_latency)

        self.kpi_prompt_tokens = QLabel("📥 Prompt: 0 tok")
        self.kpi_prompt_tokens.setStyleSheet(f"color: {C.TEXT_SECONDARY}; font-size: 11px;")
        kpi_layout.addWidget(self.kpi_prompt_tokens)

        self.kpi_eval_tokens = QLabel("📤 Eval: 0 tok")
        self.kpi_eval_tokens.setStyleSheet(f"color: {C.TEXT_SECONDARY}; font-size: 11px;")
        kpi_layout.addWidget(self.kpi_eval_tokens)

        self.kpi_speed = QLabel("⚡ Speed: 0 tok/s")
        self.kpi_speed.setStyleSheet(f"color: {C.EMERALD_DARK}; font-size: 11px; font-weight: 600;")
        kpi_layout.addWidget(self.kpi_speed)

        kpi_layout.addStretch()

        self.copy_btn = QPushButton("Copy Trace")
        self.copy_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.copy_btn.clicked.connect(self._on_copy_trace)
        kpi_layout.addWidget(self.copy_btn)

        self.clear_btn = QPushButton("Clear")
        self.clear_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.clear_btn.clicked.connect(self.clear)
        kpi_layout.addWidget(self.clear_btn)

        layout.addWidget(kpi_card)

        self.browser = QPlainTextEdit()
        self.browser.setReadOnly(True)
        self.browser.setLineWrapMode(QPlainTextEdit.NoWrap)
        self.browser.setStyleSheet(
            f"""
            QPlainTextEdit {{
                background-color: {C.TEXT_STRONG};
                color: {C.BORDER_SUBTLE};
                font-family: Consolas, 'Courier New', monospace;
                font-size: 11px;
                border: 1px solid {C.TEXT_BODY};
                border-radius: 6px;
                padding: 6px;
            }}
        """
        )
        self.browser.setPlaceholderText("Raw Ollama REST requests and responses will be recorded here.")
        layout.addWidget(self.browser, stretch=1)

    def add_exchange_event(self, event: dict) -> None:
        stamp = event.get("timestamp") or ""
        direction = str(event.get("direction") or "event")
        labels = {
            "request": "APP -> OLLAMA",
            "response_headers": "OLLAMA -> APP [HEADERS]",
            "response": "OLLAMA -> APP",
            "response_chunk": "OLLAMA -> APP [STREAM CHUNK]",
        }
        label = labels.get(direction, direction.upper())
        if direction == "request":
            body = {k: event.get(k) for k in ("method", "endpoint", "payload") if k in event}
        elif direction == "response_headers":
            body = {k: event.get(k) for k in ("status_code", "reason", "headers") if k in event}
        elif direction == "response_chunk":
            raw = event.get("raw", "")
            try:
                body = json.loads(raw)
            except Exception:
                body = raw
        else:
            body = event.get("payload", event)

        if isinstance(body, (dict, list)):
            rendered = json.dumps(body, indent=2, ensure_ascii=False, default=str)
        else:
            rendered = str(body)

        self.browser.appendPlainText(f"[{stamp}] {label}\n{rendered}\n")

    def update_metrics(self, total_seconds: float, prompt_tokens: int, eval_tokens: int, eval_seconds: float) -> None:
        self.kpi_latency.setText(f"⏱️ Latency: {total_seconds:.1f}s")
        self.kpi_prompt_tokens.setText(f"📥 Prompt: {prompt_tokens} tok")
        self.kpi_eval_tokens.setText(f"📤 Eval: {eval_tokens} tok")
        speed = (eval_tokens / eval_seconds) if eval_seconds > 0 else 0.0
        self.kpi_speed.setText(f"⚡ Speed: {speed:.1f} tok/s")

    def _on_copy_trace(self) -> None:
        text = self.browser.toPlainText()
        if text:
            QApplication.clipboard().setText(text)
            orig = self.copy_btn.text()
            self.copy_btn.setText("✓ Copied!")
            QTimer.singleShot(1500, lambda: self.copy_btn.setText(orig))

    def clear(self) -> None:
        self.browser.clear()
        self.kpi_latency.setText("⏱️ Latency: 0.0s")
        self.kpi_prompt_tokens.setText("📥 Prompt: 0 tok")
        self.kpi_eval_tokens.setText("📤 Eval: 0 tok")
        self.kpi_speed.setText("⚡ Speed: 0 tok/s")
