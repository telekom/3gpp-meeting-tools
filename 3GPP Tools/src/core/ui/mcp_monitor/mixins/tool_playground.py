"""Tab 3 (Interactive Tool Playground) tool/action selector and executor mixin."""

from __future__ import annotations

import json
import time

from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtWidgets import (
    QApplication,
    QComboBox,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QSplitter,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from core.ui.mcp_monitor.presets import DEFAULT_PLAYGROUND_ARGUMENTS
from core.ui.theme import T
from modules.assistant.core.tools.external_agent_tools import (
    PROFILE_TOOLS,
    build_external_agent_tool_registry,
)
from core.ui.theme import C


class McpToolPlaygroundMixin:
    """Lets delegates test-drive domain MCP tools and inspect live JSON telemetry."""

    def _setup_playground_tab(self, parent: QWidget):
        layout = QVBoxLayout(parent)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)

        intro = QLabel(
            "Test-drive any of the 12 domain tools and 58 actions directly within the desktop application. "
            "Select a tool, modify arguments as needed, and hit 'Execute Test Call' to inspect live output."
        )
        intro.setStyleSheet(f"color: {C.TEXT_SECONDARY}; font-size: 12px;")
        layout.addWidget(intro)

        # Tool & Action Selector
        sel_row = QHBoxLayout()
        sel_row.addWidget(QLabel("Tool:"))
        self.play_tool_combo = QComboBox()
        for tname in PROFILE_TOOLS.get("all", []):
            self.play_tool_combo.addItem(tname, tname)
        self.play_tool_combo.currentIndexChanged.connect(self._on_play_tool_changed)
        sel_row.addWidget(self.play_tool_combo, stretch=1)

        sel_row.addWidget(QLabel("Action:"))
        self.play_action_combo = QComboBox()
        self.play_action_combo.currentIndexChanged.connect(self._on_play_action_changed)
        sel_row.addWidget(self.play_action_combo, stretch=1)

        self.play_exec_btn = QPushButton("▶️ Execute Test Call")
        self.play_exec_btn.setStyleSheet(f"""
            QPushButton {{
                background: {C.EMERALD_PRIMARY}; color: white; border: none; border-radius: 4px;
                padding: 6px 18px; font-weight: bold; font-size: 12px;
            }}
            QPushButton:hover {{ background: {C.EMERALD_DARK}; }}
        """)
        self.play_exec_btn.clicked.connect(self._execute_playground_call)
        sel_row.addWidget(self.play_exec_btn)

        layout.addLayout(sel_row)

        # Splitter: Input Arguments JSON (Left) and Output JSON (Right)
        play_splitter = QSplitter(Qt.Horizontal)

        in_widget = QWidget()
        in_layout = QVBoxLayout(in_widget)
        in_layout.setContentsMargins(0, 0, 0, 0)
        in_layout.addWidget(QLabel("Arguments JSON:"))
        self.play_args_edit = QTextEdit()
        self.play_args_edit.setStyleSheet(f"font-family: Consolas, 'Courier New', monospace; font-size: {T.SIZE_BODY}px;")
        self.play_args_edit.setText("{\n  \"action\": \"search_text\"\n}")
        in_layout.addWidget(self.play_args_edit)
        play_splitter.addWidget(in_widget)

        out_widget = QWidget()
        out_layout = QVBoxLayout(out_widget)
        out_layout.setContentsMargins(0, 0, 0, 0)
        self.play_out_title = QLabel("Execution Output & Telemetry:")
        out_layout.addWidget(self.play_out_title)
        self.play_output_edit = QTextEdit()
        self.play_output_edit.setReadOnly(True)
        self.play_output_edit.setStyleSheet(f"font-family: Consolas, 'Courier New', monospace; font-size: {T.SIZE_BODY}px;")
        self.play_output_edit.setPlaceholderText("Execution result will appear here...")
        out_layout.addWidget(self.play_output_edit)
        play_splitter.addWidget(out_widget)

        play_splitter.setSizes([380, 560])
        layout.addWidget(play_splitter, stretch=1)

        # Trigger initial tool populating
        self._on_play_tool_changed(0)

    def _ensure_registry(self):
        if self._tool_registry is None:
            self._tool_registry = build_external_agent_tool_registry(profile="all", read_only=False)
        return self._tool_registry

    def _on_play_tool_changed(self, _index: int):
        tool_name = self.play_tool_combo.currentText()
        registry = self._ensure_registry()
        tool = registry.get(tool_name)
        self.play_action_combo.blockSignals(True)
        self.play_action_combo.clear()
        if tool:
            for arg in tool.arguments:
                if arg.name == "action" and arg.enum:
                    for act in arg.enum:
                        self.play_action_combo.addItem(act, act)
                    break
        self.play_action_combo.blockSignals(False)
        self._on_play_action_changed(0)

    def _on_play_action_changed(self, _index: int):
        tool_name = self.play_tool_combo.currentText()
        action = self.play_action_combo.currentText()
        preset = DEFAULT_PLAYGROUND_ARGUMENTS.get(tool_name, {}).get(action, {})
        full_args = {"action": action, **preset}
        self.play_args_edit.setText(json.dumps(full_args, indent=2))

    def _execute_playground_call(self):
        tool_name = self.play_tool_combo.currentText()
        raw_text = self.play_args_edit.toPlainText().strip()
        try:
            args = json.loads(raw_text) if raw_text else {}
        except Exception as exc:
            QMessageBox.critical(self, "Invalid JSON", f"Arguments must be valid JSON:\n{exc}")
            return

        self.play_exec_btn.setEnabled(False)
        self.play_out_title.setText("⏳ Executing tool...")
        QApplication.processEvents()

        t0 = time.perf_counter()
        registry = self._ensure_registry()
        res = registry.execute(tool_name, args)
        elapsed_ms = (time.perf_counter() - t0) * 1000.0

        self.play_exec_btn.setEnabled(True)
        if res.ok:
            is_err = isinstance(res.data, dict) and res.data.get("ok") is False
            status_emoji = "🟠 Logic Note" if is_err else "🟢 Success"
            self.play_out_title.setText(f"{status_emoji} | Finished in {elapsed_ms:.1f} ms")
            self.play_output_edit.setText(json.dumps(res.data, ensure_ascii=False, indent=2, default=str))
        else:
            err_msg = getattr(res, "error", None) or getattr(res, "message", "Tool execution failed")
            self.play_out_title.setText(f"🔴 Failed in {elapsed_ms:.1f} ms")
            self.play_output_edit.setText(f"Error:\n{err_msg}")

        # Trigger activity refresh so this test call appears in the table
        QTimer.singleShot(200, self.refresh_all)
