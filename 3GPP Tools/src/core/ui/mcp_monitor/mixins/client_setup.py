"""Tab 2 (AI Agent Client Setup) configuration cards and CLI launcher mixin."""

from __future__ import annotations

import os
from pathlib import Path
from typing import Dict, Optional

from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtWidgets import (
    QApplication,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QTabWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from core.ai.cli_detector import detect_cli_agents, format_active_model_badge
from core.ai.cli_launcher import launch_cli_in_terminal
from core.ai.mcp_activity import get_mcp_client_configurations
from core.ui.cli_context_dialog import CliContextEditorDialog
from core.ui.theme import C, T
from core.ui.ui_components import BUTTON_STYLE_TOOLBAR_SECONDARY
from core.utils.paths import get_project_root


class McpClientSetupMixin:
    """Builds installed CLI status strip, MCP client snippet tabs, and copy helpers."""

    def _setup_config_tab(self, parent: QWidget):
        layout = QVBoxLayout(parent)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(10)

        info_box = QLabel(
            "External AI agents communicate with 3GPP Tools via standard input/output (stdio) JSON-RPC. "
            "Copy the pre-configured snippets below into your AI client settings. All absolute Python "
            "and application paths are automatically resolved."
        )
        info_box.setWordWrap(True)
        info_box.setStyleSheet(f"background: {C.BG_CANVAS}; border: 1px solid {C.BORDER_STRONG}; border-radius: 5px; padding: 8px;")
        layout.addWidget(info_box)

        # Installed LLM CLIs live detection strip
        cli_strip = QFrame()
        cli_strip.setStyleSheet(
            f"QFrame {{ background: {C.BG_SURFACE}; border: 1px solid {C.BORDER_STRONG}; border-radius: 6px; padding: 4px; }}"
        )
        cli_strip_layout = QHBoxLayout(cli_strip)
        cli_strip_layout.setContentsMargins(10, 6, 10, 6)
        cli_strip_layout.setSpacing(10)
        cli_strip_title = QLabel("<b>Installed LLM CLIs:</b>")
        cli_strip_title.setStyleSheet(f"color: {C.TEXT_STRONG}; border: none;")
        cli_strip_layout.addWidget(cli_strip_title)

        self._cli_status_widgets: Dict[str, tuple] = {}
        for aid in ("codex", "claude", "gemini"):
            badge = QLabel("...")
            badge.setStyleSheet(f"font-size: 11px; font-weight: 600; color: {C.TEXT_MUTED}; border: none;")
            cli_strip_layout.addWidget(badge)

            launch_btn = QPushButton("🚀 Launch")
            launch_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
            launch_btn.setVisible(False)
            launch_btn.clicked.connect(lambda _=False, target_aid=aid: self._launch_cli_from_cockpit(target_aid))
            cli_strip_layout.addWidget(launch_btn)
            self._cli_status_widgets[aid] = (badge, launch_btn)

        cli_strip_layout.addStretch()
        self.cli_settings_btn = QPushButton("🧠 Models && Company Context...")
        self.cli_settings_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.cli_settings_btn.setToolTip(
            "Configure default Model & Reasoning Effort for Codex, Claude, and Gemini CLIs, "
            "and edit standing company/delegate context."
        )
        self.cli_settings_btn.clicked.connect(self._open_cli_models_and_context_dialog)
        cli_strip_layout.addWidget(self.cli_settings_btn)

        self.cli_rescan_btn = QPushButton("🔄 Re-scan")
        self.cli_rescan_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.cli_rescan_btn.setToolTip("Re-scan PATH and user directories for codex, claude, and gemini CLIs.")
        self.cli_rescan_btn.clicked.connect(lambda _=False: self._refresh_cli_agents_status(force_refresh=True))
        cli_strip_layout.addWidget(self.cli_rescan_btn)
        layout.addWidget(cli_strip)
        self._refresh_cli_agents_status(force_refresh=False)

        cfg_tabs = QTabWidget()

        # Generate configuration dictionary
        cfgs = get_mcp_client_configurations()
        self._client_cfgs = cfgs

        # 1. Claude Desktop Card
        claude_card = self._create_client_card(
            title="Claude Desktop",
            snippet=cfgs["claude_desktop"]["config_json"],
            instructions=cfgs["claude_desktop"]["instructions"],
            target_path=cfgs["claude_desktop"]["target_file"],
            is_json=True,
        )
        cfg_tabs.addTab(claude_card, "Claude Desktop")

        # 2. Codex CLI Card
        codex_card = self._create_client_card(
            title="OpenAI Codex CLI",
            snippet=cfgs["codex_cli"]["config_json"],
            instructions=cfgs["codex_cli"]["instructions"],
            target_path=cfgs["codex_cli"].get("target_file"),
            command_line=cfgs["codex_cli"]["cli_command"],
            is_json=False,
        )
        cfg_tabs.addTab(codex_card, "Codex CLI")

        # 3. Claude CLI Card
        claude_cli_card = self._create_client_card(
            title="Anthropic Claude CLI",
            snippet=cfgs["claude_cli"]["cli_command"],
            instructions=cfgs["claude_cli"]["instructions"],
            is_json=False,
        )
        cfg_tabs.addTab(claude_cli_card, "Claude CLI")

        # 3b. Google Gemini CLI Card
        gemini_cli_card = self._create_client_card(
            title="Google Gemini CLI",
            snippet=cfgs["gemini_cli"]["config_json"],
            instructions=cfgs["gemini_cli"]["instructions"],
            target_path=cfgs["gemini_cli"]["target_file"],
            command_line=cfgs["gemini_cli"]["cli_command"],
            is_json=True,
        )
        cfg_tabs.addTab(gemini_cli_card, "Gemini CLI")

        # 4. Cursor / Windsurf Card
        cursor_card = self._create_client_card(
            title="Cursor / Windsurf",
            snippet=cfgs["cursor"]["command_line"],
            instructions=cfgs["cursor"]["instructions"],
            is_json=False,
        )
        cfg_tabs.addTab(cursor_card, "Cursor / Windsurf")

        # 5. Windows Batch Launcher Card
        batch_card = self._create_client_card(
            title="Zero-Config Batch Launcher (Launch_MCP.cmd)",
            snippet=cfgs["claude_desktop"]["cmd_config_json"],
            instructions="Use this variant if you prefer running via the bundled Launch_MCP.cmd launcher script.",
            target_path=cfgs["paths"]["cmd_launcher"],
            is_json=True,
        )
        cfg_tabs.addTab(batch_card, "Batch Launcher (.cmd)")

        layout.addWidget(cfg_tabs, stretch=1)

    def _create_client_card(
        self,
        title: str,
        snippet: str,
        instructions: str,
        target_path: Optional[str] = None,
        command_line: Optional[str] = None,
        is_json: bool = True,
    ) -> QWidget:
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)

        inst_lbl = QLabel(instructions)
        inst_lbl.setWordWrap(True)
        inst_lbl.setStyleSheet(f"font-size: 12px; color: {C.TEXT_BODY}; font-weight: 500;")
        layout.addWidget(inst_lbl)

        if target_path:
            path_row = QHBoxLayout()
            p_lbl = QLabel(f"<b>Target Config File:</b> <code>{target_path}</code>")
            p_lbl.setTextInteractionFlags(Qt.TextSelectableByMouse)
            path_row.addWidget(p_lbl)
            path_row.addStretch()

            open_dir_btn = QPushButton("📂 Show in Explorer")
            open_dir_btn.setStyleSheet("padding: 2px 8px; font-size: 11px;")
            open_dir_btn.clicked.connect(lambda: self._open_in_explorer(target_path))
            path_row.addWidget(open_dir_btn)
            layout.addLayout(path_row)

        snippet_edit = QTextEdit()
        snippet_edit.setReadOnly(True)
        snippet_edit.setText(snippet)
        snippet_edit.setStyleSheet(
            f"background: {C.TEXT_PRIMARY}; color: {C.BORDER_SUBTLE}; border-radius: 4px; padding: 6px; "
            f"font-family: Consolas, 'Courier New', monospace; font-size: {T.SIZE_BODY}px;"
        )
        layout.addWidget(snippet_edit, stretch=1)

        btn_row = QHBoxLayout()
        copy_btn = QPushButton("📋 Copy Configuration Snippet")
        copy_btn.setStyleSheet(f"""
            QPushButton {{
                background: {C.ACCENT_ALT}; color: white; border: none; border-radius: 4px;
                padding: 6px 16px; font-weight: bold;
            }}
            QPushButton:hover {{ background: {C.INFO}; }}
        """)
        copy_btn.clicked.connect(lambda: self._copy_snippet_with_feedback(snippet, copy_btn))
        btn_row.addWidget(copy_btn)

        if command_line:
            copy_cmd_btn = QPushButton("📋 Copy CLI Command")
            copy_cmd_btn.setStyleSheet("padding: 6px 12px; font-weight: 600;")
            copy_cmd_btn.clicked.connect(lambda: self._copy_snippet_with_feedback(command_line, copy_cmd_btn))
            btn_row.addWidget(copy_cmd_btn)

        btn_row.addStretch()
        layout.addLayout(btn_row)
        return widget

    def _refresh_cli_agents_status(self, force_refresh: bool = False):
        statuses = detect_cli_agents(force_refresh=force_refresh)
        for aid, (badge, launch_btn) in getattr(self, "_cli_status_widgets", {}).items():
            st = statuses.get(aid)
            if not st:
                continue
            if st.installed:
                mcp_tag = " · MCP ✓" if st.mcp_configured else ""
                model_badge = format_active_model_badge(aid)
                badge.setText(f"🟢 {st.label} ({model_badge}){mcp_tag}")
                badge.setStyleSheet(f"font-size: 11px; font-weight: 600; color: {C.SUCCESS_DARK}; border: none;")
                badge.setToolTip(
                    f"Executable: {st.executable_path}\nModel: {model_badge}\n"
                    f"Source: {st.source}\nMCP configured: {st.mcp_configured}"
                )
                launch_btn.setVisible(True)
                launch_btn.setToolTip(f"Launch {st.label} [{model_badge}] in an interactive terminal.")
            else:
                badge.setText(f"⚪ {st.label}: Not Found")
                badge.setStyleSheet(f"font-size: 11px; font-weight: 500; color: {C.TEXT_MUTED}; border: none;")
                badge.setToolTip(f"Not detected.\nInstall: {st.install_hint}")
                launch_btn.setVisible(False)

    def _open_cli_models_and_context_dialog(self):
        statuses = detect_cli_agents()
        first_installed = next((aid for aid, st in statuses.items() if st.installed), "codex")
        dlg = CliContextEditorDialog(
            agent_id=first_installed,
            settings_only=True,
            parent=self,
        )
        dlg.exec_()
        self._refresh_cli_agents_status(force_refresh=False)

    def _launch_cli_from_cockpit(self, agent_id: str):
        proj = get_project_root()
        target_dir = proj.parent if proj.name == "src" else proj
        launch_cli_in_terminal(agent_id, target_dir, bootstrap_prompt="", parent=self)

    def _open_in_explorer(self, target_path: str):
        try:
            p = Path(target_path)
            target_dir = p.parent if p.is_file() or not p.exists() else p
            if not target_dir.exists():
                target_dir.mkdir(parents=True, exist_ok=True)
            os.startfile(str(target_dir))
        except Exception as exc:
            QMessageBox.warning(self, "Explorer Error", f"Could not open directory:\n{exc}")

    def _copy_snippet_with_feedback(self, text: str, btn: QPushButton):
        QApplication.clipboard().setText(text)
        orig_text = btn.text()
        btn.setText("✓ Copied to Clipboard!")
        btn.setStyleSheet(f"background: {C.EMERALD_PRIMARY}; color: white; border: none; border-radius: 4px; padding: 6px 16px; font-weight: bold;")
        QTimer.singleShot(1500, lambda: self._reset_copy_btn(btn, orig_text))

    def _reset_copy_btn(self, btn: QPushButton, orig_text: str):
        btn.setText(orig_text)
        btn.setStyleSheet(f"background: {C.ACCENT_ALT}; color: white; border: none; border-radius: 4px; padding: 6px 16px; font-weight: bold;")
