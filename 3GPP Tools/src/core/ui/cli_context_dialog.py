# --- File: src/core/ui/cli_context_dialog.py ---
"""
Interactive CLI Context, Company Guidance & Model/Effort Editor Dialog.

Allows delegates to:
- Configure per-CLI Model & Reasoning Effort (populated from `~/.codex/models_cache.json`
  plus curated/custom models for Claude CLI and Gemini CLI).
- Edit persistent per-TDoc / workspace `_USER_NOTES.md` files (with automatic aggregation
  preview for multi-TDoc selections).
- Edit global standing company/delegate guidance (`src/config/cli_user_context.md`)
  without modifying any auto-generated `_CLI_CONTEXT.md` structure.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

from PyQt5.QtWidgets import (
    QComboBox,
    QDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QPlainTextEdit,
    QPushButton,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from core.ai.cli_detector import (
    CLI_AGENT_SPECS,
    CliModelOption,
    discover_cli_models,
    get_cli_agent_model_config,
    set_cli_agent_model_config,
)
from core.ai.cli_launcher import (
    USER_NOTES_FILENAME,
    load_standing_cli_context,
    load_workspace_user_notes,
    save_standing_cli_context,
    save_workspace_user_notes,
)
from core.ui.theme import C
from core.ui.ui_components import BUTTON_STYLE_TOOLBAR_SECONDARY


DEFAULT_TDOC_TASK_OPTIONS: Tuple[Tuple[str, str], ...] = (
    ("interactive", "💬 Interactive Session (Read & Stand By)"),
    ("summarize", "⚡ Summarize & Check Spec Impact"),
    ("compare", "⚖️ Compare Selected / Revisions"),
    ("custom", "✏️ Custom Question"),
)


@dataclass
class CliContextDialogResult:
    """Structured result returned after `CliContextEditorDialog` saves or launches."""

    launch_requested: bool
    agent_id: str
    task_mode: str
    custom_question: str
    model: str
    reasoning_effort: str
    workspace_notes: str
    standing_context: str


class CliContextEditorDialog(QDialog):
    """
    Modal dialog for editing persistent CLI model/reasoning settings,
    workspace/TDoc `_USER_NOTES.md`, and standing company context.
    """

    def __init__(
        self,
        *,
        agent_id: str = "codex",
        scope_label: str = "",
        target_dir: Optional[Path] = None,
        task_options: Optional[Sequence[Tuple[str, str]]] = None,
        initial_task_mode: str = "interactive",
        initial_custom_question: str = "",
        initial_workspace_notes: Optional[str] = None,
        per_tdoc_notes: Optional[Dict[str, str]] = None,
        settings_only: bool = False,
        home_dir: Optional[Path] = None,
        parent: Optional[QWidget] = None,
    ) -> None:
        super().__init__(parent)
        self._agent_id = (agent_id or "codex").strip().lower()
        if self._agent_id not in CLI_AGENT_SPECS:
            self._agent_id = "codex"
        self._scope_label = scope_label.strip()
        self._target_dir = Path(target_dir) if target_dir else None
        self._task_options = list(task_options or DEFAULT_TDOC_TASK_OPTIONS)
        self._per_tdoc_notes = dict(per_tdoc_notes or {})
        self._settings_only = settings_only
        self._home_dir = home_dir
        self._launch_requested = False
        self._models_cache: List[CliModelOption] = []

        self.setWindowTitle(
            "CLI Models & Standing Company Context"
            if self._settings_only
            else "Edit CLI Context, Notes & Model"
        )
        self.resize(760, 620)
        self.setModal(True)

        self._setup_ui(
            initial_task_mode=initial_task_mode,
            initial_custom_question=initial_custom_question,
            initial_workspace_notes=initial_workspace_notes,
        )

    def _setup_ui(
        self,
        *,
        initial_task_mode: str,
        initial_custom_question: str,
        initial_workspace_notes: Optional[str],
    ) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(10)

        # Header card: Agent, Model, Reasoning Effort
        header_card = QFrame()
        header_card.setStyleSheet(
            f"QFrame {{ background: {C.BG_SURFACE}; border: 1px solid {C.BORDER_STRONG}; border-radius: 6px; }}"
        )
        header_layout = QVBoxLayout(header_card)
        header_layout.setContentsMargins(12, 10, 12, 10)
        header_layout.setSpacing(8)

        title_row = QHBoxLayout()
        title_txt = (
            "⚙️ Default CLI Models & Standing Company Context"
            if self._settings_only
            else f"📝 CLI Context & Priorities — {self._scope_label or 'Workspace'}"
        )
        title_lbl = QLabel(f"<b>{title_txt}</b>")
        title_lbl.setStyleSheet(f"color: {C.TEXT_PRIMARY}; font-size: 13px; border: none;")
        title_row.addWidget(title_lbl)
        title_row.addStretch()
        header_layout.addLayout(title_row)

        model_row = QHBoxLayout()
        model_row.setSpacing(8)

        lbl_agent = QLabel("CLI Agent:")
        lbl_agent.setStyleSheet(f"color: {C.TEXT_STRONG}; font-weight: 600; border: none;")
        model_row.addWidget(lbl_agent)

        self.agent_combo = QComboBox()
        for aid, spec in CLI_AGENT_SPECS.items():
            self.agent_combo.addItem(f"{spec.icon} {spec.label}", aid)
        idx_a = self.agent_combo.findData(self._agent_id)
        if idx_a >= 0:
            self.agent_combo.setCurrentIndex(idx_a)
        self.agent_combo.setEnabled(self._settings_only)
        self.agent_combo.currentIndexChanged.connect(self._on_agent_changed)
        model_row.addWidget(self.agent_combo)

        lbl_model = QLabel("Model:")
        lbl_model.setStyleSheet(f"color: {C.TEXT_STRONG}; font-weight: 600; border: none;")
        model_row.addWidget(lbl_model)

        self.model_combo = QComboBox()
        self.model_combo.setEditable(True)
        self.model_combo.setInsertPolicy(QComboBox.NoInsert)
        self.model_combo.setMinimumWidth(220)
        self.model_combo.currentIndexChanged.connect(self._on_model_changed)
        model_row.addWidget(self.model_combo, stretch=1)

        lbl_effort = QLabel("Reasoning Effort:")
        lbl_effort.setStyleSheet(f"color: {C.TEXT_STRONG}; font-weight: 600; border: none;")
        model_row.addWidget(lbl_effort)

        self.effort_combo = QComboBox()
        self.effort_combo.setMinimumWidth(140)
        model_row.addWidget(self.effort_combo)

        header_layout.addLayout(model_row)

        if not self._settings_only:
            task_row = QHBoxLayout()
            task_row.setSpacing(8)
            lbl_task = QLabel("Task Mode:")
            lbl_task.setStyleSheet(f"color: {C.TEXT_STRONG}; font-weight: 600; border: none;")
            task_row.addWidget(lbl_task)

            self.task_combo = QComboBox()
            for mode_id, mode_label in self._task_options:
                self.task_combo.addItem(mode_label, mode_id)
            idx_t = self.task_combo.findData(initial_task_mode)
            if idx_t >= 0:
                self.task_combo.setCurrentIndex(idx_t)
            self.task_combo.currentIndexChanged.connect(self._on_task_mode_changed)
            task_row.addWidget(self.task_combo, stretch=1)
            header_layout.addLayout(task_row)

            self.custom_prompt_edit = QPlainTextEdit()
            self.custom_prompt_edit.setPlaceholderText("Optional: enter your specific question or analysis instruction for this launch...")
            self.custom_prompt_edit.setFixedHeight(64)
            self.custom_prompt_edit.setPlainText(initial_custom_question or "")
            self.custom_prompt_edit.setVisible(initial_task_mode == "custom")
            header_layout.addWidget(self.custom_prompt_edit)
        else:
            self.task_combo = None
            self.custom_prompt_edit = None

        layout.addWidget(header_card)

        # Tabs: Workspace / TDoc Notes + Standing Company Context
        self.tabs = QTabWidget()

        if not self._settings_only:
            notes_tab = QWidget()
            notes_layout = QVBoxLayout(notes_tab)
            notes_layout.setContentsMargins(10, 10, 10, 10)
            notes_layout.setSpacing(6)

            loc_hint = (
                f"Stored in <code>{self._target_dir / USER_NOTES_FILENAME}</code>"
                if self._target_dir
                else f"Saved into the batch <code>{USER_NOTES_FILENAME}</code> file for this run"
            )
            lbl_notes_hint = QLabel(
                f"Persistent notes for this TDoc / workspace ({loc_hint}). "
                "Never overwritten when <code>_CLI_CONTEXT.md</code> regenerates."
            )
            lbl_notes_hint.setWordWrap(True)
            lbl_notes_hint.setStyleSheet(f"color: {C.TEXT_SECONDARY}; font-size: 11px;")
            notes_layout.addWidget(lbl_notes_hint)

            self.workspace_notes_edit = QPlainTextEdit()
            self.workspace_notes_edit.setPlaceholderText(
                "Example: Focus on section 5.2 and UE policy precedence. "
                "I definitely see points XYZ as very critical for our company deployment."
            )
            loaded_ws_notes = (
                initial_workspace_notes
                if initial_workspace_notes is not None
                else (load_workspace_user_notes(self._target_dir) if self._target_dir else "")
            )
            self.workspace_notes_edit.setPlainText(loaded_ws_notes)
            notes_layout.addWidget(self.workspace_notes_edit, stretch=2)

            if self._per_tdoc_notes:
                lbl_agg = QLabel(
                    f"<b>📌 Included Automatically from Selected TDocs ({len(self._per_tdoc_notes)}):</b>"
                )
                lbl_agg.setStyleSheet(f"color: {C.TEXT_STRONG}; font-size: 11px;")
                notes_layout.addWidget(lbl_agg)

                self.per_tdoc_preview = QPlainTextEdit()
                self.per_tdoc_preview.setReadOnly(True)
                self.per_tdoc_preview.setMaximumHeight(120)
                preview_blocks = [f"### {tid}\n{txt}" for tid, txt in self._per_tdoc_notes.items()]
                self.per_tdoc_preview.setPlainText("\n\n".join(preview_blocks))
                notes_layout.addWidget(self.per_tdoc_preview, stretch=1)

            self.tabs.addTab(notes_tab, "📌 TDoc / Workspace Notes")
        else:
            self.workspace_notes_edit = None

        standing_tab = QWidget()
        standing_layout = QVBoxLayout(standing_tab)
        standing_layout.setContentsMargins(10, 10, 10, 10)
        standing_layout.setSpacing(6)

        lbl_standing = QLabel(
            "Standing guidance applied to <b>every</b> CLI session across single TDocs, "
            "multi-TDoc batches, and Cross-WG Topic Dossiers (stored in "
            "<code>src/config/cli_user_context.md</code>)."
        )
        lbl_standing.setWordWrap(True)
        lbl_standing.setStyleSheet(f"color: {C.TEXT_SECONDARY}; font-size: 11px;")
        standing_layout.addWidget(lbl_standing)

        self.standing_edit = QPlainTextEdit()
        self.standing_edit.setPlaceholderText(
            "Example: We represent an MNO (Deutsche Telekom). Always highlight impacts on "
            "roaming, core network signaling load, backward compatibility with Rel-18 UEs, "
            "and proprietary/multi-vendor interoperability risks."
        )
        self.standing_edit.setPlainText(load_standing_cli_context())
        standing_layout.addWidget(self.standing_edit, stretch=1)
        self.tabs.addTab(standing_tab, "🏢 Standing Company Context (All Sessions)")

        layout.addWidget(self.tabs, stretch=1)

        # Footer action buttons
        btn_row = QHBoxLayout()
        btn_row.addStretch()

        self.cancel_btn = QPushButton("Cancel")
        self.cancel_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.cancel_btn.clicked.connect(self.reject)
        btn_row.addWidget(self.cancel_btn)

        self.save_only_btn = QPushButton(
            "💾 Save Settings" if self._settings_only else "💾 Save Notes Only"
        )
        self.save_only_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.save_only_btn.clicked.connect(self._on_save_only)
        btn_row.addWidget(self.save_only_btn)

        if not self._settings_only:
            self.save_launch_btn = QPushButton("🚀 Save && Launch CLI")
            self.save_launch_btn.setStyleSheet(
                f"QPushButton {{ background: {C.ACCENT_PRIMARY}; color: {C.TEXT_ON_ACCENT}; "
                f"border: none; border-radius: 4px; padding: 6px 14px; font-weight: bold; }}"
                f"QPushButton:hover {{ background: {C.ACCENT_HOVER}; }}"
            )
            self.save_launch_btn.clicked.connect(self._on_save_and_launch)
            btn_row.addWidget(self.save_launch_btn)
        else:
            self.save_launch_btn = None

        layout.addLayout(btn_row)

        self._populate_models_for_agent(self._agent_id)

    def _on_agent_changed(self) -> None:
        # Persist previous agent's selected model/effort before switching
        self._persist_current_agent_model()
        aid = str(self.agent_combo.currentData() or "codex").strip().lower()
        self._agent_id = aid
        self._populate_models_for_agent(aid)

    def _populate_models_for_agent(self, agent_id: str) -> None:
        models, cli_def_model, cli_def_effort = discover_cli_models(agent_id, home_dir=self._home_dir)
        self._models_cache = models
        saved_model, saved_effort = get_cli_agent_model_config(agent_id)

        self.model_combo.blockSignals(True)
        self.model_combo.clear()
        def_label = f"(Use CLI Default: {cli_def_model})" if cli_def_model else "(Use CLI Default)"
        self.model_combo.addItem(def_label, "")
        for opt in models:
            lbl = f"{opt.display_name} ({opt.slug})" if opt.display_name != opt.slug else opt.slug
            self.model_combo.addItem(lbl, opt.slug)

        if saved_model:
            idx = self.model_combo.findData(saved_model)
            if idx >= 0:
                self.model_combo.setCurrentIndex(idx)
            else:
                self.model_combo.setEditText(saved_model)
        else:
            self.model_combo.setCurrentIndex(0)
        self.model_combo.blockSignals(False)

        self._populate_efforts(selected_effort=saved_effort, cli_default_effort=cli_def_effort)

    def _current_model_slug(self) -> str:
        idx = self.model_combo.currentIndex()
        if idx >= 0 and self.model_combo.currentText() == self.model_combo.itemText(idx):
            return str(self.model_combo.itemData(idx) or "").strip()
        txt = self.model_combo.currentText().strip()
        if txt.startswith("(Use CLI Default"):
            return ""
        return txt

    def _populate_efforts(self, *, selected_effort: str = "", cli_default_effort: str = "") -> None:
        slug = self._current_model_slug()
        levels: Sequence[str] = ()
        for opt in self._models_cache:
            if opt.slug == slug:
                levels = opt.reasoning_levels
                break
        if not levels and self._agent_id in ("codex", "claude"):
            # Fallback effort set when custom model slug is entered
            for opt in self._models_cache:
                if opt.reasoning_levels:
                    levels = opt.reasoning_levels
                    break

        self.effort_combo.clear()
        def_eff_label = f"(CLI Default: {cli_default_effort})" if cli_default_effort else "(CLI Default)"
        self.effort_combo.addItem(def_eff_label, "")
        for lvl in levels:
            self.effort_combo.addItem(lvl, lvl)

        if selected_effort:
            idx = self.effort_combo.findData(selected_effort)
            if idx >= 0:
                self.effort_combo.setCurrentIndex(idx)
            else:
                self.effort_combo.addItem(selected_effort, selected_effort)
                self.effort_combo.setCurrentIndex(self.effort_combo.count() - 1)
        else:
            self.effort_combo.setCurrentIndex(0)
        self.effort_combo.setEnabled(self.effort_combo.count() > 1)

    def _on_model_changed(self) -> None:
        _, _, cli_def_effort = discover_cli_models(self._agent_id, home_dir=self._home_dir)
        cur_eff = str(self.effort_combo.currentData() or "").strip().lower()
        self._populate_efforts(selected_effort=cur_eff, cli_default_effort=cli_def_effort)

    def _on_task_mode_changed(self) -> None:
        if self.task_combo is None or self.custom_prompt_edit is None:
            return
        mode = str(self.task_combo.currentData() or "interactive")
        self.custom_prompt_edit.setVisible(mode == "custom")

    def _persist_current_agent_model(self) -> Tuple[str, str]:
        model = self._current_model_slug()
        effort = str(self.effort_combo.currentData() or "").strip().lower()
        set_cli_agent_model_config(self._agent_id, model, effort)
        return model, effort

    def save_inputs(self) -> CliContextDialogResult:
        """Persist edited model/effort, standing company context, and workspace `_USER_NOTES.md`."""
        model, effort = self._persist_current_agent_model()
        standing_txt = self.standing_edit.toPlainText().strip()
        save_standing_cli_context(standing_txt)

        ws_notes = ""
        if self.workspace_notes_edit is not None:
            ws_notes = self.workspace_notes_edit.toPlainText().strip()
            if self._target_dir is not None:
                save_workspace_user_notes(self._target_dir, ws_notes)

        task_mode = str(self.task_combo.currentData()) if self.task_combo is not None else "interactive"
        custom_q = (
            self.custom_prompt_edit.toPlainText().strip()
            if self.custom_prompt_edit is not None and task_mode == "custom"
            else ""
        )
        return CliContextDialogResult(
            launch_requested=self._launch_requested,
            agent_id=self._agent_id,
            task_mode=task_mode,
            custom_question=custom_q,
            model=model,
            reasoning_effort=effort,
            workspace_notes=ws_notes,
            standing_context=standing_txt,
        )

    def _on_save_only(self) -> None:
        self._launch_requested = False
        self.save_inputs()
        self.accept()

    def _on_save_and_launch(self) -> None:
        self._launch_requested = True
        self.save_inputs()
        self.accept()

    def get_result(self) -> CliContextDialogResult:
        return CliContextDialogResult(
            launch_requested=self._launch_requested,
            agent_id=self._agent_id,
            task_mode=str(self.task_combo.currentData()) if self.task_combo is not None else "interactive",
            custom_question=(
                self.custom_prompt_edit.toPlainText().strip()
                if self.custom_prompt_edit is not None
                else ""
            ),
            model=self._current_model_slug(),
            reasoning_effort=str(self.effort_combo.currentData() or "").strip().lower(),
            workspace_notes=self.workspace_notes_edit.toPlainText().strip() if self.workspace_notes_edit is not None else "",
            standing_context=self.standing_edit.toPlainText().strip(),
        )
