"""Considered Inputs viewer launch, dossier export worker, and CLI integration mixin."""

import os
import sys
from pathlib import Path
from typing import Dict, Optional

from PyQt5.QtWidgets import QDialog, QMessageBox

from core.ai.cli_detector import detect_cli_agents, format_active_model_badge
from core.ai.cli_launcher import (
    add_model_selector_submenu,
    launch_cli_in_terminal,
    open_mcp_agent_setup_dialog,
    write_cli_context_file,
)
from core.ui.cli_context_dialog import CliContextEditorDialog
from core.utils.folder_opener import open_folder as _open_folder
from core.utils.paths import strip_extended_path, to_extended_path
from modules.meetings.core.exporters.export_manifest import load_manifest
from modules.meetings.core.exporters.topic_dossier_exporter import TopicDossierExporterThread


def _resolve_shim_attr(name: str, default):
    """Forward calls through topic_dossier_dialog shim when monkeypatched in tests."""
    shim_mod = sys.modules.get("modules.meetings.ui.dialogs.topic_dossier_dialog")
    return getattr(shim_mod, name, default) if shim_mod else default


class TopicDossierExportCliMixin:
    """Manages TopicInputsViewerDialog launch, TopicDossierExporterThread, and LLM CLIs."""

    def _on_view_considered_inputs(self):
        if not self.active_topic:
            QMessageBox.warning(self, "No Topic", "Please select a topic.")
            return
        if not self.selected_cohorts:
            QMessageBox.warning(self, "No Meetings", "No meetings found in the selected timeframe.")
            return

        from modules.meetings.ui.dialogs.topic_inputs_dialog import TopicInputsViewerDialog

        tf_label = self.combo_cycle_window.currentText()
        if "Custom" in tf_label:
            tf_label = f"{self.edit_date_from.date().toString('yyyy-MM-dd')} to {self.edit_date_to.date().toString('yyyy-MM-dd')}"

        dlg = TopicInputsViewerDialog(
            meetings_db=self.meetings_db,
            topic=self.active_topic,
            selected_rounds=self.selected_cohorts,
            scope_mode="full_dossier",
            timeframe_label=tf_label,
            parent=None,
        )
        dlg.export_requested.connect(self._on_start_export)
        dlg.show()
        dlg.raise_()
        dlg.activateWindow()
        self._inputs_dlg = dlg

    def _on_start_export(self):
        if self._worker and self._worker.isRunning():
            self.raise_()
            self.activateWindow()
            return
        if not self.active_topic:
            QMessageBox.warning(self, "No Topic", "Please select a topic.")
            return
        if not self.selected_cohorts:
            QMessageBox.warning(self, "No Meetings", "No meetings found in the selected timeframe.")
            return

        if getattr(self, "_inputs_dlg", None) is not None:
            try:
                if self._inputs_dlg.isVisible():
                    self._inputs_dlg.close()
            except Exception:
                pass

        dest_path = Path(self.edit_dest_folder.text().strip())
        parent_dir = dest_path.parent
        target_name = dest_path.name

        self.btn_export.setEnabled(False)
        self.prog_bar.setVisible(True)
        self.prog_bar.setValue(0)
        self.lbl_status.setText("Starting export...")
        self.raise_()
        self.activateWindow()

        self._worker = _resolve_shim_attr("TopicDossierExporterThread", TopicDossierExporterThread)(
            meetings_db=self.meetings_db,
            topic=self.active_topic,
            selected_cohorts=self.selected_cohorts,
            export_scope_mode="full_dossier",
            output_base_dir=parent_dir,
            target_folder_name=target_name,
            include_discussions=self.chk_discussions.isChecked(),
            deduplicate_revisions=self.chk_dedup.isChecked(),
            generate_consolidated_md=self.chk_consolidated.isChecked(),
            allow_network_download=self.chk_download.isChecked(),
        )
        self._worker.progress_msg.connect(self.lbl_status.setText)
        self._worker.progress_step.connect(self.prog_bar.setValue)
        self._worker.session_finished.connect(self._on_export_finished)
        self._worker.start()

    def _on_export_finished(self, success: bool, target_dir: str, msg: str):
        self.btn_export.setEnabled(True)
        self.prog_bar.setVisible(False)
        if success:
            self.lbl_status.setText("Export completed successfully!")
            self.last_target_dir = target_dir
            self._check_incremental_status()
            if target_dir and os.path.isdir(to_extended_path(target_dir)):
                self.btn_open_folder.setEnabled(True)
                self.btn_launch_cli.setEnabled(True)
            QMessageBox.information(self, "Export Complete", msg)
        else:
            self.lbl_status.setText("Export failed.")
            QMessageBox.critical(self, "Export Error", msg)

    def _on_open_output(self):
        dest_text = self.edit_dest_folder.text().strip()
        last_dir = str(getattr(self, "last_target_dir", "") or "").strip()
        target = dest_text if (dest_text and os.path.isdir(to_extended_path(dest_text))) else (last_dir or dest_text)
        if not target:
            QMessageBox.warning(self, "No Folder", "No destination folder is selected.")
            return
        clean_target = os.path.normpath(strip_extended_path(target))
        if os.path.isdir(to_extended_path(clean_target)):
            _resolve_shim_attr("_open_folder", _open_folder)(Path(clean_target))
        else:
            self.btn_open_folder.setEnabled(False)
            self.btn_launch_cli.setEnabled(False)
            QMessageBox.information(
                self,
                "Folder Not Found",
                f"Output folder does not exist yet:\n{clean_target}\n\nPlease run 'Export / Update Topic Dossier' first.",
            )

    def _collect_topic_manifest_summary(self, folder: Path) -> Dict[str, object]:
        """Inspect `.export_manifest.json` in `folder` for Stage-1 and Stage-2 status."""
        summary: Dict[str, object] = {
            "analyzed_meetings": [],
            "stale_meetings": [],
            "pending_meetings": [],
            "synthesis_status": "pending",
        }
        manifest = load_manifest(folder)
        if not manifest or not manifest.exported_meetings:
            return summary

        for m_key, m_state in manifest.exported_meetings.items():
            if not isinstance(m_state, dict):
                continue
            lc = str(m_state.get("lifecycle", "") or "").strip().lower()
            d_c = int(m_state.get("decision_count", 0) or 0)
            t_c = int(m_state.get("tdoc_count", 0) or 0)
            if lc == "no_activity" or (d_c == 0 and t_c == 0):
                continue

            a_file = Path(to_extended_path(folder / m_key / "_company_positions_analysis.md"))
            st = str(m_state.get("analysis_status") or "").strip().lower()
            if not a_file.is_file():
                summary["pending_meetings"].append(m_key)
            elif st == "stale":
                summary["stale_meetings"].append(m_key)
            else:
                summary["analyzed_meetings"].append(m_key)

        syn_state = manifest.synthesis_state or {}
        summary["synthesis_status"] = str(syn_state.get("status") or "pending").lower()
        return summary

    def _populate_cli_menu(self):
        self.cli_menu.clear()
        self.cli_menu.setToolTipsVisible(True)

        statuses = detect_cli_agents()
        installed = [st for st in statuses.values() if st.installed]
        missing = [st for st in statuses.values() if not st.installed]

        if installed:
            for st in installed:
                badge = format_active_model_badge(st.agent_id)
                sub = self.cli_menu.addMenu(f"{st.icon} {st.label}  [{badge}]")
                sub.setToolTipsVisible(True)
                sub.setToolTip(f"{st.executable_path}\nModel: {badge}")

                act_launch = sub.addAction(f"🚀 Launch {st.label} in Topic Folder")
                act_launch.setToolTip(
                    f"Open {st.label} interactively in the topic dossier folder without running any automatic analysis."
                )
                act_launch.triggered.connect(lambda _=False, aid=st.agent_id: self._launch_cli_for_topic(aid, "topic_idle"))

                act_ctx = sub.addAction("📝 Edit Topic Notes && Launch...")
                act_ctx.setToolTip(
                    "Edit persistent topic _USER_NOTES.md / standing company guidance and select Model / Reasoning Effort before launching."
                )
                act_ctx.triggered.connect(lambda _=False, aid=st.agent_id: self._launch_cli_for_topic(aid, "edit_context"))

                sub.addSeparator()
                add_model_selector_submenu(sub, st.agent_id)
        else:
            no_cli = self.cli_menu.addAction("⚪ No LLM CLI Detected (codex / claude / gemini)")
            no_cli.setEnabled(False)

        if missing:
            self.cli_menu.addSeparator()
            for st in missing:
                act_miss = self.cli_menu.addAction(f"⚪ {st.label} (Not Found)")
                act_miss.setEnabled(False)
                act_miss.setToolTip(f"Not detected.\nInstall: {st.install_hint}")

        self.cli_menu.addSeparator()
        act_cfg = self.cli_menu.addAction("⚙️ Configure CLI Agents && MCP...")
        act_cfg.triggered.connect(lambda _=False: open_mcp_agent_setup_dialog(self))

    def _launch_cli_for_topic(self, agent_id: str, task_mode: str = "topic_idle"):
        dest_text = self.edit_dest_folder.text().strip()
        if not dest_text or not os.path.isdir(to_extended_path(dest_text)):
            QMessageBox.warning(self, "Folder Not Found", "Please export the topic dossier before launching a CLI agent.")
            return

        target_dir = Path(strip_extended_path(dest_text))
        topic_name = self.active_topic.name if self.active_topic else target_dir.name
        model_override: Optional[str] = None
        effort_override: Optional[str] = None

        if task_mode == "edit_context":
            dlg = CliContextEditorDialog(
                agent_id=agent_id,
                scope_label=f"Topic: {topic_name}",
                target_dir=target_dir,
                task_options=[],
                initial_task_mode="topic_idle",
                parent=self,
            )
            if dlg.exec_() != QDialog.Accepted:
                return
            res = dlg.get_result()
            if not res.launch_requested:
                return
            model_override = res.model
            effort_override = res.reasoning_effort
            task_mode = "topic_idle"

        summary = self._collect_topic_manifest_summary(target_dir)
        write_cli_context_file(
            target_dir,
            task_mode=task_mode,
            topic_name=topic_name,
            manifest_summary=summary,
        )
        # Do not send an automatic bootstrap prompt on cross-WG/cross-meeting corpora
        _resolve_shim_attr("launch_cli_in_terminal", launch_cli_in_terminal)(
            agent_id,
            target_dir,
            bootstrap_prompt="",
            parent=self,
            model=model_override,
            reasoning_effort=effort_override,
        )
