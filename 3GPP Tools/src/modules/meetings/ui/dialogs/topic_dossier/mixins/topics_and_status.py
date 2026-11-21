"""Topic selection, cohort preview, CRUD dialogs, and incremental status mixin."""

import os
from pathlib import Path
from typing import Dict, Optional, Set

from PyQt5.QtWidgets import QDialog, QFileDialog, QListWidgetItem, QMessageBox

from core.ui.theme import C
from core.utils.paths import to_extended_path
from modules.meetings.core.cohort_resolver import MeetingCohortResolver
from modules.meetings.core.exporters.export_manifest import load_manifest
from modules.meetings.core.exporters.topic_digest_builder import STAGE2_SYNTHESIS_FILENAME
from modules.meetings.core.exporters.topic_dossier_exporter import (
    check_and_refresh_dossier,
    discover_stage2_packages,
    get_default_topic_export_base,
    sanitize_folder_name,
)
from modules.meetings.core.topic_manager import MetaTopicDefinition
from modules.meetings.ui.dialogs.topic_dossier.editor_dialog import TopicEditorDialog


class TopicDossierTopicsStatusMixin:
    """Handles topic CRUD, cohort resolution, incremental manifest checkup, and status badges."""

    def _load_topics_list(self, preferred_id: Optional[str] = None, preferred_index: Optional[int] = None):
        self.list_topics.blockSignals(True)
        self.list_topics.clear()
        topics = self.topic_manager.list_topics()
        selected_row = -1
        for idx, t in enumerate(topics):
            item = QListWidgetItem(f"{t.name}")
            item.setToolTip(f"{t.description}\nWGs: {', '.join(t.target_wgs)}")
            self.list_topics.addItem(item)
            if preferred_id and t.id == preferred_id:
                selected_row = idx

        self.list_topics.blockSignals(False)

        if topics:
            self.btn_edit.setEnabled(True)
            self.btn_clone.setEnabled(True)
            self.btn_del.setEnabled(True)
            self.btn_export.setEnabled(True)
            if hasattr(self, "btn_view_inputs"):
                self.btn_view_inputs.setEnabled(True)
            if selected_row >= 0:
                self.list_topics.setCurrentRow(selected_row)
            elif preferred_index is not None and 0 <= preferred_index < len(topics):
                self.list_topics.setCurrentRow(preferred_index)
            else:
                self.list_topics.setCurrentRow(0)
        else:
            self.active_topic = None
            self.lbl_topic_title.setText("No Topics Defined")
            self.lbl_topic_desc.setText("Click '➕ New' or '🔄 Restore Built-in Topics...' to get started.")
            self.edit_dest_folder.clear()
            self.selected_cohorts = []
            self.btn_edit.setEnabled(False)
            self.btn_clone.setEnabled(False)
            self.btn_del.setEnabled(False)
            self.btn_export.setEnabled(False)
            if hasattr(self, "btn_view_inputs"):
                self.btn_view_inputs.setEnabled(False)
            self.lbl_cohorts_summary.setText("No active topic selected.")

    def _on_topic_selected(self, row: int):
        topics = self.topic_manager.list_topics()
        if 0 <= row < len(topics):
            self.active_topic = topics[row]
            self.lbl_topic_title.setText(f"🌐 {self.active_topic.name}")
            wgs_str = ", ".join(self.active_topic.target_wgs)
            self.lbl_topic_desc.setText(f"<b>WGs:</b> {wgs_str} | {self.active_topic.description}")

            # Update default destination in user meeting cache
            base_dir = get_default_topic_export_base()
            default_name = f"Topic_{sanitize_folder_name(self.active_topic.name)}"
            self.edit_dest_folder.setText(str(base_dir / default_name))

            self._update_cohorts_preview()
            self._check_incremental_status()
            self.btn_edit.setEnabled(True)
            self.btn_clone.setEnabled(True)
            self.btn_del.setEnabled(True)
            self.btn_export.setEnabled(True)
            if hasattr(self, "btn_view_inputs"):
                self.btn_view_inputs.setEnabled(True)
        else:
            self.active_topic = None
            self.btn_edit.setEnabled(False)
            self.btn_clone.setEnabled(False)
            self.btn_del.setEnabled(False)
            self.btn_export.setEnabled(False)
            if hasattr(self, "btn_view_inputs"):
                self.btn_view_inputs.setEnabled(False)

    def _on_timeframe_preset_changed(self, idx: int):
        is_custom = (idx == 6)  # "📅 Custom Date Range..."
        self.date_range_widget.setVisible(is_custom)
        self._update_cohorts_preview()

    def _update_cohorts_preview(self):
        if not self.active_topic:
            return

        idx = self.combo_cycle_window.currentIndex()
        topic_copy = MetaTopicDefinition.from_dict(self.active_topic.to_dict())

        if idx == 0:  # Past 4 Meeting Rounds (~1 Year)
            topic_copy.cycle_mode = "past_n_cycles"
            topic_copy.past_n_cycles = 4
        elif idx == 1:  # Past 6 Meeting Rounds (~1.5 Years)
            topic_copy.cycle_mode = "past_n_cycles"
            topic_copy.past_n_cycles = 6
        elif idx == 2:  # Past 8 Meeting Rounds (~2 Years)
            topic_copy.cycle_mode = "past_n_cycles"
            topic_copy.past_n_cycles = 8
        elif idx == 3:  # Release 20 Timeframe (2025–2027)
            topic_copy.cycle_mode = "date_range"
            topic_copy.date_from = "2025-01-01"
            topic_copy.date_to = "2027-12-31"
        elif idx == 4:  # Release 19 Timeframe (2023–2025)
            topic_copy.cycle_mode = "date_range"
            topic_copy.date_from = "2023-01-01"
            topic_copy.date_to = "2025-12-31"
        elif idx == 5:  # Release 18 Timeframe (2021–2023)
            topic_copy.cycle_mode = "date_range"
            topic_copy.date_from = "2021-01-01"
            topic_copy.date_to = "2023-12-31"
        elif idx == 6:  # 📅 Custom Date Range...
            topic_copy.cycle_mode = "date_range"
            topic_copy.date_from = self.edit_date_from.date().toString("yyyy-MM-dd")
            topic_copy.date_to = self.edit_date_to.date().toString("yyyy-MM-dd")
        else:  # All Available Meetings
            topic_copy.cycle_mode = "cycle_range"
            topic_copy.from_cycle = "2018-Q1"
            topic_copy.to_cycle = "2035-Q4"

        self.selected_cohorts = MeetingCohortResolver.resolve_cohorts_for_topic(self.meetings_db, topic_copy)
        if self.selected_cohorts:
            first_c = self.selected_cohorts[0].label
            last_c = self.selected_cohorts[-1].label
            total_mtgs = sum(len(c.meetings) for c in self.selected_cohorts)
            mtg_names = []
            for c in self.selected_cohorts[-2:]:
                for m in c.meetings:
                    mtg_names.append(f"{m.get('wg_name')} #{m.get('meeting_number')}")
            preview_str = ", ".join(mtg_names[:6])
            if len(mtg_names) > 6:
                preview_str += "..."
            self.lbl_cohorts_summary.setText(
                f"Resolved: <b>{len(self.selected_cohorts)} Synchronized Meeting Round(s)</b> ({first_c} → {last_c}) with <b>{total_mtgs} meetings</b> across {', '.join(sorted(self.active_topic.target_wgs))}.<br/>"
                f"<span style='color: {C.TEXT_MUTED};'>Recent meetings: {preview_str}</span>"
            )
        else:
            self.lbl_cohorts_summary.setText("No meetings found matching this timeframe and target WGs.")

    def _check_incremental_status(self):
        folder_str = self.edit_dest_folder.text().strip()
        if not folder_str:
            self.lbl_inc_status.setText("⚠️ Please specify a destination folder.")
            self.btn_open_folder.setEnabled(False)
            self.btn_open_folder.setToolTip("Destination folder is empty.")
            self.btn_launch_cli.setEnabled(False)
            self.btn_check_status.setEnabled(False)
            return

        folder_exists = os.path.isdir(to_extended_path(folder_str))
        self.btn_open_folder.setEnabled(folder_exists)
        self.btn_launch_cli.setEnabled(folder_exists)
        self.btn_open_folder.setToolTip(
            f"Open in File Explorer: {folder_str}"
            if folder_exists
            else "Destination folder does not exist yet. Run an export first."
        )

        p = Path(folder_str)
        manifest_p = p / ".export_manifest.json"
        has_manifest = manifest_p.is_file()
        self.btn_check_status.setEnabled(has_manifest)
        if not has_manifest:
            self.lbl_inc_status.setText("🆕 <b>New Directory:</b> Initial complete export.")
            self.lbl_inc_status.setStyleSheet(f"color: {C.ACCENT_ALT}; font-size: 11px;")
            return

        manifest = load_manifest(p)
        if not manifest or not manifest.exported_meetings:
            self.lbl_inc_status.setText(
                "🔄 <b>Existing Corpus Detected:</b> Incremental update mode active (skips already exported TDocs, writes delta feed)."
            )
            self.lbl_inc_status.setStyleSheet(f"color: {C.EMERALD_DARK}; font-size: 11px;")
            return

        analyzed_cnt = 0
        stale_cnt = 0
        pending_cnt = 0
        no_act_cnt = 0
        dec_cnt = 0
        doc_cnt = 0
        stub_cnt = 0
        max_input_mtime = 0.0
        excluded_wgs: Set[str] = set()

        for m_key, m_state in manifest.exported_meetings.items():
            if not isinstance(m_state, dict):
                continue
            if m_state.get("wg"):
                excluded_wgs.add(str(m_state["wg"]).strip())
            d_c = int(m_state.get("decision_count", 0) or 0)
            t_c = int(m_state.get("tdoc_count", 0) or 0)
            s_c = int(m_state.get("stub_count", 0) or 0)
            dec_cnt += d_c
            doc_cnt += t_c
            stub_cnt += s_c

            lc = str(m_state.get("lifecycle", "") or "").strip().lower()
            if lc == "no_activity" or (d_c == 0 and t_c == 0):
                no_act_cnt += 1
                continue

            cur_snap = m_state.get("current_snapshot") or {}
            pkt_mtime = float(cur_snap.get("packet_mtime", 0.0) or 0.0)
            max_input_mtime = max(max_input_mtime, pkt_mtime)

            a_file = Path(to_extended_path(p / m_key / "_company_positions_analysis.md"))
            if not a_file.is_file():
                pending_cnt += 1
                continue

            st = m_state.get("analysis_status")
            ana_snap = m_state.get("analyzed_snapshot") or {}
            a_mtime = float(os.path.getmtime(a_file))
            max_input_mtime = max(max_input_mtime, a_mtime)

            if st == "stale" and a_mtime > pkt_mtime > 0.0:
                analyzed_cnt += 1
            elif st in ("analyzed", "stale"):
                if st == "analyzed" and ana_snap and cur_snap and ana_snap.get("fingerprint") != cur_snap.get("fingerprint") and a_mtime <= pkt_mtime:
                    stale_cnt += 1
                elif st == "stale":
                    stale_cnt += 1
                else:
                    analyzed_cnt += 1
            else:
                # Legacy manifest without snapshot fields
                if pkt_mtime > 0.0 and a_mtime < pkt_mtime - 5.0:
                    stale_cnt += 1
                else:
                    analyzed_cnt += 1

        packages = discover_stage2_packages(p, excluded_dirs=excluded_wgs)
        completed_pkgs = [pkg for pkg in packages if pkg.get("status") == "completed"]
        in_prog_pkgs = [pkg for pkg in packages if pkg.get("status") == "in_progress"]
        syn_state_dict = manifest.synthesis_state or {}
        syn_state = str(syn_state_dict.get("status") or "pending").lower()

        if not completed_pkgs:
            if in_prog_pkgs:
                syn_badge = f"🏗️ Stage-2 In Progress [{in_prog_pkgs[-1]['folder_name']}]"
            else:
                syn_badge = "⏳ Stage-2 Pending"
        else:
            latest_pkg = completed_pkgs[-1]
            dated_lbl = f"{latest_pkg.get('as_of') or 'dated'} ({latest_pkg.get('revision') or 'r1'})"
            folder_tag = f" [{latest_pkg['folder_name']}]" if latest_pkg.get("folder_name") else " [root]"
            runs_tag = f" · {len(completed_pkgs)} runs" if len(completed_pkgs) > 1 else ""
            latest_mtime = float(latest_pkg.get("report_mtime", 0.0) or 0.0)
            is_stale = (
                syn_state == "stale"
                or stale_cnt > 0
                or (latest_mtime > 0.0 and max_input_mtime > latest_mtime + 0.5)
            )
            if is_stale:
                chg_cnt = len(syn_state_dict.get("changed_meetings") or [])
                if chg_cnt:
                    syn_badge = f"🔄 Stage-2 Stale: {chg_cnt} updated since {dated_lbl}{folder_tag}{runs_tag}"
                else:
                    syn_badge = f"🔄 Stage-2 Stale (since {dated_lbl}{folder_tag}){runs_tag}"
            else:
                syn_badge = f"✅ Stage-2: {dated_lbl}{folder_tag}{runs_tag}"

        total_mtgs = len(manifest.exported_meetings)
        stage1_parts = [f"✅ {analyzed_cnt} Analyzed"]
        if stale_cnt:
            stage1_parts.append(f"🔄 {stale_cnt} Stale")
        stage1_parts.append(f"⏳ {pending_cnt} Pending")
        if no_act_cnt:
            stage1_parts.append(f"⚪ {no_act_cnt} No Activity")

        stub_note = f" ({stub_cnt} stubs)" if stub_cnt else ""
        self.lbl_inc_status.setText(
            f"🔄 <b>Existing Corpus Detected:</b> {total_mtgs} meetings · {dec_cnt} decisions · "
            f"{doc_cnt} docs{stub_note}<br>"
            f"<b>LLM Stage-1:</b> {' · '.join(stage1_parts)} &nbsp;|&nbsp; <b>{syn_badge}</b>"
        )
        self.lbl_inc_status.setStyleSheet(f"color: {C.EMERALD_DARK}; font-size: 11px;")

    def _on_check_status(self, silent: bool = False) -> Dict[str, object]:
        folder_str = self.edit_dest_folder.text().strip()
        if not folder_str:
            if not silent:
                QMessageBox.warning(self, "No Destination", "Please specify a destination folder first.")
            return {"ok": False, "error": "Empty destination"}

        res = check_and_refresh_dossier(Path(folder_str), topic=self.active_topic)
        self._check_incremental_status()

        if not res.get("ok"):
            if not silent:
                QMessageBox.warning(
                    self,
                    "Checkup Failed",
                    str(res.get("error") or "Could not inspect destination corpus."),
                )
            return res

        syn = str(res.get("synthesis_status") or "pending").lower()
        s2_state = res.get("synthesis_state") or {}
        latest_pkg = s2_state.get("latest_package") or {}
        all_pkgs = list(s2_state.get("packages") or [])
        completed_cnt = sum(1 for pk in all_pkgs if pk.get("status") == "completed")

        if latest_pkg:
            dated_s2 = f"{latest_pkg.get('as_of') or 'dated'} ({latest_pkg.get('revision') or 'r1'})"
            s2_folder = latest_pkg.get("folder_name") or "(root)"
            if syn == "stale":
                chg_cnt = len(s2_state.get("changed_meetings") or [])
                syn_short = f"🔄 stale ({chg_cnt} updated since {dated_s2})" if chg_cnt else f"🔄 stale (since {dated_s2})"
            else:
                syn_short = f"✅ {dated_s2} [{s2_folder}]"
        elif syn == "in_progress":
            syn_short = f"🏗️ in-progress [{s2_state.get('folder_name') or 'stage2'}]"
        else:
            syn_short = "⏳ pending"

        self.lbl_status.setText(
            f"Index refreshed: ✅ {res['analyzed']} analyzed · 🔄 {res['stale']} stale · "
            f"⏳ {res['pending']} pending · ⚪ {res['no_activity']} no-activity · Stage-2: {syn_short}"
        )

        if not silent:
            promo = res.get("newly_promoted", 0)
            promo_line = f"\n• Newly Linked / Promoted Analyses: +{promo}" if promo else ""
            if latest_pkg:
                val_note = (
                    f" | Validation: {'PASSED' if latest_pkg.get('validation_passed') else 'ISSUES'}"
                    if latest_pkg.get("validation_passed") is not None
                    else ""
                )
                claims_note = f" ({latest_pkg['claims']} claims)" if latest_pkg.get("claims") is not None else ""
                s2_detail = (
                    f"• Stage-2 Synthesis: {syn.upper()} — {s2_folder} "
                    f"[{latest_pkg.get('package_id', '')}, cut-off {latest_pkg.get('as_of', 'N/A')} {latest_pkg.get('revision', '')}]"
                    f"{val_note}{claims_note}"
                    + (f" ({completed_cnt} package(s) total)" if completed_cnt > 1 else "")
                )
                chg_list = list(s2_state.get("changed_meetings") or [])
                if chg_list:
                    preview_chg = ", ".join(chg_list[:5]) + ("..." if len(chg_list) > 5 else "")
                    s2_detail += f"\n  ↳ Updated since baseline ({len(chg_list)}): {preview_chg}"
            else:
                s2_detail = f"• Stage-2 E2E Synthesis ({STAGE2_SYNTHESIS_FILENAME} / stage2_*): {syn.upper()}"

            QMessageBox.information(
                self,
                "Corpus Status & Index Refreshed",
                f"Checked '{res.get('topic_name', '')}' ({res['total_meetings']} meetings):\n"
                f"• Decisions: {res['total_decisions']} | Active Docs: {res['total_active_docs']} (Stubs: {res['total_stubs']})\n"
                f"• Stage-1 Analyses: ✅ {res['analyzed']} Analyzed | 🔄 {res['stale']} Stale (Delta Ready) | "
                f"⏳ {res['pending']} Pending | ⚪ {res['no_activity']} No Activity"
                f"{promo_line}\n"
                f"{s2_detail}\n\n"
                f"Updated index.md, company_positions_e2e_prompt.md, per-meeting briefs, and .export_manifest.json.",
            )
        return res

    def _on_browse_dest(self):
        cur = self.edit_dest_folder.text().strip() or str(get_default_topic_export_base())
        folder = QFileDialog.getExistingDirectory(self, "Select Output Directory", cur)
        if folder:
            self.edit_dest_folder.setText(folder)

    def _on_new_topic(self):
        dlg = TopicEditorDialog(meetings_db=self.meetings_db, parent=self)
        if dlg.exec_() == QDialog.Accepted:
            self.topic_manager.save_topic(dlg.result_topic)
            self._load_topics_list(preferred_id=dlg.result_topic.id)

    def _on_edit_topic(self):
        if not self.active_topic:
            return
        dlg = TopicEditorDialog(topic=self.active_topic, meetings_db=self.meetings_db, parent=self)
        if dlg.exec_() == QDialog.Accepted:
            self.topic_manager.save_topic(dlg.result_topic)
            self._load_topics_list(preferred_id=dlg.result_topic.id)

    def _on_clone_topic(self):
        if not self.active_topic:
            return
        new_topic = self.topic_manager.clone_topic(self.active_topic.id)
        preferred_id = new_topic.id if new_topic else None
        self._load_topics_list(preferred_id=preferred_id)

    def _on_delete_topic(self):
        if not self.active_topic:
            return
        cur_row = self.list_topics.currentRow()
        ret = QMessageBox.question(
            self,
            "Confirm Delete",
            f"Are you sure you want to delete topic '{self.active_topic.name}'?\n\nThis will permanently remove it from your saved topics.",
            QMessageBox.Yes | QMessageBox.No,
        )
        if ret == QMessageBox.Yes:
            self.topic_manager.delete_topic(self.active_topic.id)
            remaining = self.topic_manager.list_topics()
            next_idx = max(0, min(cur_row, len(remaining) - 1)) if remaining else None
            self._load_topics_list(preferred_index=next_idx)

    def _on_restore_defaults(self):
        ret = QMessageBox.question(
            self,
            "Restore Built-in Topics",
            "Would you like to restore standard built-in 3GPP topics (Ambient IoT, AI/ML Phase 2, 6G Architecture, Network Energy Savings, Data Framework) to your list?\n\nAny custom topics you created will be preserved.",
            QMessageBox.Yes | QMessageBox.No,
        )
        if ret == QMessageBox.Yes:
            count = self.topic_manager.restore_default_topics(overwrite=False)
            self._load_topics_list()
            if count > 0:
                QMessageBox.information(self, "Topics Restored", f"{count} built-in topic(s) were restored to your list.")
            else:
                QMessageBox.information(self, "Topics Restored", "All built-in topics are already present in your list.")
