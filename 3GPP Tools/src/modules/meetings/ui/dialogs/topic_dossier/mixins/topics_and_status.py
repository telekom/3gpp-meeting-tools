"""Topic selection, cohort preview, CRUD dialogs, and incremental status mixin."""

import datetime
import html
import os
from pathlib import Path
from typing import Dict, List, Optional, Set

from PyQt5.QtCore import QDate
from PyQt5.QtWidgets import QDialog, QFileDialog, QListWidgetItem, QMessageBox

from core.ui.theme import C
from core.utils.paths import to_extended_path
from modules.meetings.core.cohort_resolver import (
    DEFAULT_UPCOMING_LOOKAHEAD_DAYS,
    MeetingCohortResolver,
)
from modules.meetings.core.exporters.export_manifest import load_manifest
from modules.meetings.core.exporters.topic_dossier_exporter import (
    check_and_refresh_dossier,
    discover_stage2_packages,
    get_default_topic_export_base,
    sanitize_folder_name,
)
from modules.meetings.core.topic_manager import MetaTopicDefinition
from modules.meetings.ui.dialogs.topic_dossier.checkup_worker import TopicCheckupWorker
from modules.meetings.ui.dialogs.topic_dossier.editor_dialog import TopicEditorDialog
from modules.meetings.ui.dialogs.topic_dossier.status_report_dialog import CorpusStatusReportDialog
from modules.meetings.ui.dialogs.topic_dossier.styles import (
    build_incremental_banner_html,
    format_meeting_label,
    format_meeting_num,
    format_wg_badges_html,
)


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
            wg_badges = format_wg_badges_html(self.active_topic.target_wgs)
            desc_txt = html.escape(self.active_topic.description) if self.active_topic.description else ""
            sep = " &nbsp;·&nbsp; " if (wg_badges and desc_txt) else ""
            self.lbl_topic_desc.setText(f"{wg_badges}{sep}{desc_txt}")

            # Update default destination in user meeting cache
            base_dir = get_default_topic_export_base()
            default_name = f"Topic_{sanitize_folder_name(self.active_topic.name)}"
            dest_folder = str(base_dir / default_name)
            self.edit_dest_folder.setText(dest_folder)

            self._apply_topic_default_timeframe(dest_folder)
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

    def _current_timeframe_key(self) -> str:
        data = self.combo_cycle_window.currentData()
        return str(data) if data else "since_date"

    def _select_timeframe_key(self, key: str) -> None:
        idx = self.combo_cycle_window.findData(key)
        if idx >= 0:
            self.combo_cycle_window.setCurrentIndex(idx)

    def _sync_date_range_controls(self) -> None:
        key = self._current_timeframe_key()
        is_since = key == "since_date"
        is_custom = key == "custom_range"
        self.date_range_widget.setVisible(is_since or is_custom)
        self.lbl_date_to.setVisible(is_custom)
        self.edit_date_to.setVisible(is_custom)
        self.lbl_open_ended_hint.setVisible(is_since)
        self.chk_include_future.setVisible(is_since)

    @staticmethod
    def _detect_corpus_start_date(folder_str: str) -> str:
        if not folder_str:
            return ""
        manifest = load_manifest(Path(folder_str))
        if not manifest or not manifest.exported_meetings:
            return ""
        dates = [
            str(m.get("start_date") or "").strip()[:10]
            for m in manifest.exported_meetings.values()
            if isinstance(m, dict) and str(m.get("start_date") or "").strip()
        ]
        valid = [d for d in dates if len(d) == 10 and d[4] == "-" and d[7] == "-"]
        return min(valid) if valid else ""

    def _apply_topic_default_timeframe(self, dest_folder: str) -> None:
        if not self.active_topic:
            return
        widgets = (
            self.combo_cycle_window,
            self.edit_date_from,
            self.edit_date_to,
            self.chk_include_future,
        )
        for w in widgets:
            w.blockSignals(True)
        try:
            self._anchored_corpus_date = ""
            d_from_str = str(self.active_topic.date_from or "").strip()[:10]
            d_to_str = str(self.active_topic.date_to or "").strip()[:10]
            q_from = QDate.fromString(d_from_str, "yyyy-MM-dd")
            q_to = QDate.fromString(d_to_str, "yyyy-MM-dd")

            if self.active_topic.cycle_mode == "since_date" and q_from.isValid():
                self._select_timeframe_key("since_date")
                self.edit_date_from.setDate(q_from)
                self.chk_include_future.setChecked(False)
            elif self.active_topic.cycle_mode == "date_range" and q_from.isValid() and q_to.isValid():
                self._select_timeframe_key("custom_range")
                self.edit_date_from.setDate(q_from)
                self.edit_date_to.setDate(q_to)
            else:
                earliest = self._detect_corpus_start_date(dest_folder)
                q_earliest = QDate.fromString(earliest, "yyyy-MM-dd") if earliest else QDate()
                self._select_timeframe_key("since_date")
                self.chk_include_future.setChecked(False)
                if q_earliest.isValid():
                    self.edit_date_from.setDate(q_earliest)
                    self._anchored_corpus_date = earliest
                else:
                    self.edit_date_from.setDate(QDate.currentDate().addYears(-1))
        finally:
            for w in widgets:
                w.blockSignals(False)
        self._sync_date_range_controls()

    def _on_timeframe_preset_changed(self, _idx: int):
        self._sync_date_range_controls()
        self._update_cohorts_preview()

    def _update_cohorts_preview(self):
        if not self.active_topic:
            return

        key = self._current_timeframe_key()
        topic_copy = MetaTopicDefinition.from_dict(self.active_topic.to_dict())
        upcoming_days: Optional[int] = DEFAULT_UPCOMING_LOOKAHEAD_DAYS

        if key == "since_date":
            topic_copy.cycle_mode = "since_date"
            topic_copy.date_from = self.edit_date_from.date().toString("yyyy-MM-dd")
            topic_copy.date_to = ""
            if self.chk_include_future.isChecked():
                upcoming_days = None
        elif key == "past_4":
            topic_copy.cycle_mode = "past_n_cycles"
            topic_copy.past_n_cycles = 4
        elif key == "past_6":
            topic_copy.cycle_mode = "past_n_cycles"
            topic_copy.past_n_cycles = 6
        elif key == "past_8":
            topic_copy.cycle_mode = "past_n_cycles"
            topic_copy.past_n_cycles = 8
        elif key == "rel20":
            topic_copy.cycle_mode = "date_range"
            topic_copy.date_from = "2025-01-01"
            topic_copy.date_to = "2027-12-31"
        elif key == "rel19":
            topic_copy.cycle_mode = "date_range"
            topic_copy.date_from = "2023-01-01"
            topic_copy.date_to = "2025-12-31"
        elif key == "custom_range":
            topic_copy.cycle_mode = "date_range"
            topic_copy.date_from = self.edit_date_from.date().toString("yyyy-MM-dd")
            topic_copy.date_to = self.edit_date_to.date().toString("yyyy-MM-dd")
        elif key == "rel18":
            topic_copy.cycle_mode = "date_range"
            topic_copy.date_from = "2021-01-01"
            topic_copy.date_to = "2023-12-31"
        else:
            topic_copy.cycle_mode = "cycle_range"
            topic_copy.from_cycle = "2018-Q1"
            topic_copy.to_cycle = "2035-Q4"

        self.selected_cohorts = MeetingCohortResolver.resolve_cohorts_for_topic(
            self.meetings_db,
            topic_copy,
            upcoming_days=upcoming_days,
        )
        if self.selected_cohorts:
            first_c = self.selected_cohorts[0].label
            last_c = self.selected_cohorts[-1].label
            total_mtgs = sum(len(c.meetings) for c in self.selected_cohorts)
            today_iso = datetime.date.today().strftime("%Y-%m-%d")
            by_wg: Dict[str, List[Dict[str, object]]] = {}
            for c in self.selected_cohorts:
                for m in c.meetings:
                    wg = str(m.get("wg_name") or "").strip().upper()
                    if wg:
                        by_wg.setdefault(wg, []).append(m)

            ordered_wgs = [w.upper() for w in self.active_topic.target_wgs if w.strip()]
            for wg in sorted(by_wg.keys()):
                if wg not in ordered_wgs:
                    ordered_wgs.append(wg)

            latest_parts: List[str] = []
            for wg in ordered_wgs:
                mtgs = by_wg.get(wg)
                if not mtgs:
                    continue
                mtgs_sorted = sorted(
                    mtgs,
                    key=lambda x: (str(x.get("start_date") or ""), str(x.get("folder_name") or "")),
                )
                past_or_cur = [m for m in mtgs_sorted if str(m.get("start_date") or "")[:10] <= today_iso]
                future = [m for m in mtgs_sorted if str(m.get("start_date") or "")[:10] > today_iso]
                if past_or_cur:
                    lbl = format_meeting_label(past_or_cur[-1])
                    if future:
                        lbl += f" → #{format_meeting_num(future[-1])}"
                    latest_parts.append(lbl)
                elif future:
                    latest_parts.append(format_meeting_label(future[-1]))

            anchored_date = getattr(self, "_anchored_corpus_date", "")
            anchor_note = ""
            if key == "since_date" and anchored_date and self.edit_date_from.date().toString("yyyy-MM-dd") == anchored_date:
                anchor_note = f" (anchored to existing corpus start {anchored_date})"

            tooltip_lines = [
                f"{c.label} ({c.key}): " + ", ".join(format_meeting_label(m) for m in c.meetings)
                for c in self.selected_cohorts
            ]
            self.lbl_cohorts_summary.setToolTip("\n".join(tooltip_lines))
            self.lbl_cohorts_summary.setText(
                f"Resolved: <b>{len(self.selected_cohorts)} Synchronized Meeting Round(s)</b> ({first_c} → {last_c}) "
                f"with <b>{total_mtgs} meetings</b> across {', '.join(sorted(self.active_topic.target_wgs))}{anchor_note}.<br/>"
                f"<span style='color: {C.TEXT_MUTED};'>Latest per WG: {', '.join(latest_parts)} (hover for all rounds)</span>"
            )
        else:
            self.lbl_cohorts_summary.setToolTip("")
            self.lbl_cohorts_summary.setText("No meetings found matching this timeframe and target WGs.")

    def _check_incremental_status(self):
        folder_str = self.edit_dest_folder.text().strip()
        self.edit_dest_folder.setToolTip(folder_str)
        if not folder_str:
            self.lbl_inc_status.setText("⚠️ Please specify a destination folder.")
            self.lbl_inc_status.setStyleSheet(f"color: {C.WARNING_DARK}; font-size: 11px;")
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
                syn_color = C.INFO_DARK
            else:
                syn_badge = "⏳ Stage-2 Pending"
                syn_color = C.TEXT_MUTED
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
                syn_color = C.WARNING_DARK
            else:
                syn_badge = f"✅ Stage-2: {dated_lbl}{folder_tag}{runs_tag}"
                syn_color = C.SUCCESS_DARK

        self.lbl_inc_status.setText(
            build_incremental_banner_html(
                total_mtgs=len(manifest.exported_meetings),
                dec_cnt=dec_cnt,
                doc_cnt=doc_cnt,
                stub_cnt=stub_cnt,
                analyzed_cnt=analyzed_cnt,
                stale_cnt=stale_cnt,
                pending_cnt=pending_cnt,
                no_act_cnt=no_act_cnt,
                syn_badge=syn_badge,
                syn_color=syn_color,
            )
        )
        self.lbl_inc_status.setStyleSheet(f"color: {C.TEXT_BODY}; font-size: 11px;")

    def _set_operation_busy(self, busy: bool, status_text: Optional[str] = None) -> None:
        self.btn_export.setEnabled((not busy) and (self.active_topic is not None))
        if busy:
            self.btn_check_status.setEnabled(False)
            self.prog_bar.setRange(0, 100)
            self.prog_bar.setValue(0)
        self.prog_bar.setVisible(busy)
        if hasattr(self, "btn_cancel_op"):
            self.btn_cancel_op.setVisible(busy)
            self.btn_cancel_op.setEnabled(busy)
        if status_text is not None:
            self.lbl_status.setText(status_text)

    def _apply_checkup_result(self, res: Dict[str, object], show_dialog: bool = False) -> None:
        self._check_incremental_status()
        if not res.get("ok"):
            return

        syn = str(res.get("synthesis_status") or "pending").lower()
        s2_state = res.get("synthesis_state") or {}
        latest_pkg = s2_state.get("latest_package") or {}

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

        if show_dialog:
            if getattr(self, "_status_report_dlg", None) is None:
                self._status_report_dlg = CorpusStatusReportDialog(parent=None)
            self._status_report_dlg.update_report(res)
            self._status_report_dlg.show()
            self._status_report_dlg.raise_()
            self._status_report_dlg.activateWindow()

    def _on_checkup_finished(self, res: Dict[str, object]) -> None:
        self._set_operation_busy(False)
        if res.get("cancelled"):
            self._check_incremental_status()
            self.lbl_status.setText("Checkup cancelled.")
            return
        if not res.get("ok"):
            self._check_incremental_status()
            self.lbl_status.setText("Checkup failed.")
            QMessageBox.warning(
                self,
                "Checkup Failed",
                str(res.get("error") or "Could not inspect destination corpus."),
            )
            return
        self._apply_checkup_result(res, show_dialog=True)

    def _on_check_status(self, silent: bool = False) -> Dict[str, object]:
        folder_str = self.edit_dest_folder.text().strip()
        if not folder_str:
            if not silent:
                QMessageBox.warning(self, "No Destination", "Please specify a destination folder first.")
            return {"ok": False, "error": "Empty destination"}

        if silent:
            res = check_and_refresh_dossier(Path(folder_str), topic=self.active_topic)
            if not res.get("ok"):
                self._check_incremental_status()
                return res
            self._apply_checkup_result(res, show_dialog=False)
            return res

        if (self._worker and self._worker.isRunning()) or (
            getattr(self, "_checkup_worker", None) and self._checkup_worker.isRunning()
        ):
            self.raise_()
            self.activateWindow()
            return {"ok": False, "error": "Another operation is already running."}

        self._set_operation_busy(True, "Scanning corpus manifest...")
        self._checkup_worker = TopicCheckupWorker(Path(folder_str), topic=self.active_topic)
        self._checkup_worker.progress_msg.connect(self.lbl_status.setText)
        self._checkup_worker.progress_step.connect(self.prog_bar.setValue)
        self._checkup_worker.checkup_finished.connect(self._on_checkup_finished)
        self._checkup_worker.start()
        return {"ok": True, "started": True}

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
