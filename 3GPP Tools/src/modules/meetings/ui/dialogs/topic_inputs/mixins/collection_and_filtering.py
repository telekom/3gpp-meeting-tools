"""Worker collection callbacks, KPI updates, and pill/search filtering mixin."""

from typing import Any, Dict, List

from PyQt5.QtWidgets import QMessageBox

from modules.meetings.core.cross_wg_ls_tracker import CrossWgLiaisonTracker
from modules.meetings.core.topic_decision import TopicDecision
from modules.meetings.ui.dialogs.topic_inputs.collector_thread import InputsCollectorThread


class TopicInputsCollectionFilteringMixin:
    """Manages background scanning lifecycle, KPI counters, and interactive filters."""

    def _start_collection(self):
        if self._collector_worker and self._collector_worker.isRunning():
            self._collector_worker.cancel()
            self._collector_worker.wait(500)

        self.prog_frame.setVisible(True)
        self.prog_bar.setValue(0)
        self.lbl_prog_status.setText(f"Scanning meeting records across {len(self.rounds)} rounds...")
        self.btn_refresh.setEnabled(False)

        self._collector_worker = InputsCollectorThread(
            meetings_db=self.meetings_db,
            topic=self.topic,
            rounds=self.rounds,
            scope_mode=self.scope_mode,
            parent=None,
        )
        self._collector_worker.progress_status.connect(self._on_scan_progress)
        self._collector_worker.collection_finished.connect(self._on_scan_finished)
        self._collector_worker.collection_failed.connect(self._on_scan_failed)
        self._collector_worker.start()

    def _on_scan_progress(self, msg: str, cur_idx: int, total_count: int):
        self.lbl_prog_status.setText(msg)
        pct = int((cur_idx / max(1, total_count)) * 100)
        self.prog_bar.setValue(pct)

    def _on_scan_finished(self, decisions: List[TopicDecision], contributions: List[Dict[str, Any]], ls_tracker: CrossWgLiaisonTracker):
        self.all_decisions = decisions
        self.all_contributions = contributions
        self.ls_tracker = ls_tracker

        self.prog_frame.setVisible(False)
        self.btn_refresh.setEnabled(True)

        self._update_kpi_strip()
        self._apply_filters()
        self._populate_distribution_table()

    def _on_scan_failed(self, error_msg: str):
        self.prog_frame.setVisible(False)
        self.btn_refresh.setEnabled(True)
        QMessageBox.warning(self, "Collection Warning", f"Could not complete scan:\n{error_msg}")

    def _on_wg_pill_toggled(self, wg: str, checked: bool):
        if checked:
            idx = self.combo_wg_filter.findText(wg)
            if idx >= 0:
                self.combo_wg_filter.blockSignals(True)
                self.combo_wg_filter.setCurrentIndex(idx)
                self.combo_wg_filter.blockSignals(False)
            self._apply_filters()

    def _on_cat_pill_toggled(self, cat_key: str, checked: bool):
        if checked:
            idx = self.combo_type_filter.findText(cat_key)
            if idx >= 0:
                self.combo_type_filter.blockSignals(True)
                self.combo_type_filter.setCurrentIndex(idx)
                self.combo_type_filter.blockSignals(False)

            # Auto-switch to matching view for fast navigation
            if cat_key in ("🟢 Decisions Only (Agreements, pCRs, CRs)", "Text Agreements Only", "Approved pCRs Only", "Approved CRs Only"):
                if self.tabs.currentIndex() != 0:
                    self.tabs.setCurrentIndex(0)
            elif cat_key == "📬 Liaison Statements Only":
                if self.tabs.currentIndex() != 2:
                    self.tabs.setCurrentIndex(2)
            elif cat_key == "📄 Discussion Papers Only":
                if self.tabs.currentIndex() != 1:
                    self.tabs.setCurrentIndex(1)

            self._apply_filters()

    def _reset_filters(self):
        self.edit_filter.clear()
        self.combo_round_filter.setCurrentIndex(0)
        if hasattr(self, "wg_pills") and "All WGs" in self.wg_pills:
            self.wg_pills["All WGs"].setChecked(True)
        if hasattr(self, "cat_pills") and "All Inputs" in self.cat_pills:
            self.cat_pills["All Inputs"].setChecked(True)
        self._apply_filters()

    def _update_kpi_strip(self):
        tot_cnt = len(self.all_contributions)
        dec_cnt = len(self.all_decisions)
        ls_cnt = sum(1 for c in self.all_contributions if "ls" in c.get("type", "").lower())
        disc_cnt = sum(1 for c in self.all_contributions if "discussion" in c.get("type", "").lower())
        wgs_seen = set(c.get("wg") for c in self.all_contributions if c.get("wg"))

        self.kpi_total.setText(f"📊 Total Inputs: <b>{tot_cnt}</b>")
        self.kpi_decisions.setText(f"🟢 Decisions: <b>{dec_cnt}</b>")
        self.kpi_ls.setText(f"📬 Liaisons: <b>{ls_cnt}</b>")
        self.kpi_disc.setText(f"📄 Discussions: <b>{disc_cnt}</b>")
        self.kpi_wgs.setText(f"👥 WGs: <b>{len(wgs_seen)}</b>")

        self.tabs.setTabText(0, f"⚖️ Technical Decisions && Agreements ({dec_cnt})")
        self.tabs.setTabText(1, f"📄 All Considered Contributions ({tot_cnt})")
        self.tabs.setTabText(2, f"📬 Inter-WG Liaison Statements ({ls_cnt})")

        # Update WG pill counts
        if hasattr(self, "wg_pills"):
            if "All WGs" in self.wg_pills:
                self.wg_pills["All WGs"].set_count(tot_cnt)
            for wg, pill in self.wg_pills.items():
                if wg == "All WGs":
                    continue
                cnt = sum(1 for c in self.all_contributions if str(c.get("wg", "")).upper() == wg.upper())
                pill.set_count(cnt)

        # Update Category pill counts
        if hasattr(self, "cat_pills"):
            if "All Inputs" in self.cat_pills:
                self.cat_pills["All Inputs"].set_count(tot_cnt)
            if "🟢 Decisions Only (Agreements, pCRs, CRs)" in self.cat_pills:
                self.cat_pills["🟢 Decisions Only (Agreements, pCRs, CRs)"].set_count(dec_cnt)
            if "Text Agreements Only" in self.cat_pills:
                self.cat_pills["Text Agreements Only"].set_count(sum(1 for d in self.all_decisions if d.decision_type == "Text Agreement"))
            if "Approved pCRs Only" in self.cat_pills:
                self.cat_pills["Approved pCRs Only"].set_count(sum(1 for d in self.all_decisions if d.decision_type == "Approved pCR"))
            if "Approved CRs Only" in self.cat_pills:
                self.cat_pills["Approved CRs Only"].set_count(sum(1 for d in self.all_decisions if d.decision_type == "Approved CR"))
            if "📬 Liaison Statements Only" in self.cat_pills:
                self.cat_pills["📬 Liaison Statements Only"].set_count(ls_cnt)
            if "📄 Discussion Papers Only" in self.cat_pills:
                self.cat_pills["📄 Discussion Papers Only"].set_count(disc_cnt)

    def _apply_filters(self):
        query = self.edit_filter.text().strip().lower()
        selected_wg = self.combo_wg_filter.currentText()
        selected_cat = self.combo_type_filter.currentText()
        selected_round = self.combo_round_filter.currentText()

        # 1. Filter Decisions
        filtered_decs = []
        for d in self.all_decisions:
            if selected_wg != "All WGs" and d.wg.upper() != selected_wg.upper():
                continue
            if selected_round != "All Meeting Rounds":
                # Check cohort match
                matching_round = next((r for r in self.rounds if r.key == d.cohort_key), None)
                if not matching_round or matching_round.label != selected_round:
                    continue
            if selected_cat == "Text Agreements Only" and d.decision_type != "Text Agreement":
                continue
            if selected_cat == "Approved pCRs Only" and d.decision_type != "Approved pCR":
                continue
            if selected_cat == "Approved CRs Only" and d.decision_type != "Approved CR":
                continue
            if selected_cat == "📬 Liaison Statements Only" and d.decision_type != "Agreed LS":
                continue
            if selected_cat == "📄 Discussion Papers Only":
                continue

            if query:
                haystack = f"{d.tdoc_id or ''} {d.title_or_summary} {d.source} {d.ai} {d.target_spec}".lower()
                if query not in haystack:
                    continue
            filtered_decs.append(d)

        self._populate_decisions_table(filtered_decs)

        # 2. Filter Contributions
        filtered_contribs = []
        for c in self.all_contributions:
            if selected_wg != "All WGs" and c.get("wg", "").upper() != selected_wg.upper():
                continue
            if selected_round != "All Meeting Rounds" and c.get("round_label") != selected_round:
                continue
            if selected_cat == "🟢 Decisions Only (Agreements, pCRs, CRs)" and not c.get("is_decision"):
                continue
            if selected_cat == "Approved pCRs Only" and c.get("decision_type") != "Approved pCR":
                continue
            if selected_cat == "Approved CRs Only" and c.get("decision_type") != "Approved CR":
                continue
            if selected_cat == "Text Agreements Only" and c.get("decision_type") != "Text Agreement":
                continue
            if selected_cat == "📬 Liaison Statements Only" and "ls" not in c.get("type", "").lower():
                continue
            if selected_cat == "📄 Discussion Papers Only" and "discussion" not in c.get("type", "").lower():
                continue

            if query:
                haystack = f"{c.get('tdoc_id', '')} {c.get('title', '')} {c.get('source', '')} {c.get('ai', '')} {c.get('ai_desc', '')}".lower()
                if query not in haystack:
                    continue
            filtered_contribs.append(c)

        self._populate_contributions_table(filtered_contribs)

        # 3. Filter Liaison Statements
        filtered_ls = [c for c in filtered_contribs if "ls" in c.get("type", "").lower()]
        self._populate_ls_table(filtered_ls)
