"""Filter helpers, worker execution, live filtering, and KPI refresh mixin."""

from typing import List

from PyQt5.QtCore import QDate, Qt
from PyQt5.QtWidgets import QMessageBox

from modules.meetings.core.agreement_tracker_worker import AgreementTrackerWorker
from modules.meetings.ui.contributions.contribution_widgets import FilterTokenChip


class _TrackerSearchMixin:
    """Manages filter tokens, watchlist preset buttons, async search, and KPI updates."""

    def set_active_wgs(self, wgs: List[str]):
        target_upper = {w.strip().upper() for w in (wgs or []) if w.strip()}
        for i in range(1, self.wg_filter.model().rowCount()):
            item = self.wg_filter.model().item(i)
            if item:
                item.setCheckState(
                    Qt.Checked if item.text().strip().upper() in target_upper else Qt.Unchecked
                )
        self.wg_filter.update()

    def set_date_range(self, d_from: QDate, d_to: QDate):
        if d_from and d_from.isValid():
            self.date_from.setDate(d_from)
        if d_to and d_to.isValid():
            self.date_to.setDate(d_to)

    def add_wi_tokens(self, tokens: List[str]):
        for tok in (tokens or []):
            clean = str(tok).strip()
            if clean and clean not in self.wi_token_widget.selected_tokens:
                self.wi_token_widget.selected_tokens.add(clean)
                chip = FilterTokenChip(clean, self.wi_token_widget)
                chip.removed.connect(self.wi_token_widget._remove_token)
                self.wi_token_widget.chips_layout.addWidget(chip)
        self.wi_token_widget.tokens_changed.emit()

    def _clear_wi_tokens(self):
        for tok in list(self.wi_token_widget.selected_tokens):
            self.wi_token_widget._remove_token(tok)

    def _load_followed_wis(self):
        followed = self.wi_settings.get_followed_wis()
        if not followed:
            QMessageBox.information(
                self, "No Followed WIs", "You have not marked any Work Items as Followed yet."
            )
            return
        self.add_wi_tokens(followed)

    def _load_starred_wis(self):
        starred = self.wi_settings.get_starred_wis()
        if not starred:
            QMessageBox.information(
                self, "No Starred WIs", "You have not starred any Work Items yet."
            )
            return
        self.add_wi_tokens(starred)

    # =========================================================================
    # Search Execution & Lifecycle
    # =========================================================================

    def _cleanup_threads(self):
        if self.worker and self.worker.isRunning():
            self.worker.cancel()
            self.worker.requestInterruption()
            self.worker.quit()
            if not self.worker.wait(1000):
                try:
                    self.worker.terminate()
                    self.worker.wait(500)
                except Exception:
                    pass
        for th in list(self.active_tdoc_threads.values()):
            if th and th.isRunning():
                th.requestInterruption()
                th.quit()
                if not th.wait(500):
                    try:
                        th.terminate()
                        th.wait(200)
                    except Exception:
                        pass
        self.active_tdoc_threads.clear()

    def _start_search(self):
        selected_wgs = self.wg_filter.getCheckedItems()
        target_wis = self.wi_token_widget.get_tokens()
        target_ais = self.ai_token_widget.get_tokens()
        keyword_query = self.keyword_edit.text().strip()
        wi_ai_or = self.chk_wi_ai_or.isChecked()
        date_from = self.date_from.date().toString("yyyy-MM-dd")
        date_to = self.date_to.date().toString("yyyy-MM-dd")

        if not (
            self.chk_agreements.isChecked()
            or self.chk_ffs.isChecked()
            or self.chk_agreed_tdocs.isChecked()
        ):
            QMessageBox.warning(
                self, "No Item Types Selected", "Please select at least one item type to include."
            )
            return

        summary_parts = []
        if selected_wgs:
            summary_parts.append("WG: " + ", ".join(selected_wgs))
        summary_parts.append(f"{date_from} → {date_to}")
        if target_wis:
            summary_parts.append("WI/SI: " + ", ".join(target_wis))
        if target_ais:
            summary_parts.append("AI: " + ", ".join(target_ais))
        if keyword_query:
            summary_parts.append(f"Keyword: '{keyword_query}'")
        self.filter_summary_lbl.setText("  •  ".join(summary_parts))

        self.btn_search.setEnabled(False)
        self.btn_cancel.setEnabled(True)
        self.btn_export_md.setEnabled(False)
        self.btn_export_excel.setEnabled(False)
        self.log_viewer.clear()
        self.progress_bar.setValue(0)

        self.worker = AgreementTrackerWorker(
            meetings_db=self.meetings_db,
            wgs=selected_wgs,
            date_from=date_from,
            date_to=date_to,
            target_wis=target_wis,
            target_ais=target_ais,
            keyword_query=keyword_query,
            wi_ai_or=wi_ai_or,
            include_agreements=self.chk_agreements.isChecked(),
            include_ffs=self.chk_ffs.isChecked(),
            include_agreed_tdocs=self.chk_agreed_tdocs.isChecked(),
            include_prev_recaps=self.chk_prev_recaps.isChecked(),
            auto_download_missing=self.chk_auto_download.isChecked(),
            bypass_cache=self.chk_bypass_cache.isChecked(),
            cache_dir=self.settings.cache_dir,
            parent=self,
        )
        self.worker.stage_progress.connect(self._on_stage_progress)
        self.worker.log_msg.connect(self._append_log)
        self.worker.results_ready.connect(self._on_results_ready)
        self.worker.finished.connect(self._on_search_finished)
        self.worker.start()

    def _cancel_search(self):
        if self.worker and self.worker.isRunning():
            self.worker.cancel()
            self.btn_cancel.setEnabled(False)

    def _on_stage_progress(self, desc: str, current: int, total: int):
        self.progress_bar.setMaximum(max(total, 1))
        self.progress_bar.setValue(current)
        self.setWindowTitle(
            f"[{current}/{total}] {desc} - 3GPP Study Item & Topic Agreement Tracker"
        )

    def _append_log(self, msg: str):
        self.log_viewer.append(msg)
        self.log_viewer.verticalScrollBar().setValue(self.log_viewer.verticalScrollBar().maximum())

    def _on_results_ready(self, results: List[dict]):
        self.table_model.update_data(results)
        self._refresh_kpis_and_counts()
        self._apply_column_widths()
        if self.table_model.rowCount() > 0:
            self.table_view.selectRow(0)

    def _on_live_filter_changed(self, text: str):
        self.table_model.set_live_filter(text)
        self._refresh_kpis_and_counts()
        if self.table_model.rowCount() > 0:
            self.table_view.selectRow(0)

    def _refresh_kpis_and_counts(self):
        rows = self.table_model.get_filtered_rows()
        all_count = len(self.table_model.get_all_rows())
        total = len(rows)

        if self.live_filter_edit.text().strip():
            self.lbl_count.setText(f"Showing {total} of {all_count} agreement(s).")
        else:
            self.lbl_count.setText(f"Found {total} matching agreement(s)/item(s).")

        if total == 0:
            self.kpi_total.set_value("0 Items", "No matches")
            self.kpi_agree.set_value("0", "0.0% of total")
            self.kpi_ffs.set_value("0", "0 open issues")
            self.kpi_tp.set_value("0", "0 documents")
            self.kpi_mtgs.set_value("0 Mtgs", "No meetings")
            self.btn_export_md.setEnabled(False)
            self.btn_export_excel.setEnabled(False)
            return

        wgs_set = {r.get("WG", "") for r in rows if r.get("WG")}
        mtgs_set = {r.get("Meeting", "") for r in rows if r.get("Meeting")}
        wis_set = {r.get("Work Item", "") for r in rows if r.get("Work Item")}

        agree_cnt = 0
        ffs_cnt = 0
        tp_cnt = 0
        for r in rows:
            k = str(r.get("Kind", "")).lower()
            if "ffs" in k or "open issue" in k:
                ffs_cnt += 1
            elif "agreed tdoc" in k or "pcr" in k:
                tp_cnt += 1
            else:
                agree_cnt += 1

        self.kpi_total.set_value(f"{total} Items", f"Across {len(wis_set)} WI(s)")
        self.kpi_agree.set_value(f"🟢 {agree_cnt}", "Agreements & WAs")
        self.kpi_ffs.set_value(f"🔵 {ffs_cnt}", "Open Issues / FFS")
        self.kpi_tp.set_value(f"📄 {tp_cnt}", "Agreed TPs / pCRs")
        self.kpi_mtgs.set_value(f"{len(mtgs_set)} Mtgs", f"In {', '.join(sorted(wgs_set))}")

        self.btn_export_md.setEnabled(True)
        self.btn_export_excel.setEnabled(True)

    def _on_search_finished(self, success: bool, msg: str):
        self.btn_search.setEnabled(True)
        self.btn_cancel.setEnabled(False)
        self.setWindowTitle("⚖️ 3GPP Cross-Meeting Study Item & Topic Agreement Tracker")
        if not success and "cancelled" not in msg.lower():
            QMessageBox.warning(self, "Tracker Notice", msg)
