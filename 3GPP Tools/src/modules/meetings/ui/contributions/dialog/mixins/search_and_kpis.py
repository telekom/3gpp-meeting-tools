"""Search orchestration, progress updates, and KPI calculations for `ContributionReportDialog`."""

from __future__ import annotations

import re

from PyQt5.QtWidgets import QMessageBox

from core.utils.company_sanitizer import CompanySanitizer
from modules.meetings.core.contribution_worker import ContributionSearchWorker


class ContributionDialogSearchMixin:
    """Mixin handling async contribution search execution and KPI ribbon aggregation."""

    def _start_search(self):
        selected_wgs = self.wg_filter.getCheckedItems()
        target_companies = self.company_widget.get_selected_companies()
        primary_only = self.company_widget.is_primary_only()
        include_tentative = self.company_widget.is_include_tentative()
        target_wis = self.wi_token_widget.get_tokens()
        target_ais = self.ai_token_widget.get_tokens()
        wi_ai_or = self.chk_wi_ai_or.isChecked()
        date_from = self.date_from.date().toString("yyyy-MM-dd")
        date_to = self.date_to.date().toString("yyyy-MM-dd")
        bypass = self.chk_bypass_cache.isChecked()

        summary_parts = []
        if selected_wgs:
            summary_parts.append("WG: " + ", ".join(selected_wgs))
        summary_parts.append(f"{date_from} → {date_to}")
        if target_companies:
            company_rule = "primary" if primary_only else "any co-author"
            if include_tentative:
                company_rule += ", incl. ?/[]"
            summary_parts.append("Companies: " + ", ".join(sorted(target_companies)) + f" ({company_rule})")
        wi_text = ", ".join(sorted(target_wis)) if target_wis else ""
        ai_text = ", ".join(sorted(target_ais)) if target_ais else ""
        if wi_text and ai_text:
            joiner = " OR " if wi_ai_or else " AND "
            summary_parts.append(f"WI: {wi_text}{joiner}AI: {ai_text}")
        elif wi_text:
            summary_parts.append(f"WI: {wi_text}")
        elif ai_text:
            summary_parts.append(f"AI: {ai_text}")
        self.filter_summary_lbl.setText("  •  ".join(summary_parts) if summary_parts else "All contributions")

        self.btn_search.setEnabled(False)
        self.btn_cancel.setEnabled(True)
        self.btn_export.setEnabled(False)
        self.btn_stats.setEnabled(False)
        self.log_viewer.clear()
        self.progress_bar.setValue(0)

        self.worker = ContributionSearchWorker(
            meetings_db=self.meetings_db,
            wgs=selected_wgs,
            date_from=date_from,
            date_to=date_to,
            target_companies=target_companies,
            primary_only=primary_only,
            target_wis=target_wis,
            target_ais=target_ais,
            wi_ai_or=wi_ai_or,
            bypass_cache=bypass,
            cache_dir=self.settings.cache_dir,
            include_tentative=include_tentative,
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
        self.setWindowTitle(f"[{current}/{total}] {desc} - 3GPP Contributions")

    def _append_log(self, msg: str):
        self.log_viewer.append(msg)
        self.log_viewer.verticalScrollBar().setValue(self.log_viewer.verticalScrollBar().maximum())

    def _on_results_ready(self, results: list):
        self.table_model.update_data(results)
        total = len(results)
        self.lbl_count.setText(f"Found {total} matching contribution(s).")
        self._apply_column_widths()

        if total == 0:
            self.kpi_total.set_value("0 TDocs", "No matches")
            self.kpi_agree.set_value("0 TDocs", "0.0% of total")
            self.kpi_joint.set_value("0% Joint", "0 joint")
            self.kpi_partner.set_value("None", "No allies")
            self.kpi_wi.set_value("None", "No WIs")
            self.btn_export.setEnabled(False)
            self.btn_stats.setEnabled(False)
            return

        wgs_count = len(set(r.get("WG", "") for r in results if r.get("WG")))
        self.kpi_total.set_value(f"{total} TDocs", f"Across {wgs_count} WG(s)")

        agreed_count = sum(
            1 for r in results if any(w in str(r.get("TDoc Status", "")).lower() for w in ["agreed", "approved"])
        )
        agree_pct = round((agreed_count / total) * 100, 1) if total else 0.0
        self.kpi_agree.set_value(f"{agreed_count} TDocs", f"{agree_pct}% of {total} total")

        joint_count = 0
        partner_counts = {}
        target_companies = self.company_widget.get_selected_companies()

        for r in results:
            src = str(r.get("Source", ""))
            contribs = CompanySanitizer.get_matching_contributors(src)
            if len(contribs) > 1:
                joint_count += 1

            if target_companies:
                targets_in_row = [c for c in contribs if c in target_companies]
                if targets_in_row:
                    for c in contribs:
                        if c not in target_companies:
                            partner_counts[c] = partner_counts.get(c, 0) + 1
            else:
                for c in contribs:
                    partner_counts[c] = partner_counts.get(c, 0) + 1

        joint_pct = round((joint_count / total) * 100, 1)
        solo_count = total - joint_count
        self.kpi_joint.set_value(f"{joint_pct}% Joint", f"{solo_count} solo, {joint_count} joint")

        if partner_counts:
            top_partner, top_count = max(partner_counts.items(), key=lambda x: x[1])
            self.kpi_partner.set_value(top_partner, f"{top_count} joint TDocs")
        else:
            self.kpi_partner.set_value("N/A", "No co-signers")

        wi_counts = {}
        for r in results:
            wi_str = str(r.get("Related WIs", "")).strip()
            if wi_str and wi_str.lower() not in ["", "none", "unknown", "-", "dummy"]:
                for item in re.split(r"[,;]+", wi_str):
                    item_clean = item.strip()
                    if item_clean and item_clean.lower() != "dummy":
                        wi_counts[item_clean] = wi_counts.get(item_clean, 0) + 1

        if wi_counts:
            top_wi, wi_vol = max(wi_counts.items(), key=lambda x: x[1])
            disp_wi = top_wi if len(top_wi) <= 18 else top_wi[:16] + ".."
            self.kpi_wi.set_value(disp_wi, f"{wi_vol} TDocs")
        else:
            self.kpi_wi.set_value("Unspecified", "No WI tagged")

        self.btn_export.setEnabled(True)
        self.btn_stats.setEnabled(True)

    def _on_search_finished(self, success: bool, msg: str):
        self.btn_search.setEnabled(True)
        self.btn_cancel.setEnabled(False)
        self.setWindowTitle("3GPP Company Contribution Search & Report")
        if not success:
            QMessageBox.warning(self, "Search Notice", msg)
