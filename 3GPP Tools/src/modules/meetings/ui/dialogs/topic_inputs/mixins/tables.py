"""Table population mixin for Decisions, Contributions, Liaisons, and WG Matrix tabs."""

from typing import Any, Dict, List

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor
from PyQt5.QtWidgets import QHeaderView, QTableWidgetItem

from core.ui.theme import C
from core.utils.company_sanitizer import CompanySanitizer
from modules.meetings.core.topic_decision import TopicDecision
from modules.meetings.ui.dialogs.topic_inputs.styles import BADGE_COLORS


class TopicInputsTablesMixin:
    """Populates the Decisions, All Contributions, LS, and WG Distribution tables."""

    def _populate_decisions_table(self, decisions: List[TopicDecision]):
        self.table_decisions.setUpdatesEnabled(False)
        self.table_decisions.blockSignals(True)
        try:
            self._displayed_decisions = list(decisions)
            self.table_decisions.setRowCount(len(decisions))
            self.table_decisions.setColumnCount(8)
            self.table_decisions.setHorizontalHeaderLabels([
                "Decision Type", "WG", "Meeting", "Date", "TDoc ID", "Decision / Substantive Scope", "Target Spec", "Agenda Item"
            ])

            header = self.table_decisions.horizontalHeader()
            for i in range(8):
                header.setSectionResizeMode(i, QHeaderView.Interactive)

            for row, d in enumerate(decisions):
                # 0: Decision Type Badge & Provenance
                if d.decision_type in ("Text Agreement", "Working Assumption"):
                    badge_text = f"🟢 {d.decision_type} (Notes)"
                    tip_text = (
                        f"<b>Binding Consensus Text</b> (Chairman's Notes AI {d.ai or 'N/A'})<br/>"
                        f"<b>Meeting:</b> {d.wg} #{d.meeting_number} ({d.date or 'N/A'})<br/>"
                        f"<b>Background Proposal:</b> {d.tdoc_id or 'None'} (Noted/Discussed)<br/>"
                        f"<i>In RAN WGs, only the verbatim text in Chairman's Notes is binding consensus.</i>"
                    )
                elif d.decision_type in ("Approved pCR", "Approved CR"):
                    badge_text = f"🟢 {d.decision_type} (Doc)"
                    tip_text = (
                        f"<b>Approved Specification Change</b> ({d.tdoc_id or 'N/A'})<br/>"
                        f"<b>Meeting:</b> {d.wg} #{d.meeting_number} ({d.date or 'N/A'})<br/>"
                        f"<b>Target Specification:</b> {d.target_spec or 'N/A'}<br/>"
                        f"<i>Full document attachment approved as baseline change.</i>"
                    )
                else:
                    badge_text = f"{d.badge_icon} {d.decision_type}"
                    tip_text = f"<b>{d.decision_type}</b> ({d.wg} #{d.meeting_number})<br/><b>AI:</b> {d.ai or 'N/A'}"

                badge_item = QTableWidgetItem(badge_text)
                badge_item.setTextAlignment(Qt.AlignCenter)
                badge_item.setToolTip(tip_text)
                bg_color, fg_color = BADGE_COLORS.get(d.decision_type, (C.BG_CANVAS, C.TEXT_STRONG))
                badge_item.setBackground(QColor(bg_color))
                badge_item.setForeground(QColor(fg_color))
                font = badge_item.font()
                font.setBold(True)
                badge_item.setFont(font)
                self.table_decisions.setItem(row, 0, badge_item)

                # 1: WG
                wg_item = QTableWidgetItem(d.wg)
                wg_item.setTextAlignment(Qt.AlignCenter)
                self.table_decisions.setItem(row, 1, wg_item)

                # 2: Meeting #
                mtg_item = QTableWidgetItem(f"#{d.meeting_number}")
                mtg_item.setTextAlignment(Qt.AlignCenter)
                self.table_decisions.setItem(row, 2, mtg_item)

                # 3: Date
                date_item = QTableWidgetItem(d.date)
                date_item.setTextAlignment(Qt.AlignCenter)
                self.table_decisions.setItem(row, 3, date_item)

                # 4: TDoc ID
                tid_str = d.tdoc_id or "—"
                tid_item = QTableWidgetItem(tid_str)
                tid_item.setTextAlignment(Qt.AlignCenter)
                if d.tdoc_id:
                    tid_item.setForeground(QColor(C.ACCENT_SKY_DARK))
                    f = tid_item.font()
                    f.setBold(True)
                    tid_item.setFont(f)
                self.table_decisions.setItem(row, 4, tid_item)

                # 5: Decision / Substantive Scope
                scope_item = QTableWidgetItem(d.title_or_summary)
                scope_item.setToolTip(f"<b>{d.decision_type} ({d.wg} #{d.meeting_number})</b><br/>{d.title_or_summary}")
                self.table_decisions.setItem(row, 5, scope_item)

                # 6: Target Spec
                spec_item = QTableWidgetItem(d.target_spec or "—")
                spec_item.setTextAlignment(Qt.AlignCenter)
                self.table_decisions.setItem(row, 6, spec_item)

                # 7: Agenda Item
                ai_item = QTableWidgetItem(d.ai or "—")
                ai_item.setTextAlignment(Qt.AlignCenter)
                self.table_decisions.setItem(row, 7, ai_item)

            self.table_decisions.resizeColumnsToContents()
            header.setSectionResizeMode(5, QHeaderView.Stretch)
        finally:
            self.table_decisions.blockSignals(False)
            self.table_decisions.setUpdatesEnabled(True)

        if decisions:
            self.table_decisions.selectRow(0)
        else:
            if self.tabs.currentIndex() == 0:
                self.detail_text_edit.clear()
                self.lbl_drawer_provenance.setVisible(False)
                self._current_detail_text = ""

    def _populate_contributions_table(self, contribs: List[Dict[str, Any]]):
        self.table_contributions.setUpdatesEnabled(False)
        self.table_contributions.blockSignals(True)
        try:
            self._displayed_contribs = list(contribs)
            self.table_contributions.setRowCount(len(contribs))
            self.table_contributions.setColumnCount(9)
            self.table_contributions.setHorizontalHeaderLabels([
                "TDoc ID", "WG", "Meeting", "Date", "Agenda Item", "Type", "Status", "Title", "Source Company"
            ])

            header = self.table_contributions.horizontalHeader()
            for i in range(9):
                header.setSectionResizeMode(i, QHeaderView.Interactive)

            for row, c in enumerate(contribs):
                # 0: TDoc ID
                tid_item = QTableWidgetItem(c.get("tdoc_id", ""))
                tid_item.setTextAlignment(Qt.AlignCenter)
                tid_item.setForeground(QColor(C.ACCENT_SKY_DARK))
                f = tid_item.font()
                f.setBold(True)
                tid_item.setFont(f)
                self.table_contributions.setItem(row, 0, tid_item)

                # 1: WG
                self.table_contributions.setItem(row, 1, QTableWidgetItem(c.get("wg", "")))
                # 2: Meeting #
                self.table_contributions.setItem(row, 2, QTableWidgetItem(f"#{c.get('meeting_number', '')}"))
                # 3: Date
                self.table_contributions.setItem(row, 3, QTableWidgetItem(c.get("date", "")))

                # 4: AI
                ai_item = QTableWidgetItem(c.get("ai", ""))
                ai_item.setToolTip(c.get("ai_desc", ""))
                self.table_contributions.setItem(row, 4, ai_item)

                # 5: Type
                t_type = c.get("type", "")
                t_item = QTableWidgetItem(t_type)
                t_item.setTextAlignment(Qt.AlignCenter)
                bg, fg = BADGE_COLORS.get(t_type, (C.BG_CARD, C.TEXT_BODY))
                t_item.setBackground(QColor(bg))
                t_item.setForeground(QColor(fg))
                self.table_contributions.setItem(row, 5, t_item)

                # 6: Status
                status_item = QTableWidgetItem(c.get("status", ""))
                status_item.setTextAlignment(Qt.AlignCenter)
                if any(s in c.get("status", "").lower() for s in ("approved", "agreed", "endorsed")):
                    status_item.setForeground(QColor(C.SUCCESS))
                    f_s = status_item.font()
                    f_s.setBold(True)
                    status_item.setFont(f_s)
                self.table_contributions.setItem(row, 6, status_item)

                # 7: Title
                title_item = QTableWidgetItem(c.get("title", ""))
                title_item.setToolTip(c.get("title", ""))
                self.table_contributions.setItem(row, 7, title_item)

                # 8: Source Company
                clean_co = c.get("clean_source") or CompanySanitizer.get_confirmed_source_text(c.get("source", "")) or c.get("source", "")
                co_item = QTableWidgetItem(clean_co)
                co_item.setToolTip(c.get("source", ""))
                self.table_contributions.setItem(row, 8, co_item)

            self.table_contributions.resizeColumnsToContents()
            header.setSectionResizeMode(7, QHeaderView.Stretch)
        finally:
            self.table_contributions.blockSignals(False)
            self.table_contributions.setUpdatesEnabled(True)

        if contribs and self.tabs.currentIndex() == 1:
            self.table_contributions.selectRow(0)

    def _populate_ls_table(self, ls_contribs: List[Dict[str, Any]]):
        self.table_ls.setUpdatesEnabled(False)
        self.table_ls.blockSignals(True)
        try:
            self._displayed_ls = list(ls_contribs)
            self.table_ls.setRowCount(len(ls_contribs))
            self.table_ls.setColumnCount(8)
            self.table_ls.setHorizontalHeaderLabels([
                "Direction", "TDoc ID", "WG", "Meeting", "Title", "Status", "Reply To", "Reply In"
            ])

            header = self.table_ls.horizontalHeader()
            for i in range(8):
                header.setSectionResizeMode(i, QHeaderView.Interactive)

            for row, c in enumerate(ls_contribs):
                raw = c.get("raw_row", {})
                t_type = c.get("type", "")
                is_in = "in" in t_type.lower()
                dir_text = "📥 LS IN" if is_in else "📤 LS OUT"
                dir_item = QTableWidgetItem(dir_text)
                dir_item.setTextAlignment(Qt.AlignCenter)
                bg, fg = BADGE_COLORS.get("LS in" if is_in else "LS out", (C.VIOLET_SOFT, C.PURPLE_DARK))
                dir_item.setBackground(QColor(bg))
                dir_item.setForeground(QColor(fg))
                f = dir_item.font()
                f.setBold(True)
                dir_item.setFont(f)
                self.table_ls.setItem(row, 0, dir_item)

                # TDoc
                tid_item = QTableWidgetItem(c.get("tdoc_id", ""))
                tid_item.setTextAlignment(Qt.AlignCenter)
                tid_item.setForeground(QColor(C.ACCENT_SKY_DARK))
                self.table_ls.setItem(row, 1, tid_item)

                self.table_ls.setItem(row, 2, QTableWidgetItem(c.get("wg", "")))
                self.table_ls.setItem(row, 3, QTableWidgetItem(f"#{c.get('meeting_number', '')}"))
                self.table_ls.setItem(row, 4, QTableWidgetItem(c.get("title", "")))
                self.table_ls.setItem(row, 5, QTableWidgetItem(c.get("status", "")))
                self.table_ls.setItem(row, 6, QTableWidgetItem(raw.get("Reply to", "") or raw.get("reply_to", "") or "—"))
                self.table_ls.setItem(row, 7, QTableWidgetItem(raw.get("Reply in", "") or raw.get("reply_in", "") or "—"))

            self.table_ls.resizeColumnsToContents()
            header.setSectionResizeMode(4, QHeaderView.Stretch)
        finally:
            self.table_ls.blockSignals(False)
            self.table_ls.setUpdatesEnabled(True)

        if ls_contribs and self.tabs.currentIndex() == 2:
            self.table_ls.selectRow(0)

    def _populate_distribution_table(self):
        wgs = sorted(list(self.topic.target_wgs))
        rounds = self.rounds

        self.table_dist.setRowCount(len(wgs) + 1)
        self.table_dist.setColumnCount(len(rounds) + 2)

        col_labels = ["Working Group"] + [r.label for r in rounds] + ["Total Across Scope"]
        self.table_dist.setHorizontalHeaderLabels(col_labels)

        header = self.table_dist.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        for i in range(1, len(rounds) + 1):
            header.setSectionResizeMode(i, QHeaderView.Stretch)
        header.setSectionResizeMode(len(rounds) + 1, QHeaderView.ResizeToContents)

        round_totals = {r.key: 0 for r in rounds}
        grand_total = 0

        for r_idx, wg in enumerate(wgs):
            wg_item = QTableWidgetItem(f"👥 {wg}")
            f = wg_item.font()
            f.setBold(True)
            wg_item.setFont(f)
            self.table_dist.setItem(r_idx, 0, wg_item)

            wg_row_total = 0
            for c_idx, m_round in enumerate(rounds, start=1):
                count = sum(1 for c in self.all_contributions if c.get("wg") == wg and c.get("round_key") == m_round.key)
                dec_cnt = sum(1 for d in self.all_decisions if d.wg == wg and d.cohort_key == m_round.key)
                text = f"{count} docs ({dec_cnt} dec)" if count or dec_cnt else "—"
                item = QTableWidgetItem(text)
                item.setTextAlignment(Qt.AlignCenter)
                self.table_dist.setItem(r_idx, c_idx, item)
                wg_row_total += count
                round_totals[m_round.key] += count

            tot_item = QTableWidgetItem(str(wg_row_total))
            tot_item.setTextAlignment(Qt.AlignCenter)
            f_tot = tot_item.font()
            f_tot.setBold(True)
            tot_item.setFont(f_tot)
            self.table_dist.setItem(r_idx, len(rounds) + 1, tot_item)
            grand_total += wg_row_total

        # Total Summary Row
        last_row = len(wgs)
        tot_label = QTableWidgetItem("Total Documents")
        f_t = tot_label.font()
        f_t.setBold(True)
        tot_label.setFont(f_t)
        self.table_dist.setItem(last_row, 0, tot_label)

        for c_idx, m_round in enumerate(rounds, start=1):
            r_tot = QTableWidgetItem(str(round_totals[m_round.key]))
            r_tot.setTextAlignment(Qt.AlignCenter)
            f_r = r_tot.font()
            f_r.setBold(True)
            r_tot.setFont(f_r)
            self.table_dist.setItem(last_row, c_idx, r_tot)

        grand_item = QTableWidgetItem(str(grand_total))
        grand_item.setTextAlignment(Qt.AlignCenter)
        f_g = grand_item.font()
        f_g.setBold(True)
        grand_item.setFont(f_g)
        self.table_dist.setItem(last_row, len(rounds) + 1, grand_item)
