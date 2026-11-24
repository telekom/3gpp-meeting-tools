"""Filter dropdowns, search matching, table population, and detail drawer mixin."""

from __future__ import annotations

import html

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QBrush, QColor
from PyQt5.QtWidgets import QTableWidgetItem

from core.ui.theme import C


class SpecTDocReportTableMixin:
    """Manages filter combo population, text/facet filtering, table rendering, and detail drawer."""

    def _populate_filters(self):
        """Populates the Company, Type, and Status dropdowns."""
        self.company_combo.blockSignals(True)
        self.type_combo.blockSignals(True)
        self.status_combo.blockSignals(True)

        self.company_combo.clear()
        self.type_combo.clear()
        self.status_combo.clear()

        # Companies
        companies = sorted(
            {r.get("source", "").strip() for r in self.all_entries if r.get("source", "").strip()}
        )
        self.company_combo.addItem("All Companies", "")
        for comp in companies:
            cnt = sum(1 for r in self.all_entries if r.get("source", "").strip() == comp)
            self.company_combo.addItem(f"{comp} ({cnt})", comp)

        # Types
        types = sorted(
            {r.get("type", "").strip() for r in self.all_entries if r.get("type", "").strip()}
        )
        self.type_combo.addItem("All Types", "")
        for t in types:
            cnt = sum(1 for r in self.all_entries if r.get("type", "").strip() == t)
            self.type_combo.addItem(f"{t} ({cnt})", t)

        # Statuses
        statuses = sorted(
            {
                r.get("status", "").strip().lower()
                for r in self.all_entries
                if r.get("status", "").strip()
            }
        )
        self.status_combo.addItem("All Statuses", "")
        for s in statuses:
            cnt = sum(1 for r in self.all_entries if r.get("status", "").strip().lower() == s)
            self.status_combo.addItem(f"{s.capitalize()} ({cnt})", s)

        self.company_combo.blockSignals(False)
        self.type_combo.blockSignals(False)
        self.status_combo.blockSignals(False)

    def _apply_filters(self):
        """Filters the entries based on search input and filter combos."""
        q = self.search_input.text().strip().lower()
        sel_comp = self.company_combo.currentData() or ""
        sel_type = self.type_combo.currentData() or ""
        sel_status = self.status_combo.currentData() or ""

        filtered = []
        for r in self.all_entries:
            if sel_comp and r.get("source", "").strip() != sel_comp:
                continue
            if sel_type and r.get("type", "").strip() != sel_type:
                continue
            if sel_status and r.get("status", "").strip().lower() != sel_status:
                continue

            if q:
                tokens = [
                    r.get("tdoc", ""),
                    r.get("title", ""),
                    r.get("source", ""),
                    r.get("agenda_item", ""),
                    r.get("ai_desc", ""),
                    r.get("abstract", ""),
                    r.get("wi", ""),
                    r.get("remarks", ""),
                ]
                match_text = " ".join(t.lower() for t in tokens if t)
                if not all(term in match_text for term in q.split()):
                    continue

            filtered.append(r)

        self.filtered_entries = filtered
        self._refresh_table()

    def _refresh_table(self):
        """Re-populates the table widget with filtered records."""
        self.table.setSortingEnabled(False)
        self.table.setRowCount(0)

        for row_idx, r in enumerate(self.filtered_entries):
            self.table.insertRow(row_idx)

            # 0. TDoc ID
            tdoc_str = r.get("tdoc", "")
            tdoc_item = QTableWidgetItem(tdoc_str)
            f_tdoc = tdoc_item.font()
            f_tdoc.setBold(True)
            tdoc_item.setFont(f_tdoc)
            tdoc_item.setForeground(QBrush(QColor(C.ACCENT_SKY_DARK)))
            tdoc_item.setData(Qt.UserRole, r)
            self.table.setItem(row_idx, 0, tdoc_item)

            # 1. Type
            type_str = r.get("type") or "pCR"
            type_item = QTableWidgetItem(type_str)
            type_item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row_idx, 1, type_item)

            # 2. Status
            status_str = r.get("status") or "—"
            status_item = QTableWidgetItem(status_str.capitalize())
            status_item.setTextAlignment(Qt.AlignCenter)
            if "approv" in status_str or "agree" in status_str:
                status_item.setForeground(QBrush(QColor(C.SUCCESS_DARK)))
            elif "noted" in status_str:
                status_item.setForeground(QBrush(QColor(C.ACCENT_SKY_DARK)))
            elif "postpone" in status_str or "revise" in status_str:
                status_item.setForeground(QBrush(QColor(C.WARNING)))
            self.table.setItem(row_idx, 2, status_item)

            # 3. Source (Company)
            source_item = QTableWidgetItem(r.get("source", ""))
            self.table.setItem(row_idx, 3, source_item)

            # 4. Agenda Item
            ai_item = QTableWidgetItem(r.get("agenda_item", ""))
            ai_item.setTextAlignment(Qt.AlignCenter)
            if r.get("ai_desc"):
                ai_item.setToolTip(f"AI {r.get('agenda_item')}: {r.get('ai_desc')}")
            self.table.setItem(row_idx, 4, ai_item)

            # 5. Title
            title_item = QTableWidgetItem(r.get("title", ""))
            self.table.setItem(row_idx, 5, title_item)

            # 6. Abstract / Remarks Preview
            abs_text = r.get("abstract") or r.get("remarks") or ""
            clean_abs = " ".join(abs_text.split())
            abs_item = QTableWidgetItem(clean_abs)
            abs_item.setToolTip(abs_text)
            self.table.setItem(row_idx, 6, abs_item)

        self.table.setSortingEnabled(True)
        self.count_lbl.setText(
            f"Showing {len(self.filtered_entries)} of {len(self.all_entries)} TDocs"
        )

    def _on_selection_changed(self):
        """Updates the bottom drawer when a table row is selected."""
        selected_rows = self.table.selectionModel().selectedRows()
        if not selected_rows:
            self.drawer_text.clear()
            self.drawer_title_lbl.setText("<b>Technical Details &amp; Proposed Solution</b>")
            return

        row_idx = selected_rows[0].row()
        item = self.table.item(row_idx, 0)
        if not item:
            return

        r = item.data(Qt.UserRole)
        if not r:
            return

        tdoc = r.get("tdoc", "")
        title = r.get("title", "No Title")
        source = r.get("source", "Unknown")
        ai = r.get("agenda_item", "")
        ai_desc = r.get("ai_desc", "")
        wi = r.get("wi", "")
        status = r.get("status", "")
        meeting = r.get("meeting", "")
        abstract = r.get("abstract", "")
        remarks = r.get("remarks", "")
        contact = r.get("contact", "")

        self.drawer_title_lbl.setText(f"<b>{tdoc}: {html.escape(title)}</b>")

        html_parts = []
        meta_spans = []
        if source:
            meta_spans.append(f"<b>Source:</b> {html.escape(source)}")
        if meeting:
            meta_spans.append(f"<b>Meeting:</b> {html.escape(meeting)}")
        if ai:
            meta_spans.append(
                f"<b>Agenda Item:</b> {html.escape(ai)}{f' ({html.escape(ai_desc)})' if ai_desc else ''}"
            )
        if status:
            meta_spans.append(f"<b>Status:</b> {html.escape(status.capitalize())}")
        if wi:
            meta_spans.append(f"<b>Work Item:</b> {html.escape(wi)}")
        if contact:
            meta_spans.append(f"<b>Contact:</b> {html.escape(contact)}")

        html_parts.append(
            f"<div style='margin-bottom: 8px; color: {C.TEXT_SECONDARY};'>{' &nbsp;|&nbsp; '.join(meta_spans)}</div>"
        )

        if abstract:
            html_parts.append(
                f"<div style='font-weight: bold; margin-bottom: 4px; color: {C.TEXT_PRIMARY};'>Contribution Abstract / Technical Scope:</div>"
            )
            html_parts.append(
                f"<div style='background-color: {C.BG_CANVAS}; border-left: 3px solid {C.ACCENT_SKY}; padding: 6px 10px; margin-bottom: 8px; white-space: pre-wrap;'>{html.escape(abstract)}</div>"
            )

        if remarks:
            html_parts.append(
                f"<div style='font-weight: bold; margin-bottom: 4px; color: {C.TEXT_PRIMARY};'>Secretary Remarks / Outcome:</div>"
            )
            html_parts.append(
                f"<div style='background-color: {C.WARNING_SOFT_ALT}; border-left: 3px solid {C.WARNING_AMBER}; padding: 6px 10px; color: {C.WARNING_DARK};'>{html.escape(remarks)}</div>"
            )

        if not abstract and not remarks:
            html_parts.append(
                f"<div style='color: {C.TEXT_FAINT}; font-style: italic;'>No extended abstract or secretary remarks in the meeting spreadsheet for this document.</div>"
            )

        self.drawer_text.setHtml("".join(html_parts))
