"""Detail reader drawer and document/portal/export action handlers mixin."""

import csv
import logging
import os
from pathlib import Path
from typing import Optional

from PyQt5.QtCore import QTimer
from PyQt5.QtWidgets import (
    QApplication,
    QFileDialog,
    QMessageBox,
    QTableWidget,
    QTableWidgetItem,
)

from core.ui.theme import C
from modules.meetings.core.tdocs.tdoc_url_resolver import resolve_tdoc_download_url

logger = logging.getLogger("modules.meetings.ui.dialogs.topic_inputs_dialog")


class TopicInputsDrawerActionsMixin:
    """Manages detail reader rendering, document/portal opening, and Markdown/CSV exports."""

    def _toggle_drawer(self):
        is_hidden = self.detail_text_edit.isHidden()
        self.detail_text_edit.setVisible(is_hidden)
        self.btn_toggle_drawer.setText("▲ Collapse" if is_hidden else "▼ Expand")

    def _on_copy_detail_text(self):
        text = self._current_detail_text or self.detail_text_edit.toPlainText().strip()
        if text:
            app = QApplication.instance()
            if app:
                app.clipboard().setText(text)
            self.lbl_drawer_title.setText("📖 Decision & Agreement Detail Reader — 📋 Copied!")
            QTimer.singleShot(2000, lambda: self.lbl_drawer_title.setText("📖 Decision & Agreement Detail Reader"))

    def _on_decision_selection_changed(self):
        row = self.table_decisions.currentRow()
        if row < 0 or row >= len(self._displayed_decisions):
            return
        d = self._displayed_decisions[row]
        self._current_detail_text = d.title_or_summary

        if d.decision_type in ("Text Agreement", "Working Assumption"):
            prov_text = f"🟢 {d.decision_type} (Chairman's Notes)"
            prov_style = f"background-color: {C.SUCCESS_SOFT_ALT}; color: {C.SUCCESS_DARK}; font-weight: bold; font-size: 11px; padding: 2px 8px; border-radius: 4px; border: 1px solid {C.SUCCESS_BORDER};"
        elif d.decision_type in ("Approved pCR", "Approved CR"):
            prov_text = f"🟢 {d.decision_type} (Approved Document)"
            prov_style = f"background-color: {C.WG_RAN_BG}; color: {C.WG_RAN_TEXT}; font-weight: bold; font-size: 11px; padding: 2px 8px; border-radius: 4px; border: 1px solid {C.WG_RAN_BORDER};"
        else:
            prov_text = f"{d.badge_icon} {d.decision_type}"
            prov_style = f"background-color: {C.BG_CANVAS}; color: {C.TEXT_BODY}; font-weight: bold; font-size: 11px; padding: 2px 8px; border-radius: 4px; border: 1px solid {C.BORDER_STRONG};"

        self.lbl_drawer_provenance.setText(prov_text)
        self.lbl_drawer_provenance.setStyleSheet(prov_style)
        self.lbl_drawer_provenance.setVisible(True)

        meta_parts = [
            f"<b>WG:</b> {d.wg}",
            f"<b>Meeting:</b> #{d.meeting_number} ({d.date or 'N/A'})",
            f"<b>AI:</b> <code>{d.ai or 'N/A'}</code>",
        ]
        if d.target_spec:
            meta_parts.append(f"<b>Spec:</b> <code>{d.target_spec}</code>")
        if d.wi:
            meta_parts.append(f"<b>WI:</b> <code>{d.wi}</code>")
        if d.tdoc_id:
            meta_parts.append(f"<b>TDoc / Background Proposal:</b> <b>{d.tdoc_id}</b> ({d.status})")

        html_lines = [
            f"<p style='margin:0 0 6px 0; font-size:11px;'>{' &nbsp;|&nbsp; '.join(meta_parts)}</p>"
        ]

        if d.decision_type in ("Text Agreement", "Working Assumption"):
            html_lines.append(
                f"<div style='background-color:{C.SUCCESS_SOFT}; border-left:3px solid {C.GREEN_ACCENT}; padding:4px 8px; margin-bottom:8px; font-size:11px; color:{C.SUCCESS_DARK};'>"
                "<b>Binding Consensus:</b> In 3GPP RAN Working Groups, only the verbatim text below recorded in the Chairman's Notes represents formal consensus. Cited background proposal was Noted.</div>"
            )
        elif d.decision_type in ("Approved pCR", "Approved CR"):
            html_lines.append(
                f"<div style='background-color:{C.WG_RAN_BG}; border-left:3px solid {C.EMERALD_PRIMARY}; padding:4px 8px; margin-bottom:8px; font-size:11px; color:{C.WG_RAN_TEXT};'>"
                f"<b>Approved Specification Change:</b> The full document attachment of <b>{d.tdoc_id}</b> was formally approved as a baseline change for <code>{d.target_spec}</code>.</div>"
            )

        verbatim_html = d.title_or_summary.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\n", "<br/>")
        html_lines.append(f"<div style='background-color:{C.BG_CARD}; border:1px solid {C.BORDER_STRONG}; border-radius:4px; padding:8px; font-family:Consolas, monospace; font-size:11px; line-height:1.4;'>{verbatim_html}</div>")

        self.detail_text_edit.setHtml("".join(html_lines))

    def _on_contribution_selection_changed(self):
        row = self.table_contributions.currentRow()
        if row < 0 or row >= len(self._displayed_contribs):
            return
        c = self._displayed_contribs[row]
        tid = c.get("tdoc_id", "")
        raw = c.get("raw_row", {})
        abstract = raw.get("Abstract", "") or raw.get("abstract", "") or raw.get("Secretary Remarks", "")

        text_parts = [f"{tid}: {c.get('title', '')}", f"Source: {c.get('source', '')}", f"Type: {c.get('type', '')} | Status: {c.get('status', '')}"]
        if abstract:
            text_parts.append(f"\nAbstract / Remarks:\n{abstract}")
        self._current_detail_text = "\n".join(text_parts)

        self.lbl_drawer_provenance.setText(f"📄 {c.get('type', 'TDoc')} ({c.get('status', 'N/A')})")
        self.lbl_drawer_provenance.setStyleSheet(f"background-color: {C.BG_CANVAS}; color: {C.TEXT_BODY}; font-weight: bold; font-size: 11px; padding: 2px 8px; border-radius: 4px; border: 1px solid {C.BORDER_STRONG};")
        self.lbl_drawer_provenance.setVisible(True)

        meta_parts = [
            f"<b>TDoc:</b> <b>{tid}</b>",
            f"<b>WG:</b> {c.get('wg')} #{c.get('meeting_number')}",
            f"<b>AI:</b> <code>{c.get('ai')}</code>",
            f"<b>Source:</b> {c.get('source')}",
            f"<b>Status:</b> {c.get('status')}",
        ]
        if c.get("wi"):
            meta_parts.append(f"<b>WI:</b> <code>{c.get('wi')}</code>")
        if c.get("spec"):
            meta_parts.append(f"<b>Spec:</b> <code>{c.get('spec')}</code>")

        html_lines = [
            f"<p style='margin:0 0 4px 0; font-size:11px;'>{' &nbsp;|&nbsp; '.join(meta_parts)}</p>",
            f"<p style='margin:0 0 6px 0; font-size:12px; font-weight:bold; color:{C.TEXT_STRONG};'>{c.get('title')}</p>",
        ]
        if abstract:
            abs_html = abstract.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\n", "<br/>")
            html_lines.append(f"<div style='background-color:{C.BG_CARD}; border:1px solid {C.BORDER_STRONG}; border-radius:4px; padding:8px; font-size:11px;'><b>Abstract / Remarks:</b><br/>{abs_html}</div>")

        self.detail_text_edit.setHtml("".join(html_lines))

    def _on_ls_selection_changed(self):
        row = self.table_ls.currentRow()
        if row < 0 or row >= len(self._displayed_ls):
            return
        c = self._displayed_ls[row]
        tid = c.get("tdoc_id", "")
        raw = c.get("raw_row", {})
        reply_to = raw.get("Reply to", "") or raw.get("reply_to", "") or "—"
        reply_in = raw.get("Reply in", "") or raw.get("reply_in", "") or "—"

        self._current_detail_text = (
            f"Liaison Statement: {tid}\nTitle: {c.get('title')}\nSource: {c.get('source')}\n"
            f"Reply To: {reply_to}\nReply In: {reply_in}\nStatus: {c.get('status')}"
        )

        t_type = c.get("type", "")
        is_in = "in" in t_type.lower()
        dir_text = "📥 LS IN" if is_in else "📤 LS OUT"

        self.lbl_drawer_provenance.setText(f"{dir_text} ({c.get('status', 'N/A')})")
        self.lbl_drawer_provenance.setStyleSheet(f"background-color: {C.VIOLET_SOFT}; color: {C.VIOLET_DARK}; font-weight: bold; font-size: 11px; padding: 2px 8px; border-radius: 4px; border: 1px solid {C.VIOLET_BORDER};")
        self.lbl_drawer_provenance.setVisible(True)

        meta_parts = [
            f"<b>Liaison Statement:</b> <b>{tid}</b> ({dir_text})",
            f"<b>WG:</b> {c.get('wg')} #{c.get('meeting_number')}",
            f"<b>Source:</b> {c.get('source')}",
            f"<b>Status:</b> {c.get('status')}",
        ]
        html_lines = [
            f"<p style='margin:0 0 4px 0; font-size:11px;'>{' &nbsp;|&nbsp; '.join(meta_parts)}</p>",
            f"<p style='margin:0 0 6px 0; font-size:12px; font-weight:bold; color:{C.TEXT_STRONG};'>{c.get('title')}</p>",
            f"<p style='margin:0; font-size:11px; color:{C.TEXT_SECONDARY};'><b>Reply To:</b> <code>{reply_to}</code> &nbsp;|&nbsp; <b>Reply In:</b> <code>{reply_in}</code></p>",
        ]
        self.detail_text_edit.setHtml("".join(html_lines))

    def _on_tab_changed(self, idx: int):
        if idx == 0:
            self._on_decision_selection_changed()
        elif idx == 1:
            self._on_contribution_selection_changed()
        elif idx == 2:
            self._on_ls_selection_changed()
        else:
            self.detail_text_edit.clear()
            self.lbl_drawer_provenance.setVisible(False)
            self._current_detail_text = ""

    def _get_active_table(self) -> Optional[QTableWidget]:
        idx = self.tabs.currentIndex()
        if idx == 0:
            return self.table_decisions
        elif idx == 1:
            return self.table_contributions
        elif idx == 2:
            return self.table_ls
        return None

    def _get_selected_tdoc_id(self) -> Optional[str]:
        table = self._get_active_table()
        if not table:
            return None
        row = table.currentRow()
        if row < 0:
            return None
        if table == self.table_decisions:
            val = table.item(row, 4)
            return val.text().strip() if val and val.text() != "—" else None
        elif table == self.table_contributions:
            val = table.item(row, 0)
            return val.text().strip() if val else None
        elif table == self.table_ls:
            val = table.item(row, 1)
            return val.text().strip() if val else None
        return None

    def _on_table_item_double_clicked(self, item: QTableWidgetItem):
        self._on_open_selected_document()

    def _on_open_selected_document(self):
        tid = self._get_selected_tdoc_id()
        if not tid:
            QMessageBox.information(self, "No Document Selected", "Please select a row with a valid TDoc ID.")
            return

        # Find in candidate contributions
        contrib = next((c for c in self.all_contributions if c.get("tdoc_id") == tid), None)
        if not contrib:
            QMessageBox.information(self, "TDoc Info", f"TDoc ID: {tid}")
            return

        from modules.meetings.core.settings import MeetingsSettings
        cache_dir = MeetingsSettings().cache_dir
        folder_name = contrib.get("folder_name", "")
        mtg_dir = Path(cache_dir) / folder_name if cache_dir and folder_name else None

        # Check local files
        if mtg_dir and mtg_dir.is_dir():
            for p in mtg_dir.rglob(f"{tid}*"):
                if p.is_file() and not p.name.endswith(".json"):
                    try:
                        os.startfile(str(p))
                        return
                    except Exception as e:
                        logger.debug(f"Failed to open {p}: {e}")

        # If not cached locally, offer portal open
        ret = QMessageBox.question(
            self,
            "Document Not Cached",
            f"TDoc {tid} is not yet downloaded locally in meeting folder '{folder_name}'.\n\nWould you like to open it on the 3GPP Portal?",
            QMessageBox.Yes | QMessageBox.No,
        )
        if ret == QMessageBox.Yes:
            self._on_open_selected_portal()

    def _on_open_selected_portal(self):
        tid = self._get_selected_tdoc_id()
        if not tid:
            QMessageBox.information(self, "No Document Selected", "Please select a row with a valid TDoc ID.")
            return
        contrib = next((c for c in self.all_contributions if c.get("tdoc_id") == tid), None)
        folder_name = contrib.get("folder_name", "") if contrib else ""
        url = resolve_tdoc_download_url(tid, folder_name)
        if url:
            import webbrowser
            webbrowser.open(url)
        else:
            QMessageBox.warning(self, "URL Error", f"Could not construct 3GPP Portal URL for {tid}.")

    def _on_copy_markdown_table(self):
        table = self._get_active_table()
        if not table or table.rowCount() == 0:
            QMessageBox.information(self, "Empty Table", "No rows available to copy.")
            return

        cols = table.columnCount()
        headers = [table.horizontalHeaderItem(c).text() for c in range(cols)]
        lines = []
        lines.append("| " + " | ".join(headers) + " |")
        lines.append("| " + " | ".join([":---"] * cols) + " |")

        for r in range(table.rowCount()):
            row_vals = []
            for c in range(cols):
                it = table.item(r, c)
                val = it.text().replace("\n", " ").replace("|", "\\|").strip() if it else ""
                row_vals.append(val)
            lines.append("| " + " | ".join(row_vals) + " |")

        md_text = "\n".join(lines)
        QApplication.clipboard().setText(md_text)
        QMessageBox.information(self, "Copied", f"Copied {table.rowCount()} rows formatted as Markdown to clipboard.")

    def _on_export_csv(self):
        table = self._get_active_table()
        if not table or table.rowCount() == 0:
            QMessageBox.information(self, "Empty Table", "No rows available to export.")
            return

        path, _ = QFileDialog.getSaveFileName(
            self,
            "Export Table to CSV",
            f"Topic_Inputs_{self.topic.id}.csv",
            "CSV Files (*.csv);;All Files (*.*)",
        )
        if not path:
            return

        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                headers = [table.horizontalHeaderItem(c).text() for c in range(table.columnCount())]
                writer.writerow(headers)
                for r in range(table.rowCount()):
                    row_vals = [table.item(r, c).text().strip() if table.item(r, c) else "" for c in range(table.columnCount())]
                    writer.writerow(row_vals)
            QMessageBox.information(self, "Export Successful", f"Saved {table.rowCount()} rows to:\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "Export Failed", f"Could not write CSV:\n{e}")

    def _on_proceed_to_export(self):
        # Stop and wait for the preview collector before starting the exporter so both threads
        # never download or touch the same meeting folders concurrently.
        self.close()
        self.export_requested.emit()
