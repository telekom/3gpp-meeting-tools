"""Filtering, detail rendering, TDoc localization, clipboard copy, and Markdown export."""

from __future__ import annotations

import html
import logging
from pathlib import Path
from typing import Dict, List

from PyQt5.QtCore import QModelIndex, QTimer
from PyQt5.QtWidgets import QApplication, QFileDialog, QMessageBox

from core.ui.theme import C

logger = logging.getLogger(__name__)


class MeetingAgreementsActionsMixin:
    """Mixin handling selection updates, HTML detail rendering, navigation, and Markdown export."""

    def _on_filter_changed(self):
        query = self.search_input.text()
        kind = self.combo_kind.currentData()
        ai = self.combo_ai.currentData()
        self.model.apply_filter(search_text=query, kind_filter=kind, ai_filter=ai)
        self._update_status_count()
        if self.model.rowCount() > 0:
            self.table.selectRow(0)
            self._on_table_row_selected(self.model.index(0, 0))
        else:
            self.viewer.setHtml(
                f"<p style='color: {C.TEXT_MUTED};'><i>No agreements match active search filters.</i></p>"
            )

    def _update_status_count(self):
        curr = self.model.rowCount()
        tot = len(self.raw_agreements)
        self.status_count_lbl.setText(f"Showing {curr} of {tot} agreements")

    def _on_selection_changed(self, selected, deselected):
        indexes = self.table.selectionModel().selectedRows()
        if indexes:
            self._on_table_row_selected(indexes[0])

    def _on_table_row_selected(self, index: QModelIndex):
        ag = self.model.get_row(index.row())
        if not ag:
            return

        tdoc = ag.get("tdoc") or ag.get("source_tdoc") or "AI Level"
        ai = ag.get("ai") or "-"
        wi = ag.get("wi") or "-"
        title = ag.get("title") or ""
        company = ag.get("company") or ""
        kind = str(ag.get("kind", "Agreement") or "Agreement").strip()
        file_src = ag.get("file") or ""
        text = ag.get("text", "")

        is_wa = "working assumption" in kind.lower()
        is_ffs = "ffs" in kind.lower() or "open issue" in kind.lower()
        is_pcr = "pcr" in kind.lower() or "p-cr" in kind.lower()
        is_cr = kind.lower().endswith("cr") and not is_pcr
        is_ls = "ls" in kind.lower()

        if is_pcr:
            tag_bg, tag_fg, tag_border = C.SUCCESS_SOFT_ALT, C.SUCCESS_DARK, C.SUCCESS_BORDER_MEDIUM
            tag_label = "📝 3GPP AGREED PSEUDO-CR (pCR)"
        elif is_cr:
            tag_bg, tag_fg, tag_border = C.SUCCESS_SOFT_ALT, C.SUCCESS_DARK, C.SUCCESS_BORDER_MEDIUM
            tag_label = "📄 3GPP APPROVED CHANGE REQUEST (CR)"
        elif is_ls:
            tag_bg, tag_fg, tag_border = C.INDIGO_SOFT, C.WG_SA_TEXT, C.WG_SA_BORDER
            tag_label = "📤 3GPP APPROVED OUTGOING LS"
        elif is_wa:
            tag_bg, tag_fg, tag_border = C.WARNING_SOFT_ALT, C.WARNING_DARK, C.WARNING_BORDER_STRONG
            tag_label = "🟡 WORKING ASSUMPTION"
        elif is_ffs:
            tag_bg, tag_fg, tag_border = C.INFO_SOFT_STRONG, C.INFO_DARK, C.INFO_BORDER_MEDIUM
            tag_label = "🔵 FFS / OPEN ISSUE"
        else:
            tag_bg, tag_fg, tag_border = C.SUCCESS_SOFT_ALT, C.SUCCESS_DARK, C.SUCCESS_BORDER_MEDIUM
            tag_label = "🟢 3GPP TECHNICAL AGREEMENT"

        self.detail_title_lbl.setText(
            f"<b>Agreement Details:</b> <span style='color: {C.ACCENT_ALT};'>"
            f"{html.escape(tdoc)}</span> (AI {html.escape(ai)})"
        )

        meta_rows = []
        if tdoc and tdoc != "AI Level":
            meta_rows.append(
                f"<b>TDoc:</b> <span style='font-family: monospace; font-size: 13px; "
                f"font-weight: bold;'>{html.escape(tdoc)}</span>"
            )
        if ai and ai != "-":
            meta_rows.append(f"<b>Agenda Item:</b> {html.escape(ai)}")
        if wi and wi != "-" and wi.lower() != "general":
            meta_rows.append(f"<b>Work Item:</b> <code>{html.escape(wi)}</code>")
        if title:
            meta_rows.append(f"<b>Contribution Title:</b> {html.escape(title)}")
        if company:
            meta_rows.append(f"<b>Source Company:</b> {html.escape(company)}")
        if file_src:
            meta_rows.append(f"<b>Notes Document:</b> {html.escape(Path(file_src).name)}")

        meta_html = " &nbsp;•&nbsp; ".join(meta_rows)
        escaped_txt = html.escape(text)

        html_content = f"""
        <html>
        <body style="font-family: 'Segoe UI', Arial, sans-serif; line-height: 1.5; color: {C.TEXT_PRIMARY};">
            <div style="margin-bottom: 8px;">
                <span style="background-color: {tag_bg}; color: {tag_fg}; border: 1px solid {tag_border}; border-radius: 4px; padding: 2px 8px; font-weight: bold; font-size: 11px;">
                    {tag_label}
                </span>
            </div>
            <div style="background-color: {C.BG_CANVAS}; border-radius: 4px; padding: 6px 10px; margin-bottom: 12px; font-size: 11px; color: {C.TEXT_BODY};">
                {meta_html}
            </div>
            <div style="background-color: {C.BG_CARD}; border: 1px solid {C.BORDER_STRONG}; border-left: 4px solid {tag_fg}; border-radius: 4px; padding: 12px 14px; font-size: 13px; white-space: pre-wrap; font-family: 'Consolas', 'Segoe UI', monospace;">{escaped_txt}</div>
        </body>
        </html>
        """
        self.viewer.setHtml(html_content)

    def _on_table_double_clicked(self, index: QModelIndex):
        """Double clicking a row locates it in the parent meeting table."""
        self._locate_current_tdoc()

    def _locate_current_tdoc(self):
        indexes = self.table.selectionModel().selectedRows()
        if not indexes:
            return
        ag = self.model.get_row(indexes[0].row())
        if not ag:
            return
        tdoc = ag.get("tdoc") or ag.get("source_tdoc")
        if tdoc and tdoc != "AI Level":
            self.locate_tdoc_requested.emit(tdoc)
            parent = self.parent()
            if parent and hasattr(parent, "locate_tdoc"):
                parent.locate_tdoc(tdoc)

    def _copy_current_agreement(self):
        indexes = self.table.selectionModel().selectedRows()
        if not indexes:
            return
        ag = self.model.get_row(indexes[0].row())
        if not ag:
            return

        tdoc = ag.get("tdoc") or ag.get("source_tdoc") or "AI Level"
        ai = ag.get("ai") or ""
        wi = ag.get("wi") or ""
        kind = ag.get("kind") or "Agreement"
        text = ag.get("text", "")

        citation_header = f"[{kind}] {self.meeting_name}"
        if ai:
            citation_header += f" | AI {ai}"
        if wi and wi.lower() != "general":
            citation_header += f" ({wi})"
        if tdoc and tdoc != "AI Level":
            citation_header += f" | {tdoc}"

        full_copy = f"{citation_header}:\n{text}"
        QApplication.clipboard().setText(full_copy)
        self.btn_copy_single.setText("✓ Copied!")
        QTimer.singleShot(2000, lambda: self.btn_copy_single.setText("📋 Copy Agreement"))

    def _copy_all_markdown(self):
        md_text = self._build_markdown_content()
        if not md_text:
            QMessageBox.information(self, "No Agreements", "No agreements to copy.")
            return
        QApplication.clipboard().setText(md_text)
        self.btn_copy_all.setText("✓ Copied All!")
        QTimer.singleShot(2500, lambda: self.btn_copy_all.setText("📋 Copy All (Markdown)"))

    def _build_markdown_content(self) -> str:
        """Builds a structured GFM Markdown document containing all currently filtered agreements."""
        rows = self.model._filtered_agreements
        if not rows:
            return ""

        lines = [
            f"# {self.meeting_name} — Agreements & Working Assumptions",
            f"*Generated by 3GPP Tools • Total Agreements: {len(rows)}*",
            "",
            "---",
            "",
        ]

        by_ai: Dict[str, List[dict]] = {}
        for ag in rows:
            ai = str(ag.get("ai") or "General").strip()
            by_ai.setdefault(ai, []).append(ag)

        for ai, ag_list in by_ai.items():
            lines.append(f"## Agenda Item {ai}")
            lines.append("")

            for ag in ag_list:
                kind = ag.get("kind", "Agreement")
                tdoc = ag.get("tdoc") or ag.get("source_tdoc") or "AI Level"
                wi = ag.get("wi", "")
                title = ag.get("title", "")
                company = ag.get("company", "")
                text = ag.get("text", "").strip()

                meta_parts = []
                if tdoc and tdoc != "AI Level":
                    meta_parts.append(f"**TDoc:** `{tdoc}`")
                if wi and wi.lower() != "general":
                    meta_parts.append(f"**WI:** `{wi}`")
                if company:
                    meta_parts.append(f"**Source:** {company}")

                badge = (
                    "🟡 WORKING ASSUMPTION"
                    if "working assumption" in str(kind).lower()
                    else ("🔵 FFS" if "ffs" in str(kind).lower() else "🟢 AGREEMENT")
                )
                lines.append(f"### {badge}: {tdoc}")
                if meta_parts:
                    lines.append(f"> {' • '.join(meta_parts)}")
                    lines.append(">")
                if title:
                    lines.append(f"> *{title}*")
                    lines.append("")

                for t_line in text.splitlines():
                    lines.append(f"    {t_line}")
                lines.append("")

            lines.append("---")
            lines.append("")

        return "\n".join(lines).strip()

    def _export_to_markdown_file(self):
        md_text = self._build_markdown_content()
        if not md_text:
            QMessageBox.information(self, "No Agreements", "No agreements to export.")
            return

        default_dir = str(self.meeting_dir / "Export") if self.meeting_dir else ""
        default_file = f"{self.meeting_name.replace(' ', '_')}_Agreements.md"
        save_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Meeting Agreements to Markdown",
            str(Path(default_dir) / default_file) if default_dir else default_file,
            "Markdown Files (*.md);;All Files (*.*)",
        )
        if save_path:
            try:
                Path(save_path).parent.mkdir(parents=True, exist_ok=True)
                with open(save_path, "w", encoding="utf-8") as f:
                    f.write(md_text)
                QMessageBox.information(
                    self,
                    "Export Successful",
                    f"Successfully exported {self.model.rowCount()} agreement(s) to:\n{save_path}",
                )
            except Exception as e:
                logger.error(f"Failed to export agreements markdown: {e}")
                QMessageBox.warning(self, "Export Failed", f"Could not save file:\n{e}")

    def focus_tdoc(self, tdoc_id: str):
        """Highlights and selects the agreement corresponding to a given TDoc ID."""
        if not tdoc_id:
            return
        tdoc_clean = tdoc_id.strip().upper()
        for row_idx, ag in enumerate(self.model._filtered_agreements):
            t = str(ag.get("tdoc") or ag.get("source_tdoc") or "").strip().upper()
            if t == tdoc_clean:
                self.table.selectRow(row_idx)
                self.table.scrollTo(self.model.index(row_idx, 0))
                self._on_table_row_selected(self.model.index(row_idx, 0))
                return
