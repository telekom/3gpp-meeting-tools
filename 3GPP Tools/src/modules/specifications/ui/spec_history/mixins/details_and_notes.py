"""Rich HTML notes rendering, anchor links, and notes drawer context menu mixin."""

import html
from typing import Any, Dict, List

from PyQt5.QtCore import QPoint, QUrl
from PyQt5.QtWidgets import QMenu

from core.ui.theme import C
from modules.meetings.ui.tdocs_menus import _copy_to_clipboard
from modules.specifications.core.spec_history_extractor import RE_TDOC_TOKEN
from modules.specifications.core.spec_tdoc_opener import SpecTDocOpener


class _HistoryDetailsMixin:
    """Formats rapporteur notes into clickable HTML and handles selection/anchor menus."""

    def _format_record_to_html(self, rec: Dict[str, Any]) -> str:
        """Formats the selected entry's metadata and rapporteur notes into interactive HTML with clickable TDoc badges."""
        if not rec:
            return (
                "<html><body style=\"font-family: 'Segoe UI', -apple-system, sans-serif; "
                f"font-size: 12px; color: {C.TEXT_FAINT}; background-color: {C.BG_SURFACE}; margin: 10px;\">"
                "<i>Select an entry above to view details, rapporteur notes, and clickable TDoc links.</i>"
                "</body></html>"
            )

        subject_text = rec.get("subject", "").strip()
        role = rec.get("role", "")
        score = rec.get("score", 0)
        evidence_badges = rec.get("evidence_badges", [])
        aff_clauses = rec.get("affected_clauses", "")
        primary = rec.get("source_tdoc", "").strip()
        if primary == "-":
            primary = ""
        meeting = rec.get("meeting", "—")
        version = rec.get("version", "—")
        cr_num = rec.get("cr_number", "—")
        cat = rec.get("category", "—")

        all_tdocs = rec.get("tdocs", [])
        raw_tokens: List[str] = []
        if primary:
            raw_tokens.append(primary)
        for t in all_tdocs:
            tok = t.get("raw_tdoc", "")
            if tok and tok not in raw_tokens:
                raw_tokens.append(tok)

        def _make_badge(token: str, trailing_nbsp: bool = True) -> str:
            t_sp = "&nbsp;" if trailing_nbsp else ""
            return (
                f'<a href="tdoc:{token}" title="Click to open {token} (Right-click for options)" '
                f'style="color: {C.ACCENT_SKY_DARK}; background-color: {C.SKY_SOFT}; text-decoration: none; '
                f'font-weight: bold; font-family: Consolas, monospace;">'
                f'&nbsp;{token}{t_sp}</a>'
            )

        if subject_text:
            escaped = html.escape(subject_text)

            def _repl_tdoc(match):
                token = match.group(1)
                next_char = escaped[match.end():match.end()+1]
                trailing = next_char not in (",", ";", ":", ".", "!", "?", ")", "]", '"')
                return _make_badge(token, trailing_nbsp=trailing)

            html_body = RE_TDOC_TOKEN.sub(_repl_tdoc, escaped)
            lines = html_body.split("\n")
            formatted_lines = []
            for line in lines:
                l_str = line.strip()
                if not l_str:
                    formatted_lines.append("<div style='height: 6px;'></div>")
                elif l_str.startswith(("- ", "* ", "• ")):
                    clean_b = l_str[2:].strip()
                    formatted_lines.append(f"<div style='margin-left: 14px; margin-bottom: 4px;'>• {clean_b}</div>")
                else:
                    formatted_lines.append(f"<div style='margin-bottom: 5px; line-height: 1.5;'>{line}</div>")
            notes_html = "\n".join(formatted_lines)
        else:
            notes_html = f"<div style='color: {C.TEXT_MUTED};'><i>(No rapporteur notes recorded for this entry)</i></div>"

        # Top metadata row
        primary_html = ""
        if primary and primary not in subject_text:
            primary_html = f"&nbsp;&nbsp;|&nbsp;&nbsp;<b>Primary TDoc:</b> {_make_badge(primary, False)}"

        count_html = ""
        if len(raw_tokens) > 0:
            count_html = f"&nbsp;&nbsp;|&nbsp;&nbsp;<b style='color: {C.ACCENT_SKY};'>{len(raw_tokens)} TDoc(s)</b>"

        meta_row = (
            f"<div style='padding-bottom: 8px; margin-bottom: 10px; border-bottom: 1px solid {C.BORDER_SUBTLE}; "
            f"font-size: 11px; color: {C.TEXT_SECONDARY};'>"
            f"<b>Meeting:</b> {meeting} &nbsp;&nbsp;|&nbsp;&nbsp; "
            f"<b>Version:</b> v{version} &nbsp;&nbsp;|&nbsp;&nbsp; "
            f"<b>CR:</b> {cr_num} (Cat {cat})"
            f"{primary_html}"
            f"{count_html}"
            f"</div>"
        )

        role_banner = ""
        if role or score:
            badges_html = " ".join(
                f'<span style="background-color: {C.SKY_SOFT}; color: {C.ACCENT_SKY_DARK}; border: 1px solid {C.SKY_BORDER}; border-radius: 3px; padding: 2px 6px; font-size: 11px; margin-right: 4px;">{html.escape(b)}</span>'
                for b in evidence_badges
            )
            aff_html = f'<br/><span style="color: {C.TEXT_SECONDARY}; font-size: 11px;"><b>Affected Clauses:</b> <code>{html.escape(aff_clauses)}</code></span>' if aff_clauses else ""
            role_banner = f'''
            <div style="background-color: {C.SKY_SURFACE}; border: 1px solid {C.SKY_BORDER}; border-radius: 6px; padding: 8px 12px; margin-bottom: 10px;">
                <div style="font-size: 13px; font-weight: bold; color: {C.ACCENT_SKY}; margin-bottom: 4px;">
                    {html.escape(role)} <span style="font-size: 11px; font-weight: normal; color: {C.TEXT_MUTED};">(Relevance Score: <b>{score}</b>/100)</span>
                </div>
                <div>{badges_html}</div>
                {aff_html}
            </div>
            '''

        full_html = f"""
        <html>
        <head>
        <style>
          body {{
            font-family: 'Segoe UI', -apple-system, sans-serif;
            font-size: 12px;
            color: {C.TEXT_STRONG};
            background-color: {C.BG_SURFACE};
            margin: 6px 8px;
          }}
        </style>
        </head>
        <body>
          {role_banner}
          {meta_row}
          {notes_html}
        </body>
        </html>
        """
        return full_html

    def _on_table_selection_changed(self):
        """Populates the bottom drawer with the selected row's interactive notes and clickable TDocs."""
        selected = self.table.selectedItems()
        if not selected:
            self.notes_browser.setHtml(self._format_record_to_html({}))
            self.drawer_title_lbl.setText("<b>Selected Entry Details:</b>")
            self.btn_open_primary.setText("📄 Open TDoc")
            self.btn_open_primary.setEnabled(False)
            self.btn_copy_entry_tdocs.setEnabled(False)
            return

        row_idx = self.table.currentRow()
        if row_idx < 0 or row_idx >= len(self.filtered_records):
            return

        rec = self.filtered_records[row_idx]

        # Collect unique raw tokens
        all_tdocs = rec.get("tdocs", [])
        primary = rec.get("source_tdoc", "")
        raw_tokens: List[str] = []
        if primary and primary != "-":
            raw_tokens.append(primary)
        for t in all_tdocs:
            tok = t.get("raw_tdoc", "")
            if tok and tok not in raw_tokens:
                raw_tokens.append(tok)

        self.notes_browser.setHtml(self._format_record_to_html(rec))

        # Update drawer header title
        m_name = rec.get("meeting", "")
        v_name = f"v{rec.get('version', '')}" if rec.get("version") else ""
        cr_text = f"CR {rec.get('cr_number')}" if rec.get("cr_number") and rec.get("cr_number") != "-" else "p-CRs / Draft"
        header_text = f"<b>Selected Entry:</b> {f'{v_name} · ' if v_name else ''}{f'{m_name} · ' if m_name else ''}{cr_text}"
        if raw_tokens:
            header_text += f" <span style='color: {C.TEXT_MUTED};'>({len(raw_tokens)} TDoc{'s' if len(raw_tokens) > 1 else ''})</span>"
        self.drawer_title_lbl.setText(header_text)

        self.btn_open_primary.setEnabled(bool(raw_tokens))
        if raw_tokens:
            self.btn_open_primary.setText(f"📄 Open {raw_tokens[0]}")
        else:
            self.btn_open_primary.setText("📄 Open TDoc")

        self.btn_copy_entry_tdocs.setEnabled(bool(raw_tokens))

    def _on_notes_anchor_clicked(self, url: QUrl):
        """Dispatches left-click navigation on interactive links in the notes browser."""
        url_str = url.toString()
        if url_str.startswith("tdoc:"):
            tdoc_id = url_str.split(":", 1)[1]
            row_idx = self.table.currentRow()
            meeting = ""
            if 0 <= row_idx < len(self.filtered_records):
                meeting = self.filtered_records[row_idx].get("meeting", "")
            self._open_tdoc(tdoc_id, meeting)
        elif url.scheme() in ("http", "https"):
            from PyQt5.QtGui import QDesktopServices
            QDesktopServices.openUrl(url)

    def _show_notes_context_menu(self, pos: QPoint):
        """Shows context menu on right-click inside the notes browser, targeting TDoc links or selection."""
        anchor = self.notes_browser.anchorAt(pos)
        row_idx = self.table.currentRow()
        rec = self.filtered_records[row_idx] if 0 <= row_idx < len(self.filtered_records) else {}
        meeting_name = rec.get("meeting", "")

        menu = QMenu(self)

        if anchor and anchor.startswith("tdoc:"):
            tdoc_id = anchor.split(":", 1)[1]
            act_open = menu.addAction(f"📄 Open {tdoc_id} Document")
            act_open.triggered.connect(lambda: self._open_tdoc(tdoc_id, meeting_name))

            act_jump = menu.addAction(f"🚀 Jump to {tdoc_id} in Meetings Tab")
            act_jump.triggered.connect(lambda: SpecTDocOpener.jump_to_meetings_tab(tdoc_id))

            act_portal = menu.addAction(f"🌐 Open {tdoc_id} on 3GPP Portal (3GU)")
            act_portal.triggered.connect(lambda: SpecTDocOpener.open_tdoc_in_browser(tdoc_id))

            menu.addSeparator()
            act_copy_cit = menu.addAction(f"📋 Copy Citation ({tdoc_id})")
            act_copy_cit.triggered.connect(
                lambda: SpecTDocOpener.copy_tdoc_citation(
                    tdoc_id, self.spec_number, rec.get("version", ""), meeting_name, rec.get("subject", ""), parent=self
                )
            )

            act_copy_id = menu.addAction(f"📋 Copy Identifier ({tdoc_id})")
            act_copy_id.triggered.connect(
                lambda: _copy_to_clipboard(tdoc_id, label=tdoc_id, parent=self)
            )
        else:
            if self.notes_browser.textCursor().hasSelection():
                act_copy_sel = menu.addAction("📋 Copy Selection")
                act_copy_sel.triggered.connect(self.notes_browser.copy)

            act_select_all = menu.addAction("Select All")
            act_select_all.triggered.connect(self.notes_browser.selectAll)

            if rec:
                menu.addSeparator()
                act_copy_all = menu.addAction("📋 Copy All TDocs for this Entry")
                act_copy_all.triggered.connect(lambda: self._copy_entry_tdocs(rec))

        menu.exec_(self.notes_browser.viewport().mapToGlobal(pos))
