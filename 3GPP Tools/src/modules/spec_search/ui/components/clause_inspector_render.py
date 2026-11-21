"""Clause HTML highlighting, match excerpting, clipboard copy, and clear helpers."""

import html
import re
from typing import Optional

from PyQt5.QtCore import QTimer
from PyQt5.QtGui import QTextCursor, QTextDocument

from core.ui.theme import C
from modules.spec_search.core.spec_clause_exporter import copy_clause_to_clipboard


class SpecClauseInspectorRenderMixin:
    """Rendering, search navigation, and clipboard formatting methods for SpecClauseInspector."""

    def display_clause(
        self,
        clause_number: str,
        clause_title: str,
        spec_number: str,
        version: str,
        content: str,
        release_date: Optional[str] = None,
        search_query: str = "",
        first_added_ver: Optional[str] = None,
        prev_ver: Optional[str] = None,
        spec_title: Optional[str] = None,
    ):
        self._current_spec = spec_number
        self._current_version = version
        self._current_clause_num = clause_number
        self._current_clause_title = clause_title
        self._current_content = content
        self._current_release_date = release_date or ""
        self._current_search_query = search_query.strip()
        self._first_added_ver = first_added_ver or ""
        self._prev_ver = prev_ver or ""
        self._current_spec_title = spec_title or ""

        date_str = f" ({release_date})" if release_date else ""
        if spec_title:
            header_title = f"TS {spec_number} ({spec_title}) v{version}{date_str} — Clause {clause_number}: {clause_title}"
            self.title_lbl.setToolTip(f"TS {spec_number}: {spec_title}")
        else:
            header_title = f"TS {spec_number} v{version}{date_str} — Clause {clause_number}: {clause_title}"
            self.title_lbl.setToolTip("")
        self.title_lbl.setText(header_title)

        self._last_citation_text = f"3GPP TS {spec_number} v{version}{date_str}, Clause {clause_number} ({clause_title}):\n\n{content}"
        self.btn_copy.setEnabled(True)
        self.btn_extract.setEnabled(True)
        self.btn_llm_diff.setEnabled(True)
        self.btn_open_doc.setEnabled(True)
        self.btn_show_in_specs.setEnabled(True)
        self.btn_version_changes.setEnabled(True)
        self.btn_version_changes.setText(f"📜 v{version} Changes")
        self.btn_version_changes.setToolTip(f"View all CRs, p-CRs, and TDocs integrated in TS {spec_number} v{version}, and directly open the actual TDoc documents.")
        self.btn_origin_tdoc.setEnabled(True)
        self.btn_origin_tdoc.setToolTip(f"Identify which CR or p-CR added or modified Clause {clause_number} in v{version}")

        # Update Quick Diff button
        if self._prev_ver:
            self.btn_quick_diff.setEnabled(True)
            self.btn_quick_diff.setToolTip(f"Compare clause {clause_number} between preceding v{self._prev_ver} and current v{version}")
        else:
            self.btn_quick_diff.setEnabled(False)
            self.btn_quick_diff.setToolTip("No preceding release indexed to compare against")

        # Update First Introduced chip
        if self._first_added_ver:
            clean_first = self._first_added_ver.lstrip("v")
            clean_curr = version.lstrip("v")
            if clean_first == clean_curr:
                self.first_added_badge.setText(f"🟢 First Added in v{clean_first}")
            else:
                self.first_added_badge.setText(f"🟢 Introduced in v{clean_first} (Click to View)")
            self.first_added_badge.setVisible(True)
        else:
            self.first_added_badge.setVisible(False)

        match_count = 0
        if self._current_search_query:
            match_count = len(re.findall(re.escape(self._current_search_query), content, re.IGNORECASE))

        if match_count > 0:
            self.match_badge.setText(f"🎯 {match_count} Match(es)")
            self.match_badge.setVisible(True)
            self.btn_prev.setEnabled(True)
            self.btn_next.setEnabled(True)
        else:
            self.match_badge.setVisible(False)
            self.btn_prev.setEnabled(False)
            self.btn_next.setEnabled(False)

        excerpt_html = ""
        if self._current_search_query and match_count > 0:
            excerpt_html = self._generate_match_excerpt(content, self._current_search_query, clause_title)

        # Deduplicate redundant first line if it duplicates the clause title banner
        cleaned_body = content.strip()
        lines = cleaned_body.splitlines()
        if lines:
            first_line = lines[0].strip()
            norm_header = f"{clause_number} {clause_title}".strip().lower()
            if (
                first_line.lower() == norm_header
                or first_line.lower() == clause_title.lower()
                or first_line.lower() == clause_number.lower()
            ):
                cleaned_body = "\n".join(lines[1:]).strip()

        escaped_content = html.escape(cleaned_body)
        if self._current_search_query:
            pattern = re.compile(re.escape(html.escape(self._current_search_query)), re.IGNORECASE)
            escaped_content = pattern.sub(
                lambda m: f'<mark style="background-color: {C.HIGHLIGHT_YELLOW}; font-weight: bold; padding: 1px 4px; border-radius: 2px; color: {C.TEXT_PRIMARY};">{m.group(0)}</mark>',
                escaped_content,
            )

        formatted_html = f"""
        <html>
        <body style="font-family: 'Segoe UI', Arial, sans-serif; font-size: 12px; line-height: 1.6; color: {C.TEXT_STRONG}; margin: 4px;">
            {excerpt_html}
            <div style="background-color: {C.BG_CANVAS}; border-left: 4px solid {C.ACCENT_SKY}; border-radius: 4px; padding: 6px 12px; margin: 8px 0 10px 0;">
                <span style="color: {C.ACCENT_SKY_DARK}; font-weight: bold; font-size: 13px;">Clause {clause_number}: {html.escape(clause_title)}</span>
            </div>
            <pre style="white-space: pre-wrap; font-family: 'Segoe UI', Arial, sans-serif; font-size: 12px; color: {C.TEXT_BODY}; line-height: 1.6;">{escaped_content}</pre>
        </body>
        </html>
        """
        self.browser.setHtml(formatted_html)

        if self._current_search_query and match_count > 0:
            self._find_next()


    def _generate_match_excerpt(self, full_text: str, query: str, clause_title: str = "") -> str:
        paragraphs = [p.strip() for p in full_text.split("\n") if p.strip()]
        matched_paras = [p for p in paragraphs if query.lower() in p.lower()]

        # Filter out paragraphs that simply echo the clause title or clause number + title
        body_matches = [
            p for p in matched_paras
            if p.lower() != clause_title.lower() and not p.lower().endswith(clause_title.lower())
        ]

        if body_matches:
            target_excerpt = body_matches[0]
            escaped_excerpt = html.escape(target_excerpt)
            pattern = re.compile(re.escape(html.escape(query)), re.IGNORECASE)
            highlighted_excerpt = pattern.sub(
                lambda m: f'<mark style="background-color: {C.HIGHLIGHT_YELLOW}; font-weight: bold; padding: 1px 4px; border-radius: 2px; color: {C.TEXT_PRIMARY};">{m.group(0)}</mark>',
                escaped_excerpt,
            )
            return f"""
            <div style="background-color: {C.BG_SURFACE}; border: 1px solid {C.BORDER_SUBTLE}; border-left: 4px solid {C.ACCENT_SKY}; border-radius: 6px; padding: 8px 12px; margin-bottom: 10px;">
                <div style="font-weight: 600; color: {C.ACCENT_SKY_DARK}; font-size: 11px; margin-bottom: 3px;">💡 Key Match Excerpt & Context:</div>
                <div style="color: {C.TEXT_STRONG}; font-size: 12px; line-height: 1.55;">{highlighted_excerpt}</div>
            </div>
            """
        elif clause_title and query.lower() in clause_title.lower():
            escaped_title = html.escape(clause_title)
            pattern = re.compile(re.escape(html.escape(query)), re.IGNORECASE)
            highlighted_title = pattern.sub(
                lambda m: f'<mark style="background-color: {C.HIGHLIGHT_YELLOW}; font-weight: bold; padding: 1px 4px; border-radius: 2px; color: {C.TEXT_PRIMARY};">{m.group(0)}</mark>',
                escaped_title,
            )
            return f"""
            <div style="background-color: {C.BG_SURFACE}; border: 1px solid {C.BORDER_SUBTLE}; border-left: 4px solid {C.ACCENT_SKY}; border-radius: 6px; padding: 8px 12px; margin-bottom: 10px;">
                <div style="font-weight: 600; color: {C.ACCENT_SKY_DARK}; font-size: 11px; margin-bottom: 3px;">💡 Match Location:</div>
                <div style="color: {C.TEXT_STRONG}; font-size: 12px; line-height: 1.55;">Query matches the <b>Clause Title</b>: {highlighted_title}</div>
            </div>
            """
        elif matched_paras:
            escaped_excerpt = html.escape(matched_paras[0])
            pattern = re.compile(re.escape(html.escape(query)), re.IGNORECASE)
            highlighted_excerpt = pattern.sub(
                lambda m: f'<mark style="background-color: {C.HIGHLIGHT_YELLOW}; font-weight: bold; padding: 1px 4px; border-radius: 2px; color: {C.TEXT_PRIMARY};">{m.group(0)}</mark>',
                escaped_excerpt,
            )
            return f"""
            <div style="background-color: {C.BG_SURFACE}; border: 1px solid {C.BORDER_SUBTLE}; border-left: 4px solid {C.ACCENT_SKY}; border-radius: 6px; padding: 8px 12px; margin-bottom: 10px;">
                <div style="font-weight: 600; color: {C.ACCENT_SKY_DARK}; font-size: 11px; margin-bottom: 3px;">💡 Key Match Excerpt & Context:</div>
                <div style="color: {C.TEXT_STRONG}; font-size: 12px; line-height: 1.55;">{highlighted_excerpt}</div>
            </div>
            """
        return ""

    def _find_next(self):
        if not self._current_search_query:
            return
        found = self.browser.find(self._current_search_query)
        if not found:
            self.browser.moveCursor(QTextCursor.Start)
            self.browser.find(self._current_search_query)

    def _find_prev(self):
        if not self._current_search_query:
            return
        found = self.browser.find(self._current_search_query, QTextDocument.FindBackward)
        if not found:
            self.browser.moveCursor(QTextCursor.End)
            self.browser.find(self._current_search_query, QTextDocument.FindBackward)

    def _copy_formatted(self, mode: str = "email"):
        if not self._current_spec or not self._current_clause_num or not self._current_content:
            return
        copy_clause_to_clipboard(
            self._current_spec,
            self._current_version,
            self._current_clause_num,
            self._current_clause_title,
            self._current_content,
            self._current_release_date,
            mode=mode,
        )
        self.btn_copy.setText("✅ Copied!")
        QTimer.singleShot(1500, lambda: self.btn_copy.setText("📋 Copy ▾"))

    def _copy_citation(self):
        self._copy_formatted("citation")

    def clear_display(self):
        self.title_lbl.setText("No clause selected")
        self.match_badge.setVisible(False)
        self.first_added_badge.setVisible(False)
        self.btn_prev.setEnabled(False)
        self.btn_next.setEnabled(False)
        self.btn_copy.setEnabled(False)
        self.btn_extract.setEnabled(False)
        self.btn_open_doc.setEnabled(False)
        self.btn_show_in_specs.setEnabled(False)
        self.btn_quick_diff.setEnabled(False)
        self.btn_llm_diff.setEnabled(False)
        self.btn_version_changes.setEnabled(False)
        self.btn_version_changes.setText("📜 Version Changes")
        self.btn_origin_tdoc.setEnabled(False)
        self.browser.clear()
        self._current_search_query = ""
        self._last_citation_text = ""
        self._current_spec = ""
        self._current_version = ""
        self._current_clause_num = ""
        self._current_clause_title = ""
        self._current_content = ""
        self._current_release_date = ""
        self._first_added_ver = ""
        self._prev_ver = ""
