# --- File: src/core/utils/company_sanitization/formatters.py ---
"""HTML, tooltip, Markdown, and LLM triage formatters for TDoc Source attribution."""

import html
from typing import List


class CompanyFormatterMixin:
    """Mixin formatting confirmed vs tentative sources for Qt HTML, Markdown, and LLM prompts."""

    @classmethod
    def format_source_html(cls, original_source: str) -> str:
        """
        Renders `original_source` as rich HTML where confirmed contributors appear in standard text
        and unconfirmed/tentative contributors (`?`, `(?)`, `[...]`) are rendered greyed-out and italicized.
        """
        if not original_source:
            return ""
        s = str(original_source).strip()
        if not cls.has_tentative_sources(s):
            return html.escape(s)

        parts = []
        for seg_text, is_tent in cls.tokenize_source_segments(s):
            esc = html.escape(seg_text)
            if is_tent:
                parts.append(f'<span style="color: #94A3B8; font-style: italic;">{esc}</span>')
            else:
                parts.append(esc)
        return "".join(parts)

    @classmethod
    def build_source_tooltip_html(cls, original_source: str) -> str:
        """
        Builds a rich, self-explanatory HTML tooltip for a TDoc Source string.
        When tentative sources (`?` or `[...]`) are present, breaks down Confirmed vs. Unconfirmed
        companies and explains the 3GPP/SA2 convention and statistical exclusion rule.
        """
        if not original_source:
            return ""
        s = str(original_source).strip()
        if not cls.has_tentative_sources(s):
            return f"<div style='max-width: 420px; white-space: pre-wrap;'>{html.escape(s)}</div>"

        confirmed, tentative = cls.parse_contributors_detailed(s)
        formatted_inline = cls.format_source_html(s)

        conf_str = ", ".join(confirmed) if confirmed else "<i>None recognized</i>"
        tent_str = ", ".join(tentative) if tentative else "<i>Unconfirmed / bracketed segment(s)</i>"

        return (
            f"<div style='max-width: 440px;'>"
            f"<b>Source Field:</b><br>{formatted_inline}"
            f"<hr style='border: none; border-top: 1px solid #CBD5E1; margin: 6px 0;'>"
            f"<b>✅ Confirmed Contributors ({len(confirmed)}):</b> {conf_str}<br>"
            f"<b>❔ Unconfirmed / Carry-over ({len(tentative)}):</b> "
            f"<span style='color: #64748B; font-style: italic;'>{tent_str}</span>"
            f"<hr style='border: none; border-top: 1px solid #CBD5E1; margin: 6px 0;'>"
            f"<span style='color: #475569; font-size: 11px;'>"
            f"ℹ️ <b>How to read <code>?</code> and <code>[...]</code> in Source:</b><br>"
            f"In 3GPP revisions (especially SA2 CRs/pCRs), the source list is carried over from the "
            f"original contribution, but companies marked with <code>?</code> or <code>[...]</code> have "
            f"<b>not confirmed</b> co-signing this revision (and may even oppose the change).<br>"
            f"• <i>Unconfirmed companies are excluded from contribution counts and statistics.</i>"
            f"</span>"
            f"</div>"
        )

    @classmethod
    def get_confirmed_source_text(cls, original_source: str) -> str:
        """
        Returns clean comma-separated confirmed source segments (preserving original
        casing, qualifiers, and role brackets like '[pen holder]'), excluding tentative
        ('?' or '[...]') segments.
        """
        if not original_source:
            return ""
        if not cls.has_tentative_sources(original_source):
            return str(original_source).strip()

        confirmed_parts = []
        for seg_text, is_tent in cls.tokenize_source_segments(original_source):
            if not is_tent:
                for sub in cls._COMMA_SEMICOLON_SPLIT_RE.split(seg_text):
                    sub_s = sub.strip()
                    if sub_s:
                        confirmed_parts.append(sub_s)
        return ", ".join(confirmed_parts)

    @classmethod
    def get_tentative_source_text(cls, original_source: str) -> str:
        """
        Returns clean comma-separated unconfirmed/tentative source segments,
        stripped of outer brackets ('[...]'), question marks ('?'), and parenthetical question marks.
        Deduplicates against confirmed sources to prevent repeated names.
        """
        if not original_source or not cls.has_tentative_sources(original_source):
            return ""

        confirmed_norm = {
            c.lower()
            for c in cls._CONFIRMED_NORM_SPLIT_RE.split(cls.get_confirmed_source_text(original_source))
            if c
        }
        tentative_parts = []
        seen = set()

        for seg_text, is_tent in cls.tokenize_source_segments(original_source):
            if is_tent:
                cleaned = cls._TENTATIVE_OUTER_STRIP_RE.sub('', seg_text)
                for sub in cls._TENTATIVE_INNER_SPLIT_RE.split(cleaned):
                    s = sub.strip()
                    if s and s.lower() not in confirmed_norm and s.lower() not in seen:
                        seen.add(s.lower())
                        tentative_parts.append(s)
        return ", ".join(tentative_parts)

    @classmethod
    def format_source_markdown_header_lines(cls, original_source: str) -> List[str]:
        """
        Formats source lines for Markdown metadata headers using the Two-Line Header pattern (Approach B):
        - '- **Source:** <confirmed>'
        - If tentative sources exist, adds:
          '- **Unconfirmed / Carry-over Sources:** <tentative> *(carry-over from earlier revision; unconfirmed co-signers)*'
        """
        if not original_source:
            return ["- **Source:** Unknown"]

        if not cls.has_tentative_sources(original_source):
            return [f"- **Source:** {str(original_source).strip()}"]

        confirmed = cls.get_confirmed_source_text(original_source)
        tentative = cls.get_tentative_source_text(original_source)

        lines = []
        if confirmed:
            lines.append(f"- **Source:** {confirmed}")
        else:
            lines.append("- **Source:** *(none confirmed)*")

        if tentative:
            lines.append(
                f"- **Unconfirmed / Carry-over Sources:** {tentative} "
                f"*(carry-over from earlier revision; unconfirmed co-signers)*"
            )
        return lines

    @classmethod
    def format_source_markdown_table_cell(cls, original_source: str) -> str:
        """
        Formats a clean, single-line representation of the source for Markdown tables
        (e.g. in index.md and Company summary reports).
        - If tentative sources exist: 'Confirmed *(carry-over: Tentative)*'
        - Escapes markdown table pipes ('|') and removes newlines.
        """
        if not original_source:
            return ""

        if not cls.has_tentative_sources(original_source):
            return str(original_source).replace('\n', ' ').replace('\r', '').replace('|', '\\|').strip()

        confirmed = cls.get_confirmed_source_text(original_source)
        tentative = cls.get_tentative_source_text(original_source)

        if confirmed and tentative:
            cell = f"{confirmed} *(carry-over: {tentative})*"
        elif confirmed:
            cell = confirmed
        elif tentative:
            cell = f"*(carry-over only: {tentative})*"
        else:
            cell = str(original_source).strip()

        return cell.replace('\n', ' ').replace('\r', '').replace('|', '\\|')

    @classmethod
    def format_source_triage_context(cls, original_source: str) -> str:
        """
        Formats the Source parameter passed to the LLM for TDoc technical triage.
        Prevents LLMs from falsely attributing proposals to unconfirmed carry-over co-sponsors.
        """
        if not original_source:
            return "Unknown"

        if not cls.has_tentative_sources(original_source):
            return str(original_source).strip()

        confirmed = cls.get_confirmed_source_text(original_source)
        tentative = cls.get_tentative_source_text(original_source)

        if confirmed and tentative:
            return (
                f"{confirmed} "
                f"[Note: {tentative} are carry-over co-sponsors from an earlier draft and have not confirmed co-signing this revision]"
            )
        elif confirmed:
            return confirmed
        elif tentative:
            return f"None confirmed (carry-over co-sponsors: {tentative})"
        return str(original_source).strip()
