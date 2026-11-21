"""TDoc hyperlink formatting, Chairman's Notes tagging, and rich HTML cards/tooltips."""

import html
import re

from core.ui.theme import C


class TDocsAgreementsHtmlMixin:
    """Detects agreement rows and builds styled HTML snippets for cells and hover tooltips."""

    def _linkify(self, prefix: str, text: str, html: bool) -> str:
        if not text:
            return ""
        if not html:
            return f"{prefix}: {text}" if prefix else text

        def repl(match):
            tdoc = match.group(0)
            tdoc_upper = tdoc.upper()
            is_local = (tdoc in self.valid_tdocs) or (tdoc_upper in self.valid_tdocs)

            if not is_local:
                base_match = re.search(r'^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$', tdoc_upper, re.IGNORECASE)
                if base_match:
                    base_tdoc = base_match.group(1)
                    if base_tdoc in self.valid_tdocs:
                        is_local = True

            if is_local:
                return f'<a href="{tdoc}" style="color: {C.OFFICE_BLUE}; font-weight: bold; text-decoration: underline;">{tdoc}</a>'
            else:
                return f'<a href="{tdoc}" style="color: {C.WG_CT_TEXT}; text-decoration: underline;">{tdoc}</a>'

        linked_text = re.sub(r'[a-zA-Z0-9]+-\d+(?:r\d+[a-zA-Z]?)?', repl, text, flags=re.IGNORECASE)

        if prefix:
            return f"<span style='color: {C.TEXT_SECONDARY};'><b>{prefix}:</b></span> {linked_text}"
        return linked_text

    def _format_related_tdocs(self, row_data: dict, html=False) -> str:
        parts = []
        if r_rev_of := row_data.get("Is revision of"):
            parts.append(self._linkify("⬅️ Rev of", r_rev_of, html))
        if r_rev_to := row_data.get("Revised to"):
            parts.append(self._linkify("➡️ Rev to", r_rev_to, html))
        if r_orig := row_data.get("Original LS"):
            parts.append(self._linkify("✉️ Orig LS", r_orig, html))
        if r_reply := row_data.get("Reply in"):
            parts.append(self._linkify("↩️ Reply", r_reply, html))
        return ("<br>" if html else "\n").join(parts)

    def set_chair_agreements(self, agreements: list):
        """Stores structured agreements and tags rows with agreement metadata."""
        self.chair_agreements = agreements or []
        self._detect_and_tag_agreements(self._data, self.chair_agreements)
        self.layoutChanged.emit()

    def _detect_and_tag_agreements(self, rows: list, chair_agreements: list = None):
        """Precomputes _has_agreement, _agreement_kind, and _agreements on row dicts."""
        tdoc_ag_map = {}
        if chair_agreements:
            for ag in chair_agreements:
                if ag.get("is_previous_meeting_recap"):
                    continue
                tid = str(ag.get("tdoc") or ag.get("source_tdoc") or "").strip().upper()
                if tid and tid != "AI LEVEL":
                    tdoc_ag_map.setdefault(tid, []).append(ag)
                ai = str(ag.get("ai") or "").strip()
                if ai:
                    syn_id = f"AI-{ai}-AGREEMENTS".upper()
                    tdoc_ag_map.setdefault(syn_id, []).append(ag)
                    syn_id2 = f"AI-{ai}-CONCLUSIONS".upper()
                    tdoc_ag_map.setdefault(syn_id2, []).append(ag)

        for row in rows:
            tdoc = str(row.get("TDoc", "")).strip().upper()
            matched_ags = tdoc_ag_map.get(tdoc, [])
            if matched_ags:
                row['_has_agreement'] = True
                row['_agreements'] = matched_ags
                kinds = [str(a.get("kind", "")).lower() for a in matched_ags]
                if any("working assumption" in k for k in kinds):
                    row['_agreement_kind'] = "Working Assumption"
                elif any("ffs" in k for k in kinds):
                    row['_agreement_kind'] = "FFS"
                else:
                    row['_agreement_kind'] = "Agreement"
            else:
                remarks = str(row.get("Secretary Remarks") or row.get("Comments") or "")
                m = re.search(r'(?:🟢|🟡|🔵)?\s*\[\s*(AGREEMENT|WORKING ASSUMPTION|CONCLUSION|DECISION|FFS)\s*\]', remarks, re.IGNORECASE)
                if m:
                    row['_has_agreement'] = True
                    k_str = m.group(1).title()
                    row['_agreement_kind'] = "Working Assumption" if "working" in k_str.lower() else ("FFS" if "ffs" in k_str.lower() else "Agreement")
                elif re.search(r'(?:^|\\n)\\s*(?:Agreements?|Working\\s+Assumptions?)\\s*[:\\n]', remarks, re.IGNORECASE):
                    row['_has_agreement'] = True
                    row['_agreement_kind'] = "Agreement"
                elif str(row.get("Type", "")).lower() == "chair notes agreement" or (tdoc.startswith("AI-") and "AGREEMENT" in tdoc):
                    row['_has_agreement'] = True
                    row['_agreement_kind'] = "Agreement"


            if row.get('_has_agreement'):
                status = str(row.get('TDoc Status') or '').strip()
                kind = row.get('_agreement_kind', 'Agreement')
                icon = "🟡" if "working assumption" in kind.lower() else ("🔵" if "ffs" in kind.lower() else "🟢")
                if not status or status.lower() in ('unknown', '', '-', 'not handled'):
                    row['TDoc Status'] = f"{icon} {kind}"
                elif status.lower() in ('noted', 'available'):
                    row['TDoc Status'] = f"{status} ({icon} {kind})"

    def _format_agreement_remarks_html(self, val_str: str, row: dict) -> str:
        """Formats executive pill badge and clean 2-line preview for agreement rows in Secretary Remarks."""
        kind = str(row.get('_agreement_kind', 'Agreement') or 'Agreement').strip()
        is_wa = "working assumption" in kind.lower()
        is_ffs = "ffs" in kind.lower() or "open issue" in kind.lower()
        is_pcr = "pcr" in kind.lower() or "p-cr" in kind.lower()
        is_cr = "cr" in kind.lower() and not is_pcr
        is_ls = "ls" in kind.lower()

        if is_wa:
            badge_html = f'<span style="background-color: {C.WARNING_SOFT_ALT}; color: {C.WARNING_DARK}; font-weight: bold; border-radius: 3px; padding: 1px 5px; font-size: 10px; border: 1px solid {C.WARNING_BORDER_STRONG};">⚖️ WORKING ASSUMPTION</span>'
            title_col = C.WARNING_DARK
        elif is_ffs:
            badge_html = f'<span style="background-color: {C.INFO_SOFT_STRONG}; color: {C.INFO_DARK}; font-weight: bold; border-radius: 3px; padding: 1px 5px; font-size: 10px; border: 1px solid {C.INFO_BORDER_MEDIUM};">⚖️ FFS / OPEN ISSUE</span>'
            title_col = C.INFO_DARK
        elif is_pcr:
            badge_html = f'<span style="background-color: {C.SUCCESS_SOFT_ALT}; color: {C.SUCCESS_DARK}; font-weight: bold; border-radius: 3px; padding: 1px 5px; font-size: 10px; border: 1px solid {C.SUCCESS_BORDER_MEDIUM};">📝 AGREED pCR</span>'
            title_col = C.SUCCESS_DARK
        elif is_cr:
            badge_html = f'<span style="background-color: {C.SUCCESS_SOFT_ALT}; color: {C.SUCCESS_DARK}; font-weight: bold; border-radius: 3px; padding: 1px 5px; font-size: 10px; border: 1px solid {C.SUCCESS_BORDER_MEDIUM};">📄 APPROVED CR</span>'
            title_col = C.SUCCESS_DARK
        elif is_ls:
            badge_html = f'<span style="background-color: {C.INDIGO_SOFT}; color: {C.WG_SA_TEXT}; font-weight: bold; border-radius: 3px; padding: 1px 5px; font-size: 10px; border: 1px solid {C.WG_SA_BORDER};">📤 APPROVED LS</span>'
            title_col = C.WG_SA_TEXT
        else:
            badge_html = f'<span style="background-color: {C.SUCCESS_SOFT_ALT}; color: {C.SUCCESS_DARK}; font-weight: bold; border-radius: 3px; padding: 1px 5px; font-size: 10px; border: 1px solid {C.SUCCESS_BORDER_MEDIUM};">⚖️ AGREEMENT</span>'
            title_col = C.SUCCESS_DARK

        txt = val_str
        if row.get('_agreements'):
            txt = row['_agreements'][0].get("text", "") or val_str
        if not txt.strip():
            txt = str(row.get("Title", "") or "")

        t = re.sub(r'^(?:🟢|🟡|🔵)?\s*\[\s*(?:AGREEMENT|WORKING ASSUMPTION|CONCLUSION|DECISION|FFS)\s*\]\s*:?', '', txt, flags=re.IGNORECASE)
        t = re.sub(r'^\s*(?:Agreements?|Working\s+Assumptions?)\s*[:\n]?', '', t, flags=re.IGNORECASE)

        raw_lines = [l.strip() for l in t.splitlines() if l.strip()]
        line1 = raw_lines[0] if raw_lines else str(row.get("Title", "") or "Agreed contribution")
        line2 = raw_lines[1] if len(raw_lines) > 1 else ""

        # For SA contributions without separate remarks, show title and source/WI
        if is_pcr or is_cr or is_ls:
            title_txt = str(row.get("Title", "") or "")
            src_txt = str(row.get("Source", "") or "")
            wi_txt = str(row.get("Related WIs", "") or "")
            line1 = title_txt if title_txt else line1
            line2 = f"Source: {src_txt}" if src_txt else (f"WI: {wi_txt}" if wi_txt else "")

        esc1 = html.escape(line1[:110] + ('...' if len(line1) > 110 else ''))
        if line2:
            esc2 = html.escape(line2[:130] + ('...' if len(line2) > 130 else ''))
            return f'<div style="line-height: 1.25;"><div>{badge_html} <span style="color: {title_col}; font-weight: bold; font-size: 11px; margin-left: 4px;">{esc1}</span></div><div style="color: {C.TEXT_BODY}; font-size: 11px; margin-top: 1px;">{esc2} <span style="color: {C.ACCENT_ALT}; font-size: 10px; margin-left: 4px;"><i>(🔍 Double-click to view)</i></span></div></div>'
        else:
            return f'<div style="line-height: 1.25;"><div>{badge_html} <span style="color: {title_col}; font-weight: bold; font-size: 11px; margin-left: 4px;">{esc1}</span></div><div style="margin-top: 1px;"><span style="color: {C.ACCENT_ALT}; font-size: 10px;"><i>(🔍 Double-click to view)</i></span></div></div>'

    def _format_agreement_tooltip_html(self, row: dict) -> str:
        """Renders rich HTML agreement card for cell hover tooltips."""
        tdoc = html.escape(str(row.get("TDoc", "") or "TDoc"))
        ai = html.escape(str(row.get("Agenda Item", "") or "-"))
        wi = html.escape(str(row.get("Related WIs", "") or "-"))
        kind = str(row.get("_agreement_kind", "Agreement") or "Agreement").strip()
        is_wa = "working assumption" in kind.lower()
        is_ffs = "ffs" in kind.lower() or "open issue" in kind.lower()

        if is_wa:
            tag_bg, tag_fg, border, card_bg = C.WARNING_SOFT_ALT, C.WARNING_DARK, C.WARNING_BORDER_STRONG, C.WARNING_SOFT
            tag_label = "🟡 3GPP Working Assumption"
        elif is_ffs:
            tag_bg, tag_fg, border, card_bg = C.INFO_SOFT_STRONG, C.INFO_DARK, C.INFO_BORDER_MEDIUM, C.INFO_SOFT
            tag_label = "🔵 3GPP FFS / Open Issue"
        else:
            tag_bg, tag_fg, border, card_bg = C.SUCCESS_SOFT_ALT, C.SUCCESS_DARK, C.SUCCESS_BORDER_MEDIUM, C.SUCCESS_SOFT
            tag_label = "🟢 3GPP Technical Agreement"

        raw_txt = ""
        if row.get('_agreements'):
            raw_txt = "\n\n".join(a.get("text", "") for a in row['_agreements'] if a.get("text"))
        if not raw_txt:
            raw_txt = str(row.get("Secretary Remarks") or row.get("Comments") or "")
            raw_txt = re.sub(r'^(?:🟢|🟡|🔵)?\s*\[\s*(?:AGREEMENT|WORKING ASSUMPTION|CONCLUSION|DECISION|FFS)\s*\]\s*:?', '', raw_txt, flags=re.IGNORECASE)
            raw_txt = re.sub(r'^\\s*(?:Agreements?|Working\\s+Assumptions?)\\s*[:\\n]?', '', raw_txt, flags=re.IGNORECASE)

        esc_txt = html.escape(raw_txt.strip())
        if len(esc_txt) > 1200:
            esc_txt = esc_txt[:1200] + "\n\n[... truncated in tooltip — double-click cell to view full agreement text ...]"

        meta_parts = []
        if ai and ai != "-":
            meta_parts.append(f"AI {ai}")
        if wi and wi != "-" and wi.lower() != "general":
            meta_parts.append(f"WI: {wi}")
        meta_str = f" • {' • '.join(meta_parts)}" if meta_parts else ""

        return f"""<div style="width: 520px; font-family: 'Segoe UI', Arial, sans-serif;">
<div style="background-color: {card_bg}; border: 1.5px solid {border}; border-radius: 6px; padding: 10px 12px;">
  <div style="margin-bottom: 6px;">
    <span style="background-color: {tag_bg}; color: {tag_fg}; font-weight: bold; border-radius: 4px; padding: 2px 7px; font-size: 11px;">{tag_label}</span>
    <span style="color: {C.TEXT_PRIMARY}; font-weight: bold; font-size: 12px; margin-left: 8px;">{tdoc}</span>
    <span style="color: {C.TEXT_MUTED}; font-size: 11px;">{meta_str}</span>
  </div>
  <div style="color: {C.TEXT_STRONG}; font-size: 12px; line-height: 1.45; white-space: pre-wrap; max-height: 320px;">{esc_txt}</div>
  <div style="color: {C.TEXT_MUTED}; font-size: 10px; margin-top: 8px; border-top: 1px dashed {C.BORDER_STRONG}; padding-top: 5px;">
    💡 <b>Tip:</b> Double-click this cell or click <b>⚖️ Agreements</b> on the toolbar to open the full Agreements Viewer.
  </div>
</div>
</div>"""
