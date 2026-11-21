"""Color-coded and explicit paragraph agreement extraction mixin (RAN3)."""

from __future__ import annotations

import re
from typing import Dict, List, Tuple

from docx.text.paragraph import Paragraph

from .text_and_results import (
    _TDOC_ID_RE,
    _TRIVIAL_STATUS_SET,
    _classify_agreement_kind,
    _resolve_inherited_field,
)


class ChairNotesParagraphAgreementsMixin:
    """Extracts green/blue color-coded and explicit agreements from RAN3 paragraph buffers."""

    @classmethod
    def _extract_colored_and_explicit_agreements_from_paragraphs(
        cls,
        paras: List[Tuple[str, str, Paragraph]],
        is_banner_row: bool,
        from_prev_mtg_re: re.Pattern,
        current_ai: str,
        current_sub_item: str,
        ai_titles: Dict[str, str],
        ai_wis: Dict[str, str],
        ai_releases: Dict[str, str],
        wg: str,
        mtg_num: str,
        mtg_date: str,
        source_file: str,
        default_tdoc_id: str,
        agenda_data: Dict[str, dict],
        agreements: List[dict],
        order_counter_ref: List[int],
    ):
        """
        Extracts both RAN3 color-coded agreements (Green = Agreement/WA, Blue = FFS)
        and explicit textual `Agreement:` / `Working Assumption:` blocks from a list of paragraphs.
        """
        in_prev_mtg_recap = False
        referenced_summary_tdoc = default_tdoc_id

        # First pass: check if this merged discussion cell references a Moderator/Offline Summary TDoc
        # e.g., `Summary of offline R3-264513 Noted` or `Moderator summary R3-264501`
        if not referenced_summary_tdoc:
            for txt, _, _ in paras:
                m_sum = re.search(
                    r'(?:summary|offline|moderator|rapporteur|tp\b).*?\b([A-Z][A-Z0-9]{1,3}-\d{4,8}(?:[rR]\d{1,2}[a-zA-Z]?)?)\b',
                    txt,
                    re.IGNORECASE,
                )
                if m_sum:
                    referenced_summary_tdoc = m_sum.group(1).upper()
                    break

        green_buffer: List[str] = []
        blue_buffer: List[str] = []
        explicit_buffer: List[str] = []
        in_explicit_block = False

        resolved_wi = _resolve_inherited_field(current_ai, ai_wis)
        resolved_title = _resolve_inherited_field(current_ai, ai_titles) or current_sub_item
        resolved_rel = _resolve_inherited_field(current_ai, ai_releases)

        def _emit_item(lines: List[str], default_kind: str, is_recap: bool):
            if not lines:
                return
            text_joined = "\n".join(lines).strip()
            if not text_joined:
                return
            low_j = text_joined.lower().rstrip(".")
            if low_j in _TRIVIAL_STATUS_SET:
                return
            # Filter out short disposition lines like `Rev in R3-264569 Agreed` or `Summary of offline R3-264513 Noted`
            if len(lines) == 1 and len(text_joined) < 55 and any(
                w in low_j for w in ("noted", "withdrawn", "postponed", "not treated", "rev in ", "revised in ")
            ):
                # However, if it says `Rev in R3-264569 Agreed` or `TP in R3-264569 Agreed`, record as Agreed TDoc / TP!
                if "agreed" in low_j or "approved" in low_j or "endorsed" in low_j:
                    m_td = _TDOC_ID_RE.search(text_joined)
                    linked = m_td.group(1).upper() if m_td else referenced_summary_tdoc
                    agreements.append({
                        "wg": wg,
                        "meeting_number": mtg_num,
                        "meeting_date": mtg_date,
                        "ai": current_ai if current_ai != "N/A" else "",
                        "ai_title": resolved_title,
                        "wi": resolved_wi,
                        "release": resolved_rel,
                        "kind": "Agreed TDoc / TP / pCR",
                        "text": text_joined,
                        "tdoc_id": linked,
                        "source_file": source_file,
                        "is_previous_meeting_recap": is_recap,
                    })
                return

            kind = _classify_agreement_kind(text_joined, default_kind)
            linked_tdoc = referenced_summary_tdoc
            if not linked_tdoc:
                m_td = _TDOC_ID_RE.search(text_joined)
                if m_td:
                    linked_tdoc = m_td.group(1).upper()

            agreements.append({
                "wg": wg,
                "meeting_number": mtg_num,
                "meeting_date": mtg_date,
                "ai": current_ai if current_ai != "N/A" else "",
                "ai_title": resolved_title,
                "wi": resolved_wi,
                "release": resolved_rel,
                "kind": kind,
                "text": text_joined,
                "tdoc_id": linked_tdoc,
                "source_file": source_file,
                "is_previous_meeting_recap": is_recap,
            })

            if is_recap:
                return

            # Also surface in agenda_data so `TDocsWindow` displays the agreement!
            icon = "🔵" if "FFS" in kind else ("🟡" if kind == "Working Assumption" else "🟢")
            formatted_block = f"{icon} [{kind.upper()}]:\n{text_joined}"

            if linked_tdoc and linked_tdoc in agenda_data:
                existing_c = agenda_data[linked_tdoc].get("Comments", "")
                if text_joined not in existing_c:
                    agenda_data[linked_tdoc]["Comments"] = (
                        f"{existing_c}\n\n{formatted_block}" if existing_c else formatted_block
                    )
                curr_res = agenda_data[linked_tdoc].get("Result", "")
                if "🟢" not in curr_res and kind in ("Agreement", "Working Assumption", "Conclusion"):
                    agenda_data[linked_tdoc]["Result"] = (
                        f"{curr_res} ({icon} {kind})" if curr_res else f"{icon} {kind}"
                    )
            elif current_ai and current_ai != "N/A" and not default_tdoc_id:
                syn_id = referenced_summary_tdoc or f"AI-{current_ai}-CONCLUSIONS"
                if syn_id not in agenda_data:
                    order_counter_ref[0] += 1
                    agenda_data[syn_id] = {
                        "Title": (
                            f"[AI {current_ai} Conclusions] {resolved_title}"
                            if resolved_title
                            else f"[AI {current_ai} Conclusions]"
                        ),
                        "Source": "Chair / Offline Session",
                        "Type": "Chair Notes Agreement",
                        "For": "Agreement",
                        "Comments": formatted_block,
                        "Result": f"{icon} {kind}",
                        "Agenda Item": current_ai,
                        "Sub-Item": current_sub_item,
                        "Related WIs": resolved_wi,
                        "Release": resolved_rel,
                        "_agenda_order": order_counter_ref[0],
                        "_sub_item_order": order_counter_ref[1],
                        "e-mail_Discussion": "",
                    }
                else:
                    existing_c = agenda_data[syn_id].get("Comments", "")
                    if text_joined not in existing_c:
                        agenda_data[syn_id]["Comments"] = (
                            f"{existing_c}\n\n{formatted_block}" if existing_c else formatted_block
                        )

        for txt, col, _ in paras:
            if from_prev_mtg_re.match(txt):
                _emit_item(green_buffer, "Agreement", in_prev_mtg_recap)
                green_buffer = []
                _emit_item(blue_buffer, "FFS / Open Issue", in_prev_mtg_recap)
                blue_buffer = []
                in_prev_mtg_recap = True
                continue

            if col == "GREEN":
                _emit_item(blue_buffer, "FFS / Open Issue", in_prev_mtg_recap)
                blue_buffer = []
                _emit_item(explicit_buffer, "Agreement", in_prev_mtg_recap)
                explicit_buffer = []
                in_explicit_block = False

                # If a new "Proposal X:" or "Agreement:" starts, flush previous green buffer
                if green_buffer and re.match(r'^(?:proposal\s*\d*|agreement\s*\d*|working\s+assumption|conclusion|wa\s*:)', txt, re.IGNORECASE):
                    _emit_item(green_buffer, "Agreement", in_prev_mtg_recap)
                    green_buffer = []
                green_buffer.append(txt)
                continue

            if col == "BLUE":
                _emit_item(green_buffer, "Agreement", in_prev_mtg_recap)
                green_buffer = []
                _emit_item(explicit_buffer, "Agreement", in_prev_mtg_recap)
                explicit_buffer = []
                in_explicit_block = False

                if blue_buffer and re.match(r'^(?:ffs|open\s+issue|to\s+be\s+continued)', txt, re.IGNORECASE):
                    _emit_item(blue_buffer, "FFS / Open Issue", in_prev_mtg_recap)
                    blue_buffer = []
                blue_buffer.append(txt)
                continue

            # Non-colored paragraph: flush color buffers
            _emit_item(green_buffer, "Agreement", in_prev_mtg_recap)
            green_buffer = []
            _emit_item(blue_buffer, "FFS / Open Issue", in_prev_mtg_recap)
            blue_buffer = []

            # Check for uncolored explicit Agreement / Working Assumption blocks (e.g. SA2 or uncolored notes)
            if re.match(r'^(?:agreements?|working\s+assumptions?|conclusions?)\s*[:\-]', txt, re.IGNORECASE):
                _emit_item(explicit_buffer, "Agreement", in_prev_mtg_recap)
                explicit_buffer = [txt]
                in_explicit_block = True
            elif in_explicit_block:
                if txt.startswith(("-", "•", "*", "1", "2", "3", "4", "5", "a)", "b)", "c)")):
                    explicit_buffer.append(txt)
                else:
                    _emit_item(explicit_buffer, "Agreement", in_prev_mtg_recap)
                    explicit_buffer = []
                    in_explicit_block = False

        _emit_item(green_buffer, "Agreement", in_prev_mtg_recap)
        _emit_item(blue_buffer, "FFS / Open Issue", in_prev_mtg_recap)
        _emit_item(explicit_buffer, "Agreement", in_prev_mtg_recap)
