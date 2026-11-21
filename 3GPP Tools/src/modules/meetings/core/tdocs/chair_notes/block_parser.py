"""Document-order block parser mixin (RAN1 / RAN2 tab-delimited & boxed notes)."""

from __future__ import annotations

import re
from typing import Dict, List, Optional, Tuple

from docx.document import Document as _Document
from docx.table import Table

from .text_and_results import (
    _PUA_RE,
    _RAN2_PAREN_WI_RE,
    _RAN3_WID_BRACKET_RE,
    _RELEASE_RE,
    _TDOC_ID_RE,
    _TRIVIAL_STATUS_SET,
    _classify_agreement_kind,
    _resolve_inherited_field,
    clean_chair_notes_text,
    iter_block_items,
)


class ChairNotesBlockParserMixin:
    """Extracts agreements and TDoc disposition records from document-order Word blocks."""

    @classmethod
    def _parse_block_based_docx(
        cls,
        doc: _Document,
        source_file: str,
        mtg_info: dict,
        agenda_map: dict,
    ) -> Tuple[Dict[str, dict], List[dict]]:
        agenda_data: Dict[str, dict] = {}
        agreements: List[dict] = []

        wg = str(mtg_info.get("wg_name", "")).strip().upper()
        mtg_num = str(mtg_info.get("meeting_number", "")).strip()
        mtg_date = str(mtg_info.get("start_date", "")).strip()

        current_ai = "N/A"
        current_sub_item = ""
        current_tdoc: Optional[str] = None
        last_tdoc_in_ai: Optional[str] = None
        order_in_agenda = 0
        sub_item_order_counter = 0

        ai_titles: Dict[str, str] = {}
        ai_wis: Dict[str, str] = {}
        ai_releases: Dict[str, str] = {}

        # Pre-seed from agenda_map if available
        for ai_k, item in (agenda_map or {}).items():
            if getattr(item, "description", None):
                ai_titles[str(ai_k)] = str(item.description).strip()
            if getattr(item, "acronym", None):
                ai_wis[str(ai_k)] = str(item.acronym).strip()

        tdoc_line_re = re.compile(
            r'^([A-Z][A-Z0-9]{1,4}-\d{4,8}(?:[rR]\d{1,2}[a-zA-Z]?)?)\s*(?:\t+|\s{2,}|\s*[-–:]\s*)(.*)$'
        )
        heading_ai_re = re.compile(
            r'^(?:agenda\s*item\s*)?([0-9]+(?:\.[0-9]+)*)\s*(.*)$',
            re.IGNORECASE,
        )
        agreement_header_re = re.compile(
            r'^(?:agreements?|working\s+assumptions?|conclusions?)(?:\s+(?:on|for|regarding)\b.*)?\s*:?$',
            re.IGNORECASE,
        )

        # State for multi-paragraph agreement blocks (Pattern B in RAN2)
        in_para_agreement_block = False
        para_block_header = ""
        para_block_lines: List[str] = []
        para_block_tdoc: Optional[str] = None
        para_block_ai: str = "N/A"
        consecutive_empty_lines = 0

        def _flush_para_agreement_block():
            nonlocal in_para_agreement_block, para_block_header, para_block_lines, para_block_tdoc, consecutive_empty_lines
            if not in_para_agreement_block:
                return
            if para_block_lines:
                while para_block_lines and not para_block_lines[-1].strip():
                    para_block_lines.pop()
                full_text = "\n".join(para_block_lines).strip()
                if full_text:
                    header_prefix = f"{para_block_header}\n" if para_block_header else ""
                    combined = f"{header_prefix}{full_text}".strip()
                    kind = _classify_agreement_kind(para_block_header or full_text, "Agreement")
                    _record_agreement(
                        ai_num=para_block_ai,
                        text=combined,
                        kind=kind,
                        tdoc_id=para_block_tdoc or "",
                    )
            in_para_agreement_block = False
            para_block_header = ""
            para_block_lines = []
            para_block_tdoc = None
            consecutive_empty_lines = 0

        def _record_agreement(ai_num: str, text: str, kind: str, tdoc_id: str = ""):
            clean_txt = clean_chair_notes_text(text, preserve_newlines=True)
            if not clean_txt:
                return
            resolved_wi = _resolve_inherited_field(ai_num, ai_wis)
            resolved_title = _resolve_inherited_field(ai_num, ai_titles)
            resolved_rel = _resolve_inherited_field(ai_num, ai_releases)

            # Check if text references a specific TDoc if tdoc_id is empty
            linked_tdoc = tdoc_id
            if not linked_tdoc:
                m_td = _TDOC_ID_RE.search(clean_txt)
                if m_td:
                    linked_tdoc = m_td.group(1).upper()

            if not resolved_wi and linked_tdoc and linked_tdoc in agenda_data:
                resolved_wi = str(agenda_data[linked_tdoc].get("Related WIs", "")).strip()
            if not resolved_rel and linked_tdoc and linked_tdoc in agenda_data:
                resolved_rel = str(agenda_data[linked_tdoc].get("Release", "")).strip()

            agreements.append({
                "wg": wg,
                "meeting_number": mtg_num,
                "meeting_date": mtg_date,
                "ai": ai_num if ai_num != "N/A" else "",
                "ai_title": resolved_title,
                "wi": resolved_wi,
                "release": resolved_rel,
                "kind": kind,
                "text": clean_txt,
                "tdoc_id": linked_tdoc,
                "source_file": source_file,
                "is_previous_meeting_recap": False,
            })

            # Also attach to the TDoc in agenda_data (or create a synthetic AI conclusions entry)
            target_tid = tdoc_id or last_tdoc_in_ai
            tag_label = "WORKING ASSUMPTION" if kind == "Working Assumption" else kind.upper()
            formatted_remark = f"🟢 [{tag_label}]:\n{clean_txt}"

            if target_tid and target_tid in agenda_data:
                existing_comm = agenda_data[target_tid].get("Comments", "")
                if clean_txt not in existing_comm:
                    agenda_data[target_tid]["Comments"] = (
                        f"{existing_comm}\n\n{formatted_remark}" if existing_comm else formatted_remark
                    )
                curr_res = agenda_data[target_tid].get("Result", "")
                if not curr_res or curr_res.lower() in ("noted", "unknown", ""):
                    agenda_data[target_tid]["Result"] = (
                        f"{curr_res} (🟢 {kind})" if curr_res else f"🟢 {kind}"
                    )
                if resolved_wi and not agenda_data[target_tid].get("Related WIs"):
                    agenda_data[target_tid]["Related WIs"] = resolved_wi
            elif ai_num and ai_num != "N/A":
                syn_id = f"AI-{ai_num}-AGREEMENTS"
                if syn_id not in agenda_data:
                    nonlocal order_in_agenda
                    order_in_agenda += 1
                    agenda_data[syn_id] = {
                        "Title": f"[Chair Notes Agreements] AI {ai_num}: {resolved_title}" if resolved_title else f"[Chair Notes Agreements] AI {ai_num}",
                        "Source": "Chair / Session Report",
                        "Type": "Chair Notes Agreement",
                        "For": "Agreement",
                        "Comments": formatted_remark,
                        "Result": f"🟢 {kind}",
                        "Agenda Item": ai_num,
                        "Sub-Item": current_sub_item,
                        "Related WIs": resolved_wi,
                        "Release": resolved_rel,
                        "_agenda_order": order_in_agenda,
                        "_sub_item_order": sub_item_order_counter,
                        "e-mail_Discussion": "",
                    }
                else:
                    existing_comm = agenda_data[syn_id].get("Comments", "")
                    if clean_txt not in existing_comm:
                        agenda_data[syn_id]["Comments"] = f"{existing_comm}\n\n{formatted_remark}"

        for block in iter_block_items(doc):
            if isinstance(block, Table):
                _flush_para_agreement_block()
                # Pattern A in RAN2: 1x1 or 2x1 boxed table containing Agreements
                box_texts = []
                for row in block.rows:
                    unique_cells = []
                    for c in row.cells:
                        ct = clean_chair_notes_text(c.text, preserve_newlines=True)
                        if ct and (not unique_cells or ct != unique_cells[-1]):
                            unique_cells.append(ct)
                    if unique_cells:
                        box_texts.append("\n".join(unique_cells))
                full_box = "\n".join(box_texts).strip()
                if not full_box:
                    continue
                first_line = full_box.split("\n", 1)[0].strip()
                first_low = first_line.lower()
                full_low = full_box.lower()

                # Recognize 1-column boxed consensus/guidance/agreements tables (excluding IPR / antitrust boilerplate)
                is_boilerplate = any(w in full_low for w in ("ipr declarations", "essential patent", "antitrust", "consensus principles reminder"))
                is_agreement_box = not is_boilerplate and (
                    agreement_header_re.match(first_line)
                    or any(w in first_low for w in ("agreement", "working assumption", "conclusion", "guidance", "terminology", "decision", "outcome", "way forward", "recommendation"))
                    or any(w in full_low for w in ("agreements", "working assumptions", "from ran2 point of view", "terminology to consider"))
                    or first_low.endswith("agreements")
                )
                if is_agreement_box:
                    kind = _classify_agreement_kind(first_line, "Agreement")
                    _record_agreement(
                        ai_num=current_ai,
                        text=full_box,
                        kind=kind,
                        tdoc_id=current_tdoc or last_tdoc_in_ai or "",
                    )
                continue

            # Paragraph block
            p = block
            text = clean_chair_notes_text(p.text, preserve_newlines=False)
            if not text:
                if in_para_agreement_block:
                    consecutive_empty_lines += 1
                    para_block_lines.append("")
                continue

            style = (p.style.name.lower() if p.style and p.style.name else "")

            # 1. Check if this paragraph is a Heading (Agenda Item)
            if "heading" in style or text.lower().startswith(("agenda item ", "ai ")):
                _flush_para_agreement_block()
                hm = heading_ai_re.match(text)
                if hm:
                    current_ai = hm.group(1).strip()
                    ai_title_clean = hm.group(2).strip().lstrip("-–: \t")
                    if ai_title_clean:
                        ai_titles[current_ai] = ai_title_clean
                        m_rel_h = _RELEASE_RE.search(ai_title_clean)
                        if m_rel_h:
                            ai_releases[current_ai] = f"Rel-{m_rel_h.group(1)}"
                        if "6gr" in ai_title_clean.lower() or "6g radio" in ai_title_clean.lower():
                            ai_wis[current_ai] = "FS_6G_Radio"
                    current_sub_item = ai_title_clean
                    sub_item_order_counter += 1
                    current_tdoc = None
                    last_tdoc_in_ai = None
                    continue
                elif "heading" in style:
                    # Non-numbered sub-heading (e.g. "General:", " LTM:")
                    current_sub_item = text
                    sub_item_order_counter += 1
                    current_tdoc = None
                    continue

            # 2. Check if this paragraph is a RAN2 parenthetical WI metadata line right under a heading
            # e.g. `(NR_AIML_air_Ph2, leading WG: RAN1; REL-20; WID: RP-260511)`
            m_wi_paren = _RAN2_PAREN_WI_RE.match(text)
            if m_wi_paren and current_ai != "N/A":
                wi_code = m_wi_paren.group(1).strip()
                ai_wis[current_ai] = wi_code
                m_rel = _RELEASE_RE.search(text)
                if m_rel:
                    ai_releases[current_ai] = f"Rel-{m_rel.group(1)}"
                continue

            # Also check for RAN2 `New WID: Study on 6G Radio...` or RAN3-style bracketed WID/SID in paragraph text
            if text.lower().startswith(("new wid:", "new sid:", "wid:", "sid:")) and current_ai != "N/A":
                m_rel_w = _RELEASE_RE.search(text)
                if m_rel_w:
                    ai_releases[current_ai] = f"Rel-{m_rel_w.group(1)}"
                if "6g radio" in text.lower() or "6gr" in text.lower():
                    ai_wis[current_ai] = "FS_6G_Radio"

            m_wid_br = _RAN3_WID_BRACKET_RE.search(text)
            if m_wid_br and current_ai != "N/A":
                ai_wis[current_ai] = m_wid_br.group(1).strip()

            # 3. Check if this paragraph is a TDoc entry line (e.g. `Doc-title` style or `R2-2607187\tTitle\tSource...`)
            raw_p_text = _PUA_RE.sub("", p.text or "").strip()
            tm = tdoc_line_re.match(raw_p_text)
            if tm and ("doc-title" in style or "\t" in raw_p_text or len(raw_p_text.split()) >= 3):
                _flush_para_agreement_block()
                tdoc_id = tm.group(1).upper()
                rest = tm.group(2).split("\t")
                parts = [part.strip() for part in rest if part.strip()]
                title = parts[0] if len(parts) > 0 else ""
                source = parts[1] if len(parts) > 1 else ""
                doc_type = parts[2] if len(parts) > 2 else ""
                doc_for = parts[3] if len(parts) > 3 else ""

                # In RAN2 Doc-title lines, columns 4 and 5 often hold Release (`Rel-20`) and WI (`FS_6G_Radio`, `NR_Mob_Ph5-Core`)
                row_rel = ""
                row_wi = ""
                for extra_part in parts[3:]:
                    m_r = _RELEASE_RE.match(extra_part)
                    if m_r:
                        row_rel = f"Rel-{m_r.group(1)}"
                    elif (
                        "_" in extra_part
                        or extra_part.upper().startswith(("FS_", "NR_", "LTE_", "TEI", "6GR"))
                    ) and not extra_part.lower().startswith(("to:", "cc:", "from:")):
                        row_wi = extra_part

                if row_wi and current_ai != "N/A" and current_ai not in ai_wis:
                    ai_wis[current_ai] = row_wi
                if row_rel and current_ai != "N/A" and current_ai not in ai_releases:
                    ai_releases[current_ai] = row_rel

                order_in_agenda += 1
                resolved_wi = row_wi or _resolve_inherited_field(current_ai, ai_wis)
                resolved_rel = row_rel or _resolve_inherited_field(current_ai, ai_releases)

                if tdoc_id not in agenda_data:
                    agenda_data[tdoc_id] = {
                        "Title": title,
                        "Source": source,
                        "Type": doc_type,
                        "For": doc_for,
                        "Comments": "",
                        "Result": "",
                        "Agenda Item": current_ai,
                        "Sub-Item": current_sub_item,
                        "Related WIs": resolved_wi,
                        "Release": resolved_rel,
                        "_agenda_order": order_in_agenda,
                        "_sub_item_order": sub_item_order_counter,
                        "e-mail_Discussion": "",
                    }
                current_tdoc = tdoc_id
                last_tdoc_in_ai = tdoc_id
                continue

            low_t = text.lower()

            # Subheadings and section breakers in RAN2 notes that delimit agreement blocks
            is_section_breaker = (
                low_t.startswith(("discussion", "discussion on", "discussion:", "offline discussion", "not treated"))
                or low_t.startswith(("proposal ", "proposal:", "proposals:", "observation ", "observation:"))
                or "emaildiscussion" in style
                or (consecutive_empty_lines >= 1 and (
                    low_t in _TRIVIAL_STATUS_SET
                    or low_t in ("consumer/end points", "ran decodability", "transport protocol", "model transfer", "enhancements for data transfer", "ue-side data collection", "cn vs. oam decodability")
                    or low_t.startswith(("ran awareness for standardized data collection", "transport protocol", "cp (non-ip) vs. up", "not treated"))
                ))
                or (consecutive_empty_lines >= 2 and len(text) < 60 and not any(text.startswith(b) for b in ("-", "\uf0d8", "\uf06e", "*", "•", "1", "2", "3", "4", "5", "6", "7", "8", "9", "a)", "b)", "c)", "d)")))
            )

            if in_para_agreement_block and is_section_breaker:
                _flush_para_agreement_block()

            # 4. Check if we are currently accumulating a multi-paragraph `Agreements` block (Pattern B)
            if in_para_agreement_block:
                # Does this line continue the agreement block?
                # Continues if style is doc-text / agreement / list / normal and not a trivial single-word status for a new TDoc
                if text.lower() in _TRIVIAL_STATUS_SET and "agreement" in style:
                    _flush_para_agreement_block()
                elif agreement_header_re.match(text):
                    # Starting a new agreement block right after the previous one
                    _flush_para_agreement_block()
                    in_para_agreement_block = True
                    para_block_header = text
                    para_block_lines = []
                    para_block_tdoc = current_tdoc or last_tdoc_in_ai
                    para_block_ai = current_ai
                    consecutive_empty_lines = 0
                    continue
                else:
                    para_block_lines.append(text)
                    consecutive_empty_lines = 0
                    continue

            # 5. Check if this paragraph starts a new `Agreements` / `Agreements on ...` / `Working Assumption` block (Pattern B)
            if agreement_header_re.match(text):
                in_para_agreement_block = True
                para_block_header = text
                para_block_lines = []
                para_block_tdoc = current_tdoc or last_tdoc_in_ai
                para_block_ai = current_ai
                consecutive_empty_lines = 0
                continue

            # 6. Check if this paragraph has `style == 'Agreement'` (Pattern C in RAN2)
            if "agreement" in style or "decision" in style:
                low_t = text.lower().rstrip(".")
                is_procedural_status = (
                    low_t in _TRIVIAL_STATUS_SET
                    or low_t.startswith((
                        "revised to ", "revised into ", "revised in ", "merged into ",
                        "postponed", "the cr is postponed", "cr is postponed",
                        "withdrawn", "not treated", "not handled", "not pursued",
                        "review of email", "continue the discussion", "the ls is approved",
                        "ls is approved", "the report is approved", "report is approved",
                        "the report is noted", "no tp is expected", "guidance is endorsed",
                    ))
                )
                is_agreed_tp_or_pcr = (
                    any(w in low_t for w in ("tp is endorsed", "tp is agreed", "tp is approved", "pcr is agreed", "cr is agreed"))
                )
                if current_tdoc and current_tdoc in agenda_data:
                    if not agenda_data[current_tdoc]["Result"]:
                        agenda_data[current_tdoc]["Result"] = text
                    elif text not in agenda_data[current_tdoc]["Result"]:
                        agenda_data[current_tdoc]["Result"] += f" | {text}"

                if is_agreed_tp_or_pcr:
                    _record_agreement(
                        ai_num=current_ai,
                        text=text,
                        kind="Agreed TDoc / TP / pCR",
                        tdoc_id=current_tdoc or last_tdoc_in_ai or "",
                    )
                elif not is_procedural_status and len(text) > 15:
                    kind = _classify_agreement_kind(text, "Agreement")
                    _record_agreement(
                        ai_num=current_ai,
                        text=text,
                        kind=kind,
                        tdoc_id=current_tdoc or last_tdoc_in_ai or "",
                    )
                continue

            # 7. Regular comment/discussion paragraph under `current_tdoc`
            if current_tdoc and current_tdoc in agenda_data:
                low_t = text.lower()
                is_status_line = low_t.startswith((
                    "agreed", "approved", "noted", "revised to", "revised into",
                    "postponed", "withdrawn", "endorsed", "decision:", "status:", "result:"
                ))
                if is_status_line and len(text) < 60:
                    if not agenda_data[current_tdoc]["Result"]:
                        agenda_data[current_tdoc]["Result"] = text
                    elif text not in agenda_data[current_tdoc]["Result"]:
                        agenda_data[current_tdoc]["Result"] += f" | {text}"
                else:
                    existing_comm = agenda_data[current_tdoc]["Comments"]
                    agenda_data[current_tdoc]["Comments"] = (
                        f"{existing_comm}\n{text}" if existing_comm else text
                    )

        _flush_para_agreement_block()
        return agenda_data, agreements
