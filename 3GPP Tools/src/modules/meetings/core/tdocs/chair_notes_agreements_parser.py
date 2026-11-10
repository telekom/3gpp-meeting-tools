# --- File: src/modules/meetings/core/tdocs/chair_notes_agreements_parser.py ---
"""
Unified Chairman's Notes & Agreements Parser across 3GPP Working Groups (RAN1, RAN2, RAN3, SA2, etc.).

Handles the three distinct Chairman's Notes authoring idioms used across 3GPP:
1. RAN2 / RAN1 Document-Order Block Format:
   - Tab-delimited TDoc paragraphs (`R2-2607187\tTitle\tSource\tType\tFor`),
   - Agenda Item headings (`Heading 1..5`) with `(WI_Code, leading WG: ..., REL-20; WID: RP-...)` metadata,
   - 1-column (or 2-column) boxed Word tables containing `Agreements` / `Working Assumption` blocks,
   - Paragraph sequences starting with `Agreements` / `Agreements on ...` / `Working Assumption`,
   - Standalone `(Agreement)` styled paragraphs.
2. RAN3 Color-Coded & Banner-Segmented 3-Column Word Table (`TDoc | Title | Comments`):
   - Shaded merged 3-column Agenda Item banners declaring `AI #`, `Title`, `WID [code]: RP-...` / `SID [code]: RP-...`,
   - `From RAN3#<prev>:` recap lines inside banner rows (distinguished from current-meeting agreements),
   - Merged 3-column moderator summary / offline discussion rows at the end of Agenda Items,
   - Color-coded runs/paragraphs:
     * Green (`#008000`, `#00B050`, `#538135`) -> Official Agreements, Working Assumptions (`WA:`), Conclusions
     * Blue (`#0000FF`, `#3333FF`) -> Open Issues / `FFS` / "To be continued"
3. SA2 / General Multi-Column TDoc Table (`TD# | Title | Source | ... | Comments | Result`):
   - Extracts explicit agreements, conclusions, and agreed pCRs/CRs/LSs from Remarks/Result cells.
"""
import json
import logging
import os
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import docx
from docx.document import Document as _Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph

from core.utils.paths import to_extended_path

logger = logging.getLogger(__name__)

# Private Use Area characters (Wingdings bullets like \uf0d8, \uf0b7, \uf0a7)
_PUA_RE = re.compile(r'[\uf000-\uf8ff]')

# Trivial 1-word TDoc disposition statuses that are NOT substantive technical agreements on their own
_TRIVIAL_STATUS_SET = {
    "noted",
    "approved",
    "agreed",
    "endorsed",
    "postponed",
    "withdrawn",
    "revised",
    "not treated",
    "not handled",
    "not pursued",
    "no agreement",
    "tbd",
    "ffs",
    "handled",
    "merged",
    "seen",
}

_TDOC_ID_RE = re.compile(
    r'\b([A-Z][A-Z0-9]{1,3}-\d{4,8}(?:[rR]\d{1,2}[a-zA-Z]?)?)\b'
)

# Matches RAN3 WID/SID bracket tags: e.g., `SID [FS_6G_Radio]: RP-261559` or `WID [NR_AIML_NGRAN_Ph3-Core]: RP-260773`
_RAN3_WID_BRACKET_RE = re.compile(
    r'(?:WID|SID|WI|SI)\s*\[\s*([A-Za-z0-9_\-\/]+)\s*\]',
    re.IGNORECASE,
)

# Matches RAN2 parenthetical WI metadata under headings:
# e.g. `(NR_AIML_air_Ph2, leading WG: RAN1; REL-20; WID: RP-260511)`
# or `(FS_NR_ish_redcap, leading WG: RAN2, REL-19, SID: RP-240899)`
_RAN2_PAREN_WI_RE = re.compile(
    r'^\(\s*([A-Za-z0-9]+_[A-Za-z0-9_\-]+)\s*[,;)]'
)

# Matches Release tags like REL-19, Rel-20, Release 19
_RELEASE_RE = re.compile(r'\b(?:REL|Release)[-_\s]*(\d{1,2})\b', re.IGNORECASE)

# Green hex colors used in RAN3 Chair Notes for Agreements / Working Assumptions
_GREEN_HEX_PREFIXES = ("008000", "00B050", "538135", "385723", "006600", "009900", "2E7D32")
# Blue hex colors used in RAN3 Chair Notes for FFS / To be continued
_BLUE_HEX_PREFIXES = ("0000FF", "3333FF", "0070C0", "002060", "2F5597")


def clean_chair_notes_text(text: str, preserve_newlines: bool = True) -> str:
    """Strips Wingdings Private Use Area bullets and normalizes whitespace."""
    if not text:
        return ""
    s = _PUA_RE.sub("", str(text))
    s = s.replace("\r\n", "\n").replace("\r", "\n")
    if preserve_newlines:
        lines = []
        for line in s.split("\n"):
            cl = re.sub(r'[ \t]+', ' ', line).strip()
            if cl:
                lines.append(cl)
        return "\n".join(lines).strip()
    return re.sub(r'\s+', ' ', s).strip()


def iter_block_items(parent):
    """
    Yields each paragraph and table child within `parent` (Document or _Cell)
    in true document order.
    """
    if isinstance(parent, _Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        raise ValueError(f"Unsupported parent type: {type(parent)}")

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


def _get_paragraph_dominant_color(p: Paragraph) -> str:
    """
    Inspects OpenXML run colors (`w:rPr/w:color/@w:val`) in paragraph `p`.
    Returns:
      - `"GREEN"` if green runs dominate the non-whitespace text,
      - `"BLUE"` if blue runs dominate,
      - `""` otherwise.
    """
    green_chars = 0
    blue_chars = 0
    other_chars = 0

    for run in p.runs:
        txt = _PUA_RE.sub("", run.text or "").strip()
        if not txt:
            continue
        n = len(txt)
        color_hex = ""
        try:
            if run.font and run.font.color and run.font.color.rgb:
                color_hex = str(run.font.color.rgb).upper()
        except Exception:
            color_hex = ""

        if not color_hex:
            # Check XML directly in case python-docx misses theme/val
            try:
                rPr = run._r.rPr
                if rPr is not None:
                    c_elms = rPr.findall("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}color")
                    if c_elms:
                        color_hex = (c_elms[0].get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val") or "").upper()
            except Exception:
                pass

        if color_hex in _GREEN_HEX_PREFIXES:
            green_chars += n
        elif color_hex in _BLUE_HEX_PREFIXES:
            blue_chars += n
        else:
            other_chars += n

    total = green_chars + blue_chars + other_chars
    if total == 0:
        return ""
    if green_chars > 0 and green_chars >= (blue_chars + other_chars):
        return "GREEN"
    if blue_chars > 0 and blue_chars >= (green_chars + other_chars):
        return "BLUE"
    # Also handle lines where a company prefix or bullet is black but the body is green
    if green_chars >= 15 and green_chars >= total * 0.45:
        return "GREEN"
    if blue_chars >= 15 and blue_chars >= total * 0.45:
        return "BLUE"
    return ""


def _classify_agreement_kind(text: str, default_kind: str = "Agreement") -> str:
    """Classifies whether a text block is an Agreement, Working Assumption, Conclusion, or FFS."""
    low = text.strip().lower()
    if low.startswith(("working assumption", "wa:", "wa ")):
        return "Working Assumption"
    if low.startswith(("conclusion", "conclusions")):
        return "Conclusion"
    if low.startswith(("ffs", "open issue", "to be continued", "for further study")):
        return "FFS / Open Issue"
    return default_kind


def _resolve_inherited_field(ai_num: str, mapping: Dict[str, str]) -> str:
    """
    Looks up `ai_num` in `mapping` (e.g. `10.2.1`), falling back up the parent hierarchy
    (`10.2` -> `10`) when child sub-items inherit their parent Agenda Item's WI/SI or Title.
    """
    if not ai_num or ai_num == "N/A":
        return ""
    if ai_num in mapping and mapping[ai_num]:
        return mapping[ai_num]
    parts = ai_num.split(".")
    for i in range(len(parts) - 1, 0, -1):
        prefix = ".".join(parts[:i])
        if prefix in mapping and mapping[prefix]:
            return mapping[prefix]
    return ""


class ChairNotesAgreementsParser:
    """
    Parses `.docx` and `.htm`/`.html` Chairman's Notes into:
      1. `agenda_data`: Dict[tdoc_id, dict] compatible with `TDocsTableModel.merge_agenda_data`
      2. `agreements`: List[dict] of structured `MeetingAgreement` records
    """

    @classmethod
    def parse_docx_with_agreements(
        cls,
        docx_path: str,
        mtg_info: Optional[dict] = None,
        agenda_map: Optional[dict] = None,
    ) -> Tuple[Dict[str, dict], List[dict]]:
        path = Path(docx_path)
        if not path.exists():
            logger.error(f"[ChairNotesAgreementsParser] File not found: {docx_path}")
            return {}, []

        try:
            doc = docx.Document(str(path))
        except Exception as e:
            logger.error(f"[ChairNotesAgreementsParser] Failed to open docx '{docx_path}': {e}")
            return {}, []

        mtg_info = mtg_info or {}
        agenda_map = agenda_map or {}
        source_file = path.name

        # Detect whether this document has a main multi-row TDoc table (RAN3 / SA2)
        # or is a document-order Heading + Paragraph + 1x1 Boxed Table document (RAN2 / RAN1)
        main_tables = cls._find_main_tdoc_tables(doc)

        if main_tables:
            agenda_data, agreements = cls._parse_table_based_docx(
                doc, main_tables, source_file, mtg_info, agenda_map
            )
        else:
            agenda_data, agreements = cls._parse_block_based_docx(
                doc, source_file, mtg_info, agenda_map
            )

        logger.info(
            f"[ChairNotesAgreementsParser] Parsed {len(agenda_data)} TDoc entries and "
            f"{len(agreements)} structured agreements from {source_file}"
        )
        return agenda_data, agreements

    @staticmethod
    def _find_main_tdoc_tables(doc: _Document) -> List[Tuple[Table, dict, int]]:
        """
        Finds tables in `doc` that represent a multi-row TDoc schedule/notes table
        (having >= 3 rows and a recognizable `tdoc` header column).
        """
        from modules.meetings.core.tdocs.tdocs_parser import TDocsParser

        found = []
        for table in doc.tables:
            if len(table.rows) < 2:
                continue
            col_map = {}
            header_row_idx = 0
            max_scan = min(5, len(table.rows))
            for r_idx in range(max_scan):
                candidate_cells = [cell.text for cell in table.rows[r_idx].cells]
                candidate_map = TDocsParser._map_docx_header_row(candidate_cells)
                if "tdoc" in candidate_map and len(candidate_map) >= 2:
                    col_map = candidate_map
                    header_row_idx = r_idx
                    break
            if not col_map:
                candidate_cells = [cell.text for cell in table.rows[0].cells]
                candidate_map = TDocsParser._map_docx_header_row(candidate_cells)
                if "tdoc" in candidate_map:
                    col_map = candidate_map
                    header_row_idx = 0
            if col_map and "tdoc" in col_map:
                found.append((table, col_map, header_row_idx))
        return found

    # =========================================================================
    # 1. RAN2 / RAN1 Document-Order Block Parser
    # =========================================================================

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

        def _flush_para_agreement_block():
            nonlocal in_para_agreement_block, para_block_header, para_block_lines, para_block_tdoc, para_block_ai
            if not in_para_agreement_block:
                return
            if para_block_lines:
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
                if (
                    agreement_header_re.match(first_line)
                    or full_box.lower().startswith(("agreement", "working assumption", "conclusion"))
                ):
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
                _flush_para_agreement_block()
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
                    continue
                else:
                    para_block_lines.append(text)
                    continue

            # 5. Check if this paragraph starts a new `Agreements` / `Agreements on ...` / `Working Assumption` block (Pattern B)
            if agreement_header_re.match(text):
                in_para_agreement_block = True
                para_block_header = text
                para_block_lines = []
                para_block_tdoc = current_tdoc or last_tdoc_in_ai
                para_block_ai = current_ai
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

    # =========================================================================
    # 2. RAN3 & SA2 Table-Based Parser (with OpenXML Color & Banner Intelligence)
    # =========================================================================

    @classmethod
    def _parse_table_based_docx(
        cls,
        doc: _Document,
        main_tables: List[Tuple[Table, dict, int]],
        source_file: str,
        mtg_info: dict,
        agenda_map: dict,
    ) -> Tuple[Dict[str, dict], List[dict]]:
        agenda_data: Dict[str, dict] = {}
        agreements: List[dict] = []

        wg = str(mtg_info.get("wg_name", "")).strip().upper()
        mtg_num = str(mtg_info.get("meeting_number", "")).strip()
        mtg_date = str(mtg_info.get("start_date", "")).strip()

        ai_titles: Dict[str, str] = {}
        ai_wis: Dict[str, str] = {}
        ai_releases: Dict[str, str] = {}

        for ai_k, item in (agenda_map or {}).items():
            if getattr(item, "description", None):
                ai_titles[str(ai_k)] = str(item.description).strip()
            if getattr(item, "acronym", None):
                ai_wis[str(ai_k)] = str(item.acronym).strip()

        order_in_agenda = 0
        sub_item_order_counter = 0

        # Regex for RAN3 numbered banner headers: e.g. `10. Study on 6G Radio`, `10.2.1. Data collection`
        ran3_banner_ai_re = re.compile(
            r'^(\d+(?:\.\d+)*)\.?\s+(.+)$'
        )
        explicit_ai_banner_re = re.compile(
            r'(?:agenda\s*item|agenda\s*no\.?|ai)\s*[:#]?\s*([0-9]+(?:\.[0-9]+)*[a-zA-Z]?)\s*[-–:]?\s*(.*)',
            re.IGNORECASE,
        )
        from_prev_mtg_re = re.compile(
            r'^from\s+(?:ran\d?|sa\d?|ct\d?|wg\d?)\s*#?\s*\d+',
            re.IGNORECASE,
        )

        for table, col_map, header_row_idx in main_tables:
            current_ai = "N/A"
            current_sub_item = ""
            last_tdoc_in_ai = ""

            for row in table.rows[header_row_idx + 1:]:
                raw_cells = row.cells
                if not raw_cells or len(raw_cells) <= col_map["tdoc"]:
                    continue

                cells_text = [clean_chair_notes_text(c.text, preserve_newlines=True) for c in raw_cells]
                raw_tdoc = cells_text[col_map["tdoc"]].strip()

                # Check if this is a merged banner/discussion row (or non-TDoc row)
                td_match = None
                if raw_tdoc and raw_tdoc != "-":
                    # Only treat col 0 as a TDoc ID cell if the cell is short (not a merged full-width discussion paragraph!)
                    first_line_td = raw_tdoc.split("\n", 1)[0].strip()
                    if len(first_line_td) <= 35:
                        td_match = _TDOC_ID_RE.search(first_line_td)

                is_merged_full_row = (
                    len(raw_cells) >= 2
                    and raw_cells[0]._tc is raw_cells[-1]._tc
                ) or (
                    len(cells_text) >= 2
                    and cells_text[0] == cells_text[-1]
                    and len(cells_text[0]) > 0
                )

                if not td_match or is_merged_full_row:
                    # Non-TDoc row! In RAN3 / SA2 this is either:
                    # (a) An Agenda Item / Sub-Item Header Banner row, OR
                    # (b) A Standalone Moderator Summary / Offline Session Conclusion row!
                    cell_obj = raw_cells[0]
                    paras = [
                        (clean_chair_notes_text(p.text, preserve_newlines=False), _get_paragraph_dominant_color(p), p)
                        for p in cell_obj.paragraphs
                    ]
                    paras = [(txt, col, p) for txt, col, p in paras if txt]
                    if not paras:
                        continue

                    first_txt = paras[0][0]

                    # Check if first line defines a new Agenda Item banner
                    new_ai_detected = False
                    m_exp = explicit_ai_banner_re.search(first_txt)
                    if m_exp:
                        current_ai = m_exp.group(1).strip()
                        rest_title = m_exp.group(2).strip()
                        if rest_title:
                            ai_titles[current_ai] = rest_title
                            current_sub_item = rest_title
                        else:
                            current_sub_item = ""
                        new_ai_detected = True
                        last_tdoc_in_ai = ""
                    else:
                        m_num = ran3_banner_ai_re.match(first_txt)
                        # Ensure it's a real AI number (like `10.` or `10.2.1.`) and not a numbered agreement list (`1. RAN3 agrees...`)
                        if m_num:
                            cand_ai = m_num.group(1).strip()
                            cand_title = m_num.group(2).strip()
                            is_heading_like = (
                                "." in cand_ai
                                or (
                                    paras[0][1] != "GREEN"
                                    and not cand_title.lower().startswith(("ran", "sa", "ct", "agree", "the ", "for ", "if ", "when ", "in "))
                                    and len(cand_title) < 120
                                )
                            )
                            if is_heading_like:
                                current_ai = cand_ai
                                ai_titles[current_ai] = cand_title
                                current_sub_item = cand_title
                                sub_item_order_counter += 1
                                new_ai_detected = True
                                last_tdoc_in_ai = ""

                    # Extract any WID / SID / Release metadata present in the banner
                    full_cell_text = "\n".join(t for t, _, _ in paras)
                    m_wid = _RAN3_WID_BRACKET_RE.search(full_cell_text)
                    if m_wid and current_ai != "N/A":
                        ai_wis[current_ai] = m_wid.group(1).strip()
                    m_rel = _RELEASE_RE.search(full_cell_text)
                    if m_rel and current_ai != "N/A":
                        ai_releases[current_ai] = f"Rel-{m_rel.group(1)}"

                    # Check if this row contains green agreements or blue FFS items!
                    cls._extract_colored_and_explicit_agreements_from_paragraphs(
                        paras=paras,
                        is_banner_row=new_ai_detected,
                        from_prev_mtg_re=from_prev_mtg_re,
                        current_ai=current_ai,
                        current_sub_item=current_sub_item,
                        ai_titles=ai_titles,
                        ai_wis=ai_wis,
                        ai_releases=ai_releases,
                        wg=wg,
                        mtg_num=mtg_num,
                        mtg_date=mtg_date,
                        source_file=source_file,
                        default_tdoc_id="",
                        agenda_data=agenda_data,
                        agreements=agreements,
                        order_counter_ref=[order_in_agenda, sub_item_order_counter],
                    )
                    order_in_agenda = max(order_in_agenda, len(agenda_data))

                    if not new_ai_detected and not is_merged_full_row:
                        title_candidate = (
                            cells_text[col_map["title"]].strip()
                            if "title" in col_map and col_map["title"] < len(cells_text)
                            else ""
                        )
                        if title_candidate and title_candidate != "-" and not title_candidate.lower().startswith("all agenda"):
                            current_sub_item = title_candidate
                            sub_item_order_counter += 1
                    continue

                # Valid TDoc Row!
                tdoc_id = td_match.group(1).upper()
                order_in_agenda += 1
                last_tdoc_in_ai = tdoc_id

                if "ai" in col_map and col_map["ai"] < len(cells_text):
                    ai_val = cells_text[col_map["ai"]].strip()
                    if ai_val and ai_val != "-":
                        if ai_val != current_ai:
                            current_ai = ai_val
                            current_sub_item = ""

                title = cells_text[col_map["title"]] if "title" in col_map and col_map["title"] < len(cells_text) else ""
                source = cells_text[col_map["source"]] if "source" in col_map and col_map["source"] < len(cells_text) else ""
                doc_type = cells_text[col_map["type"]] if "type" in col_map and col_map["type"] < len(cells_text) else ""
                doc_for = cells_text[col_map["for"]] if "for" in col_map and col_map["for"] < len(cells_text) else ""
                comments = cells_text[col_map["comments"]] if "comments" in col_map and col_map["comments"] < len(cells_text) else ""
                result = cells_text[col_map["result"]] if "result" in col_map and col_map["result"] < len(cells_text) else ""

                # In RAN3 3-column tables (`TDoc | Title | Comments`), there is no separate `Result` or `Source` column:
                # - `Title` often includes `(Source, Type)` or TDoc metadata comes from Excel
                # - The final line of `Comments` (often in green `#008000`) is the disposition (`Noted`, `Agreed`, `Rev in R3-264569 Agreed`)
                if "result" not in col_map and "comments" in col_map and col_map["comments"] < len(raw_cells):
                    comm_cell = raw_cells[col_map["comments"]]
                    comm_paras = [
                        (clean_chair_notes_text(p.text, preserve_newlines=False), _get_paragraph_dominant_color(p), p)
                        for p in comm_cell.paragraphs
                    ]
                    comm_paras = [(t, c, p) for t, c, p in comm_paras if t]
                    if comm_paras:
                        last_line_txt, last_line_col, _ = comm_paras[-1]
                        if (
                            len(last_line_txt) <= 55
                            and any(
                                kw in last_line_txt.lower()
                                for kw in ("noted", "agreed", "approved", "endorsed", "postponed", "withdrawn", "revised", "rev in ")
                            )
                        ):
                            result = last_line_txt

                    # Also check if the TDoc's Comments cell contains substantive green agreements or explicit Agreements!
                    cls._extract_colored_and_explicit_agreements_from_paragraphs(
                        paras=comm_paras,
                        is_banner_row=False,
                        from_prev_mtg_re=from_prev_mtg_re,
                        current_ai=current_ai,
                        current_sub_item=current_sub_item,
                        ai_titles=ai_titles,
                        ai_wis=ai_wis,
                        ai_releases=ai_releases,
                        wg=wg,
                        mtg_num=mtg_num,
                        mtg_date=mtg_date,
                        source_file=source_file,
                        default_tdoc_id=tdoc_id,
                        agenda_data=agenda_data,
                        agreements=agreements,
                        order_counter_ref=[order_in_agenda, sub_item_order_counter],
                    )
                elif "comments" in col_map and col_map["comments"] < len(raw_cells):
                    # SA2 / general table with separate Comments and Result columns
                    comm_cell = raw_cells[col_map["comments"]]
                    comm_paras = [
                        (clean_chair_notes_text(p.text, preserve_newlines=False), _get_paragraph_dominant_color(p), p)
                        for p in comm_cell.paragraphs
                    ]
                    comm_paras = [(t, c, p) for t, c, p in comm_paras if t]
                    if comm_paras:
                        cls._extract_colored_and_explicit_agreements_from_paragraphs(
                            paras=comm_paras,
                            is_banner_row=False,
                            from_prev_mtg_re=from_prev_mtg_re,
                            current_ai=current_ai,
                            current_sub_item=current_sub_item,
                            ai_titles=ai_titles,
                            ai_wis=ai_wis,
                            ai_releases=ai_releases,
                            wg=wg,
                            mtg_num=mtg_num,
                            mtg_date=mtg_date,
                            source_file=source_file,
                            default_tdoc_id=tdoc_id,
                            agenda_data=agenda_data,
                            agreements=agreements,
                            order_counter_ref=[order_in_agenda, sub_item_order_counter],
                        )

                resolved_wi = _resolve_inherited_field(current_ai, ai_wis)
                resolved_rel = _resolve_inherited_field(current_ai, ai_releases)

                if tdoc_id in agenda_data:
                    existing = agenda_data[tdoc_id]
                    if comments and comments != existing.get("Comments", ""):
                        if existing.get("Comments") and comments not in existing["Comments"]:
                            comments = f"{existing['Comments']}\n---\n{comments}"
                        else:
                            comments = existing.get("Comments") or comments
                    else:
                        comments = comments or existing.get("Comments", "")

                    result = result or existing.get("Result", "")
                    title = title or existing.get("Title", "")
                    source = source or existing.get("Source", "")
                    doc_type = doc_type or existing.get("Type", "")
                    doc_for = doc_for or existing.get("For", "")
                    if current_ai == "N/A" and existing.get("Agenda Item"):
                        current_ai = existing.get("Agenda Item")
                    if not current_sub_item and existing.get("Sub-Item"):
                        current_sub_item = existing.get("Sub-Item")
                    resolved_wi = resolved_wi or existing.get("Related WIs", "")
                    resolved_rel = resolved_rel or existing.get("Release", "")

                agenda_data[tdoc_id] = {
                    "Title": title,
                    "Source": source,
                    "Type": doc_type,
                    "For": doc_for,
                    "Comments": comments,
                    "Result": result,
                    "Agenda Item": current_ai,
                    "Sub-Item": current_sub_item,
                    "Related WIs": resolved_wi,
                    "Release": resolved_rel,
                    "_agenda_order": order_in_agenda,
                    "_sub_item_order": sub_item_order_counter,
                    "e-mail_Discussion": "",
                }

        return agenda_data, agreements

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

    # =========================================================================
    # 3. Persistence & Cache Helpers (`Agenda/chair_notes_agreements.json`)
    # =========================================================================

    @staticmethod
    def save_cached_chair_notes_data(
        meeting_dir: Path,
        agenda_data: Dict[str, dict],
        agreements: List[dict],
    ):
        """
        Persists merged `agenda_data` and `agreements` into `Agenda/chair_notes_cache.json`
        inside `meeting_dir` so opening `TDocsWindow` or running `PostMeetingExporterThread`
        or `AgreementTrackerWorker` can load them instantaneously.
        """
        if not meeting_dir:
            return
        try:
            agenda_dir = Path(meeting_dir) / "Agenda"
            os.makedirs(to_extended_path(agenda_dir), exist_ok=True)
            cache_path = agenda_dir / "chair_notes_cache.json"

            existing_agenda = {}
            existing_agreements = []
            if os.path.exists(to_extended_path(cache_path)):
                try:
                    with open(to_extended_path(cache_path), "r", encoding="utf-8") as f:
                        loaded = json.load(f)
                        existing_agenda = loaded.get("agenda_data", {})
                        existing_agreements = loaded.get("agreements", [])
                except Exception:
                    pass

            # Merge agenda_data
            for tid, info in (agenda_data or {}).items():
                if tid in existing_agenda:
                    old = existing_agenda[tid]
                    for k, v in info.items():
                        if v:
                            old[k] = v
                else:
                    existing_agenda[tid] = info

            # Merge agreements (deduplicate by (ai, kind, text[:120]))
            seen_keys = {
                (a.get("ai", ""), a.get("kind", ""), a.get("text", "")[:120])
                for a in existing_agreements
            }
            for ag in (agreements or []):
                key = (ag.get("ai", ""), ag.get("kind", ""), ag.get("text", "")[:120])
                if key not in seen_keys:
                    seen_keys.add(key)
                    existing_agreements.append(ag)

            payload = {
                "agenda_data": existing_agenda,
                "agreements": existing_agreements,
            }
            with open(to_extended_path(cache_path), "w", encoding="utf-8") as f:
                json.dump(payload, f, indent=2, ensure_ascii=False)
        except Exception as e:
            logger.warning(f"[ChairNotesAgreementsParser] Failed to save chair notes cache: {e}")

    @staticmethod
    def load_cached_chair_notes_data(meeting_dir: Path) -> Tuple[Dict[str, dict], List[dict]]:
        """Loads cached `(agenda_data, agreements)` from `Agenda/chair_notes_cache.json` if available."""
        if not meeting_dir:
            return {}, []
        cache_path = Path(meeting_dir) / "Agenda" / "chair_notes_cache.json"
        if not os.path.exists(to_extended_path(cache_path)):
            return {}, []
        try:
            with open(to_extended_path(cache_path), "r", encoding="utf-8") as f:
                loaded = json.load(f)
                return loaded.get("agenda_data", {}), loaded.get("agreements", [])
        except Exception as e:
            logger.warning(f"[ChairNotesAgreementsParser] Failed to load chair notes cache: {e}")
            return {}, []

    @classmethod
    def parse_chair_notes_file(
        cls,
        file_path: Path,
        mtg_info: Optional[dict] = None,
        agenda_map: Optional[dict] = None,
    ) -> List[dict]:
        """Convenience method parsing a single Chairman's Notes file and returning extracted agreements."""
        file_path = Path(file_path)
        if not file_path.exists() or file_path.name.startswith("~$"):
            return []
        ext = file_path.suffix.lower()
        if ext == ".docx":
            _, agreements = cls.parse_docx_with_agreements(str(file_path), mtg_info=mtg_info, agenda_map=agenda_map)
            return agreements
        return []

    @classmethod
    def scan_and_cache_chair_notes(
        cls,
        meeting_dir: Path,
        mtg_info: Optional[dict] = None,
        agenda_map: Optional[dict] = None,
        force_rescan: bool = False,
    ) -> Tuple[Dict[str, dict], List[dict], int]:
        """
        Scans local Chairman's Notes files (.docx, .doc, .zip, .htm, .html) across
        `meeting_dir/Agenda/ChairNotes/` and `meeting_dir/Agenda/`, parses them for
        TDoc remarks and structured agreements, persists them to `Agenda/chair_notes_cache.json`,
        and returns `(agenda_data, agreements, files_scanned_count)`.

        If `force_rescan` is False and `chair_notes_cache.json` exists and is newer than
        all candidate notes files, returns the cached data immediately.
        """
        if not meeting_dir:
            return {}, [], 0

        meeting_dir = Path(meeting_dir)
        agenda_dir = meeting_dir / "Agenda"
        chair_notes_dir = agenda_dir / "ChairNotes"

        candidate_dirs = []
        if chair_notes_dir.is_dir():
            candidate_dirs.append(chair_notes_dir)
        if agenda_dir.is_dir():
            candidate_dirs.append(agenda_dir)

        # 1. Discover all candidate files
        candidate_files: List[Path] = []
        seen_names = set()

        for d in candidate_dirs:
            for ext in ("*.docx", "*.doc", "*.zip", "*.htm", "*.html"):
                for p in d.glob(ext):
                    if p.name.startswith("~$"):
                        continue
                    if p.name not in seen_names:
                        seen_names.add(p.name)
                        candidate_files.append(p)

        # 2. Extract any ZIP archives if needed
        unpacked_files: List[Path] = []
        for f in list(candidate_files):
            if f.suffix.lower() == ".zip":
                try:
                    import zipfile
                    target_extract = chair_notes_dir if chair_notes_dir.is_dir() else agenda_dir
                    os.makedirs(to_extended_path(target_extract), exist_ok=True)
                    with zipfile.ZipFile(to_extended_path(f), "r") as zf:
                        for zinfo in zf.infolist():
                            zname = zinfo.filename
                            if zname.startswith("~$") or "/" in zname or "\\" in zname:
                                continue
                            zext = Path(zname).suffix.lower()
                            if zext in (".docx", ".doc", ".htm", ".html"):
                                out_p = target_extract / zname
                                if not out_p.exists():
                                    zf.extract(zinfo, path=to_extended_path(target_extract))
                                if out_p.name not in seen_names:
                                    seen_names.add(out_p.name)
                                    unpacked_files.append(out_p)
                except Exception as e:
                    logger.warning(f"[ChairNotesAgreementsParser] Failed extracting ZIP '{f.name}': {e}")

        all_files = [f for f in (candidate_files + unpacked_files) if f.suffix.lower() != ".zip"]

        if not all_files:
            cached_ag, cached_list = cls.load_cached_chair_notes_data(meeting_dir)
            return cached_ag, cached_list, 0

        # 3. Check cache freshness if not forcing rescan
        cache_path = agenda_dir / "chair_notes_cache.json"
        if not force_rescan and os.path.exists(to_extended_path(cache_path)):
            try:
                cache_mtime = os.path.getmtime(to_extended_path(cache_path))
                latest_mtime = max(os.path.getmtime(to_extended_path(p)) for p in all_files)
                if cache_mtime >= latest_mtime:
                    cached_ag, cached_list = cls.load_cached_chair_notes_data(meeting_dir)
                    if cached_ag or cached_list:
                        return cached_ag, cached_list, len(all_files)
            except Exception:
                pass

        # 4. Parse files and merge
        merged_agenda: Dict[str, dict] = {}
        merged_agreements: List[dict] = []
        scanned_count = 0

        # Sort so clean docx files take precedence over tracked changes or interim files
        def _file_sort_key(p: Path):
            name_lower = p.name.lower()
            is_clean = 1 if "clean" in name_lower else 0
            is_final = 1 if "final" in name_lower or "eom" in name_lower else 0
            return (is_clean, is_final, os.path.getmtime(to_extended_path(p)))

        sorted_files = sorted(all_files, key=_file_sort_key, reverse=True)

        for p in sorted_files:
            ext = p.suffix.lower()
            try:
                if ext == ".docx":
                    ad, ag = cls.parse_docx_with_agreements(str(p), mtg_info=mtg_info, agenda_map=agenda_map)
                    scanned_count += 1
                    merged_agenda.update(ad)
                    merged_agreements.extend(ag)
                elif ext in (".htm", ".html"):
                    from modules.meetings.core.tdocs.tdocs_parser import TDocsParser
                    ad = TDocsParser.parse_tdocs_by_agenda(str(p))
                    if ad:
                        scanned_count += 1
                        merged_agenda.update(ad)
                elif ext == ".doc":
                    # Check if converted .docx already exists
                    docx_candidate = p.with_suffix(".docx")
                    if docx_candidate.exists() and docx_candidate in sorted_files:
                        continue
                    try:
                        from modules.word_tools.word_comparator import is_libreoffice_available, convert_doc_to_docx_libreoffice
                        if is_libreoffice_available():
                            out_docx = convert_doc_to_docx_libreoffice(p, docx_candidate)
                            if out_docx and out_docx.exists():
                                ad, ag = cls.parse_docx_with_agreements(str(out_docx), mtg_info=mtg_info, agenda_map=agenda_map)
                                scanned_count += 1
                                merged_agenda.update(ad)
                                merged_agreements.extend(ag)
                    except Exception as e_lo:
                        logger.debug(f"[ChairNotesAgreementsParser] Could not convert legacy .doc '{p.name}': {e_lo}")
            except Exception as e_parse:
                logger.warning(f"[ChairNotesAgreementsParser] Error parsing '{p.name}': {e_parse}", exc_info=True)

        # 5. Persist merged cache if anything was extracted
        if merged_agenda or merged_agreements:
            cls.save_cached_chair_notes_data(meeting_dir, merged_agenda, merged_agreements)

        return merged_agenda, merged_agreements, scanned_count
