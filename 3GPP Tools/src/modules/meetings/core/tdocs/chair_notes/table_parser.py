"""Table-based Chairman's Notes parser mixin (RAN3 color-coded & SA2 multi-column)."""

from __future__ import annotations

import re
from typing import Dict, List, Tuple

from docx.document import Document as _Document
from docx.table import Table

from .text_and_results import (
    _RAN3_WID_BRACKET_RE,
    _RELEASE_RE,
    _TDOC_ID_RE,
    _get_paragraph_dominant_color,
    _resolve_inherited_field,
    clean_chair_notes_text,
)


class ChairNotesTableParserMixin:
    """Locates main TDoc tables and extracts banner, row, and moderator-summary records."""

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
                        ai_cell = (
                            cells_text[col_map["ai"]].strip()
                            if "ai" in col_map and col_map["ai"] < len(cells_text)
                            else ""
                        )
                        title_candidate = (
                            cells_text[col_map["title"]].strip()
                            if "title" in col_map and col_map["title"] < len(cells_text)
                            else ""
                        )
                        if ai_cell and ai_cell not in ("-", "0") and ai_cell != current_ai:
                            current_ai = ai_cell
                            current_sub_item = ""
                            if title_candidate and title_candidate != "-" and not title_candidate.lower().startswith("all agenda"):
                                ai_titles[current_ai] = title_candidate
                        elif title_candidate and title_candidate != "-" and not title_candidate.lower().startswith("all agenda"):
                            current_sub_item = title_candidate
                            sub_item_order_counter += 1
                    continue

                # Valid TDoc Row!
                tdoc_id = td_match.group(1).upper()
                order_in_agenda += 1

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

                row_wi = cells_text[col_map["wi"]].strip() if "wi" in col_map and col_map["wi"] < len(cells_text) else ""
                if row_wi in ("-", "—", "N/A", "None"):
                    row_wi = ""
                row_rel = cells_text[col_map["rel"]].strip() if "rel" in col_map and col_map["rel"] < len(cells_text) else ""
                if row_rel in ("-", "—", "N/A", "None"):
                    row_rel = ""

                resolved_wi = row_wi or _resolve_inherited_field(current_ai, ai_wis)
                resolved_rel = row_rel or _resolve_inherited_field(current_ai, ai_releases)

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
