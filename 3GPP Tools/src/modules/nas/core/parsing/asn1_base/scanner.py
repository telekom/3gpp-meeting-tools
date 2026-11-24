"""OpenXML ASN.1 block & table scanner plus QTextBrowser HTML formatter."""

import html
import re
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

from modules.nas.core.parsing.protocol_parser_constants import (
    RE_CLAUSE_HEADER,
    RE_DESC_TABLE,
    RE_MAJOR_BOUNDARY,
    TAG_BODY,
    TAG_P,
    TAG_TBL,
    TAG_TC,
    TAG_TR,
)
from modules.nas.core.parsing.protocol_parser_utils import (
    _convert_table_to_html,
    _extract_p_text,
    _extract_tc_text,
    extract_document_root,
)

RE_ASN1_BOX_CAPTION = re.compile(
    r"^[A-Za-z0-9\-_]+\s+(?:information\s+elements?|messages?)\s*$",
    re.IGNORECASE,
)


def normalize_key(name: str) -> str:
    """Strip non-alphanumeric characters and convert to lowercase for resilient matching."""
    return re.sub(r"[^a-zA-Z0-9]", "", name).lower()


def extract_prose_summary_and_direction(html_str: str) -> Tuple[str, str]:
    """Extract plain-text prose summary and direction string from a section's HTML before any table."""
    if not html_str:
        return "", ""
    before_table = html_str.split("<table", 1)[0]
    before_table = re.sub(r"<pre[^>]*>.*?</pre>", "", before_table, flags=re.DOTALL | re.IGNORECASE)
    p_texts = re.findall(r"<(?:p|div)[^>]*>(.*?)</(?:p|div)>", before_table, flags=re.DOTALL | re.IGNORECASE)
    summary_parts: List[str] = []
    direction = ""
    for raw_p in p_texts:
        clean_p = html.unescape(re.sub(r"<[^>]+>", "", raw_p)).strip()
        clean_p = re.sub(r"\s+", " ", clean_p)
        if not clean_p:
            continue
        if RE_ASN1_BOX_CAPTION.match(clean_p):
            continue
        low_p = clean_p.lower()
        if low_p.startswith(("direction:", "direction :")):
            direction = re.sub(r"(?i)^direction\s*:\s*", "", clean_p).strip()
        elif not low_p.startswith((
            "note", "table ", "figure ",
            "signalling radio bearer:", "rlc-sap:", "logical channel:", "logical channels:",
        )):
            summary_parts.append(clean_p)
        elif low_p.startswith(("signalling radio bearer:", "rlc-sap:", "logical channel:", "logical channels:")):
            summary_parts.append(clean_p)
    return " ".join(summary_parts).strip(), direction


def fallback_extract_asn1(docx_paths: List[Path]) -> str:
    """Fallback paragraph scan when explicit ASN1START/ASN1STOP fences are missing."""
    collected = []
    for docx_path in docx_paths:
        root = extract_document_root(docx_path)
        if root is not None:
            body = root.find(TAG_BODY)
            if body is not None:
                for p in body.findall(TAG_P):
                    t = _extract_p_text(p)
                    if any(kw in t for kw in ("::=", "SEQUENCE", "CHOICE", "PROTOCOL-IES")):
                        collected.append(t)
    return "\n".join(collected)


def scan_xml_documents(
    docx_paths: List[Path],
    progress_callback: Optional[Callable[[str, int], None]] = None,
) -> Tuple[str, Dict[str, str], Dict[str, Dict[str, str]], Dict[str, str]]:
    """Scan ordered specification DOCX files for ASN.1 modules, tables, and clause mappings."""
    raw_asn1_blocks: List[str] = []
    field_desc_tables: Dict[str, str] = {}
    field_individual_descs: Dict[str, Dict[str, str]] = {}
    clause_map: Dict[str, str] = {}
    total_files = len(docx_paths)
    current_clause_num = "6.2"

    for file_idx, docx_path in enumerate(docx_paths):
        if progress_callback:
            progress = 10 + int(file_idx / max(1, total_files) * 40)
            progress_callback(f"Scanning {docx_path.name} ({file_idx + 1}/{total_files})...", progress)

        root = extract_document_root(docx_path)
        if root is None:
            continue

        body = root.find(TAG_BODY)
        if body is None:
            continue

        in_asn1 = False
        current_asn1_lines: List[str] = []
        last_p_text = ""
        current_heading_name = ""
        current_section_def: Optional[Dict[str, Any]] = None

        def _finalize_section():
            nonlocal current_section_def
            if not current_section_def:
                return
            s_name = current_section_def["name"]
            s_clause = current_section_def["clause"]
            s_prose = current_section_def["prose"]
            s_tables = current_section_def["tables"]

            if s_name and (s_prose or s_tables):
                norm_k = normalize_key(s_name)
                lower_k = s_name.lower()

                clause_map[lower_k] = s_clause
                if norm_k:
                    clause_map[norm_k] = s_clause

                html_content = "".join(s_prose) + "".join(s_tables)
                field_desc_tables[lower_k] = html_content
                if norm_k:
                    field_desc_tables[norm_k] = html_content

            current_section_def = None

        for elem in list(body):
            if elem.tag == TAG_P:
                p_text = _extract_p_text(elem)
                if not p_text:
                    continue
                last_p_text = p_text

                if "-- ASN1START" in p_text:
                    _finalize_section()
                    in_asn1 = True
                    current_asn1_lines = []
                    continue
                elif "-- ASN1STOP" in p_text:
                    in_asn1 = False
                    raw_asn1_blocks.append("\n".join(current_asn1_lines))
                    current_asn1_lines = []
                    continue

                if in_asn1:
                    current_asn1_lines.append(p_text)
                    continue

                match_clause = RE_CLAUSE_HEADER.match(p_text)
                if match_clause:
                    _finalize_section()
                    current_clause_num = match_clause.group(1).strip()
                    clause_title = match_clause.group(2).strip()

                    if clause_title:
                        current_heading_name = clause_title
                        norm_t = normalize_key(clause_title)
                        clause_map[clause_title.lower()] = current_clause_num
                        if norm_t:
                            clause_map[norm_t] = current_clause_num

                        if current_clause_num.startswith(("9.1", "9.2", "9.3", "6.2", "6.3", "6.6", "D.6")):
                            current_section_def = {
                                "clause": current_clause_num,
                                "name": clause_title,
                                "prose": [],
                                "tables": [],
                            }
                    continue

                if p_text.startswith("–") or p_text.startswith("-"):
                    _finalize_section()
                    current_heading_name = p_text.lstrip("–- ").strip()
                    if current_heading_name:
                        norm_h = normalize_key(current_heading_name)
                        clause_map[current_heading_name.lower()] = current_clause_num
                        if norm_h:
                            clause_map[norm_h] = current_clause_num

                        current_section_def = {
                            "clause": current_clause_num,
                            "name": current_heading_name,
                            "prose": [],
                            "tables": [],
                        }
                    continue

                if RE_MAJOR_BOUNDARY.match(p_text) and not p_text.startswith(("9.", "6.")):
                    _finalize_section()
                    continue

                if current_section_def:
                    if p_text.startswith(("Direction:", "Direction :")):
                        current_section_def["prose"].append(
                            f'<p style="margin: 4px 0 6px 0; font-weight: bold; color: #0284C7; '
                            f'font-size: 11px;">{html.escape(p_text)}</p>'
                        )
                    elif p_text.startswith("NOTE"):
                        current_section_def["prose"].append(
                            f'<div style="background-color: #F0F4F8; border-left: 3px solid #90A4AE; '
                            f'padding: 4px 8px; margin: 4px 0; font-size: 11px; color: #455A64;">'
                            f'{html.escape(p_text)}</div>'
                        )
                    elif (
                        not p_text.startswith(("Table ", "Figure "))
                        and not RE_ASN1_BOX_CAPTION.match(p_text)
                        and len(current_section_def["prose"]) < 8
                    ):
                        current_section_def["prose"].append(
                            f'<p style="margin: 3px 0; line-height: 1.4; color: #334155; '
                            f'font-size: 11px;">{html.escape(p_text)}</p>'
                        )

            elif elem.tag == TAG_TBL:
                tbl_html = _convert_table_to_html(elem)

                rows = elem.findall(TAG_TR)
                first_cells = [_extract_tc_text(tc).strip() for tc in rows[0].findall(TAG_TC)] if rows else []
                header_text = " ".join(first_cells)
                match_tbl = RE_DESC_TABLE.search(last_p_text) or RE_DESC_TABLE.search(header_text)
                rrc_target = match_tbl.group(1).strip() if match_tbl else ""
                if rrc_target and tbl_html:
                    rk = rrc_target.lower()
                    norm_rk = normalize_key(rrc_target)
                    existing_rk = field_desc_tables.get(rk, "")
                    if tbl_html not in existing_rk:
                        field_desc_tables[rk] = existing_rk + tbl_html
                    if norm_rk and norm_rk != rk:
                        existing_norm = field_desc_tables.get(norm_rk, "")
                        if tbl_html not in existing_norm:
                            field_desc_tables[norm_rk] = existing_norm + tbl_html

                if current_section_def and tbl_html:
                    current_section_def["tables"].append(tbl_html)

                if rows:
                    header_cells = [c.lower() for c in first_cells]
                    name_col, desc_col = 0, -1
                    for idx, htext in enumerate(header_cells):
                        if any(k in htext for k in ("semantics", "description", "explanation")):
                            desc_col = idx
                            break
                    if desc_col == -1 and len(header_cells) >= 2:
                        desc_col = 1

                    field_dict: Dict[str, str] = {}
                    start_row = (
                        1
                        if len(rows) > 1
                        and any("field descriptions" in c.lower() or "conditional presence" in c.lower() for c in first_cells)
                        else 0
                    )
                    for r in rows[start_row:]:
                        cells = r.findall(TAG_TC)
                        if len(cells) == 1:
                            p_elements = cells[0].findall(TAG_P)
                            p_texts = [_extract_p_text(p).strip() for p in p_elements if _extract_p_text(p).strip()]
                            if len(p_texts) >= 2:
                                fname = p_texts[0].strip().lstrip("> ")
                                fdesc = "\n".join(p_texts[1:]).strip()
                                if fname and fdesc and "field descriptions" not in fname.lower():
                                    field_dict[fname.lower()] = fdesc
                                    norm_fn = normalize_key(fname)
                                    if norm_fn:
                                        field_dict[norm_fn] = fdesc
                        elif len(cells) > max(name_col, desc_col) and desc_col != -1:
                            fname = _extract_tc_text(cells[name_col]).strip().lstrip("> ")
                            fdesc = _extract_tc_text(cells[desc_col]).strip()
                            if (
                                fname
                                and fdesc
                                and "field descriptions" not in fname.lower()
                                and not any(k in fname.lower() for k in ("conditional presence", "explanation"))
                            ):
                                field_dict[fname.lower()] = fdesc
                                norm_fn = normalize_key(fname)
                                if norm_fn:
                                    field_dict[norm_fn] = fdesc

                    if field_dict:
                        target_name = rrc_target or (current_section_def["name"] if current_section_def else current_heading_name)
                        if target_name:
                            norm_t = normalize_key(target_name)
                            field_individual_descs.setdefault(target_name.lower(), {}).update(field_dict)
                            if norm_t:
                                field_individual_descs.setdefault(norm_t, {}).update(field_dict)

        _finalize_section()

    full_asn1_module = "\n".join(raw_asn1_blocks)
    if not full_asn1_module.strip():
        full_asn1_module = fallback_extract_asn1(docx_paths)

    return full_asn1_module, field_desc_tables, field_individual_descs, clause_map


def build_inspector_html(
    type_name: str,
    assigned_clause: str,
    raw_asn1: str,
    desc_table_html: str,
) -> str:
    """Construct HTML presentation combining ASN.1 definitions and 3GPP tabular specifications."""
    clause_str = f" (Clause {assigned_clause})" if assigned_clause and assigned_clause != type_name else ""

    header_html = (
        f'<h2 style="color: #0369A1; margin-top: 2px; margin-bottom: 8px; '
        f'font-family: Segoe UI, sans-serif; font-size: 15px; border-bottom: 1px solid #E2E8F0; padding-bottom: 4px;">'
        f'{html.escape(type_name)}{html.escape(clause_str)}</h2>'
    )

    raw_asn1_html = ""
    if raw_asn1 and raw_asn1.strip():
        raw_asn1_html = (
            f'<div style="margin-bottom: 10px;">'
            f'<div style="font-size: 11px; font-weight: bold; color: #475569; margin-bottom: 3px;">ASN.1 Definition:</div>'
            f'<pre style="background-color: #F8FAFC; border: 1px solid #E2E8F0; padding: 8px; '
            f'border-radius: 4px; font-family: Consolas, monospace; font-size: 11px; color: #0F172A; '
            f'white-space: pre-wrap; margin: 0;">{html.escape(raw_asn1)}</pre>'
            f'</div>'
        )

    spec_section_html = ""
    if desc_table_html and desc_table_html.strip():
        spec_section_html = (
            f'<div style="margin-top: 8px;">'
            f'<div style="font-size: 11px; font-weight: bold; color: #475569; margin-bottom: 4px;">'
            f'Specification Details & Tabular Definition:</div>'
            f'{desc_table_html}'
            f'</div>'
        )

    return f"{header_html}{raw_asn1_html}{spec_section_html}"
