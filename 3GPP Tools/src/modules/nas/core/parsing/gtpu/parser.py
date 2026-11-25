"""TS 29.281 (GTPv1-U) DOCX parser implementation."""

from __future__ import annotations

import copy
import html
import json
import logging
import re
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

from modules.nas.core.parsing.gtpu.builders import (
    build_clause_7_messages,
    build_common_header_message,
    build_gpdu_message,
    sanitize_message_name,
    section_header_html,
)
from modules.nas.core.parsing.gtpu.catalogs import (
    CLAUSE_8_EMBEDDED_FIGURE_HTML,
    DEFAULT_EXT_HEADER_REGISTRY,
    DEFAULT_TYPE_REGISTRY,
    RE_GTPU_ANY_HEADING,
    RE_GTPU_CLAUSE_HEADER,
    RE_GTPU_FIG_511,
    RE_GTPU_FIG_5213,
    RE_GTPU_TABLE_CAPTION,
)
from modules.nas.core.parsing.protocol_parser_constants import (
    TAG_BODY,
    TAG_P,
    TAG_PPR,
    TAG_TBL,
    TAG_TC,
    TAG_TR,
    W_NS,
    extract_part_index,
)
from modules.nas.core.parsing.protocol_parser_utils import (
    _convert_table_to_html,
    _extract_p_text,
    _extract_tc_text,
    extract_document_root,
)

TAG_PSTYLE = f"{W_NS}pStyle"
ATTR_VAL = f"{W_NS}val"


def _get_p_style(p_elem: Any) -> str:
    p_pr = p_elem.find(TAG_PPR)
    if p_pr is None:
        return ""
    st = p_pr.find(TAG_PSTYLE)
    return st.attrib.get(ATTR_VAL, "") if st is not None else ""


def _is_heading_style(p_style: str) -> bool:
    if not p_style:
        return True
    low = p_style.lower()
    return low.startswith("heading") or bool(re.match(r"^h[1-9]\b", low))


class GTPUDocxParser:
    """Dedicated parser for 3GPP TS 29.281 (GTPv1-U) specifications."""

    def __init__(self, docx_paths: List[Path], spec_number: str = "29.281"):
        self.docx_paths = sorted(docx_paths, key=self._extract_part_index)
        self.spec_number = spec_number
        self.logger = logging.getLogger(__name__)

    @staticmethod
    def _extract_part_index(path: Any) -> int:
        return extract_part_index(path)

    @staticmethod
    def _normalize_key(name: str) -> str:
        return re.sub(r"[^a-zA-Z0-9]", "", name).lower()

    @staticmethod
    def _clean_clause_number(clause_ref: str) -> str:
        return clause_ref.split("-")[0].strip()

    @staticmethod
    def _sanitize_message_name(caption: str) -> str:
        return sanitize_message_name(caption)

    def parse(
        self, progress_callback: Optional[Callable[[str, int], None]] = None
    ) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
        valid_paths = [p for p in self.docx_paths if p.exists()]
        if not valid_paths:
            self.logger.error("No valid specification docx files found: %s", self.docx_paths)
            raise FileNotFoundError(f"Specification file(s) not found: {self.docx_paths}")

        total_files = len(valid_paths)
        raw_tables: List[Tuple[str, str, str, Any]] = []
        ie_definitions_dict: Dict[str, Dict[str, Any]] = {}
        type_registry: Dict[str, Dict[str, str]] = copy.deepcopy(DEFAULT_TYPE_REGISTRY)
        ext_header_registry: Dict[str, Dict[str, Any]] = copy.deepcopy(DEFAULT_EXT_HEADER_REGISTRY)
        clause_7_prose_map: Dict[str, List[str]] = {}
        fig_511_tbl: Any = None
        fig_5213_tbl: Any = None

        for file_idx, docx_path in enumerate(valid_paths):
            if progress_callback:
                progress = 10 + int(file_idx / max(1, total_files) * 45)
                progress_callback(
                    f"Scanning GTP-U {docx_path.name} ({file_idx + 1}/{total_files})...",
                    progress,
                )

            root = extract_document_root(docx_path)
            if root is None:
                continue
            body = root.find(TAG_BODY)
            if body is None:
                continue

            body_elems = list(body)
            has_scope = any(
                el.tag == TAG_P and re.match(r"^1\.?\s+Scope$", _extract_p_text(el).strip(), re.IGNORECASE)
                for el in body_elems
            )
            in_body = not has_scope

            last_caption_info: Optional[Tuple[str, str, str]] = None
            last_tbl_elem: Any = None
            current_ie_clause = ""
            current_ie_name = ""
            current_ie_html: List[str] = []
            current_ie_has_table = False
            current_c7_clause = ""

            def _finalize_current_ie() -> None:
                nonlocal current_ie_clause, current_ie_name, current_ie_html, current_ie_has_table
                if current_ie_clause and current_ie_name and current_ie_html:
                    if not current_ie_has_table and current_ie_clause in CLAUSE_8_EMBEDDED_FIGURE_HTML:
                        current_ie_html.append(CLAUSE_8_EMBEDDED_FIGURE_HTML[current_ie_clause])
                        if current_ie_clause == "8.5":
                            current_ie_html.append(
                                '<p style="margin: 4px 0; line-height: 1.4; color: #475569; font-size: 11px;">'
                                'See Clause 5.2.1 (Figure 5.2.1-3) for the full Extension Header Type registry.</p>'
                            )
                    norm_k = self._normalize_key(current_ie_name)
                    if norm_k not in ie_definitions_dict:
                        ie_definitions_dict[norm_k] = {
                            "clause": current_ie_clause,
                            "name": current_ie_name,
                            "ie_name": current_ie_name,
                            "raw_description": "".join(current_ie_html),
                            "structure_table": json.dumps([]),
                        }
                current_ie_clause = ""
                current_ie_name = ""
                current_ie_html = []
                current_ie_has_table = False

            for elem in body_elems:
                if elem.tag == TAG_P:
                    p_text = _extract_p_text(elem).strip()
                    if not p_text:
                        continue

                    if not in_body:
                        if re.match(r"^1\.?\s+Scope$", p_text, re.IGNORECASE):
                            in_body = True
                        continue

                    p_style = _get_p_style(elem)
                    heading_ok = _is_heading_style(p_style)
                    match_header = RE_GTPU_CLAUSE_HEADER.match(p_text) if heading_ok else None

                    if match_header:
                        cl_title_check = match_header.group(2).strip()
                        if not p_style and cl_title_check and not re.match(r"^[A-Z(]", cl_title_check):
                            match_header = None

                    if match_header:
                        cl_num = match_header.group(1).strip()
                        cl_title = match_header.group(2).strip()
                        is_target_section = (
                            (cl_num == "5.1")
                            or (cl_num == "5.2.1")
                            or (cl_num.startswith("5.2.2") and bool(cl_title))
                            or (cl_num.startswith("8.") and cl_num not in ("8", "8.1") and bool(cl_title))
                        )
                        if is_target_section:
                            _finalize_current_ie()
                            current_c7_clause = ""
                            last_tbl_elem = None
                            current_ie_clause = cl_num
                            if cl_num == "5.1":
                                current_ie_name = "GTP-U Header"
                            elif cl_num == "5.2.1":
                                current_ie_name = "GTP-U Extension Header Format"
                            else:
                                current_ie_name = cl_title
                            current_ie_html = [section_header_html(current_ie_name, cl_num)]
                            continue
                        if cl_num.startswith(("7.2", "7.3")) and cl_title:
                            _finalize_current_ie()
                            last_tbl_elem = None
                            current_c7_clause = cl_num
                            clause_7_prose_map.setdefault(current_c7_clause, [])
                            continue
                        _finalize_current_ie()
                        current_c7_clause = ""
                        last_tbl_elem = None
                        continue
                    elif heading_ok and RE_GTPU_ANY_HEADING.match(p_text):
                        if p_style or re.match(r"^(?:\d+(?:\.[0-9A-Za-z]+)*\s+[A-Z]|Annex\s+[A-Z])", p_text):
                            _finalize_current_ie()
                            current_c7_clause = ""
                            last_tbl_elem = None
                            continue

                    if RE_GTPU_FIG_511.match(p_text) and last_tbl_elem is not None and fig_511_tbl is None:
                        fig_511_tbl = last_tbl_elem
                    elif RE_GTPU_FIG_5213.match(p_text) and last_tbl_elem is not None and fig_5213_tbl is None:
                        fig_5213_tbl = last_tbl_elem

                    match_cap = RE_GTPU_TABLE_CAPTION.match(p_text)
                    if match_cap:
                        last_caption_info = (match_cap.group(1).strip(), match_cap.group(2).strip(), p_text)
                    elif not p_text.startswith("Table "):
                        last_caption_info = None

                    if current_c7_clause and not p_text.startswith(("Table ", "Figure ", "NOTE")):
                        prose_list = clause_7_prose_map.setdefault(current_c7_clause, [])
                        if len(prose_list) < 4:
                            prose_list.append(p_text)

                    if current_ie_clause:
                        self._append_ie_paragraph(current_ie_html, p_text)

                elif elem.tag == TAG_TBL:
                    if not in_body:
                        continue
                    last_tbl_elem = elem
                    if last_caption_info:
                        clause_ref, cap_name, full_cap = last_caption_info
                        raw_tables.append((clause_ref, cap_name, full_cap, elem))
                        last_caption_info = None

                    if current_ie_clause:
                        is_bit_diagram = any(
                            "bit" in _extract_tc_text(c).lower() or "octet" in _extract_tc_text(c).lower()
                            for c in elem.iter(TAG_TC)
                        )
                        tbl_html = _convert_table_to_html(elem, is_figure_diagram=is_bit_diagram)
                        if tbl_html:
                            current_ie_html.append(tbl_html)
                            current_ie_has_table = True

            _finalize_current_ie()

        if progress_callback:
            progress_callback("Analyzing Table 8.1-1 and Extension Header registries...", 60)

        for clause_ref, cap_name, _, tbl_elem in raw_tables:
            cap_low = cap_name.lower()
            if "8.1-1" in clause_ref or "information elements" in cap_low:
                self._parse_clause_8_type_registry(tbl_elem, type_registry)
            elif fig_5213_tbl is None and ("5.2.1-3" in clause_ref or "extension header type" in cap_low):
                self._parse_extension_header_registry(tbl_elem, ext_header_registry)

        if fig_5213_tbl is not None:
            self._parse_extension_header_registry(fig_5213_tbl, ext_header_registry)
            fig_5213_html = _convert_table_to_html(fig_5213_tbl, is_figure_diagram=False)
            if fig_5213_html and "extensionheadertypes" in ie_definitions_dict:
                ie_definitions_dict["extensionheadertypes"]["raw_description"] += (
                    '<p style="font-weight: bold; color: #475569; margin-top: 8px; margin-bottom: 3px; font-size: 11px;">'
                    'Figure 5.2.1-3: Definition of Extension Header Type</p>'
                    f'{fig_5213_html}'
                )

        messages = build_clause_7_messages(
            raw_tables,
            type_registry,
            ext_header_registry,
            clause_7_prose_map,
            ie_definitions_dict,
            self._normalize_key,
        )
        messages.append(build_common_header_message(ie_definitions_dict, fig_511_tbl))
        messages.append(build_gpdu_message(ext_header_registry, ie_definitions_dict, fig_511_tbl, fig_5213_tbl))

        ie_definitions = list(ie_definitions_dict.values())
        self.logger.info(
            "Completed parsing TS %s: Returning %d messages and %d IE definitions",
            self.spec_number,
            len(messages),
            len(ie_definitions),
        )
        return messages, ie_definitions

    @staticmethod
    def _append_ie_paragraph(target_html: List[str], p_text: str) -> None:
        if p_text.startswith(("Figure 5.", "Figure 8.")):
            target_html.append(
                f'<p style="font-weight: bold; color: #475569; margin-top: 8px; margin-bottom: 3px; font-size: 11px;">'
                f'{html.escape(p_text)}</p>'
            )
        elif p_text.startswith(("Table 5.", "Table 8.")):
            target_html.append(
                f'<p style="font-weight: bold; color: #0369A1; margin-top: 10px; margin-bottom: 3px; font-size: 11px;">'
                f'{html.escape(p_text)}</p>'
            )
        elif p_text.startswith("NOTE"):
            target_html.append(
                f'<div style="background-color: #F0F4F8; border-left: 3px solid #90A4AE; '
                f'padding: 4px 8px; margin: 4px 0; font-size: 11px; color: #455A64;">'
                f'{html.escape(p_text)}</div>'
            )
        elif not p_text.startswith(("Figure", "Table")) and len(target_html) < 45:
            target_html.append(
                f'<p style="margin: 3px 0; line-height: 1.4; color: #334155; font-size: 11px;">'
                f'{html.escape(p_text)}</p>'
            )

    def _parse_clause_8_type_registry(self, tbl_elem: Any, registry: Dict[str, Dict[str, str]]) -> None:
        rows = tbl_elem.findall(TAG_TR)
        if not rows:
            return
        for row in rows[1:]:
            cells = [_extract_tc_text(tc).strip() for tc in row.findall(TAG_TC)]
            if len(cells) < 3:
                continue
            type_val = cells[0]
            fmt_val = cells[1] if len(cells) > 3 else "TLV"
            name_val = cells[2] if len(cells) > 3 else cells[1]
            ref_val = cells[3] if len(cells) > 3 else cells[2]
            clean_name = re.sub(r"(?i)\s*see\s+note.*$", "", name_val).strip()
            clean_name = re.sub(r"\.$", "", clean_name).strip()
            if clean_name and type_val.isdigit():
                norm_k = self._normalize_key(clean_name)
                registry[norm_k] = {
                    "type_id": type_val,
                    "format": fmt_val,
                    "clause": ref_val,
                    "length": "Variable" if "TLV" in fmt_val else ("1" if type_val == "14" else "4"),
                }
                if "gsn" in norm_k:
                    registry["gtpupeeraddress"] = registry[norm_k]

    def _parse_extension_header_registry(self, tbl_elem: Any, registry: Dict[str, Dict[str, Any]]) -> None:
        rows = tbl_elem.findall(TAG_TR)
        if not rows:
            return
        live_entries: Dict[str, Dict[str, Any]] = {}
        for row in rows[1:]:
            cells = [_extract_tc_text(tc).strip() for tc in row.findall(TAG_TC)]
            if len(cells) < 2:
                continue
            code_bits = cells[0].replace(" ", "")
            if not re.match(r"^[01]{8}$", code_bits):
                continue
            raw_name = cells[1]
            low_name = raw_name.lower()
            if not raw_name or "reserved" in low_name or "no more" in low_name:
                continue
            clean_name = raw_name.split(".", 1)[0]
            clean_name = re.sub(r"\s*\[\d+\](?:-\[\d+\])?", "", clean_name)
            clean_name = re.sub(r"(?i)\s*see\s+note.*$", "", clean_name).strip()
            if not clean_name:
                continue
            hex_val = f"0x{int(code_bits, 2):02X}"
            norm_k = self._normalize_key(clean_name)
            entry = live_entries.setdefault(
                norm_k,
                {
                    "iei_list": [],
                    "name": clean_name,
                    "clause": registry.get(norm_k, {}).get("clause", "5.2.2"),
                },
            )
            if hex_val not in entry["iei_list"]:
                entry["iei_list"].append(hex_val)

        for norm_k, entry in live_entries.items():
            ordered_codes = sorted(entry["iei_list"], key=lambda h: int(h, 16))
            existing = registry.get(norm_k, {})
            registry[norm_k] = {
                "iei": " / ".join(ordered_codes),
                "iei_list": ordered_codes,
                "name": existing.get("name", entry["name"]),
                "clause": existing.get("clause", entry["clause"]),
            }
