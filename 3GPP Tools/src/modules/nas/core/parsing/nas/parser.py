"""Dedicated WordprocessingML DOCX parser for 3GPP NAS & UE Policy specifications."""

from __future__ import annotations

import html
import json
import logging
import re
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

try:
    from lxml import etree as ET
except ImportError:
    import xml.etree.ElementTree as ET

from modules.nas.core.parsing.nas.catalogs import (
    ATTR_VAL,
    RE_CAPTION,
    RE_IE_HEADING_NAS,
    RE_IE_HEADING_POLICY,
    RE_MAJOR_BOUNDARY_NAS,
    RE_MAJOR_BOUNDARY_POLICY,
    RE_MSG_HEADING_NAS,
    RE_PARENT_SECTION_POLICY,
    RE_TRAILING_PAGE_NUM,
    TAG_INSTR,
    TAG_PSTYLE,
    TS24514_POLICY_MESSAGES,
    TS24526_POLICY_MESSAGES,
    TS24555_POLICY_MESSAGES,
    TS24578_POLICY_MESSAGES,
    TS24588_POLICY_MESSAGES,
)
from modules.nas.core.parsing.nas.container_unroller import (
    finalize_ie_definition,
    unroll_container_sub_ies,
)
from modules.nas.core.parsing.protocol_parser_constants import (
    TAG_BODY,
    TAG_P,
    TAG_PPR,
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

_POLICY_CATALOGS = {
    "24.526": TS24526_POLICY_MESSAGES,
    "24.588": TS24588_POLICY_MESSAGES,
    "24.555": TS24555_POLICY_MESSAGES,
    "24.514": TS24514_POLICY_MESSAGES,
    "24.578": TS24578_POLICY_MESSAGES,
}


class NASDocxParser:
    """Dedicated parser for standard 3GPP NAS & UE Policy specifications (TS 24.501, TS 24.301, TS 24.526, TS 24.588, TS 24.555, TS 24.514, TS 24.578)."""

    def __init__(self, docx_paths: List[Path], spec_number: str = "24.501"):
        self.docx_paths = docx_paths
        self.spec_number = str(spec_number or "24.501").replace("TS", "").strip()
        self.is_policy_spec = self.spec_number in _POLICY_CATALOGS
        self.logger = logging.getLogger(__name__)

    @staticmethod
    def _is_toc_paragraph(p_elem: ET.Element) -> bool:
        """Detects Word Table of Contents (TOC) paragraphs to avoid indexing TOC lines as IE definitions."""
        ppr = p_elem.find(TAG_PPR)
        if ppr is not None:
            ps = ppr.find(TAG_PSTYLE)
            if ps is not None:
                val = (ps.get(ATTR_VAL) or "").lower()
                if val.startswith("toc"):
                    return True
        for instr in p_elem.iter(TAG_INSTR):
            if instr.text and ("PAGEREF" in instr.text or "TOC" in instr.text):
                return True
        return False

    def parse(
        self, progress_callback: Optional[Callable[[str, int], None]] = None
    ) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
        valid_paths = [p for p in self.docx_paths if p.exists()]
        if not valid_paths:
            raise FileNotFoundError(f"Specification file(s) not found: {self.docx_paths}")

        messages: List[Dict[str, Any]] = []
        ie_definitions: List[Dict[str, Any]] = []
        total_files = len(valid_paths)

        is_policy_spec = self.spec_number in _POLICY_CATALOGS
        re_ie_heading = RE_IE_HEADING_POLICY if is_policy_spec else RE_IE_HEADING_NAS
        re_major_boundary = RE_MAJOR_BOUNDARY_POLICY if is_policy_spec else RE_MAJOR_BOUNDARY_NAS

        parent_section_titles: Dict[str, str] = {}

        for file_idx, docx_path in enumerate(valid_paths):
            if progress_callback:
                base_progress = 10 + int((file_idx / total_files) * 80)
                progress_callback(f"Reading {docx_path.name} ({file_idx + 1}/{total_files})...", base_progress)

            root = extract_document_root(docx_path)
            if root is None:
                continue

            body = root.find(TAG_BODY)
            if body is None:
                continue

            last_caption_info: Optional[Tuple[str, str, str]] = None
            last_paragraph_text: str = ""
            current_ie_def: Optional[Dict[str, Any]] = None
            current_msg_sec: Optional[Dict[str, Any]] = None

            for elem in list(body):
                if elem.tag == TAG_P:
                    if self._is_toc_paragraph(elem):
                        continue

                    p_text = _extract_p_text(elem)
                    if not p_text:
                        continue
                    last_paragraph_text = p_text

                    if is_policy_spec:
                        mp_sec = RE_PARENT_SECTION_POLICY.match(p_text)
                        if mp_sec:
                            sec_cl = mp_sec.group(1).strip()
                            sec_title = RE_TRAILING_PAGE_NUM.sub("", mp_sec.group(2).strip()).strip()
                            parent_section_titles[sec_cl] = sec_title

                    if not is_policy_spec:
                        match_msg_h = RE_MSG_HEADING_NAS.match(p_text)
                        if match_msg_h and not p_text.startswith(("Table ", "Figure ")):
                            msg_cl = match_msg_h.group(1).strip()
                            msg_title = RE_TRAILING_PAGE_NUM.sub("", match_msg_h.group(2).strip()).strip()
                            current_msg_sec = {
                                "clause": msg_cl,
                                "title": msg_title,
                                "prose": [],
                                "direction": "",
                                "html_content": [
                                    f'<h2 style="color: #0369A1; margin-top: 4px; margin-bottom: 6px; '
                                    f'font-family: Segoe UI, sans-serif; font-size: 14px; border-bottom: 1px solid #E2E8F0; padding-bottom: 4px;">'
                                    f"{html.escape(msg_title)} (Clause {html.escape(msg_cl)})</h2>"
                                ],
                            }
                            last_caption_info = None
                            continue

                    if p_text.startswith(("Table 8.", "Table D.5.")):
                        match_cap = RE_CAPTION.search(p_text)
                        if match_cap:
                            clause = match_cap.group(1).strip()
                            name = re.sub(r"(?i)\s+message\s+content", "", match_cap.group(2)).strip()
                            last_caption_info = (clause, name, p_text)
                        else:
                            last_caption_info = None
                    else:
                        last_caption_info = None

                    if current_msg_sec and not p_text.startswith(("Table 8.", "Table D.5.", "9.", "D.6")):
                        low_p = p_text.lower()
                        if low_p.startswith("direction:"):
                            raw_dir = re.sub(r"(?i)^direction\s*:\s*", "", p_text).strip().rstrip(".")
                            norm_dir = re.sub(r"(?i)\bto\b", "→", raw_dir)
                            norm_dir = re.sub(r"(?i)\band/or\b", "↔", norm_dir)
                            current_msg_sec["direction"] = norm_dir
                            current_msg_sec["html_content"].append(
                                f'<p style="margin: 4px 0 6px 0; font-weight: bold; color: #0284C7; font-size: 11px;">'
                                f"{html.escape(p_text)}</p>"
                            )
                        elif low_p.startswith(("message type:", "significance:")):
                            current_msg_sec["html_content"].append(
                                f'<p style="margin: 3px 0; font-weight: 600; color: #475569; font-size: 11px;">'
                                f"{html.escape(p_text)}</p>"
                            )
                        elif p_text.startswith("NOTE"):
                            current_msg_sec["html_content"].append(
                                f'<div style="background-color: #F0F4F8; border-left: 3px solid #90A4AE; '
                                f'padding: 4px 8px; margin: 4px 0; font-size: 11px; color: #455A64;">{html.escape(p_text)}</div>'
                            )
                        elif not p_text.startswith(("Table ", "Figure ")) and len(current_msg_sec["prose"]) < 6:
                            current_msg_sec["prose"].append(p_text)
                            current_msg_sec["html_content"].append(
                                f'<p style="margin: 4px 0; line-height: 1.4; color: #263238; font-size: 11px;">{html.escape(p_text)}</p>'
                            )

                    match_ie = re_ie_heading.match(p_text)
                    if match_ie:
                        current_msg_sec = None
                        if current_ie_def:
                            self._finalize_ie_def(current_ie_def, ie_definitions)

                        cl = match_ie.group(1).strip()
                        ie_name = RE_TRAILING_PAGE_NUM.sub("", match_ie.group(2).strip()).strip()
                        if is_policy_spec:
                            parts = cl.split(".")
                            parent_cl = ".".join(parts[:-1]) if len(parts) > 2 else cl
                            parent_title = parent_section_titles.get(parent_cl) or parent_section_titles.get(parts[0])
                            if ie_name.lower() in ("general", "overview", "information elements coding"):
                                if parent_title:
                                    ie_name = f"{parent_title} — {ie_name}"
                                elif cl == "5.3.1":
                                    ie_name = "Encoding of UE policy part type ANDSP — General"
                                elif cl == "5.3.3.1":
                                    ie_name = "Encoding of N3AN node configuration information — General"
                                elif cl.startswith("12.2"):
                                    ie_name = "RSLPP Information Elements Coding"
                                else:
                                    ie_name = f"Clause {cl} — {ie_name}"

                        current_ie_def = {
                            "clause": cl,
                            "ie_name": ie_name,
                            "html_content": [
                                f'<h2 style="color: #0D47A1; margin-top: 4px; margin-bottom: 6px; '
                                f'font-family: Segoe UI, sans-serif;">{html.escape(ie_name)} (Clause {html.escape(cl)})</h2>'
                            ],
                        }
                    elif re_major_boundary.match(p_text) and not p_text.startswith(("9.", "D.6", "5.", "12.")):
                        if current_ie_def:
                            self._finalize_ie_def(current_ie_def, ie_definitions)
                            current_ie_def = None
                    elif current_ie_def:
                        if p_text.startswith(("Figure 9.", "Figure D.", "Figure 5.", "Figure 12.")):
                            current_ie_def["html_content"].append(
                                f'<p style="font-weight: bold; color: #37474F; margin-top: 10px; margin-bottom: 4px;">{html.escape(p_text)}</p>'
                            )
                        elif p_text.startswith(("Table 9.", "Table D.", "Table 5.", "Table 12.")):
                            current_ie_def["html_content"].append(
                                f'<p style="font-weight: bold; color: #1A237E; margin-top: 12px; margin-bottom: 4px;">{html.escape(p_text)}</p>'
                            )
                        elif p_text.startswith("NOTE"):
                            current_ie_def["html_content"].append(
                                f'<div style="background-color: #F0F4F8; border-left: 3px solid #90A4AE; '
                                f'padding: 4px 8px; margin: 4px 0; font-size: 11px; color: #455A64;">{html.escape(p_text)}</div>'
                            )
                        else:
                            current_ie_def["html_content"].append(
                                f'<p style="margin: 4px 0; line-height: 1.4; color: #263238;">{html.escape(p_text)}</p>'
                            )

                elif elem.tag == TAG_TBL:
                    handled_as_message = False
                    if last_caption_info:
                        clause, name, full_caption = last_caption_info
                        ies = self._parse_clause_8_table(elem)
                        if ies:
                            msg_desc = " ".join(current_msg_sec.get("prose", [])).strip() if current_msg_sec else ""
                            msg_dir = current_msg_sec.get("direction", "") if current_msg_sec else ""
                            messages.append(
                                {
                                    "clause": clause,
                                    "message_name": name,
                                    "table_caption": full_caption,
                                    "description": msg_desc,
                                    "direction": msg_dir,
                                    "ies": ies,
                                }
                            )
                            if current_msg_sec:
                                msg_tbl_html = _convert_table_to_html(elem, is_figure_diagram=False)
                                msg_html_parts = list(current_msg_sec.get("html_content", []))
                                msg_html_parts.append(
                                    f'<p style="font-weight: bold; color: #1A237E; margin-top: 10px; margin-bottom: 4px; font-size: 11px;">'
                                    f"{html.escape(full_caption)}</p>"
                                )
                                if msg_tbl_html:
                                    msg_html_parts.append(msg_tbl_html)
                                ie_definitions.append(
                                    {
                                        "clause": clause,
                                        "ie_name": name,
                                        "raw_description": "".join(msg_html_parts),
                                        "structure_table": json.dumps([]),
                                    }
                                )
                                current_msg_sec = None
                            handled_as_message = True
                        last_caption_info = None

                    if not handled_as_message and current_ie_def:
                        is_diagram = "figure" in last_paragraph_text.lower() or any(
                            "octet" in _extract_tc_text(c).lower() for c in elem.iter(TAG_TC)
                        )
                        tbl_html = _convert_table_to_html(elem, is_figure_diagram=is_diagram)
                        if tbl_html:
                            current_ie_def["html_content"].append(tbl_html)

            if current_ie_def:
                self._finalize_ie_def(current_ie_def, ie_definitions)

        if self.is_policy_spec and ie_definitions and not messages:
            messages.extend(_POLICY_CATALOGS.get(self.spec_number, ()))

        self._unroll_container_sub_ies(messages, ie_definitions)
        return messages, ie_definitions

    def _finalize_ie_def(self, current_ie_def: Dict[str, Any], ie_definitions: List[Dict[str, Any]]):
        finalize_ie_definition(current_ie_def, ie_definitions, self.is_policy_spec)

    def _unroll_container_sub_ies(
        self,
        messages: List[Dict[str, Any]],
        ie_definitions: List[Dict[str, Any]],
    ):
        unroll_container_sub_ies(messages, ie_definitions)

    def _parse_clause_8_table(self, tbl_elem: ET.Element) -> List[Dict[str, str]]:
        rows = tbl_elem.findall(TAG_TR)
        if not rows:
            return []

        header_row_idx = -1
        for r_idx in range(min(3, len(rows))):
            cells_text = [_extract_tc_text(tc).lower().strip() for tc in rows[r_idx].findall(TAG_TC)]
            if len(cells_text) < 4:
                continue
            has_ie_col = any(c in ("information element", "information elements", "iei") for c in cells_text)
            has_meta_col = any("reference" in c or "presence" in c or "format" in c for c in cells_text)
            if has_ie_col and has_meta_col:
                header_row_idx = r_idx
                break

        if header_row_idx == -1:
            return []

        ies = []
        for row in rows[header_row_idx + 1 :]:
            cells = [_extract_tc_text(tc) for tc in row.findall(TAG_TC)]
            if len(cells) < 6:
                continue

            iei, ie_name, type_ref, presence, fmt, length = cells[:6]
            clean_name = ie_name.lower().replace(" ", "")
            if not clean_name or "informationelement" in clean_name:
                continue

            ies.append(
                {
                    "iei": iei,
                    "information_element": ie_name,
                    "field_path": ie_name,
                    "depth": 0,
                    "type_reference": type_ref.strip().rstrip("."),
                    "presence": presence,
                    "format": fmt,
                    "length": length,
                }
            )

        return ies
