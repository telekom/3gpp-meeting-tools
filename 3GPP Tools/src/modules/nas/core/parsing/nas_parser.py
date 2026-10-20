import html
import json
import logging
import re
from pathlib import Path
from typing import List, Optional, Callable, Tuple, Dict, Any

try:
    from lxml import etree as ET
except ImportError:
    import xml.etree.ElementTree as ET

from modules.nas.core.parsing.protocol_parser_utils import extract_document_root, _extract_p_text, _extract_tc_text, \
    _convert_table_to_html
from modules.nas.core.parsing.protocol_parser_constants import TAG_BODY, TAG_P, TAG_PPR, TAG_TBL, TAG_TR, TAG_TC, W_NS

TAG_PSTYLE = f"{W_NS}pStyle"
TAG_INSTR = f"{W_NS}instrText"
ATTR_VAL = f"{W_NS}val"

RE_CAPTION = re.compile(
    r"^Table\s+([8D]\.\d+(?:[\.\-/][0-9A-Za-z]+)*)\s*[:\.]\s*(.+?)(?:\s+message\s+content)?$",
    re.IGNORECASE,
)
RE_IE_HEADING = re.compile(r"^((?:9\.[2-9]|9\.1[0-9]|D\.6)(?:\.[0-9A-Za-z]+)*)\s+(.+)$")
RE_MAJOR_BOUNDARY = re.compile(r"^(?:[1-8]|10|11|12|Annex\s+[A-Z])\b")
RE_TRAILING_PAGE_NUM = re.compile(r"\s+\d{3,4}$")

# Structural umbrella headings in TS 24.501 / TS 24.301 that group subclauses rather than defining a single IE
STRUCTURAL_SECTION_CLAUSES = {
    "9.9", "9.9.1", "9.9.2", "9.9.3", "9.9.4",
    "9.11", "9.11.1", "9.11.2", "9.11.3", "9.11.4",
    "9.12", "D.6",
}

# Encapsulated payload types carried inside TS 24.501 Payload container (9.11.3.39 / Table 9.11.3.40.1)
PAYLOAD_CONTAINER_SUB_IES = [
    ("0100", "SOR transparent container", "9.11.3.51", "C", "TLV-E", "20-n"),
    ("0101", "UE policy section management list", "D.6.2", "C", "TLV-E", "5-65538"),
    ("0110", "UE parameters update transparent container", "9.11.3.53A", "C", "TLV-E", "20-n"),
    ("1000", "CIoT small data container", "9.11.3.18B", "C", "TLV-E", "5-n"),
    ("1001", "Service-level-AA container", "9.11.2.10", "C", "TLV-E", "4-65538"),
]

# Sub-parameters carried inside TS 24.501 Service-level-AA container (9.11.2.10 / Table 9.11.2.10.1)
SERVICE_LEVEL_AA_SUB_IES = [
    ("10", "Service-level device ID", "9.11.2.11", "O", "TLV", "3-n"),
    ("20", "Service-level-AA server address", "9.11.2.12", "O", "TLV", "3-n"),
    ("30", "Service-level-AA response", "9.11.2.14", "O", "TLV", "3"),
    ("40", "Service-level-AA payload type", "9.11.2.15", "O", "TLV", "3"),
    ("70", "Service-level-AA payload", "9.11.2.13", "O", "TLV-E", "4-65538"),
    ("A-", "Service-level-AA pending indication", "9.11.2.17", "O", "TV", "1"),
    ("50", "Service-level-AA service status indication", "9.11.2.18", "O", "TLV", "3"),
]

# Sub-parameters carried inside Extended protocol configuration options (9.11.4.6) defined in Clause 9
EPCO_SUB_IES = [
    ("0031", "ECS address", "9.11.4.34", "O", "TLV", "4-n"),
]


class NASDocxParser:
    """Dedicated parser for standard 3GPP NAS specifications (Clause 8 & 9 tables and IE definitions)."""

    def __init__(self, docx_paths: List[Path]):
        self.docx_paths = docx_paths
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

            for elem in list(body):
                if elem.tag == TAG_P:
                    if self._is_toc_paragraph(elem):
                        continue

                    p_text = _extract_p_text(elem)
                    if not p_text:
                        continue
                    last_paragraph_text = p_text

                    if p_text.startswith(("Table 8.", "Table D.")):
                        match_cap = RE_CAPTION.search(p_text)
                        if match_cap:
                            clause = match_cap.group(1).strip()
                            name = re.sub(r"(?i)\s+message\s+content", "", match_cap.group(2)).strip()
                            last_caption_info = (clause, name, p_text)
                    else:
                        last_caption_info = None

                    match_ie = RE_IE_HEADING.match(p_text)
                    if match_ie:
                        if current_ie_def:
                            self._finalize_ie_def(current_ie_def, ie_definitions)

                        cl = match_ie.group(1).strip()
                        ie_name = RE_TRAILING_PAGE_NUM.sub("", match_ie.group(2).strip()).strip()
                        current_ie_def = {
                            "clause": cl,
                            "ie_name": ie_name,
                            "html_content": [
                                f'<h2 style="color: #0D47A1; margin-top: 4px; margin-bottom: 6px; '
                                f'font-family: Segoe UI, sans-serif;">{html.escape(ie_name)} (Clause {html.escape(cl)})</h2>'
                            ],
                        }
                    elif RE_MAJOR_BOUNDARY.match(p_text) and not p_text.startswith("9."):
                        if current_ie_def:
                            self._finalize_ie_def(current_ie_def, ie_definitions)
                            current_ie_def = None
                    elif current_ie_def:
                        if p_text.startswith(("Figure 9.", "Figure D.")):
                            current_ie_def["html_content"].append(
                                f'<p style="font-weight: bold; color: #37474F; margin-top: 10px; margin-bottom: 4px;">{html.escape(p_text)}</p>'
                            )
                        elif p_text.startswith(("Table 9.", "Table D.")):
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
                    if last_caption_info:
                        clause, name, full_caption = last_caption_info
                        ies = self._parse_clause_8_table(elem)
                        if ies:
                            messages.append({
                                "clause": clause,
                                "message_name": name,
                                "table_caption": full_caption,
                                "ies": ies,
                            })
                        last_caption_info = None

                    elif current_ie_def:
                        is_diagram = "figure" in last_paragraph_text.lower() or any(
                            "octet" in _extract_tc_text(c).lower() for c in elem.iter(TAG_TC)
                        )
                        tbl_html = _convert_table_to_html(elem, is_figure_diagram=is_diagram)
                        if tbl_html:
                            current_ie_def["html_content"].append(tbl_html)

            if current_ie_def:
                self._finalize_ie_def(current_ie_def, ie_definitions)

        self._unroll_container_sub_ies(messages, ie_definitions)
        return messages, ie_definitions

    def _finalize_ie_def(self, current_ie_def: Dict[str, Any], ie_definitions: List[Dict[str, Any]]):
        clause = current_ie_def.get("clause", "").strip()
        ie_name = current_ie_def.get("ie_name", "").strip()
        html_blocks = current_ie_def.get("html_content", [])

        # Skip umbrella section headers, Void clauses, or empty heading stubs
        if not clause or not ie_name or len(html_blocks) <= 1:
            return
        if clause in STRUCTURAL_SECTION_CLAUSES:
            return
        lower_name = ie_name.lower()
        if lower_name.startswith("void") or lower_name in ("general", "overview", "information elements coding"):
            return

        html_str = "".join(html_blocks)
        ie_definitions.append({
            "clause": clause,
            "ie_name": ie_name,
            "raw_description": html_str,
            "structure_table": json.dumps([]),
        })

    def _unroll_container_sub_ies(
        self,
        messages: List[Dict[str, Any]],
        ie_definitions: List[Dict[str, Any]],
    ):
        """
        Unrolls encapsulated NAS container payloads (Payload container, Service-level-AA container,
        Type 6 IE container, ePCO) into depth > 0 child rows when the target Clause 9 / Annex D
        IE definitions exist in the parsed specification version.
        """
        defs_by_clause: Dict[str, Dict[str, Any]] = {
            d["clause"]: d for d in ie_definitions if d.get("clause")
        }
        if not defs_by_clause:
            return

        # Find helper table 8.2.7.54.1.1 (Registration accept type 6 IE container contents) if present
        type6_sub_ies: List[Dict[str, Any]] = []
        for m in messages:
            if m.get("clause", "").startswith("8.2.7.54"):
                type6_sub_ies = m.get("ies", [])
                break

        def _build_slaa_children(parent_path: str, base_depth: int) -> List[Dict[str, Any]]:
            children = []
            for sub_iei, fallback_name, sub_clause, pres, fmt, length in SERVICE_LEVEL_AA_SUB_IES:
                if sub_clause in defs_by_clause:
                    resolved_name = defs_by_clause[sub_clause]["ie_name"] or fallback_name
                    children.append({
                        "iei": sub_iei,
                        "information_element": resolved_name,
                        "field_path": f"{parent_path} > {resolved_name}",
                        "depth": base_depth,
                        "type_reference": f"{resolved_name} {sub_clause}",
                        "presence": pres,
                        "format": fmt,
                        "length": length,
                    })
            return children

        for msg in messages:
            # Skip unrolling inside the helper table itself
            if msg.get("clause", "").startswith("8.2.7.54"):
                continue

            expanded_ies: List[Dict[str, Any]] = []
            for ie in msg.get("ies", []):
                expanded_ies.append(ie)
                type_ref = ie.get("type_reference", "")
                parent_name = ie.get("information_element", "")

                # 1. Payload container (Clause 9.11.3.39)
                if "9.11.3.39" in type_ref:
                    for sub_iei, fallback_name, sub_clause, pres, fmt, length in PAYLOAD_CONTAINER_SUB_IES:
                        if sub_clause in defs_by_clause:
                            resolved_name = defs_by_clause[sub_clause]["ie_name"] or fallback_name
                            sub_path = f"{parent_name} > {resolved_name}"
                            expanded_ies.append({
                                "iei": sub_iei,
                                "information_element": resolved_name,
                                "field_path": sub_path,
                                "depth": 1,
                                "type_reference": f"{resolved_name} {sub_clause}",
                                "presence": pres,
                                "format": fmt,
                                "length": length,
                            })
                            if sub_clause == "9.11.2.10":
                                expanded_ies.extend(_build_slaa_children(sub_path, base_depth=2))

                # 2. Top-level Service-level-AA container (Clause 9.11.2.10)
                elif "9.11.2.10" in type_ref:
                    expanded_ies.extend(_build_slaa_children(parent_name, base_depth=1))

                # 3. Registration accept type 6 IE container (Clause 9.11.3.98)
                elif "9.11.3.98" in type_ref and type6_sub_ies:
                    for sub in type6_sub_ies:
                        sub_name = sub.get("information_element", "")
                        expanded_ies.append({
                            "iei": sub.get("iei", ""),
                            "information_element": sub_name,
                            "field_path": f"{parent_name} > {sub_name}",
                            "depth": 1,
                            "type_reference": sub.get("type_reference", ""),
                            "presence": sub.get("presence", ""),
                            "format": sub.get("format", ""),
                            "length": sub.get("length", ""),
                        })

                # 4. Extended protocol configuration options (Clause 9.11.4.6)
                elif "9.11.4.6" in type_ref:
                    for sub_iei, fallback_name, sub_clause, pres, fmt, length in EPCO_SUB_IES:
                        if sub_clause in defs_by_clause:
                            resolved_name = defs_by_clause[sub_clause]["ie_name"] or fallback_name
                            expanded_ies.append({
                                "iei": sub_iei,
                                "information_element": resolved_name,
                                "field_path": f"{parent_name} > {resolved_name}",
                                "depth": 1,
                                "type_reference": f"{resolved_name} {sub_clause}",
                                "presence": pres,
                                "format": fmt,
                                "length": length,
                            })

            msg["ies"] = expanded_ies

    def _parse_clause_8_table(self, tbl_elem: ET.Element) -> List[Dict[str, str]]:
        rows = tbl_elem.findall(TAG_TR)
        if not rows:
            return []

        header_row_idx = -1
        for r_idx in range(min(3, len(rows))):
            cells_text = [_extract_tc_text(tc).lower() for tc in rows[r_idx].findall(TAG_TC)]
            joined_header = " ".join(cells_text)
            if "information element" in joined_header or "iei" in cells_text:
                header_row_idx = r_idx
                break

        if header_row_idx == -1:
            return []

        ies = []
        for row in rows[header_row_idx + 1:]:
            cells = [_extract_tc_text(tc) for tc in row.findall(TAG_TC)]
            if len(cells) < 6:
                continue

            iei, ie_name, type_ref, presence, fmt, length = cells[:6]
            clean_name = ie_name.lower().replace(" ", "")
            if not clean_name or "informationelement" in clean_name:
                continue

            ies.append({
                "iei": iei,
                "information_element": ie_name,
                "field_path": ie_name,
                "depth": 0,
                "type_reference": type_ref,
                "presence": presence,
                "format": fmt,
                "length": length,
            })

        return ies
