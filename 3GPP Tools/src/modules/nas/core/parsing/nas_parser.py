import json
import logging
import re
from pathlib import Path
from typing import Union, List, Optional, Callable, Tuple, Dict, Any

try:
    from lxml import etree as ET
except ImportError:
    import xml.etree.ElementTree as ET

from modules.nas.core.parsing.asn1_parser import ASN1DocxParser
from modules.nas.core.parsing.protocol_parser_common import (
    TAG_BODY, TAG_P, _extract_p_text, TAG_TBL, _extract_tc_text,
    TAG_TC, _convert_table_to_html, TAG_TR, extract_document_root
)
from modules.specifications.utils.utils import file_version_to_version

RE_PART_INDEX = re.compile(r"_(\d+)_")
RE_SPEC_NUMBER = re.compile(r"(24|25|36|37|38)[._]?(301|501|331|413|423|412|473|463)")
RE_VERSION_STEM = re.compile(r"-([a-zA-Z0-9]{3})(?:_\d+.*)?$")
RE_CAPTION = re.compile(
    r"^Table\s+([8D]\.\d+(?:[\.\-/][0-9A-Za-z]+)*)\s*[:\.]\s*(.+?)(?:\s+message\s+content)?$",
    re.IGNORECASE,
)
RE_IE_HEADING = re.compile(r"^((?:9\.[2-9]|9\.1[0-9]|D\.6)(?:\.[0-9A-Za-z]+)*)\s+(.+)$")
RE_MAJOR_BOUNDARY = re.compile(r"^(?:[1-8]|10|11|12|Annex\s+[A-Z])\b")


class NASDocxParser:
    """Dedicated parser for standard 3GPP NAS specifications (Clause 8 & 9 tables and IE definitions)."""

    def __init__(self, docx_paths: List[Path]):
        self.docx_paths = docx_paths
        self.logger = logging.getLogger(__name__)

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
                        ie_name = match_ie.group(2).strip()
                        current_ie_def = {
                            "clause": cl,
                            "ie_name": ie_name,
                            "html_content": [
                                f'<h2 style="color: #0D47A1; margin-top: 4px; margin-bottom: 6px; '
                                f'font-family: Segoe UI, sans-serif;">{ie_name} (Clause {cl})</h2>'
                            ],
                        }
                    elif RE_MAJOR_BOUNDARY.match(p_text) and not p_text.startswith("9."):
                        if current_ie_def:
                            self._finalize_ie_def(current_ie_def, ie_definitions)
                            current_ie_def = None
                    elif current_ie_def:
                        if p_text.startswith(("Figure 9.", "Figure D.")):
                            current_ie_def["html_content"].append(
                                f'<p style="font-weight: bold; color: #37474F; margin-top: 10px; margin-bottom: 4px;">{p_text}</p>'
                            )
                        elif p_text.startswith(("Table 9.", "Table D.")):
                            current_ie_def["html_content"].append(
                                f'<p style="font-weight: bold; color: #1A237E; margin-top: 12px; margin-bottom: 4px;">{p_text}</p>'
                            )
                        elif p_text.startswith("NOTE"):
                            current_ie_def["html_content"].append(
                                f'<div style="background-color: #F0F4F8; border-left: 3px solid #90A4AE; '
                                f'padding: 4px 8px; margin: 4px 0; font-size: 11px; color: #455A64;">{p_text}</div>'
                            )
                        else:
                            current_ie_def["html_content"].append(
                                f'<p style="margin: 4px 0; line-height: 1.4; color: #263238;">{p_text}</p>'
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

        return messages, ie_definitions

    def _finalize_ie_def(self, current_ie_def: Dict[str, Any], ie_definitions: List[Dict[str, Any]]):
        html_str = "".join(current_ie_def["html_content"])
        ie_definitions.append({
            "clause": current_ie_def["clause"],
            "ie_name": current_ie_def["ie_name"],
            "raw_description": html_str,
            "structure_table": json.dumps([]),
        })

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


class ProtocolDocxDispatcher:
    """Unified entry point and dispatcher routing docx files to NAS or ASN.1 parsers."""

    def __init__(self, docx_paths: Union[Path, str, List[Union[Path, str]]]):
        if isinstance(docx_paths, (str, Path)):
            self.docx_paths = [Path(docx_paths)]
        else:
            self.docx_paths = [Path(p) for p in docx_paths]

        self.docx_paths.sort(key=self._extract_part_index)
        self.logger = logging.getLogger(__name__)

    @staticmethod
    def _extract_part_index(path: Path) -> int:
        match = RE_PART_INDEX.search(path.name)
        return int(match.group(1)) if match else 0

    def extract_spec_number(self) -> str:
        if not self.docx_paths:
            return "24.501"
        match = RE_SPEC_NUMBER.search(self.docx_paths[0].stem)
        return f"{match.group(1)}.{match.group(2)}" if match else "24.501"

    def extract_version_from_filename(self) -> str:
        if not self.docx_paths:
            return ""
        stem = self.docx_paths[0].stem
        match = RE_VERSION_STEM.search(stem)
        if match:
            parsed_ver = file_version_to_version(match.group(1))
            if parsed_ver:
                return parsed_ver
        return stem

    def parse(
            self, progress_callback: Optional[Callable[[str, int], None]] = None
    ) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
        spec_num = self.extract_spec_number()

        # Route ASN.1 specifications (38.331, 36.331, 38.413, 38.423, 36.413)
        if any(s in spec_num for s in ["38.331", "36.331", "38.413", "38.423", "36.413", "38.473"]):
            parser = ASN1DocxParser(self.docx_paths, spec_number=spec_num)
            return parser.parse(progress_callback=progress_callback)

        # Route NAS specifications (24.501, 24.301, 24.008)
        parser = NASDocxParser(self.docx_paths)
        messages, ie_definitions = parser.parse(progress_callback=progress_callback)

        if progress_callback:
            progress_callback(f"Extracted {len(messages)} messages and {len(ie_definitions)} definitions.", 95)

        return messages, ie_definitions


# Backward-compatibility alias
NASDocxParserAlias = ProtocolDocxDispatcher