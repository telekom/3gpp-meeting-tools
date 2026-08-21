import json
import logging
import re
import zipfile
from pathlib import Path
from typing import Union, List, Optional, Callable, Tuple, Dict, Any

from lxml import etree as ET

from modules.nas.core.parsing.asn1_parser import RRCAsn1DocxParser
from modules.nas.core.parsing.protocol_parser_common import TAG_BODY, TAG_P, _extract_p_text, TAG_TBL, _extract_tc_text, \
    TAG_TC, _convert_table_to_html, TAG_TR
from modules.specifications.utils.utils import file_version_to_version

# =========================================================================
# --- UNIFIED PARSER DISPATCHER ---
# =========================================================================
class NASDocxParser:
    """Direct XML parser for 3GPP Specifications, supporting both NAS and RRC/NGAP ASN.1."""

    def __init__(self, docx_paths: Union[Path, str, List[Union[Path, str]]]):
        if isinstance(docx_paths, (str, Path)):
            self.docx_paths = [Path(docx_paths)]
        else:
            self.docx_paths = [Path(p) for p in docx_paths]

        self.docx_paths.sort(key=lambda p: self._extract_part_index(p.name))
        self.logger = logging.getLogger(__name__)

    @staticmethod
    def _extract_part_index(filename: str) -> int:
        match = re.search(r"_(\d+)_", filename)
        return int(match.group(1)) if match else 0

    def extract_spec_number(self) -> str:
        if not self.docx_paths:
            return "24.501"
        stem = self.docx_paths[0].stem.replace(".", "").replace("-", "").replace("_", "")
        for known in ["38331", "36331", "38413", "24301", "24501"]:
            if known in stem:
                return f"{known[:2]}.{known[2:]}"
        match = re.search(r"(24|36|38)[._]?(301|501|331|413)", self.docx_paths[0].stem)
        if match:
            return f"{match.group(1)}.{match.group(2)}"
        return "24.501"

    def extract_version_from_filename(self) -> str:
        if not self.docx_paths:
            return ""
        stem = self.docx_paths[0].stem
        match = re.search(r"-([a-zA-Z0-9]{3})(?:_\d+.*)?$", stem)
        if match:
            parsed_ver = file_version_to_version(match.group(1))
            if parsed_ver:
                return parsed_ver
        return stem

    def parse(
            self, progress_callback: Optional[Callable[[str, int], None]] = None
    ) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
        spec_num = self.extract_spec_number()

        # Route ASN.1 specifications (38.331, 36.331, 38.413) to RRCAsn1DocxParser
        if any(s in spec_num for s in ["38.331", "36.331", "38.413"]):
            rrc_parser = RRCAsn1DocxParser(self.docx_paths, spec_number=spec_num)
            return rrc_parser.parse(progress_callback=progress_callback)

        # Standard NAS Clause 8/9 Parser
        valid_paths = [p for p in self.docx_paths if p.exists()]
        if not valid_paths:
            raise FileNotFoundError(f"Specification file(s) not found: {self.docx_paths}")

        messages: List[Dict[str, Any]] = []
        ie_definitions: List[Dict[str, Any]] = []

        caption_pattern = re.compile(
            r"^Table\s+([8D]\.\d+(?:[\.\-/][0-9A-Za-z]+)*)\s*[:\.]\s*(.+?)(?:\s+message\s+content)?$",
            re.IGNORECASE,
        )
        ie_heading_pattern = re.compile(
            r"^((?:9\.[2-9]|9\.1[0-9]|D\.6)(?:\.[0-9A-Za-z]+)*)\s+(.+)$"
        )
        major_boundary_pattern = re.compile(r"^(?:[1-8]|10|11|12|Annex\s+[A-Z])\b")

        total_files = len(valid_paths)

        for file_idx, docx_path in enumerate(valid_paths):
            file_weight = 1.0 / total_files
            base_file_progress = 10 + int((file_idx / total_files) * 80)

            if progress_callback:
                progress_callback(f"Reading {docx_path.name} ({file_idx + 1}/{total_files})...", base_file_progress)

            try:
                with zipfile.ZipFile(docx_path, "r") as zf:
                    if "word/document.xml" not in zf.namelist():
                        continue
                    xml_bytes = zf.read("word/document.xml")
            except Exception as e:
                self.logger.warning(f"Failed to read word/document.xml from {docx_path.name}: {e}")
                continue

            root = ET.fromstring(xml_bytes)
            body = root.find(TAG_BODY)
            if body is None:
                continue

            last_caption_info: Optional[Tuple[str, str, str]] = None
            last_paragraph_text: str = ""
            current_ie_def: Optional[Dict[str, Any]] = None

            body_elements = list(body)
            total_elements = len(body_elements)

            for idx, elem in enumerate(body_elements):
                if elem.tag == TAG_P:
                    p_text = _extract_p_text(elem)
                    if p_text:
                        last_paragraph_text = p_text

                        if p_text.startswith("Table 8.") or p_text.startswith("Table D."):
                            match_cap = caption_pattern.search(p_text)
                            if match_cap:
                                clause = match_cap.group(1).strip()
                                name = match_cap.group(2).strip()
                                name = re.sub(r"(?i)\s+message\s+content", "", name).strip()
                                last_caption_info = (clause, name, p_text)
                        else:
                            last_caption_info = None

                        match_ie = ie_heading_pattern.match(p_text)
                        if match_ie:
                            if current_ie_def:
                                self._finalize_ie_def(current_ie_def, ie_definitions)

                            cl = match_ie.group(1).strip()
                            ie_name = match_ie.group(2).strip()
                            current_ie_def = {
                                "clause": cl,
                                "ie_name": ie_name,
                                "html_content": [
                                    f'<h2 style="color: #0D47A1; margin-top: 4px; margin-bottom: 6px; font-family: Segoe UI, sans-serif;">{ie_name} (Clause {cl})</h2>'
                                ],
                            }
                        elif major_boundary_pattern.match(p_text) and not p_text.startswith("9."):
                            if current_ie_def:
                                self._finalize_ie_def(current_ie_def, ie_definitions)
                                current_ie_def = None
                        elif current_ie_def:
                            if p_text.startswith("Figure 9.") or p_text.startswith("Figure D."):
                                current_ie_def["html_content"].append(
                                    f'<p style="font-weight: bold; color: #37474F; margin-top: 10px; margin-bottom: 4px;">{p_text}</p>'
                                )
                            elif p_text.startswith("Table 9.") or p_text.startswith("Table D."):
                                current_ie_def["html_content"].append(
                                    f'<p style="font-weight: bold; color: #1A237E; margin-top: 12px; margin-bottom: 4px;">{p_text}</p>'
                                )
                            elif p_text.startswith("NOTE"):
                                current_ie_def["html_content"].append(
                                    f'<div style="background-color: #F0F4F8; border-left: 3px solid #90A4AE; padding: 4px 8px; margin: 4px 0; font-size: 11px; color: #455A64;">{p_text}</div>'
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

        if progress_callback:
            progress_callback(f"Extracted {len(messages)} messages and {len(ie_definitions)} definitions.", 95)

        return messages, ie_definitions

    def _finalize_ie_def(self, current_ie_def: Dict[str, Any], ie_definitions: List[Dict[str, Any]]):
        html_str = "".join(current_ie_def["html_content"])
        ie_definitions.append({
            "clause": current_ie_def["clause"],
            "ie_name": current_ie_def["ie_name"],
            "raw_description": html_str,
            "structure_table": json.dumps([]),
        })

    def _parse_clause_8_table(self, tbl_elem) -> List[Dict[str, str]]:
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

            iei = cells[0]
            ie_name = cells[1]
            type_ref = cells[2]
            presence = cells[3]
            fmt = cells[4]
            length = cells[5]

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
