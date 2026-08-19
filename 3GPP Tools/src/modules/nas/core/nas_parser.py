import json
import logging
from pathlib import Path
import re
from typing import Any, Callable, Dict, List, Optional, Tuple
import zipfile

# Fast C-based XML parsing via lxml with standard library fallback
try:
    from lxml import etree as ET
except ImportError:
    import xml.etree.ElementTree as ET

from modules.specifications.utils.utils import file_version_to_version

# WordprocessingML XML Namespaces
W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
TAG_BODY = f"{W_NS}body"
TAG_P = f"{W_NS}p"
TAG_TBL = f"{W_NS}tbl"
TAG_TR = f"{W_NS}tr"
TAG_TC = f"{W_NS}tc"
TAG_T = f"{W_NS}t"
TAG_TAB = f"{W_NS}tab"
TAG_BR = f"{W_NS}br"
TAG_CR = f"{W_NS}cr"
TAG_HYPHEN = f"{W_NS}noBreakHyphen"


def _extract_p_text(p_elem) -> str:
    """Extracts clean text from a <w:p> node, preserving spaces, tabs, and line breaks."""
    text_pieces = []
    for node in p_elem.iter():
        tag = node.tag
        if tag == TAG_T:
            if node.text:
                text_pieces.append(node.text)
        elif tag == TAG_TAB:
            text_pieces.append(" ")
        elif tag in (TAG_BR, TAG_CR):
            text_pieces.append(" ")
        elif tag == TAG_HYPHEN:
            text_pieces.append("-")

    raw = "".join(text_pieces)
    raw = raw.replace("\u00a0", " ").replace("\xa0", " ")
    return " ".join(raw.split())


def _extract_tc_text(tc_elem) -> str:
    """Extracts text from a table cell <w:tc>, joining multiple paragraphs with spaces."""
    p_texts = []
    for p in tc_elem.findall(TAG_P):
        pt = _extract_p_text(p)
        if pt:
            p_texts.append(pt)
    return " ".join(p_texts).strip()


def _convert_table_to_html(tbl_elem) -> str:
    """Converts a Word XML table element into clean, styled HTML markup."""
    rows = tbl_elem.findall(TAG_TR)
    if not rows:
        return ""

    html_parts = [
        '<table border="1" cellspacing="0" cellpadding="4" '
        'style="border-collapse: collapse; width: 100%; margin: 8px 0; '
        'border: 1px solid #B0BEC5; font-size: 11px; font-family: Segoe UI, sans-serif;">'
    ]

    for r_idx, row in enumerate(rows):
        html_parts.append("<tr>")
        cells = row.findall(TAG_TC)
        for cell in cells:
            cell_text = _extract_tc_text(cell)
            tag = "th" if r_idx == 0 else "td"
            style = "background-color: #ECEFF1; font-weight: bold;" if r_idx == 0 else ""
            html_parts.append(f'<{tag} style="{style} border: 1px solid #CFD8DC; padding: 4px;">{cell_text}</{tag}>')
        html_parts.append("</tr>")

    html_parts.append("</table>")
    return "".join(html_parts)


class NASDocxParser:
    """High-performance direct XML parser for 3GPP TS 24.501 specifications."""

    def __init__(self, docx_path: Path):
        self.docx_path = Path(docx_path)
        self.logger = logging.getLogger(__name__)

    def extract_version_from_filename(self) -> str:
        """Translates filenames like '24501-j30.docx' to standard versions like '19.3.0'."""
        stem = self.docx_path.stem
        match = re.search(r"-([a-zA-Z0-9]{3})$", stem)
        if match:
            parsed_ver = file_version_to_version(match.group(1))
            if parsed_ver:
                return parsed_ver
        return stem

    def parse(
            self, progress_callback: Optional[Callable[[str, int], None]] = None
    ) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
        if not self.docx_path.exists():
            raise FileNotFoundError(f"Specification file not found: {self.docx_path}")

        if progress_callback:
            progress_callback("Reading document archive into memory...", 10)

        try:
            with zipfile.ZipFile(self.docx_path, "r") as zf:
                xml_bytes = zf.read("word/document.xml")
        except Exception as e:
            raise ValueError(f"Failed to read word/document.xml from {self.docx_path.name}: {e}")

        if progress_callback:
            progress_callback("Building XML element tree...", 25)

        root = ET.fromstring(xml_bytes)
        body = root.find(TAG_BODY)
        if body is None:
            return [], []

        messages: List[Dict[str, Any]] = []
        ie_definitions: List[Dict[str, Any]] = []

        caption_pattern = re.compile(
            r"^Table\s+([8D]\.\d+(?:[\.\-/][0-9A-Za-z]+)*)\s*[:\.]\s*(.+?)(?:\s+message\s+content)?$",
            re.IGNORECASE,
        )

        # Matches all Clause 9 (e.g., 9.2, 9.7, 9.11.3.37) and Annex D (e.g., D.6.2) IE headings
        ie_heading_pattern = re.compile(
            r"^((?:9\.[2-9]|9\.1[0-9]|D\.6)(?:\.[0-9A-Za-z]+)*)\s+(.+)$"
        )

        # Identifies top-level major section transitions to close active IE capture
        major_boundary_pattern = re.compile(r"^(?:[1-8]|10|11|12|Annex\s+[A-Z])\b")

        last_caption_info: Optional[Tuple[str, str, str]] = None
        current_ie_def: Optional[Dict[str, Any]] = None

        body_elements = list(body)
        total_elements = len(body_elements)

        if progress_callback:
            progress_callback("Extracting Clause 8 tables & Clause 9 definitions...", 40)

        for idx, elem in enumerate(body_elements):
            if elem.tag == TAG_P:
                p_text = _extract_p_text(elem)
                if p_text:
                    # 1. Check for Table Captions (Clause 8 / Annex D)
                    if p_text.startswith("Table 8.") or p_text.startswith("Table D."):
                        match_cap = caption_pattern.search(p_text)
                        if match_cap:
                            clause = match_cap.group(1).strip()
                            name = match_cap.group(2).strip()
                            name = re.sub(r"(?i)\s+message\s+content", "", name).strip()
                            last_caption_info = (clause, name, p_text)
                    else:
                        last_caption_info = None

                    # 2. Check for Clause 9 / Annex D IE Heading
                    match_ie = ie_heading_pattern.match(p_text)
                    if match_ie:
                        # Flush previous active IE definition
                        if current_ie_def:
                            self._finalize_ie_def(current_ie_def, ie_definitions)

                        cl = match_ie.group(1).strip()
                        ie_name = match_ie.group(2).strip()
                        current_ie_def = {
                            "clause": cl,
                            "ie_name": ie_name,
                            "html_content": [
                                f'<h2 style="color: #1565C0; margin-bottom: 4px;">{ie_name} (Clause {cl})</h2>']
                        }
                    elif major_boundary_pattern.match(p_text) and not p_text.startswith("9."):
                        # Major clause boundary reached
                        if current_ie_def:
                            self._finalize_ie_def(current_ie_def, ie_definitions)
                            current_ie_def = None
                    elif current_ie_def:
                        # Accumulate body text and figure captions
                        if p_text.startswith("Figure 9.") or p_text.startswith("Table 9."):
                            current_ie_def["html_content"].append(
                                f'<p style="font-weight: bold; color: #37474F; margin-top: 10px; margin-bottom: 4px;">{p_text}</p>'
                            )
                        elif p_text.startswith("NOTE"):
                            current_ie_def["html_content"].append(
                                f'<p style="color: #616161; font-size: 11px; margin: 4px 0; padding-left: 10px; border-left: 3px solid #B0BEC5;">{p_text}</p>'
                            )
                        else:
                            current_ie_def["html_content"].append(
                                f'<p style="margin: 6px 0; line-height: 1.4;">{p_text}</p>'
                            )

            elif elem.tag == TAG_TBL:
                # 3. Check for Message Content Table
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

                # 4. Check for Clause 9 Structure / Description Table
                elif current_ie_def:
                    tbl_html = _convert_table_to_html(elem)
                    if tbl_html:
                        current_ie_def["html_content"].append(tbl_html)

            if idx % 500 == 0 and progress_callback:
                progress = 40 + int((idx / max(1, total_elements)) * 50)
                progress_callback(f"Scanning document ({idx}/{total_elements})...", progress)

        # Flush final IE definition
        if current_ie_def:
            self._finalize_ie_def(current_ie_def, ie_definitions)

        if progress_callback:
            progress_callback(
                f"Extracted {len(messages)} messages and {len(ie_definitions)} IE definitions.",
                95,
            )

        return messages, ie_definitions

    def _finalize_ie_def(self, current_ie_def: Dict[str, Any], ie_definitions: List[Dict[str, Any]]):
        """Compiles accumulated HTML content and stores the definition record."""
        html = "".join(current_ie_def["html_content"])
        ie_definitions.append({
            "clause": current_ie_def["clause"],
            "ie_name": current_ie_def["ie_name"],
            "raw_description": html,
            "structure_table": json.dumps([]),
        })

    def _parse_clause_8_table(self, tbl_elem) -> List[Dict[str, str]]:
        """Extracts the 6 standard columns from a message table XML node."""
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
                "type_reference": type_ref,
                "presence": presence,
                "format": fmt,
                "length": length,
            })

        return ies