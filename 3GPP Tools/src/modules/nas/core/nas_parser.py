import json
import logging
from pathlib import Path
import re
from typing import Any, Callable, Dict, List, Optional, Tuple
import zipfile

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
TAG_TCPR = f"{W_NS}tcPr"
TAG_GRIDSPAN = f"{W_NS}gridSpan"
TAG_VMERGE = f"{W_NS}vMerge"
TAG_PPR = f"{W_NS}pPr"
TAG_JC = f"{W_NS}jc"


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


def _get_tc_alignment(tc_elem) -> str:
    """Detects text alignment from the first paragraph in a table cell."""
    p = tc_elem.find(TAG_P)
    if p is not None:
        pPr = p.find(TAG_PPR)
        if pPr is not None:
            jc = pPr.find(TAG_JC)
            if jc is not None:
                val = jc.get(f"{W_NS}val")
                if val in ("center", "right", "left"):
                    return val
    return "left"


def _convert_table_to_html(tbl_elem, is_figure_diagram: bool = False) -> str:
    """Converts a Word XML table into a styled HTML table supporting colspan and vertical alignment."""
    rows = tbl_elem.findall(TAG_TR)
    if not rows:
        return ""

    table_style = (
        "border-collapse: collapse; margin: 8px 0; border: 1px solid #546E7A; "
        "font-family: 'Segoe UI', Arial, sans-serif; font-size: 11px; "
    )
    table_style += "width: 100%;" if not is_figure_diagram else "min-width: 320px; max-width: 600px;"

    html_parts = [f'<table border="1" cellspacing="0" cellpadding="4" style="{table_style}">']

    for r_idx, row in enumerate(rows):
        html_parts.append("<tr>")
        cells = row.findall(TAG_TC)

        for cell in cells:
            tcPr = cell.find(TAG_TCPR)
            colspan = 1
            is_vmerge_continue = False

            if tcPr is not None:
                # 1. Parse Column Span (gridSpan)
                gs = tcPr.find(TAG_GRIDSPAN)
                if gs is not None:
                    val = gs.get(f"{W_NS}val")
                    if val and val.isdigit():
                        colspan = int(val)

                # 2. Parse Vertical Merge (vMerge)
                vm = tcPr.find(TAG_VMERGE)
                if vm is not None:
                    val = vm.get(f"{W_NS}val")
                    if val != "restart":
                        is_vmerge_continue = True

            # Skip continuation cells for vertical merges
            if is_vmerge_continue:
                continue

            cell_text = _extract_tc_text(cell)
            align = _get_tc_alignment(cell)

            # Auto-center single digit headers (bit numbers 8, 7, 6...)
            if is_figure_diagram or (len(cell_text) <= 2 and cell_text.isdigit()):
                align = "center"

            tag = "th" if (r_idx == 0 and not is_figure_diagram) else "td"

            # Styling per cell type
            style_bits = [
                "border: 1px solid #78909C;",
                "padding: 4px 6px;",
                f"text-align: {align};",
            ]

            if r_idx == 0 and not is_figure_diagram:
                style_bits.append("background-color: #ECEFF1; font-weight: bold; color: #263238;")
            elif is_figure_diagram and r_idx == 0:
                style_bits.append("background-color: #F5F7F8; font-weight: bold; color: #37474F;")

            # Format octet index columns distinctly
            if "octet" in cell_text.lower():
                style_bits.append("font-weight: bold; background-color: #FAFAFA; white-space: nowrap;")

            colspan_attr = f' colspan="{colspan}"' if colspan > 1 else ""
            style_str = " ".join(style_bits)

            html_parts.append(f'<{tag}{colspan_attr} style="{style_str}">{cell_text}</{tag}>')

        html_parts.append("</tr>")

    html_parts.append("</table>")
    return "".join(html_parts)


class NASDocxParser:
    """High-performance direct XML parser for 3GPP TS 24.501 specifications."""

    def __init__(self, docx_path: Path):
        self.docx_path = Path(docx_path)
        self.logger = logging.getLogger(__name__)

    def extract_version_from_filename(self) -> str:
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
        ie_heading_pattern = re.compile(
            r"^((?:9\.[2-9]|9\.1[0-9]|D\.6)(?:\.[0-9A-Za-z]+)*)\s+(.+)$"
        )
        major_boundary_pattern = re.compile(r"^(?:[1-8]|10|11|12|Annex\s+[A-Z])\b")

        last_caption_info: Optional[Tuple[str, str, str]] = None
        last_paragraph_text: str = ""
        current_ie_def: Optional[Dict[str, Any]] = None

        body_elements = list(body)
        total_elements = len(body_elements)

        if progress_callback:
            progress_callback("Extracting Clause 8 tables & Clause 9 definitions...", 40)

        for idx, elem in enumerate(body_elements):
            if elem.tag == TAG_P:
                p_text = _extract_p_text(elem)
                if p_text:
                    last_paragraph_text = p_text

                    # 1. Message Table Captions
                    if p_text.startswith("Table 8.") or p_text.startswith("Table D."):
                        match_cap = caption_pattern.search(p_text)
                        if match_cap:
                            clause = match_cap.group(1).strip()
                            name = match_cap.group(2).strip()
                            name = re.sub(r"(?i)\s+message\s+content", "", name).strip()
                            last_caption_info = (clause, name, p_text)
                    else:
                        last_caption_info = None

                    # 2. IE Headings
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
                            ]
                        }
                    elif major_boundary_pattern.match(p_text) and not p_text.startswith("9."):
                        if current_ie_def:
                            self._finalize_ie_def(current_ie_def, ie_definitions)
                            current_ie_def = None
                    elif current_ie_def:
                        # Format body text, figure captions, and notes
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
                # 3. Clause 8 Message Tables
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

                # 4. Clause 9 Structure / Description Tables
                elif current_ie_def:
                    is_diagram = "figure" in last_paragraph_text.lower() or any(
                        "octet" in _extract_tc_text(c).lower() for c in elem.iter(TAG_TC)
                    )
                    tbl_html = _convert_table_to_html(elem, is_figure_diagram=is_diagram)
                    if tbl_html:
                        current_ie_def["html_content"].append(tbl_html)

            if idx % 500 == 0 and progress_callback:
                progress = 40 + int((idx / max(1, total_elements)) * 50)
                progress_callback(f"Scanning document ({idx}/{total_elements})...", progress)

        if current_ie_def:
            self._finalize_ie_def(current_ie_def, ie_definitions)

        if progress_callback:
            progress_callback(f"Extracted {len(messages)} messages and {len(ie_definitions)} IE definitions.", 95)

        return messages, ie_definitions

    def _finalize_ie_def(self, current_ie_def: Dict[str, Any], ie_definitions: List[Dict[str, Any]]):
        html = "".join(current_ie_def["html_content"])
        ie_definitions.append({
            "clause": current_ie_def["clause"],
            "ie_name": current_ie_def["ie_name"],
            "raw_description": html,
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
                "type_reference": type_ref,
                "presence": presence,
                "format": fmt,
                "length": length,
            })

        return ies