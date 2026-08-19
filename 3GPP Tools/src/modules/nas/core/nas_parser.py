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

# WordprocessingML Namespaces
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
    # Normalize non-breaking spaces (\u00a0 / \xa0) and collapse whitespace
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

        # 1. Read document.xml directly from the ZIP package
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

        # Regex for Clause 8 / Annex D table captions and Clause 9 / Annex D IE headings
        caption_pattern = re.compile(
            r"^Table\s+([8D]\.\d+(?:[\.\-/][0-9A-Za-z]+)*)\s*[:\.]\s*(.+?)(?:\s+message\s+content)?$",
            re.IGNORECASE,
        )
        ie_heading_pattern = re.compile(
            r"^((?:9\.11|D\.6)(?:\.[0-9A-Za-z]+)+)\s+(.+)$"
        )

        last_caption_info: Optional[Tuple[str, str, str]] = None
        body_elements = list(body)
        total_elements = len(body_elements)

        if progress_callback:
            progress_callback("Extracting Clause 8 tables & Clause 9 definitions...", 40)

        # 2. Single-pass linear scan through all body elements
        for idx, elem in enumerate(body_elements):
            if elem.tag == TAG_P:
                p_text = _extract_p_text(elem)
                if p_text:
                    # Check for Table Captions
                    if p_text.startswith("Table 8.") or p_text.startswith("Table D."):
                        match_cap = caption_pattern.search(p_text)
                        if match_cap:
                            clause = match_cap.group(1).strip()
                            name = match_cap.group(2).strip()
                            name = re.sub(r"(?i)\s+message\s+content", "", name).strip()
                            last_caption_info = (clause, name, p_text)
                    else:
                        # Clear caption buffer only if non-empty regular text is encountered
                        last_caption_info = None

                    # Check for Clause 9 / Annex D IE headings
                    match_ie = ie_heading_pattern.match(p_text)
                    if match_ie:
                        cl = match_ie.group(1).strip()
                        ie_name = match_ie.group(2).strip()
                        ie_definitions.append({
                            "clause": cl,
                            "ie_name": ie_name,
                            "raw_description": f"Definition and coding for '{ie_name}' (Clause {cl}).",
                            "structure_table": json.dumps([]),
                        })

            elif elem.tag == TAG_TBL:
                # If the preceding paragraph was a table caption, parse the table
                if last_caption_info:
                    clause, name, full_caption = last_caption_info
                    ies = self._parse_table_xml(elem)
                    if ies:
                        messages.append({
                            "clause": clause,
                            "message_name": name,
                            "table_caption": full_caption,
                            "ies": ies,
                        })
                    last_caption_info = None

            if idx % 500 == 0 and progress_callback:
                progress = 40 + int((idx / max(1, total_elements)) * 50)
                progress_callback(f"Scanning document ({idx}/{total_elements})...", progress)

        if progress_callback:
            progress_callback(
                f"Extracted {len(messages)} messages and {len(ie_definitions)} IE definitions.",
                95,
            )

        return messages, ie_definitions

    def _parse_table_xml(self, tbl_elem) -> List[Dict[str, str]]:
        """Extracts the 6 standard columns from a message table XML node."""
        rows = tbl_elem.findall(TAG_TR)
        if not rows:
            return []

        # Find the header row (inspect rows 0 through 2 for header tokens)
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

            # Filter out empty or repeated header rows
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