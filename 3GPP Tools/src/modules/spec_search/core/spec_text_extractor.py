"""
Extracts plain-text clauses from 3GPP .docx specification documents directly via lxml.
Accurately identifies genuine 3GPP Clause Headings while strictly rejecting body prose references,
Table of Contents (TOC) entries, and procedure steps.
"""

import logging
from pathlib import Path
import re
from typing import Any, Callable, Dict, List, Optional

try:
    from lxml import etree as ET
except ImportError:
    import xml.etree.ElementTree as ET

from modules.nas.core.parsing.protocol_parser_constants import (
    TAG_BODY,
    TAG_P,
    TAG_TBL,
    TAG_TC,
    TAG_TR,
)
from modules.nas.core.parsing.protocol_parser_utils import (
    _extract_p_text,
    _extract_tc_text,
    extract_document_root,
)

# Multi-level clause headings (e.g., '4.1', '5.2.3.1', '4.3.2.2.1')
# Strictly requires starting directly with digits and a capitalized title (no 'Clause' prefix)
RE_SUBCLAUSE_HEADER = re.compile(
    r"^((?:[1-9]\d*(?:\.\d+)+))\s+([A-Z0-9][^\n\r]*)$"
)

# Top-level single-digit clause headings (e.g., '1 Scope', '2 References', '4 General')
RE_TOPLEVEL_CLAUSE = re.compile(
    r"^([1-9]\d*)\s+([A-Z][A-Za-z0-9\s,\-\(\)\/]{1,80})$"
)

# Annex headings (e.g., 'Annex A (normative): Title', 'Annex B (informative)')
RE_ANNEX_HEADER = re.compile(
    r"^(Annex\s+[A-Z])(?:\s*\((?:normative|informative)\))?[:\s]*(.*)$",
    re.IGNORECASE,
)

# Disqualify common sentence verbs and prose starters from being treated as titles
RE_PROSE_VERBS = re.compile(
    r"^(?:specifies|defines|describes|is\s+used\s+to|provides|indicates|shall|may|refers\s+to|contains|applies\s+to|illustrates|shows|according\s+to)",
    re.IGNORECASE,
)

# Disqualify procedure steps and list items
RE_PROCEDURE_STEP_EXCLUDE = re.compile(
    r"^(?:From\s+\w+\s+to\s+\w+|The\s+\w+\s+shall|If\s+the|When\s+the|In\s+order|Step\s+\d+|[a-z]\))",
    re.IGNORECASE,
)

# Identifies Table of Contents dot leaders and trailing page numbers
RE_TOC_PAGE_NUMBER = re.compile(r"(?:\.{3,}|\t+|\s{3,})\d+$|\s+\d{1,4}$")


def _get_paragraph_style(p_elem: ET.Element) -> str:
    """Extracts the OpenXML paragraph style identifier (e.g. 'heading 1', 'heading 4', 'annex')."""
    for child in p_elem:
        if child.tag.endswith("pPr"):
            for sub in child:
                if sub.tag.endswith("pStyle"):
                    for k, v in sub.attrib.items():
                        if k.endswith("val") or k == "val":
                            return str(v).lower()
    return ""


def _is_toc_paragraph(p_elem: ET.Element, p_text: str, p_style: str) -> bool:
    """Detects whether a paragraph belongs to the Table of Contents."""
    if any(toc_key in p_style for toc_key in ["toc", "contents", "table of contents"]):
        return True

    # Check for Word field codes or dot leaders
    xml_str = ET.tostring(p_elem, encoding="utf-8").decode("utf-8", errors="ignore")
    if "TOC \\" in xml_str or "PAGEREF " in xml_str or "fldSimple" in xml_str or 'w:leader="dot"' in xml_str:
        return True

    # Check for dot leaders / trailing page numbers without heading styles
    if RE_TOC_PAGE_NUMBER.search(p_text.strip()):
        is_true_heading_style = any(
            h in p_style for h in ["heading", "h1", "h2", "h3", "h4", "h5", "h6", "h7", "h8", "h9", "annex", "ti"]
        )
        if not is_true_heading_style:
            return True

    return False


def _is_valid_clause_title(title_text: str, is_heading_style: bool) -> bool:
    """Validates that candidate text is a genuine 3GPP clause title and not body prose."""
    clean = title_text.strip()
    if not clean:
        return False

    # If explicitly styled with Heading 1-9 in Word, accept unless it's a known step pattern
    if is_heading_style:
        return not bool(RE_PROCEDURE_STEP_EXCLUDE.match(clean))

    # Strict validation for paragraphs without explicit heading styles:
    # 1. 3GPP clause titles do not end with periods, colons, or semicolons
    if clean.endswith((".", ":", ";")):
        return False

    # 2. Must not start with prose verbs (e.g. 'specifies...', 'defines...')
    if RE_PROSE_VERBS.match(clean):
        return False

    # 3. Must not match procedure step patterns
    if RE_PROCEDURE_STEP_EXCLUDE.match(clean):
        return False

    # 4. 3GPP clause titles are concise title-case phrases (rarely exceeding 100 characters)
    if len(clean) > 120:
        return False

    return True


class SpecDocxExtractor:
    """Parses .docx specifications into structured clause records for indexing."""

    def __init__(self, docx_paths: List[Path]):
        self.docx_paths = [Path(p) for p in docx_paths if Path(p).exists()]
        self.logger = logging.getLogger(__name__)

    def extract_clauses(
        self, progress_callback: Optional[Callable[[str, int], None]] = None
    ) -> List[Dict[str, Any]]:
        """Walks through the document body and splits contents cleanly by clause boundary."""
        if not self.docx_paths:
            return []

        all_clauses: List[Dict[str, Any]] = []
        total_files = len(self.docx_paths)

        for file_idx, docx_path in enumerate(self.docx_paths):
            if progress_callback:
                pct = int((file_idx / total_files) * 90)
                progress_callback(f"Extracting clauses from {docx_path.name}...", pct)

            root = extract_document_root(docx_path)
            if root is None:
                continue

            body = root.find(TAG_BODY)
            if body is None:
                continue

            current_clause: Dict[str, Any] = {
                "clause_number": "0",
                "clause_title": "Front Matter / Document Header",
                "text_fragments": [],
            }

            past_front_matter = False

            for elem in list(body):
                if elem.tag == TAG_P:
                    p_text = _extract_p_text(elem)
                    if not p_text or not p_text.strip():
                        continue

                    p_style = _get_paragraph_style(elem)

                    # Skip Table of Contents and front matter before Scope/Foreword
                    if not past_front_matter:
                        p_lower = p_text.strip().lower()
                        if p_lower.startswith("1 scope") or p_lower == "scope" or p_lower.startswith("foreword"):
                            past_front_matter = True
                        elif _is_toc_paragraph(elem, p_text, p_style):
                            continue

                    if _is_toc_paragraph(elem, p_text, p_style):
                        continue

                    is_heading_style = any(
                        h in p_style for h in ["heading", "h1", "h2", "h3", "h4", "h5", "h6", "h7", "h8", "h9", "annex", "ti"]
                    )

                    match_subclause = RE_SUBCLAUSE_HEADER.match(p_text.strip())
                    match_toplevel = RE_TOPLEVEL_CLAUSE.match(p_text.strip())
                    match_annex = RE_ANNEX_HEADER.match(p_text.strip())

                    is_new_clause = False
                    c_num = ""
                    c_title = ""

                    if match_annex:
                        c_num = match_annex.group(1).strip()
                        c_title = match_annex.group(2).strip() or "Annex"
                        c_title = RE_TOC_PAGE_NUMBER.sub("", c_title).strip()
                        if _is_valid_clause_title(c_title, is_heading_style=True):
                            is_new_clause = True

                    elif match_subclause:
                        c_num = match_subclause.group(1).strip()
                        title_cand = match_subclause.group(2).strip()
                        title_cand = RE_TOC_PAGE_NUMBER.sub("", title_cand).strip()
                        if _is_valid_clause_title(title_cand, is_heading_style):
                            is_new_clause = True
                            c_title = title_cand

                    elif match_toplevel and (is_heading_style or p_style in ["title", "h1"]):
                        c_num = match_toplevel.group(1).strip()
                        title_cand = match_toplevel.group(2).strip()
                        title_cand = RE_TOC_PAGE_NUMBER.sub("", title_cand).strip()
                        if _is_valid_clause_title(title_cand, is_heading_style=True):
                            is_new_clause = True
                            c_title = title_cand

                    if is_new_clause:
                        self._finalize_clause(current_clause, all_clauses)
                        current_clause = {
                            "clause_number": c_num,
                            "clause_title": c_title,
                            "text_fragments": [f"{c_num} {c_title}\n"],
                        }
                    else:
                        current_clause["text_fragments"].append(p_text)

                elif elem.tag == TAG_TBL:
                    tbl_lines = self._extract_table_text(elem)
                    if tbl_lines:
                        current_clause["text_fragments"].append("\n" + "\n".join(tbl_lines) + "\n")

            self._finalize_clause(current_clause, all_clauses)

        return all_clauses

    def _finalize_clause(self, current_clause: Dict[str, Any], clauses_list: List[Dict[str, Any]]):
        content = "\n".join(current_clause["text_fragments"]).strip()
        # Ensure we do not save empty stubs or zero-length fragments
        if content and len(content) > 10:
            clauses_list.append({
                "clause_number": current_clause["clause_number"],
                "clause_title": current_clause["clause_title"],
                "content": content,
            })

    def _extract_table_text(self, tbl_elem: ET.Element) -> List[str]:
        rows_text = []
        for tr in tbl_elem.findall(TAG_TR):
            cells = [_extract_tc_text(tc) for tc in tr.findall(TAG_TC)]
            row_str = " | ".join(c for c in cells if c.strip())
            if row_str.strip():
                rows_text.append(f"| {row_str} |")
        return rows_text