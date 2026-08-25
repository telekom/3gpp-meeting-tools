"""
Extracts plain-text clauses from 3GPP .docx specification documents directly via lxml.
Splits documents by Heading/Clause numbers and parses paragraph & table text.
"""

import logging
from pathlib import Path
import re
from typing import Any, Callable, Dict, List, Optional, Tuple

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

# Matches standard 3GPP numbered headings (e.g., '4.1', '5.2.3.1', 'Annex A (normative): Title')
RE_CLAUSE_HEADER = re.compile(
    r"^(?:(?:clause|subclause)\s+)?((?:[1-9]\d*|Annex\s+[A-Z])(?:\.\d+)*)\s*[:\.\t\s]+(.+)$",
    re.IGNORECASE,
)
RE_ANNEX_HEADER = re.compile(
    r"^(Annex\s+[A-Z])(?:\s*\((?:normative|informative)\))?[:\s]*(.*)$",
    re.IGNORECASE,
)


class SpecDocxExtractor:
    """Parses .docx specifications into structured clause records for indexing."""

    def __init__(self, docx_paths: List[Path]):
        self.docx_paths = [Path(p) for p in docx_paths if Path(p).exists()]
        self.logger = logging.getLogger(__name__)

    def extract_clauses(
        self, progress_callback: Optional[Callable[[str, int], None]] = None
    ) -> List[Dict[str, Any]]:
        """Walks through the document body and splits contents by clause number."""
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

            for elem in list(body):
                if elem.tag == TAG_P:
                    p_text = _extract_p_text(elem)
                    if not p_text or not p_text.strip():
                        continue

                    # Check for Heading/Clause boundary
                    match_clause = RE_CLAUSE_HEADER.match(p_text)
                    match_annex = RE_ANNEX_HEADER.match(p_text)

                    if match_clause:
                        c_num = match_clause.group(1).strip()
                        c_title = match_clause.group(2).strip()
                        self._finalize_clause(current_clause, all_clauses)
                        current_clause = {
                            "clause_number": c_num,
                            "clause_title": c_title,
                            "text_fragments": [f"{c_num} {c_title}\n"],
                        }
                    elif match_annex:
                        c_num = match_annex.group(1).strip()
                        c_title = match_annex.group(2).strip() or "Annex"
                        self._finalize_clause(current_clause, all_clauses)
                        current_clause = {
                            "clause_number": c_num,
                            "clause_title": c_title,
                            "text_fragments": [f"{p_text}\n"],
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
        if content:
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