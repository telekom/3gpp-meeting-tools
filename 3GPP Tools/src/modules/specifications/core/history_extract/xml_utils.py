"""OpenXML XML parsing helpers, Change History header detection, and CR cover-sheet extraction."""

import io
import logging
import re
import zipfile
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union

from lxml import etree

from core.utils.paths import to_extended_path

logger = logging.getLogger(__name__)

# OpenXML Namespaces
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

# Audited 3GPP TDoc token pattern:
# Captures 5-to-7 digit WG and Plenary TDocs, tolerates missing spaces in rapporteur typos
# (e.g. 'R3-261549R3-261553'), and captures revision suffixes (e.g. 'S2-2606265r01').
RE_TDOC_TOKEN = re.compile(
    r"(?<![A-Za-z-])([A-Z][A-Z0-9]{1,4}-\d{5,7}(?:[rR]\d{1,2}(?!-))?)(?![0-9])"
)

# Standard 3GPP 3-part version pattern (e.g. 16.20.0, 0.8.0, 7.1.0)
RE_VERSION = re.compile(r"\b(\d{1,2}\.\d{1,2}\.\d{1,2})\b")


def normalize_version_token(ver_str: str) -> str:
    """Extracts and normalizes clean 'X.Y.Z' version strings."""
    if not ver_str:
        return ""
    m = RE_VERSION.search(ver_str.strip())
    if m:
        return m.group(1)
    clean = ver_str.lstrip("vV").strip()
    return clean if clean != "-" else ""


def parse_cell_text(tc_elem: etree._Element) -> str:
    """
    Extracts text from a table cell (<w:tc>), preserving line breaks between
    paragraphs (<w:p>) to prevent concatenated tokens (e.g. 'S2-2609250.Documented').
    """
    paragraphs = []
    for p in tc_elem.xpath("./w:p", namespaces=NS):
        p_text = "".join(p.xpath(".//w:t/text()", namespaces=NS)).strip()
        if p_text:
            paragraphs.append(p_text)
    return "\n".join(paragraphs).strip()


def expand_row_cells(tr_elem: etree._Element) -> List[Tuple[str, int, bool]]:
    """
    Returns list of (cell_text, grid_span, is_vmerge_continue) for every cell in a row,
    expanding horizontal gridSpan so subsequent columns remain aligned.
    """
    cells = []
    for tc in tr_elem.xpath("./w:tc", namespaces=NS):
        text = parse_cell_text(tc)
        span = 1
        span_elems = tc.xpath(".//w:tcPr/w:gridSpan/@w:val", namespaces=NS)
        if span_elems and span_elems[0].isdigit():
            span = int(span_elems[0])

        vmerge_elems = tc.xpath(".//w:tcPr/w:vMerge", namespaces=NS)
        is_continue = False
        if vmerge_elems:
            val = vmerge_elems[0].get(f"{{{NS['w']}}}val", "")
            if val != "restart":
                is_continue = True

        cells.append((text, span, is_continue))
    return cells


def detect_header_columns(cell_texts: List[str]) -> Optional[Dict[str, int]]:
    """
    Determines if a row is a Change History table header and maps column roles to indices.
    Supports modern 8-column format as well as legacy/pre-Rel-8 7-to-8 column variants.
    """
    lower_cells = [c.lower().replace("\xa0", " ").strip() for c in cell_texts]
    if len(lower_cells) < 4:
        return None

    mapping: Dict[str, int] = {}
    for idx, c in enumerate(lower_cells):
        if any(k in c for k in ["date"]):
            mapping.setdefault("date", idx)
        elif any(k in c for k in ["meeting", "tsg #", "tsg-number", "tsgn-number", "tsg"]):
            mapping.setdefault("meeting", idx)
        elif any(k in c for k in ["tdoc", "tsg doc", "wg number", "doc"]):
            mapping.setdefault("tdoc", idx)
        elif c in ("cr", "cr #", "cr number"):
            mapping.setdefault("cr", idx)
        elif any(k in c for k in ["rev", "revision"]):
            mapping.setdefault("rev", idx)
        elif any(k in c for k in ["cat", "category"]):
            mapping.setdefault("cat", idx)
        elif any(k in c for k in ["subject", "comment", "title", "description"]):
            mapping.setdefault("subject", idx)
        elif any(k in c for k in ["new version", "new", "version"]):
            mapping.setdefault("version", idx)
        elif any(k in c for k in ["old version", "old"]):
            mapping.setdefault("old_version", idx)

    has_version = ("version" in mapping) or ("old_version" in mapping)
    has_date_and_cr = ("date" in mapping) and (("cr" in mapping) or ("cat" in mapping))
    has_tdoc_or_cr = ("tdoc" in mapping) or ("cr" in mapping)
    has_subject = "subject" in mapping
    has_change_hist_title = any("change history" in c for c in lower_cells)

    if (has_version or has_date_and_cr or has_change_hist_title) and has_tdoc_or_cr and has_subject:
        return mapping

    return None


def extract_cr_cover_sheet_metadata(docx_or_zip_path: Union[str, Path]) -> Dict[str, Any]:
    """
    Extracts metadata from a 3GPP Change Request (.docx or .zip containing .docx) cover sheet.
    Returns an empty dict if the file is invalid or lacks a 3GPP CR cover sheet.
    """
    path = Path(docx_or_zip_path)
    if not path.exists():
        return {}

    tree = None
    try:
        if path.suffix.lower() == ".zip":
            with zipfile.ZipFile(to_extended_path(path), "r") as zf:
                if "word/document.xml" in zf.namelist():
                    with zf.open("word/document.xml") as f:
                        tree = etree.parse(f)
                else:
                    inner_docs = [
                        n for n in zf.namelist()
                        if n.lower().endswith(".docx") and not n.startswith("~$") and not n.startswith("__MACOSX")
                    ]
                    if inner_docs:
                        inner_data = zf.read(inner_docs[0])
                        with zipfile.ZipFile(io.BytesIO(inner_data), "r") as inner_zf:
                            if "word/document.xml" in inner_zf.namelist():
                                with inner_zf.open("word/document.xml") as f:
                                    tree = etree.parse(f)
        elif path.suffix.lower() in (".docx", ".docm"):
            with zipfile.ZipFile(to_extended_path(path), "r") as zf:
                if "word/document.xml" in zf.namelist():
                    with zf.open("word/document.xml") as f:
                        tree = etree.parse(f)
    except Exception as e:
        logger.debug(f"Failed to inspect cover sheet in {path.name}: {e}")
        return {}

    if tree is None:
        return {}

    tables = tree.xpath("//w:tbl", namespaces=NS)
    if not tables:
        return {}

    result: Dict[str, Any] = {
        "affected_clauses": [],
        "affected_clauses_text": "",
        "category": "",
        "cr_number": "",
        "spec_number": "",
        "work_item": "",
        "title": "",
    }

    for tbl in tables[:2]:
        rows = tbl.xpath("./w:tr", namespaces=NS)
        for r_idx, row in enumerate(rows):
            cell_nodes = row.xpath("./w:tc", namespaces=NS)
            cell_texts = [parse_cell_text(tc) for tc in cell_nodes]

            for c_idx, text in enumerate(cell_texts):
                clean_t = text.strip()
                lower_t = clean_t.lower()

                if re.search(r"\bclauses?\s+affected\b", lower_t):
                    val = ""
                    if c_idx + 1 < len(cell_texts) and cell_texts[c_idx + 1].strip():
                        val = cell_texts[c_idx + 1].strip()
                    elif r_idx + 1 < len(rows):
                        next_row_cells = [
                            parse_cell_text(tc) for tc in rows[r_idx + 1].xpath("./w:tc", namespaces=NS)
                        ]
                        if c_idx < len(next_row_cells) and next_row_cells[c_idx].strip():
                            val = next_row_cells[c_idx].strip()

                    if val:
                        result["affected_clauses_text"] = val
                        tokens = re.findall(r"\b\d+(?:\.\d+)+\b", val)
                        if not tokens:
                            tokens = [t for t in re.findall(r"\b\d+\b", val) if len(t) <= 2]

                        unique_clauses = []
                        for tok in tokens:
                            if tok not in unique_clauses:
                                unique_clauses.append(tok)
                        result["affected_clauses"] = unique_clauses

                elif re.search(r"^\s*category\s*:?$", lower_t) or re.search(r"\bcat\b", lower_t):
                    if c_idx + 1 < len(cell_texts):
                        cat_cand = cell_texts[c_idx + 1].strip().upper()
                        m_cat = re.search(r"\b([ABCDF])\b", cat_cand)
                        if m_cat:
                            result["category"] = m_cat.group(1)

                elif "work item" in lower_t:
                    if c_idx + 1 < len(cell_texts):
                        result["work_item"] = cell_texts[c_idx + 1].strip()

                elif re.search(r"\bcr\s*#|\bcr\s+num", lower_t):
                    if c_idx + 1 < len(cell_texts):
                        cr_cand = cell_texts[c_idx + 1].strip()
                        m_cr = re.search(r"\b(\d{3,5})\b", cr_cand)
                        if m_cr:
                            result["cr_number"] = m_cr.group(1)

                elif "spec" in lower_t and ("ts" in lower_t or "3gpp" in lower_t or len(lower_t) < 15):
                    if c_idx + 1 < len(cell_texts):
                        spec_cand = cell_texts[c_idx + 1].strip()
                        m_sp = re.search(r"\b(\d{2}\.\d{3}(?:-\d{2})?)\b", spec_cand)
                        if m_sp:
                            result["spec_number"] = m_sp.group(1)

        if result["affected_clauses"] or result["category"]:
            break

    return result
