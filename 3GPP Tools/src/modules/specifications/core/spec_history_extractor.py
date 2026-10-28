# --- File: src/modules/specifications/core/spec_history_extractor.py ---
"""
Headless OpenXML Change History Extractor for 3GPP Specifications.

Parses Change History tables from .docx specification files using high-speed
lxml streaming without Word COM overhead. Extracts formal TSG CRs, pre-release
p-CRs, rapporteur adjustment notes, and integrated TDoc references.
"""

import logging
import os
import re
import zipfile
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Sequence, Set, Tuple, Union

from lxml import etree
from PyQt5.QtCore import QThread, pyqtSignal

from core.utils.paths import get_specs_root, to_extended_path
from modules.spec_search.core.spec_search_threads import is_change_mark_file
from modules.specifications.core.spec_history_db import SpecHistoryDatabase
from modules.word_tools.core.word_converter import convert_doc_to_docx
from modules.word_tools.core.word_excerpt_extractor import order_sources

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
    # Strip any leading 'v'
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

        # Vertical merge check
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

    # Change History tables must track versions or formal changes.
    # Reject solution/key-issue tables that merely list Meeting/TDoc/Subject.
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
    Inspects cover sheet tables for:
      - 'affected_clauses': List[str] of extracted clause numbers, e.g. ["4.2.3", "5.15.11"]
      - 'affected_clauses_text': str raw text from the 'Clauses affected' cell
      - 'category': str (B, C, F, A, D)
      - 'cr_number': str
      - 'spec_number': str
      - 'work_item': str
      - 'title': str
    Returns an empty dict if the file is invalid or lacks a 3GPP CR cover sheet.
    """
    path = Path(docx_or_zip_path)
    if not path.exists():
        return {}

    tree = None
    try:
        if path.suffix.lower() == ".zip":
            with zipfile.ZipFile(to_extended_path(path), "r") as zf:
                # Check if this zip is itself an OpenXML docx (contains word/document.xml)
                if "word/document.xml" in zf.namelist():
                    with zf.open("word/document.xml") as f:
                        tree = etree.parse(f)
                else:
                    # Look for inner .docx inside the zip
                    inner_docs = [
                        n for n in zf.namelist()
                        if n.lower().endswith(".docx") and not n.startswith("~$") and not n.startswith("__MACOSX")
                    ]
                    if inner_docs:
                        inner_data = zf.read(inner_docs[0])
                        import io
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

    # Inspect the first 2 tables for the standard 3GPP Change Request cover sheet
    for tbl in tables[:2]:
        rows = tbl.xpath("./w:tr", namespaces=NS)
        for r_idx, row in enumerate(rows):
            cell_nodes = row.xpath("./w:tc", namespaces=NS)
            cell_texts = [parse_cell_text(tc) for tc in cell_nodes]

            for c_idx, text in enumerate(cell_texts):
                clean_t = text.strip()
                lower_t = clean_t.lower()

                # 1. Clauses affected
                if re.search(r"\bclauses?\s+affected\b", lower_t):
                    val = ""
                    if c_idx + 1 < len(cell_texts) and cell_texts[c_idx + 1].strip():
                        val = cell_texts[c_idx + 1].strip()
                    elif r_idx + 1 < len(rows):
                        next_row_cells = [parse_cell_text(tc) for tc in rows[r_idx + 1].xpath("./w:tc", namespaces=NS)]
                        if c_idx < len(next_row_cells) and next_row_cells[c_idx].strip():
                            val = next_row_cells[c_idx].strip()

                    if val:
                        result["affected_clauses_text"] = val
                        # Extract clause tokens: numbers with dots, e.g. "4.2.3", "5.15.11"
                        tokens = re.findall(r"\b\d+(?:\.\d+)+\b", val)
                        if not tokens:
                            # Fallback to single numbers e.g. "Clause 4, 5"
                            tokens = [t for t in re.findall(r"\b\d+\b", val) if len(t) <= 2]

                        unique_clauses = []
                        for tok in tokens:
                            if tok not in unique_clauses:
                                unique_clauses.append(tok)
                        result["affected_clauses"] = unique_clauses

                # 2. Category
                elif re.search(r"^\s*category\s*:?$", lower_t) or re.search(r"\bcat\b", lower_t):
                    if c_idx + 1 < len(cell_texts):
                        cat_cand = cell_texts[c_idx + 1].strip().upper()
                        m_cat = re.search(r"\b([ABCDF])\b", cat_cand)
                        if m_cat:
                            result["category"] = m_cat.group(1)

                # 3. Work item
                elif "work item" in lower_t:
                    if c_idx + 1 < len(cell_texts):
                        result["work_item"] = cell_texts[c_idx + 1].strip()

                # 4. CR number
                elif re.search(r"\bcr\s*#|\bcr\s+num", lower_t):
                    if c_idx + 1 < len(cell_texts):
                        cr_cand = cell_texts[c_idx + 1].strip()
                        m_cr = re.search(r"\b(\d{3,5})\b", cr_cand)
                        if m_cr:
                            result["cr_number"] = m_cr.group(1)

                # 5. Spec number
                elif "spec" in lower_t and ("ts" in lower_t or "3gpp" in lower_t or len(lower_t) < 15):
                    if c_idx + 1 < len(cell_texts):
                        spec_cand = cell_texts[c_idx + 1].strip()
                        m_sp = re.search(r"\b(\d{2}\.\d{3}(?:-\d{2})?)\b", spec_cand)
                        if m_sp:
                            result["spec_number"] = m_sp.group(1)

        if result["affected_clauses"] or result["category"]:
            break

    return result


class SpecHistoryExtractor:
    """Extracts and normalizes change history records from specification Word documents."""

    extract_cr_cover_sheet_metadata = staticmethod(extract_cr_cover_sheet_metadata)

    @staticmethod
    def resolve_spec_document(
        spec_number: str,
        version: Optional[str] = None,
        specs_root: Optional[Path] = None
    ) -> Optional[Path]:
        """
        Locates the appropriate local .docx file for a specification and version.
        For multi-part specifications, prioritizes terminal parts (_sAnnexes.docx, *annex*, etc.).
        """
        root = Path(specs_root or get_specs_root())
        clean_spec = spec_number.strip()
        spec_dir = root / clean_spec
        if not spec_dir.exists() and "-" in clean_spec:
            spec_dir = root / clean_spec.split("-")[0]
        if not spec_dir.exists():
            return None

        all_docs = [
            p for p in spec_dir.glob("*.doc*")
            if not is_change_mark_file(p) and not p.name.startswith("~$")
        ]
        if not all_docs:
            return None

        # Filter by version if specified
        target_docs = all_docs
        if version:
            clean_ver = version.lstrip("vV").strip()
            # Try to match version in filename or letter code
            ver_docs = [p for p in all_docs if clean_ver in p.stem]
            if ver_docs:
                target_docs = ver_docs

        # Check for multi-part documents: prioritize annex/history files
        annex_docs = [p for p in target_docs if re.search(r"(?:annex|history)", p.stem, re.IGNORECASE)]
        if annex_docs:
            # Pick highest numbered or latest annex docx
            docx_annex = [p for p in annex_docs if p.suffix.lower() == ".docx"]
            if docx_annex:
                return docx_annex[-1]
            return annex_docs[-1]

        # Order sources naturally and pick terminal document
        ordered = order_sources([str(p) for p in target_docs])
        if ordered:
            # Check terminal document first
            for p_str in reversed(ordered):
                candidate = Path(p_str)
                if candidate.suffix.lower() == ".docx":
                    return candidate
            return Path(ordered[-1])

        return None

    @staticmethod
    def extract_from_docx(
        docx_path: Union[str, Path],
        spec_number: str,
        version_hint: Optional[str] = None
    ) -> List[Dict[str, Any]]:
        """
        Extracts all Change History rows from a specification .docx file using lxml.
        Handles carry-forward fields, p-CR extraction, and rapporteur notes.
        """
        path = Path(docx_path)
        if not path.exists():
            logger.warning(f"File does not exist for history extraction: {docx_path}")
            return []

        # If binary .doc, attempt conversion or sibling lookup
        if path.suffix.lower() == ".doc":
            sibling_docx = path.with_suffix(".docx")
            if sibling_docx.exists():
                path = sibling_docx
            else:
                try:
                    converted = convert_doc_to_docx(str(path))
                    if converted and Path(converted).exists():
                        path = Path(converted)
                    else:
                        logger.warning(f"Cannot parse legacy binary .doc without conversion: {docx_path}")
                        return []
                except Exception as conv_err:
                    logger.warning(f"Could not convert .doc for history: {conv_err}")
                    return []

        try:
            with zipfile.ZipFile(to_extended_path(path), "r") as zf:
                if "word/document.xml" not in zf.namelist():
                    return []
                with zf.open("word/document.xml") as f:
                    tree = etree.parse(f)
        except Exception as e:
            logger.error(f"Failed to read OpenXML document for {path.name}: {e}", exc_info=True)
            return []

        tables = tree.xpath("//w:tbl", namespaces=NS)
        if not tables:
            return []

        # Collect candidate Change History tables from the end of the document
        history_tables: List[Tuple[int, etree._Element, Dict[str, int]]] = []
        for idx in range(len(tables) - 1, -1, -1):
            t = tables[idx]
            rows = t.xpath("./w:tr", namespaces=NS)
            if not rows:
                continue

            mapping = None
            header_row_idx = 0
            # Inspect first 3 rows for table headers
            for r_idx in range(min(3, len(rows))):
                expanded = expand_row_cells(rows[r_idx])
                texts = [c[0] for c in expanded]
                detected = detect_header_columns(texts)
                if detected:
                    mapping = detected
                    header_row_idx = r_idx + 1
                    break

            if mapping:
                history_tables.append((header_row_idx, t, mapping))
                # If preceding table is also part of Change History annex, continue scanning
                # (e.g. TS 24.007 has 3 separate tables for GSM, CN, and CT)
            elif history_tables:
                # Stop if we hit a table that is clearly not part of the history chain
                break

        if not history_tables:
            logger.info(f"No Change History tables detected in {path.name}")
            return []

        # Reverse to process chronological tables in ascending order
        history_tables.reverse()

        records: List[Dict[str, Any]] = []
        row_order = 0

        # Carry-forward state
        last_date = ""
        last_meeting = ""
        last_version = ""

        clean_spec = spec_number.strip()

        for header_start, t_elem, col_map in history_tables:
            rows = t_elem.xpath("./w:tr", namespaces=NS)
            for r in rows[header_start:]:
                expanded = expand_row_cells(r)
                if not expanded:
                    continue

                # Flat text list
                row_texts = [c[0] for c in expanded]

                # Skip repeated header rows or title banner rows
                if len(row_texts) == 1 and "change history" in row_texts[0].lower():
                    continue
                if detect_header_columns(row_texts):
                    continue

                def get_col(role: str) -> str:
                    idx = col_map.get(role)
                    if idx is not None and idx < len(row_texts):
                        return row_texts[idx].strip()
                    return ""

                raw_date = get_col("date")
                raw_meeting = get_col("meeting")
                source_tdoc = get_col("tdoc")
                cr_number = get_col("cr")
                revision = get_col("rev")
                category = get_col("cat")
                subject = get_col("subject")
                new_ver = normalize_version_token(get_col("version"))
                old_ver = normalize_version_token(get_col("old_version"))

                # Skip completely blank rows
                if not any([raw_date, raw_meeting, source_tdoc, cr_number, subject, new_ver]):
                    continue

                # Normalization & Repair Heuristics:
                # 1. Meeting repair: if raw_date contains meeting pattern (e.g. 'CT#28', 'TP-27')
                if re.match(r"^[A-Z]{1,4}#\d+$", raw_date, re.IGNORECASE) and not raw_meeting:
                    raw_meeting = raw_date
                    raw_date = ""

                # 2. Date carry-forward across blank cells (TS 38.331)
                if raw_date:
                    last_date = raw_date
                elif last_date and (raw_meeting == last_meeting or not raw_meeting):
                    raw_date = last_date

                # 3. Meeting carry-forward or repair prefix (e.g. '#112' -> 'CT#112')
                if raw_meeting:
                    if raw_meeting.startswith("#") and clean_spec:
                        # Infer TSG prefix from spec series
                        tsg_prefix = "SA"
                        if clean_spec.startswith(("36.", "37.", "38.")):
                            tsg_prefix = "RAN"
                        elif clean_spec.startswith(("24.", "29.", "31.")):
                            tsg_prefix = "CT"
                        raw_meeting = f"{tsg_prefix}{raw_meeting}"
                    last_meeting = raw_meeting
                else:
                    raw_meeting = last_meeting

                # 4. Version carry-forward or document hint inference (TR 38.760-3)
                if new_ver:
                    last_version = new_ver
                elif last_version:
                    new_ver = last_version
                elif version_hint:
                    new_ver = normalize_version_token(version_hint)

                # 5. Extract embedded p-CRs from subject / notes
                pcr_tdocs: List[str] = []
                if subject:
                    found_tdocs = RE_TDOC_TOKEN.findall(subject)
                    for t in found_tdocs:
                        clean_t = t.strip()
                        if clean_t and clean_t != source_tdoc and clean_t not in pcr_tdocs:
                            pcr_tdocs.append(clean_t)

                row_order += 1
                record = {
                    "spec_number": clean_spec,
                    "version": new_ver,
                    "row_order": row_order,
                    "date": raw_date,
                    "meeting": raw_meeting,
                    "source_tdoc": source_tdoc if source_tdoc != "-" else "",
                    "cr_number": cr_number if cr_number != "-" else "",
                    "revision": revision if revision != "-" else "",
                    "category": category if category != "-" else "",
                    "subject": subject,
                    "old_version": old_ver,
                    "source_file": path.name,
                    "pcr_tdocs": pcr_tdocs,
                }
                records.append(record)

        logger.info(f"Extracted {len(records)} Change History rows from {path.name} for {clean_spec}")
        return records


class SpecHistoryExtractThread(QThread):
    """Background worker for extracting change history from a single specification document."""

    finished_success = pyqtSignal(str, int)  # spec_number, records_count
    error = pyqtSignal(str)

    def __init__(
        self,
        spec_number: str,
        db: SpecHistoryDatabase,
        version: Optional[str] = None,
        docx_path: Optional[Path] = None,
        specs_root: Optional[Path] = None
    ):
        super().__init__()
        self.spec_number = spec_number.strip()
        self.db = db
        self.version = version
        self.docx_path = docx_path
        self.specs_root = specs_root

    def run(self):
        try:
            target_file = self.docx_path
            if not target_file or not target_file.exists():
                target_file = SpecHistoryExtractor.resolve_spec_document(
                    self.spec_number, self.version, self.specs_root
                )

            if not target_file or not target_file.exists():
                self.error.emit(f"No local Word document found for {self.spec_number}")
                return

            records = SpecHistoryExtractor.extract_from_docx(
                target_file, self.spec_number, self.version
            )
            if records:
                count = self.db.upsert_history_records(records)
                self.finished_success.emit(self.spec_number, count)
            else:
                self.finished_success.emit(self.spec_number, 0)
        except Exception as e:
            logger.error(f"Error in SpecHistoryExtractThread for {self.spec_number}: {e}", exc_info=True)
            self.error.emit(str(e))


class SpecHistoryBatchScanThread(QThread):
    """Background worker that scans all locally cached specifications and extracts change history."""

    progress = pyqtSignal(int, int, str)     # current, total, message
    finished_success = pyqtSignal(int, int)  # specs_processed, total_records
    error = pyqtSignal(str)

    def __init__(self, db: SpecHistoryDatabase, specs_root: Optional[Path] = None):
        super().__init__()
        self.db = db
        self.specs_root = Path(specs_root or get_specs_root())

    def run(self):
        try:
            if not self.specs_root.exists():
                self.finished_success.emit(0, 0)
                return

            # Discover all specification folders in specs cache
            candidate_dirs = [
                d for d in self.specs_root.iterdir()
                if d.is_dir() and not d.name.startswith(".")
            ]
            total_dirs = len(candidate_dirs)
            if total_dirs == 0:
                self.finished_success.emit(0, 0)
                return

            total_records = 0
            specs_processed = 0

            for idx, spec_dir in enumerate(candidate_dirs):
                if self.isInterruptionRequested():
                    break

                spec_num = spec_dir.name
                self.progress.emit(idx + 1, total_dirs, f"Scanning {spec_num}...")

                doc_file = SpecHistoryExtractor.resolve_spec_document(spec_num, specs_root=self.specs_root)
                if doc_file and doc_file.exists():
                    try:
                        records = SpecHistoryExtractor.extract_from_docx(doc_file, spec_num)
                        if records:
                            count = self.db.upsert_history_records(records)
                            total_records += count
                            specs_processed += 1
                    except Exception as spec_err:
                        logger.warning(f"Skipping {spec_num} during batch history scan: {spec_err}")

            self.finished_success.emit(specs_processed, total_records)
        except Exception as e:
            logger.error(f"Error in SpecHistoryBatchScanThread: {e}", exc_info=True)
            self.error.emit(str(e))
