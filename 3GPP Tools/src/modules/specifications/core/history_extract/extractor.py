"""Specification document resolution and OpenXML Change History table extractor."""

import logging
import re
import zipfile
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union

from lxml import etree

from core.utils.paths import get_specs_root, to_extended_path
from modules.spec_search.core.spec_search_threads import is_change_mark_file
from modules.word_tools.core.word_converter import convert_doc_to_docx
from modules.word_tools.core.word_excerpt_extractor import order_sources

from .xml_utils import (
    NS,
    RE_TDOC_TOKEN,
    detect_header_columns,
    expand_row_cells,
    extract_cr_cover_sheet_metadata,
    normalize_version_token,
)

logger = logging.getLogger(__name__)


class SpecHistoryExtractor:
    """Extracts and normalizes change history records from specification Word documents."""

    extract_cr_cover_sheet_metadata = staticmethod(extract_cr_cover_sheet_metadata)

    @staticmethod
    def resolve_spec_document(
        spec_number: str,
        version: Optional[str] = None,
        specs_root: Optional[Path] = None,
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

        target_docs = all_docs
        if version:
            clean_ver = version.lstrip("vV").strip()
            ver_docs = [p for p in all_docs if clean_ver in p.stem]
            if ver_docs:
                target_docs = ver_docs

        annex_docs = [p for p in target_docs if re.search(r"(?:annex|history)", p.stem, re.IGNORECASE)]
        if annex_docs:
            docx_annex = [p for p in annex_docs if p.suffix.lower() == ".docx"]
            if docx_annex:
                return docx_annex[-1]
            return annex_docs[-1]

        ordered = order_sources([str(p) for p in target_docs])
        if ordered:
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
        version_hint: Optional[str] = None,
    ) -> List[Dict[str, Any]]:
        """
        Extracts all Change History rows from a specification .docx file using lxml.
        Handles carry-forward fields, p-CR extraction, and rapporteur notes.
        """
        path = Path(docx_path)
        if not path.exists():
            logger.warning(f"File does not exist for history extraction: {docx_path}")
            return []

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

        history_tables: List[Tuple[int, etree._Element, Dict[str, int]]] = []
        for idx in range(len(tables) - 1, -1, -1):
            t = tables[idx]
            rows = t.xpath("./w:tr", namespaces=NS)
            if not rows:
                continue

            mapping = None
            header_row_idx = 0
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
            elif history_tables:
                break

        if not history_tables:
            logger.info(f"No Change History tables detected in {path.name}")
            return []

        history_tables.reverse()

        records: List[Dict[str, Any]] = []
        row_order = 0
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

                row_texts = [c[0] for c in expanded]
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

                if not any([raw_date, raw_meeting, source_tdoc, cr_number, subject, new_ver]):
                    continue

                if re.match(r"^[A-Z]{1,4}#\d+$", raw_date, re.IGNORECASE) and not raw_meeting:
                    raw_meeting = raw_date
                    raw_date = ""

                if raw_date:
                    last_date = raw_date
                elif last_date and (raw_meeting == last_meeting or not raw_meeting):
                    raw_date = last_date

                if raw_meeting:
                    if raw_meeting.startswith("#") and clean_spec:
                        tsg_prefix = "SA"
                        if clean_spec.startswith(("36.", "37.", "38.")):
                            tsg_prefix = "RAN"
                        elif clean_spec.startswith(("24.", "29.", "31.")):
                            tsg_prefix = "CT"
                        raw_meeting = f"{tsg_prefix}{raw_meeting}"
                    last_meeting = raw_meeting
                else:
                    raw_meeting = last_meeting

                if new_ver:
                    last_version = new_ver
                elif last_version:
                    new_ver = last_version
                elif version_hint:
                    new_ver = normalize_version_token(version_hint)

                pcr_tdocs: List[str] = []
                if subject:
                    found_tdocs = RE_TDOC_TOKEN.findall(subject)
                    for t in found_tdocs:
                        clean_t = t.strip()
                        if clean_t and clean_t != source_tdoc and clean_t not in pcr_tdocs:
                            pcr_tdocs.append(clean_t)

                row_order += 1
                records.append({
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
                })

        logger.info(f"Extracted {len(records)} Change History rows from {path.name} for {clean_spec}")
        return records
