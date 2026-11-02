"""Utility for generating clean, differentiated, and filesystem-safe export filenames
for Word/Markdown clause excerpts, slicing, and specifications.
"""

import datetime
import os
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

# Supported naming schemes
SCHEME_SOURCE_CLAUSES_TIMESTAMP = "source_clauses_timestamp"
SCHEME_SOURCE_CLAUSES = "source_clauses"
SCHEME_SOURCE_CLAUSES_TITLE = "source_clauses_title"
SCHEME_SOURCE_TIMESTAMP = "source_timestamp"
SCHEME_CLAUSES_TIMESTAMP = "clauses_timestamp"
SCHEME_CLAUSES_ONLY = "clauses_only"
SCHEME_TIMESTAMP_ONLY = "timestamp_only"
SCHEME_GENERIC = "generic"

NAMING_SCHEMES = [
    (SCHEME_SOURCE_CLAUSES_TIMESTAMP, "Source + Clauses + Timestamp (Recommended)"),
    (SCHEME_SOURCE_CLAUSES, "Source + Clauses"),
    (SCHEME_SOURCE_CLAUSES_TITLE, "Source + Clauses & Title + Timestamp"),
    (SCHEME_SOURCE_TIMESTAMP, "Source + Timestamp"),
    (SCHEME_CLAUSES_TIMESTAMP, "Clauses + Timestamp"),
    (SCHEME_CLAUSES_ONLY, "Clauses Only"),
    (SCHEME_TIMESTAMP_ONLY, "Timestamp Only"),
    (SCHEME_GENERIC, "Generic (extracted_sections)"),
]

# Supported CR / pCR naming schemes
CR_SCHEME_STANDARD = "cr_standard"              # TDoc + Mode + Spec + Clauses + Timestamp (Recommended)
CR_SCHEME_NO_TIMESTAMP = "cr_no_timestamp"      # TDoc + Mode + Spec + Clauses
CR_SCHEME_WITH_TOPIC = "cr_with_topic"          # TDoc + Mode + Spec + Clauses & Topic + Timestamp
CR_SCHEME_SPEC_CLAUSES = "cr_spec_clauses"      # Mode + Spec + Clauses + Timestamp
CR_SCHEME_TDOC_SPEC = "cr_tdoc_spec"            # TDoc + Mode + Spec + Version
CR_SCHEME_GENERIC = "cr_generic"                # Generic (contribution.docx)

CR_NAMING_SCHEMES = [
    (CR_SCHEME_STANDARD, "TDoc + Mode + Spec + Clauses + Timestamp (Recommended)"),
    (CR_SCHEME_NO_TIMESTAMP, "TDoc + Mode + Spec + Clauses"),
    (CR_SCHEME_WITH_TOPIC, "TDoc + Mode + Spec + Clauses & Topic + Timestamp"),
    (CR_SCHEME_SPEC_CLAUSES, "Mode + Spec + Clauses + Timestamp"),
    (CR_SCHEME_TDOC_SPEC, "TDoc + Mode + Spec + Version"),
    (CR_SCHEME_GENERIC, "Generic (contribution.docx)"),
]


def sanitize_filename_token(text: str, max_len: int = 40) -> str:
    """Sanitizes a string to be safely used as part of a Windows/Unix filename."""
    if not text:
        return ""
    # Replace illegal filename characters: / \ : * ? " < > |
    cleaned = re.sub(r'[\/:*?"<>|]+', "_", str(text).strip())
    # Replace whitespace with underscores
    cleaned = re.sub(r"\s+", "_", cleaned)
    # Collapse multiple underscores
    cleaned = re.sub(r"_+", "_", cleaned)
    # Strip dots and underscores from edges
    cleaned = cleaned.strip("._")
    return cleaned[:max_len]


def extract_clause_identifier(heading_text: str) -> Tuple[str, str]:
    """Extracts (clause_id, clause_title) from a heading text.

    Handles standard 3GPP and Word heading formats:
    - '5.2 Key derivation for K_NF' -> ('5.2', 'Key_derivation')
    - 'Clause 6.3.1 Procedures' -> ('6.3.1', 'Procedures')
    - 'Annex A (informative): History' -> ('Annex_A', 'History')
    - 'Section 4 Architecture' -> ('4', 'Architecture')
    - 'Introduction' -> ('Introduction', '')
    """
    clean_text = heading_text.strip()
    pat = re.compile(
        r"^(?:Clause\s+|Section\s+)?((?:Annex\s+[A-Za-z](?:\.[0-9]+)*)|(?:[0-9]+(?:\.[0-9]+)*))\b(?:\s*[:.\-–—]?\s*(.*))?",
        re.IGNORECASE,
    )
    m = pat.match(clean_text)
    if m:
        clause_id = re.sub(r"\s+", "_", m.group(1))
        raw_title = (m.group(2) or "").strip()
        # Clean parentheticals from title if informative/normative
        clean_title = re.sub(r"^\((?:informative|normative)\)\s*[:.\-–—]?\s*", "", raw_title, flags=re.I)
        words = re.findall(r"[A-Za-z0-9_]+", clean_title)
        title_tag = "_".join(words[:3]) if words else ""
        return clause_id, sanitize_filename_token(title_tag, 35)

    # Non-numbered heading: sanitize first 2-3 words
    words = re.findall(r"[A-Za-z0-9_]+", clean_text)
    clause_id = "_".join(words[:2]) if words else "Section"
    title_words = words[2:4] if len(words) > 2 else []
    return sanitize_filename_token(clause_id, 30), sanitize_filename_token("_".join(title_words), 30)


def normalize_selected_headings(headings: Sequence[Any], selected_ids: Sequence[str]) -> List[Any]:
    """Filters selected headings so child headings enclosed within selected parent headings are omitted."""
    heading_map = {getattr(h, "id", str(h)): h for h in headings}
    chosen = [heading_map[x] for x in selected_ids if x in heading_map]
    if not chosen:
        return []
    # Sort by (source_index, block_index)
    chosen.sort(key=lambda h: (getattr(h, "source_index", 0), getattr(h, "block_index", 0)))
    result = []
    for h in chosen:
        src = getattr(h, "source_index", 0)
        blk = getattr(h, "block_index", 0)
        # Check if already covered by an ancestor
        if any(
            getattr(p, "source_index", 0) == src
            and getattr(p, "block_index", 0) <= blk < getattr(p, "end_index", 0)
            for p in result
        ):
            continue
        result.append(h)
    return result


def _try_format_clause_range(cids: List[str]) -> Optional[str]:
    """Detects if cids form a simple contiguous range like 5.1, 5.2, 5.3 -> '5.1-5.3'."""
    try:
        parts = [cid.split(".") for cid in cids]
        if all(len(p) == 2 and p[0] == parts[0][0] and p[1].isdigit() for p in parts):
            nums = [int(p[1]) for p in parts]
            if nums == list(range(nums[0], nums[0] + len(nums))):
                return f"{parts[0][0]}.{nums[0]}-{parts[0][0]}.{nums[-1]}"
    except Exception:
        pass
    return None


def build_clauses_token(
    headings: Sequence[Any],
    include_title: bool = False,
    max_clauses: int = 3,
) -> str:
    """Builds a concise clause token such as 'Cl_5.2', 'Cl_5.2_5.3', 'Cl_5.1-5.5', or 'Cl_5.2_plus2'."""
    if not headings:
        return "sections"

    parsed = [extract_clause_identifier(getattr(h, "text", str(h))) for h in headings]

    # De-duplicate clause IDs in order
    unique_pairs: List[Tuple[str, str]] = []
    seen = set()
    for cid, title in parsed:
        if cid and cid not in seen:
            seen.add(cid)
            unique_pairs.append((cid, title))

    if not unique_pairs:
        return "sections"

    if len(unique_pairs) == 1:
        cid, title = unique_pairs[0]
        if include_title and title:
            return f"Cl_{cid}_{title}"
        return f"Cl_{cid}"

    if len(unique_pairs) <= max_clauses:
        return "Cl_" + "_".join(cid for cid, _ in unique_pairs)

    cids = [cid for cid, _ in unique_pairs]
    range_str = _try_format_clause_range(cids)
    if range_str:
        return f"Cl_{range_str}"

    return f"Cl_{cids[0]}_{cids[1]}_plus{len(cids) - 2}"


def build_source_token(source_paths: Sequence[str]) -> str:
    """Builds a concise source document token from input paths."""
    if not source_paths:
        return "doc"
    stems = [sanitize_filename_token(Path(p).stem, 35) for p in source_paths if p]
    if not stems:
        return "doc"
    if len(stems) == 1:
        return stems[0]
    elif len(stems) == 2:
        combined = f"{stems[0]}_{stems[1]}"
        return combined if len(combined) <= 50 else f"{stems[0]}_plus1"
    else:
        return f"{stems[0]}_plus{len(stems) - 1}"


def generate_differentiated_export_name(
    source_paths: Sequence[str],
    headings: Sequence[Any],
    selected_ids: Sequence[str],
    scheme: str = SCHEME_SOURCE_CLAUSES_TIMESTAMP,
    timestamp: Optional[str] = None,
    ext: str = ".docx",
    custom_prefix: Optional[str] = None,
) -> str:
    """Generates a differentiated, highly descriptive, filesystem-safe filename for an export.

    Examples:
    - source_clauses_timestamp: 'S2-2401234_Cl_5.2_20261102_1418.docx'
    - source_clauses: 'S2-2401234_Cl_5.2.docx'
    - source_clauses_title: 'S2-2401234_Cl_5.2_Key_derivation_20261102_1418.docx'
    - source_timestamp: 'S2-2401234_excerpt_20261102_1418.docx'
    - clauses_timestamp: 'Cl_5.2_20261102_1418.docx'
    - clauses_only: 'Clause_5.2.docx'
    - timestamp_only: 'excerpt_20261102_1418.docx'
    - generic: 'extracted_sections.docx'
    """
    ext_clean = ext if ext.startswith(".") else f".{ext}"
    if scheme == SCHEME_GENERIC:
        return f"extracted_sections{ext_clean}"

    if timestamp is None:
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M")
    else:
        timestamp = sanitize_filename_token(timestamp, 20)

    # Normalize headings
    normalized_headings = normalize_selected_headings(headings, selected_ids)

    source_token = build_source_token(source_paths)
    if custom_prefix:
        clean_prefix = sanitize_filename_token(custom_prefix, 35)
        if clean_prefix:
            source_token = clean_prefix

    include_title = (scheme == SCHEME_SOURCE_CLAUSES_TITLE)
    clauses_token = build_clauses_token(normalized_headings, include_title=include_title)

    if scheme == SCHEME_SOURCE_CLAUSES_TIMESTAMP:
        base = f"{source_token}_{clauses_token}_{timestamp}"
    elif scheme == SCHEME_SOURCE_CLAUSES:
        base = f"{source_token}_{clauses_token}"
    elif scheme == SCHEME_SOURCE_CLAUSES_TITLE:
        base = f"{source_token}_{clauses_token}_{timestamp}"
    elif scheme == SCHEME_SOURCE_TIMESTAMP:
        base = f"{source_token}_excerpt_{timestamp}"
    elif scheme == SCHEME_CLAUSES_TIMESTAMP:
        base = f"{clauses_token}_{timestamp}"
    elif scheme == SCHEME_CLAUSES_ONLY:
        base = f"{clauses_token}"
    elif scheme == SCHEME_TIMESTAMP_ONLY:
        base = f"excerpt_{timestamp}"
    else:
        base = f"{source_token}_{clauses_token}_{timestamp}"

    # Clean up double underscores and trim length
    base = re.sub(r"_+", "_", base).strip("._")
    if len(base) > 100:
        base = base[:100].rstrip("._")

    return f"{base}{ext_clean}"

def generate_cr_pcr_export_name(
    mode: str,
    spec_num: str,
    spec_type: str = "TS",
    version: str = "",
    tdoc_id: Optional[str] = None,
    cr_num: Optional[str] = None,
    cr_rev: Optional[str] = None,
    work_item: Optional[str] = None,
    title: Optional[str] = None,
    headings: Sequence[Any] = (),
    selected_ids: Sequence[str] = (),
    scheme: str = CR_SCHEME_STANDARD,
    timestamp: Optional[str] = None,
    ext: str = ".docx",
) -> str:
    """Generates a clean, differentiated, 3GPP-standard filename for a CR or pCR document.

    Examples:
    - cr_standard (pCR): 'S2-2401234_pCR_TS23501_v18.2.0_Cl_5.2_20261102_1745.docx'
    - cr_standard (CR):  'S2-2401234_CR0845r1_TS23501_v18.2.0_Cl_5.2_20261102_1745.docx'
    - cr_no_timestamp:   'S2-2401234_pCR_TS23501_v18.2.0_Cl_5.2.docx'
    - cr_with_topic:     'S2-2401234_pCR_TS23501_v18.2.0_Cl_5.2_5G_ProSe_20261102_1745.docx'
    - cr_spec_clauses:   'pCR_TS23501_v18.2.0_Cl_5.2_20261102_1745.docx'
    - cr_tdoc_spec:      'S2-2401234_pCR_TS23501_v18.2.0.docx'
    - cr_generic:        'contribution.docx'
    """
    ext_clean = ext if ext.startswith(".") else f".{ext}"
    mode_clean = str(mode or "pcr").lower().strip()
    is_cr = mode_clean == "cr"

    if scheme == CR_SCHEME_GENERIC:
        return f"{mode_clean}_contribution{ext_clean}"

    if timestamp is None:
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M")
    else:
        timestamp = sanitize_filename_token(timestamp, 20)

    # 1. Mode token: CR (with CR# and rev if provided) or pCR
    if is_cr:
        clean_cr = re.sub(r"(?i)^cr\s*", "", str(cr_num or "").strip())
        if clean_cr and clean_cr not in ("<CR#>", "CR#", ""):
            clean_rev = str(cr_rev or "").strip()
            if clean_rev and clean_rev not in ("-", "", "0"):
                mode_token = f"CR{clean_cr}r{clean_rev}"
            else:
                mode_token = f"CR{clean_cr}"
        else:
            mode_token = "CR"
    else:
        mode_token = "pCR"

    # 2. TDoc ID token
    tdoc_clean = str(tdoc_id or "").strip()
    if tdoc_clean and tdoc_clean.lower() not in ("draft-tdoc", "draft", "none", "n/a", "<tdoc>"):
        tdoc_token = sanitize_filename_token(tdoc_clean, 25)
    else:
        tdoc_token = ""

    # 3. Spec & Version token
    spec_clean = str(spec_num or "").strip()
    if spec_clean:
        compact_num = spec_clean.replace(".", "")
        st = (spec_type or "TS").upper().strip()
        spec_tag = f"{st}{compact_num}"
    else:
        spec_tag = "Spec"

    ver_clean = str(version or "").lstrip("vV").strip()
    if ver_clean:
        spec_tag += f"_v{ver_clean}"

    # 4. Clauses token
    normalized_headings = normalize_selected_headings(headings, selected_ids)
    clauses_token = build_clauses_token(normalized_headings, include_title=False)
    if clauses_token == "sections":
        clauses_token = ""

    # 5. Topic token (Work Item or Title)
    topic_token = ""
    if work_item and str(work_item).strip():
        topic_token = sanitize_filename_token(str(work_item).strip(), 25)
    elif title and str(title).strip():
        words = re.findall(r"[A-Za-z0-9_]+", str(title).strip())
        topic_token = sanitize_filename_token("_".join(words[:3]), 25) if words else ""

    # 6. Assemble based on scheme
    parts: List[str] = []

    if scheme == CR_SCHEME_STANDARD:
        if tdoc_token:
            parts.append(tdoc_token)
        parts.append(mode_token)
        parts.append(spec_tag)
        if clauses_token:
            parts.append(clauses_token)
        if timestamp:
            parts.append(timestamp)

    elif scheme == CR_SCHEME_NO_TIMESTAMP:
        if tdoc_token:
            parts.append(tdoc_token)
        parts.append(mode_token)
        parts.append(spec_tag)
        if clauses_token:
            parts.append(clauses_token)

    elif scheme == CR_SCHEME_WITH_TOPIC:
        if tdoc_token:
            parts.append(tdoc_token)
        parts.append(mode_token)
        parts.append(spec_tag)
        if clauses_token:
            parts.append(clauses_token)
        if topic_token:
            parts.append(topic_token)
        if timestamp:
            parts.append(timestamp)

    elif scheme == CR_SCHEME_SPEC_CLAUSES:
        # Strictly omit TDoc ID
        parts.append(mode_token)
        parts.append(spec_tag)
        if clauses_token:
            parts.append(clauses_token)
        if timestamp:
            parts.append(timestamp)

    elif scheme == CR_SCHEME_TDOC_SPEC:
        # Strictly omit clauses
        if tdoc_token:
            parts.append(tdoc_token)
        parts.append(mode_token)
        parts.append(spec_tag)

    else:
        # Fallback to standard
        if tdoc_token:
            parts.append(tdoc_token)
        parts.append(mode_token)
        parts.append(spec_tag)
        if clauses_token:
            parts.append(clauses_token)
        if timestamp:
            parts.append(timestamp)

    base = "_".join(p for p in parts if p)
    base = re.sub(r"_+", "_", base).strip("._")
    if len(base) > 110:
        base = base[:110].rstrip("._")

    return f"{base}{ext_clean}"

