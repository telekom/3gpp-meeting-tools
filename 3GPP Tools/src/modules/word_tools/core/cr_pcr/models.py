"""Metadata models, OpenXML namespaces, and contiguous clause block grouping."""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Dict, List, Optional, Sequence, Tuple

from modules.word_tools.core.word_excerpt_extractor import HeadingInfo

# OpenXML Namespaces
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
C_PROP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
VT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"


@dataclass
class ContributionMetadata:
    """Metadata container for formal CRs and pseudo-CRs."""
    mode: str = "pcr"  # 'pcr' or 'cr'

    # Meeting Context
    meeting_wg: str = ""          # e.g. "SA2"
    meeting_num: str = ""         # e.g. "176"
    meeting_title: str = ""       # e.g. "" or "Prague"
    meeting_dates: str = ""       # e.g. "Oct 12 – 16, 2026"
    meeting_location: str = ""    # e.g. "Prague, Czech Republic"
    tdoc_id: str = ""             # e.g. "S2-2609420"

    # Contribution Details
    title: str = ""
    source_company: str = ""
    doc_for: str = "Approval"     # Approval / Discussion and Decision / Information
    agenda_item: str = ""
    work_item: str = ""
    abstract: str = ""

    # Narrative for pCR
    discussion_text: str = ""
    proposal_intro: str = ""

    # Target Specification
    spec_type: str = "TS"         # TS / TR
    spec_num: str = ""            # e.g. "23.801-04" or "24.501"
    spec_version: str = ""        # e.g. "0.2.0" or "18.4.0"
    spec_release: str = ""        # e.g. "Rel-19"

    # Formal CR Specific Fields
    cr_num: str = "<CR#>"         # e.g. "0452" or "<CR#>"
    cr_rev: str = "-"             # e.g. "-" or "1"
    category: str = "F"           # F / A / B / C / D
    reason_for_change: str = ""
    summary_of_change: str = ""
    consequences: str = ""
    clauses_affected: str = ""

    # Affects checkboxes for CR
    affects_uicc: bool = False
    affects_me: bool = False
    affects_ran: bool = False
    affects_cn: bool = False

    # Other specs affected
    other_core_specs: str = ""
    other_test_specs: str = ""
    other_om_specs: str = ""
    other_comments: str = ""


@dataclass
class ChangeBlock:
    """Represents one contiguous modification block in the contribution."""
    index: int
    source_path: str
    ranges: List[Tuple[int, int]]  # half-open (start_block, end_block)
    headings: List[HeadingInfo]
    label: str                     # e.g. "First Change" or "Next Change"
    is_all_new: bool = False       # If True, prints "Next Change (all new)" in pCR


def group_selected_headings_into_blocks(
    headings: Sequence[HeadingInfo],
    selected_ids: Sequence[str],
    all_new_map: Optional[Dict[str, bool]] = None,
) -> List[ChangeBlock]:
    """
    Groups selected headings into contiguous ChangeBlocks.
    Consecutive or adjacent headings within the same source file collapse into
    a single change section. Only gaps between unselected headings trigger
    a 'Next Change' transition.
    """
    all_new_map = all_new_map or {}
    h_by_id = {h.id: h for h in headings}
    chosen = [h_by_id[x] for x in selected_ids if x in h_by_id]
    chosen.sort(key=lambda h: (h.source_index, h.block_index))

    # Strip subheadings whose parents are already included
    normalized: List[HeadingInfo] = []
    for h in chosen:
        if any(
            p.source_index == h.source_index and p.block_index <= h.block_index < p.end_index
            for p in normalized
        ):
            continue
        normalized.append(h)

    if not normalized:
        return []

    # Group adjacent ranges within each source file
    blocks: List[ChangeBlock] = []
    current_source = normalized[0].source_path
    current_ranges: List[Tuple[int, int]] = [(normalized[0].block_index, normalized[0].end_index)]
    current_headings: List[HeadingInfo] = [normalized[0]]

    for h in normalized[1:]:
        last_range = current_ranges[-1]
        if h.source_path == current_source and h.block_index <= last_range[1]:
            # Contiguous or overlapping: extend current block
            current_ranges[-1] = (last_range[0], max(last_range[1], h.end_index))
            current_headings.append(h)
        else:
            # Discontinuity or new source file: finalize previous block
            idx = len(blocks)
            lbl = "First change" if idx == 0 else "Next change"
            is_new = any(all_new_map.get(ch.id, False) for ch in current_headings)
            blocks.append(
                ChangeBlock(
                    index=idx,
                    source_path=current_source,
                    ranges=current_ranges,
                    headings=current_headings,
                    label=lbl,
                    is_all_new=is_new,
                )
            )
            current_source = h.source_path
            current_ranges = [(h.block_index, h.end_index)]
            current_headings = [h]

    # Append trailing block
    idx = len(blocks)
    lbl = "First change" if idx == 0 else "Next change"
    is_new = any(all_new_map.get(ch.id, False) for ch in current_headings)
    blocks.append(
        ChangeBlock(
            index=idx,
            source_path=current_source,
            ranges=current_ranges,
            headings=current_headings,
            label=lbl,
            is_all_new=is_new,
        )
    )

    return blocks


def extract_clause_summary(blocks: Sequence[ChangeBlock]) -> str:
    """Extracts a concise, comma-separated list of affected clause numbers."""
    clause_numbers = []
    for block in blocks:
        for h in block.headings:
            # Extract leading clause number, e.g. "5.2.1", "6.x", "2"
            m = re.match(r"^(\d+(?:\.[0-9a-zA-Z]+)*)", h.text.strip())
            if m:
                c_num = m.group(1)
                if c_num not in clause_numbers:
                    clause_numbers.append(c_num)
            else:
                # Fallback to heading text snippet if no number prefix
                short = h.text.strip()[:30]
                if short not in clause_numbers:
                    clause_numbers.append(short)
    return ", ".join(clause_numbers) if clause_numbers else "All selected clauses"
