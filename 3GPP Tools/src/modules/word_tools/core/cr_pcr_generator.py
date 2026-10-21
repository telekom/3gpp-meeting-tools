# --- File: src/modules/word_tools/core/cr_pcr_generator.py ---
"""
High-Fidelity 3GPP Contribution Generator (Formal CR & Programmatic pCR).

This engine constructs 3GPP Change Requests (CR) and Pseudo-CRs (pCR) based on
extracted specification clauses. It enforces strict 3GPP-defined Word styles,
merges contiguous clause selections into single change sections, populates
complex DOCPROPERTY metadata and custom document properties, and ensures
safe Word COM assembly with enterprise sensitivity labeling.
"""

import io
import json
import logging
import os
import posixpath
import re
import shutil
import subprocess
import sys
import tempfile
import time
import zipfile
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Callable, Dict, List, Optional, Sequence, Set, Tuple

import pythoncom
import win32com.client
from lxml import etree
from PyQt5.QtCore import QThread, pyqtSignal

from core.utils.paths import get_temp_root, to_extended_path
from modules.word_tools.core.sensitivity_label import set_sensitivity_label
from modules.word_tools.core.styles_3gpp import (
    Style3GPP,
    ensure_3gpp_package_styles,
    get_3gpp_package_part,
    get_pseudo_cr_template_path,
    get_template_cr_path,
    install_3gpp_dotx_template,
    promote_3gpp_styles_in_word,
)
from modules.word_tools.core.word_excerpt_extractor import (
    HeadingInfo,
    WordExcerptEngine,
)

logger = logging.getLogger(__name__)

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


# ---------------------------------------------------------------------------
# Base Document Builders (Formal CR & Programmatic pCR)
# ---------------------------------------------------------------------------

def _update_cr_custom_properties(custom_xml_bytes: bytes, meta: ContributionMetadata) -> bytes:
    """Updates the 19 custom properties inside docProps/custom.xml."""
    parser = etree.XMLParser(remove_blank_text=False)
    root = etree.fromstring(custom_xml_bytes, parser=parser)

    # Clean date components
    start_date = meta.meeting_dates.split("–")[0].split("-")[0].strip() if meta.meeting_dates else ""
    end_date = meta.meeting_dates.split("–")[-1].split("-")[-1].strip() if meta.meeting_dates else ""
    cur_date_str = datetime.now().strftime("%Y-%m-%d")

    prop_values = {
        "TSG/WGRef": meta.meeting_wg,
        "MtgSeq": meta.meeting_num,
        "Location": meta.meeting_location,
        "Country": "",
        "StartDate": start_date or meta.meeting_dates,
        "EndDate": end_date or meta.meeting_dates,
        "Tdoc#": meta.tdoc_id,
        "Spec#": meta.spec_num,
        "Cr#": meta.cr_num,
        "Revision": meta.cr_rev,
        "Version": meta.spec_version,
        "SourceIfWg": meta.source_company,
        "SourceIfTsg": meta.source_company,
        "RelatedWis": meta.work_item,
        "Cat": meta.category,
        "ResDate": cur_date_str,
        "Release": meta.spec_release,
        "CrTitle": meta.title,
        "MtgTitle": meta.meeting_title,
    }

    for prop in root.findall(f"{{{C_PROP_NS}}}property"):
        p_name = prop.get("name", "")
        if p_name in prop_values:
            val_el = prop.find(f"{{{VT_NS}}}lpwstr")
            if val_el is not None:
                val_el.text = prop_values[p_name]

    return etree.tostring(root, encoding="utf-8", xml_declaration=True, standalone="yes")


def _update_cr_document_xml(
    doc_xml_bytes: bytes,
    meta: ContributionMetadata,
    change_blocks: Sequence[ChangeBlock],
) -> bytes:
    """
    Updates word/document.xml for formal CR:
    - Replaces cached DOCPROPERTY field runs
    - Populates Table 1 and Table 2 cells
    - Replaces placeholder change blocks with anchors for Word COM InsertFile
    """
    parser = etree.XMLParser(remove_blank_text=False, huge_tree=True)
    root = etree.fromstring(doc_xml_bytes, parser=parser)
    body = root.find(f"{{{W_NS}}}body")
    if body is None:
        raise RuntimeError("No <w:body> found in CR template.")

    # 1. Update DOCPROPERTY cached field runs
    field_replacements = {
        "TSG/WGRef": meta.meeting_wg,
        "MtgSeq": meta.meeting_num,
        "MtgTitle": meta.meeting_title,
        "Tdoc#": meta.tdoc_id,
        "Location": meta.meeting_location,
        "StartDate": meta.meeting_dates.split("–")[0].strip(),
        "EndDate": meta.meeting_dates.split("–")[-1].strip(),
        "Spec#": meta.spec_num,
        "Cr#": meta.cr_num,
        "Revision": meta.cr_rev,
        "Version": meta.spec_version,
        "CrTitle": meta.title,
        "SourceIfWg": meta.source_company,
        "SourceIfTsg": meta.source_company,
        "RelatedWis": meta.work_item,
        "Cat": meta.category,
        "Release": meta.spec_release,
    }

    for instr in root.xpath("//w:instrText", namespaces={"w": W_NS}):
        text = instr.text or ""
        m = re.search(r"DOCPROPERTY\s+([^\s\\]+)", text)
        if m:
            prop_key = m.group(1).strip()
            if prop_key in field_replacements:
                parent_p = instr.getparent()
                while parent_p is not None and parent_p.tag != f"{{{W_NS}}}p":
                    parent_p = parent_p.getparent()
                if parent_p is not None:
                    # Find runs between fldChar separate and end
                    in_result = False
                    for child in parent_p:
                        fld_char = child.find(f"{{{W_NS}}}fldChar")
                        if fld_char is not None:
                            f_type = fld_char.get(f"{{{W_NS}}}fldCharType")
                            if f_type == "separate":
                                in_result = True
                            elif f_type == "end":
                                in_result = False
                        elif in_result and child.tag == f"{{{W_NS}}}r":
                            t_el = child.find(f"{{{W_NS}}}t")
                            if t_el is not None:
                                t_el.text = field_replacements[prop_key]

    # 2. Update Table 1 Yellow Checkboxes (Proposed change affects)
    tables = body.findall(f"{{{W_NS}}}tbl")
    if len(tables) > 1:
        tbl_affects = tables[1]
        row_0 = tbl_affects.findall(f"{{{W_NS}}}tr")[0]
        cells = row_0.findall(f"{{{W_NS}}}tc")
        # Indices: C2=UICC, C4=ME, C6=RAN, C8=CN
        affects_map = {
            2: meta.affects_uicc,
            4: meta.affects_me,
            6: meta.affects_ran,
            8: meta.affects_cn,
        }
        for c_idx, is_checked in affects_map.items():
            if c_idx < len(cells):
                t_el = cells[c_idx].xpath(".//w:t", namespaces={"w": W_NS})
                if t_el:
                    t_el[0].text = "X" if is_checked else ""
                else:
                    p = cells[c_idx].find(f"{{{W_NS}}}p")
                    if p is not None and is_checked:
                        r = etree.SubElement(p, f"{{{W_NS}}}r")
                        r_pr = etree.SubElement(r, f"{{{W_NS}}}rPr")
                        etree.SubElement(r_pr, f"{{{W_NS}}}b")
                        t = etree.SubElement(r, f"{{{W_NS}}}t")
                        t.text = "X"

    # 3. Update Table 2 Metadata & Justifications
    if len(tables) > 2:
        tbl_meta = tables[2]
        rows = tbl_meta.findall(f"{{{W_NS}}}tr")

        def set_cell_text(row_idx: int, cell_idx: int, val: str):
            if row_idx < len(rows):
                r_cells = rows[row_idx].findall(f"{{{W_NS}}}tc")
                if cell_idx < len(r_cells):
                    p_list = r_cells[cell_idx].findall(f"{{{W_NS}}}p")
                    if p_list:
                        for p in p_list[1:]:
                            r_cells[cell_idx].remove(p)
                        target_p = p_list[0]
                        for c in list(target_p):
                            if c.tag == f"{{{W_NS}}}r":
                                target_p.remove(c)
                        r = etree.SubElement(target_p, f"{{{W_NS}}}r")
                        t = etree.SubElement(r, f"{{{W_NS}}}t")
                        t.text = val

        set_cell_text(11, 1, meta.reason_for_change)
        set_cell_text(13, 1, meta.summary_of_change)
        set_cell_text(15, 1, meta.consequences)
        set_cell_text(17, 1, meta.clauses_affected)
        set_cell_text(24, 1, meta.other_comments)

        # Other specs affected
        if meta.other_core_specs:
            set_cell_text(20, 4, meta.other_core_specs)
        if meta.other_test_specs:
            set_cell_text(21, 4, meta.other_test_specs)
        if meta.other_om_specs:
            set_cell_text(22, 4, meta.other_om_specs)

    # 4. Replace placeholder change blocks (body[9..14]) with dynamic anchors
    # Body elements [0..8] represent the Cover Page and section break
    body_children = list(body)
    # Remove old placeholder paragraphs between cover section break (child 8) and final sectPr
    final_sect_pr = None
    for child in body_children[9:]:
        if child.tag == f"{{{W_NS}}}sectPr":
            final_sect_pr = child
        body.remove(child)

    # Inject dynamic change separators and anchors
    for idx, block in enumerate(change_blocks):
        # Separator
        sep_p = etree.SubElement(body, f"{{{W_NS}}}p")
        sep_ppr = etree.SubElement(sep_p, f"{{{W_NS}}}pPr")
        p_style = etree.SubElement(sep_ppr, f"{{{W_NS}}}pStyle")
        p_style.set(f"{{{W_NS}}}val", Style3GPP.CR_SEPARATOR)
        sep_r = etree.SubElement(sep_p, f"{{{W_NS}}}r")
        sep_t = etree.SubElement(sep_r, f"{{{W_NS}}}t")
        sep_t.text = "==============First change==============" if idx == 0 else "==============Next change=============="

        # Insertion Anchor
        anc_p = etree.SubElement(body, f"{{{W_NS}}}p")
        anc_r = etree.SubElement(anc_p, f"{{{W_NS}}}r")
        anc_t = etree.SubElement(anc_r, f"{{{W_NS}}}t")
        anc_t.text = f"{{{{CHANGE_BLOCK_ANCHOR_{idx}}}}}"

    # Final End of change separator
    end_p = etree.SubElement(body, f"{{{W_NS}}}p")
    end_ppr = etree.SubElement(end_p, f"{{{W_NS}}}pPr")
    end_pstyle = etree.SubElement(end_ppr, f"{{{W_NS}}}pStyle")
    end_pstyle.set(f"{{{W_NS}}}val", Style3GPP.CR_SEPARATOR)
    end_r = etree.SubElement(end_p, f"{{{W_NS}}}r")
    end_t = etree.SubElement(end_r, f"{{{W_NS}}}t")
    end_t.text = "==============End of change=============="

    # Trailing empty paragraph
    etree.SubElement(body, f"{{{W_NS}}}p")

    # Restore final document section break
    if final_sect_pr is not None:
        body.append(final_sect_pr)
    else:
        etree.SubElement(body, f"{{{W_NS}}}sectPr")

    return etree.tostring(root, encoding="utf-8", xml_declaration=True, standalone="yes")


def build_cr_base_document(
    meta: ContributionMetadata,
    change_blocks: Sequence[ChangeBlock],
    target_path: Path,
) -> None:
    """Creates the base formal CR document with metadata and change anchors."""
    tpl_path = get_template_cr_path()
    with open(to_extended_path(tpl_path), "rb") as f:
        tpl_bytes = f.read()

    with zipfile.ZipFile(io.BytesIO(tpl_bytes), "r") as zin:
        custom_xml = zin.read("docProps/custom.xml")
        doc_xml = zin.read("word/document.xml")

        new_custom_xml = _update_cr_custom_properties(custom_xml, meta)
        new_doc_xml = _update_cr_document_xml(doc_xml, meta, change_blocks)

        out_buf = io.BytesIO()
        with zipfile.ZipFile(out_buf, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "docProps/custom.xml":
                    zout.writestr(item, new_custom_xml)
                elif item.filename == "word/document.xml":
                    zout.writestr(item, new_doc_xml)
                else:
                    zout.writestr(item, zin.read(item.filename))

    with open(to_extended_path(target_path), "wb") as f:
        f.write(out_buf.getvalue())
    ensure_3gpp_package_styles(target_path)


def build_pcr_base_document(
    meta: ContributionMetadata,
    change_blocks: Sequence[ChangeBlock],
    target_path: Path,
) -> None:
    """
    Programmatically builds the authentic 3GPP pseudo-CR (pCR) base document
    using 3GPP-defined styles (tdoc-header, Heading 1, Normal, CR_Separator).
    """
    tpl_path = get_pseudo_cr_template_path()
    with open(to_extended_path(tpl_path), "rb") as f:
        tpl_bytes = f.read()

    # Construct the body XML
    body_xml_parts = [
        f'<w:document xmlns:w="{W_NS}" xmlns:r="{R_NS}"><w:body>',
        # 1. Header Line: WG#Meeting \t TDoc
        f'<w:p><w:pPr><w:pStyle w:val="{Style3GPP.TDOC_HEADER}"/>'
        f'<w:tabs><w:tab w:val="right" w:pos="9639"/></w:tabs>'
        f'<w:spacing w:after="0"/><w:rPr><w:b/><w:sz w:val="24"/></w:rPr></w:pPr>'
        f'<w:r><w:rPr><w:b/><w:sz w:val="24"/></w:rPr><w:t>3GPP TSG-WG {meta.meeting_wg}#{meta.meeting_num}</w:t></w:r>'
        f'<w:r><w:rPr><w:b/><w:i/><w:sz w:val="24"/></w:rPr><w:tab/><w:t>{meta.tdoc_id}</w:t></w:r></w:p>',
        # 2. Date and Location
        f'<w:p><w:pPr><w:pStyle w:val="{Style3GPP.NORMAL}"/><w:spacing w:after="120"/></w:pPr>'
        f'<w:r><w:t>{meta.meeting_dates}, {meta.meeting_location}</w:t></w:r></w:p>',
        '<w:p><w:pPr><w:spacing w:after="120"/></w:pPr></w:p>',
        # 3. Metadata block
        f'<w:p><w:pPr><w:pStyle w:val="{Style3GPP.NORMAL}"/><w:spacing w:after="60"/></w:pPr>'
        f'<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Source:\t</w:t></w:r>'
        f'<w:r><w:t>{meta.source_company}</w:t></w:r></w:p>',
        f'<w:p><w:pPr><w:pStyle w:val="{Style3GPP.NORMAL}"/><w:spacing w:after="60"/></w:pPr>'
        f'<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Title:\t</w:t></w:r>'
        f'<w:r><w:t>{meta.title}</w:t></w:r></w:p>',
        f'<w:p><w:pPr><w:pStyle w:val="{Style3GPP.NORMAL}"/><w:spacing w:after="60"/></w:pPr>'
        f'<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Document for:\t</w:t></w:r>'
        f'<w:r><w:t>{meta.doc_for}</w:t></w:r></w:p>',
        f'<w:p><w:pPr><w:pStyle w:val="{Style3GPP.NORMAL}"/><w:spacing w:after="60"/></w:pPr>'
        f'<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Agenda item:\t</w:t></w:r>'
        f'<w:r><w:t>{meta.agenda_item}</w:t></w:r></w:p>',
        f'<w:p><w:pPr><w:pStyle w:val="{Style3GPP.NORMAL}"/><w:spacing w:after="60"/></w:pPr>'
        f'<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Work Item:\t</w:t></w:r>'
        f'<w:r><w:t>{meta.work_item}</w:t></w:r></w:p>',
    ]

    if meta.abstract.strip():
        body_xml_parts.append(
            f'<w:p><w:pPr><w:pStyle w:val="{Style3GPP.NORMAL}"/><w:spacing w:after="180"/></w:pPr>'
            f'<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Abstract of the contribution: </w:t></w:r>'
            f'<w:r><w:t>{meta.abstract}</w:t></w:r></w:p>'
        )

    # 4. Clause 1. Discussion
    disc_text = meta.discussion_text.strip() or (
        f"This document proposes updates to {meta.spec_type} {meta.spec_num} v{meta.spec_version} "
        f"for {meta.title}."
    )
    body_xml_parts.extend([
        f'<w:p><w:pPr><w:pStyle w:val="{Style3GPP.HEADING_1}"/><w:spacing w:before="240" w:after="120"/></w:pPr>'
        f'<w:r><w:t>1. Discussion</w:t></w:r></w:p>',
        f'<w:p><w:pPr><w:pStyle w:val="{Style3GPP.NORMAL}"/><w:spacing w:after="180"/></w:pPr>'
        f'<w:r><w:t>{disc_text}</w:t></w:r></w:p>',
    ])

    # 5. Clause 2. Proposal
    prop_text = meta.proposal_intro.strip() or (
        f"It is proposed to include the following changes in {meta.spec_type} {meta.spec_num} v{meta.spec_version}:"
    )
    body_xml_parts.extend([
        f'<w:p><w:pPr><w:pStyle w:val="{Style3GPP.HEADING_1}"/><w:spacing w:before="240" w:after="120"/></w:pPr>'
        f'<w:r><w:t>2. Proposal</w:t></w:r></w:p>',
        f'<w:p><w:pPr><w:pStyle w:val="{Style3GPP.NORMAL}"/><w:spacing w:after="180"/></w:pPr>'
        f'<w:r><w:t>{prop_text}</w:t></w:r></w:p>',
    ])

    # 6. Injected Change Separators and Anchors
    for idx, block in enumerate(change_blocks):
        if idx == 0:
            sep_label = "* * * First Change * * * *"
        else:
            sep_label = "* * * Next Change (all new)* * * *" if block.is_all_new else "* * * Next Change * * * *"

        body_xml_parts.extend([
            f'<w:p><w:pPr><w:pStyle w:val="{Style3GPP.CR_SEPARATOR}"/><w:jc w:val="center"/><w:spacing w:before="240" w:after="120"/></w:pPr>'
            f'<w:r><w:rPr><w:rStyle w:val="{Style3GPP.CR_SEPARATOR_CHAR}"/><w:b/><w:sz w:val="24"/></w:rPr>'
            f'<w:t>{sep_label}</w:t></w:r></w:p>',
            f'<w:p><w:r><w:t>{{{{CHANGE_BLOCK_ANCHOR_{idx}}}}}</w:t></w:r></w:p>',
        ])

    # Final End of Changes separator
    body_xml_parts.extend([
        f'<w:p><w:pPr><w:pStyle w:val="{Style3GPP.CR_SEPARATOR}"/><w:jc w:val="center"/><w:spacing w:before="240" w:after="120"/></w:pPr>'
        f'<w:r><w:rPr><w:rStyle w:val="{Style3GPP.CR_SEPARATOR_CHAR}"/><w:b/><w:sz w:val="24"/></w:rPr>'
        f'<w:t>* * * End of Changes * * * *</w:t></w:r></w:p>',
        '<w:p><w:pPr><w:spacing w:after="0"/></w:pPr></w:p>',
        # Page layout / section properties
        f'<w:sectPr><w:pgSz w:w="11907" w:h="16840" w:code="9"/>'
        f'<w:pgMar w:top="1418" w:right="1134" w:bottom="1134" w:left="1134" w:header="680" w:footer="567" w:gutter="0"/>'
        f'<w:cols w:space="720"/></w:sectPr>',
        '</w:body></w:document>',
    ])

    pcr_doc_xml = "".join(body_xml_parts).encode("utf-8")

    # Repackage template with new document.xml
    with zipfile.ZipFile(io.BytesIO(tpl_bytes), "r") as zin:
        out_buf = io.BytesIO()
        with zipfile.ZipFile(out_buf, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "word/document.xml":
                    zout.writestr(item, pcr_doc_xml)
                else:
                    zout.writestr(item, zin.read(item.filename))

    with open(to_extended_path(target_path), "wb") as f:
        f.write(out_buf.getvalue())
    ensure_3gpp_package_styles(target_path)


# ---------------------------------------------------------------------------
# Generator Engine & Isolated Subprocess Execution
# ---------------------------------------------------------------------------

class CrPcrGeneratorEngine:
    """
    Orchestrates end-to-end CR and pCR generation:
    1. Slices contiguous specification fragments via worker subprocess.
    2. Builds base template with 3GPP styles and change anchors.
    3. Splices fragments into anchors via Word COM Range.InsertFile.
    4. Enforces corporate sensitivity labeling and 3GPP styles lock-in.
    """

    def __init__(
        self,
        metadata: ContributionMetadata,
        source_paths: Sequence[str],
        headings: Sequence[HeadingInfo],
        selected_ids: Sequence[str],
        all_new_map: Optional[Dict[str, bool]] = None,
        progress_callback: Optional[Callable[[int, str], None]] = None,
    ):
        self.meta = metadata
        self.source_paths = list(source_paths)
        self.headings = list(headings)
        self.selected_ids = list(selected_ids)
        self.all_new_map = all_new_map or {}
        self.progress_callback = progress_callback

    def _progress(self, percent: int, msg: str):
        if self.progress_callback:
            self.progress_callback(max(0, min(100, int(percent))), msg)

    def _make_fragment(self, source: Path, ranges: List[Tuple[int, int]], target: Path) -> None:
        """Invokes isolated worker to slice the source specification."""
        helper = Path(__file__).with_name("word_excerpt_fragment_worker.py")
        if not helper.is_file():
            raise RuntimeError(f"Fragment worker not found at {helper}")

        merged = WordExcerptEngine._merge_ranges(ranges, 2**31 - 1)
        cmd = [
            sys.executable,
            str(helper),
            str(source),
            str(target),
            json.dumps(merged),
        ]
        flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        res = subprocess.run(cmd, capture_output=True, text=True, creationflags=flags, check=False)
        if res.returncode != 0:
            err = res.stderr.strip() or res.stdout.strip() or "Unknown worker error"
            raise RuntimeError(f"Fragment extraction failed for {source.name}: {err}")

    def generate(self, output_path: str) -> Path:
        """Executes full generation pipeline."""
        output = Path(output_path).resolve()
        os.makedirs(to_extended_path(output.parent), exist_ok=True)
        work_dir = Path(tempfile.mkdtemp(prefix="cr_gen_", dir=str(get_temp_root())))

        try:
            # 1. Group selected clauses into contiguous change blocks
            self._progress(5, "Analyzing clause contiguity and change blocks...")
            change_blocks = group_selected_headings_into_blocks(
                self.headings, self.selected_ids, self.all_new_map
            )
            if not change_blocks:
                raise ValueError("No valid clauses selected for contribution generation.")

            # Auto-compute affected clauses if empty
            if not self.meta.clauses_affected:
                self.meta.clauses_affected = extract_clause_summary(change_blocks)

            logger.info(
                "Grouped %d selected heading(s) into %d change block(s): %s",
                len(self.selected_ids),
                len(change_blocks),
                self.meta.clauses_affected,
            )

            # 2. Extract specification fragments
            fragment_paths = []
            total_blocks = len(change_blocks)
            for idx, block in enumerate(change_blocks):
                pct = 10 + int(40 * idx / total_blocks)
                self._progress(pct, f"Extracting change section {idx + 1}/{total_blocks}...")
                frag = work_dir / f"change_{idx:03d}.docx"
                self._make_fragment(Path(block.source_path), block.ranges, frag)
                fragment_paths.append(frag)

            # 3. Build Base Document (CR or pCR)
            self._progress(55, f"Constructing base {self.meta.mode.upper()} template...")
            base_doc_path = work_dir / f"base_{self.meta.mode}.docx"
            if self.meta.mode.lower() == "cr":
                build_cr_base_document(self.meta, change_blocks, base_doc_path)
            else:
                build_pcr_base_document(self.meta, change_blocks, base_doc_path)

            # 4. Splice fragments into anchors via Word COM
            self._progress(65, "Splicing change blocks with Microsoft Word COM...")
            merged_working_path = work_dir / "finalized_working.docx"
            shutil.copy2(to_extended_path(base_doc_path), to_extended_path(merged_working_path))

            self._splice_with_word(merged_working_path, fragment_paths)

            # 5. Lock in 3GPP styles and sensitivity label
            self._progress(85, "Applying 3GPP styles and corporate sensitivity label...")
            # Restore 3GPP formatting parts with enhanced QuickStyle visibility
            WordExcerptEngine._copy_formatting_parts(get_template_cr_path(), merged_working_path)
            ensure_3gpp_package_styles(merged_working_path)

            # Apply sensitivity label in a fresh Word session
            pythoncom.CoInitialize()
            try:
                WordExcerptEngine._apply_label_in_fresh_word(merged_working_path)
            finally:
                pythoncom.CoUninitialize()

            # Final style assertion: ensure styles.xml retains all 3GPP QuickStyles and settings.xml has linkStyles stripped
            ensure_3gpp_package_styles(merged_working_path)

            # 6. Copy finalized file to output path
            ext_output = to_extended_path(output)
            if os.path.exists(ext_output):
                try:
                    os.remove(ext_output)
                except OSError:
                    pass
            shutil.copy2(to_extended_path(merged_working_path), ext_output)

            self._progress(100, f"Successfully created {output.name}!")
            return output

        finally:
            shutil.rmtree(work_dir, ignore_errors=True)

    def _splice_with_word(self, base_doc_path: Path, fragment_paths: List[Path]) -> None:
        """Opens base document, finds each anchor, and inserts the corresponding fragment."""
        word = None
        doc = None
        com_initialized = False
        com_dead = False

        try:
            pythoncom.CoInitialize()
            com_initialized = True
            word = win32com.client.DispatchEx("Word.Application")
            WordExcerptEngine._configure_word(word)

            doc = word.Documents.Open(
                os.path.normpath(str(base_doc_path)),
                ReadOnly=False,
                AddToRecentFiles=False,
            )

            # Unprotect and pre-label to prevent enterprise editing lockouts
            try:
                if doc.ProtectionType != -1:
                    doc.Unprotect()
            except Exception:
                pass

            try:
                set_sensitivity_label(doc)
            except Exception as label_exc:
                logger.warning("Pre-labeling in COM encountered an issue: %s", label_exc)

            # Insert each fragment at its corresponding anchor
            for idx, frag_path in enumerate(fragment_paths):
                anchor_text = f"{{{{CHANGE_BLOCK_ANCHOR_{idx}}}}}"
                rng = doc.Content
                find = rng.Find
                find.ClearFormatting()
                find.Text = anchor_text
                found = find.Execute()
                if found:
                    rng.Text = ""  # Clear placeholder anchor text
                    rng.InsertFile(FileName=os.path.normpath(str(frag_path)))
                else:
                    logger.warning("Could not locate anchor %s in base document.", anchor_text)

            # Update fields so DOCPROPERTY numbers render
            try:
                doc.Fields.Update()
            except Exception:
                pass

            # Promote all 3GPP styles in live Word model
            promote_3gpp_styles_in_word(doc)

            doc.Save()
            doc.Close(SaveChanges=False)
            doc = None
            word.Quit(SaveChanges=0)
            word = None

        except Exception as exc:
            com_dead = WordExcerptEngine._is_rpc_disconnect(exc)
            raise RuntimeError(f"Word COM assembly failed: {exc}") from exc
        finally:
            if not com_dead:
                if doc is not None:
                    try:
                        doc.Close(SaveChanges=False)
                    except Exception:
                        pass
                if word is not None:
                    try:
                        word.Quit(SaveChanges=0)
                    except Exception:
                        pass
            if com_initialized:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass


# ---------------------------------------------------------------------------
# Background Worker Thread for PyQt5
# ---------------------------------------------------------------------------

class CrPcrGeneratorThread(QThread):
    """Asynchronous background worker thread for generating CR/pCR contributions."""
    finished_path = pyqtSignal(str)
    failed = pyqtSignal(str)
    progress = pyqtSignal(int, str)

    def __init__(
        self,
        metadata: ContributionMetadata,
        source_paths: Sequence[str],
        headings: Sequence[HeadingInfo],
        selected_ids: Sequence[str],
        output_path: str,
        all_new_map: Optional[Dict[str, bool]] = None,
    ):
        super().__init__()
        self.metadata = metadata
        self.source_paths = source_paths
        self.headings = headings
        self.selected_ids = selected_ids
        self.output_path = output_path
        self.all_new_map = all_new_map or {}

    def run(self):
        try:
            logger.info("Starting contribution generation (%s)...", self.metadata.mode.upper())
            engine = CrPcrGeneratorEngine(
                metadata=self.metadata,
                source_paths=self.source_paths,
                headings=self.headings,
                selected_ids=self.selected_ids,
                all_new_map=self.all_new_map,
                progress_callback=lambda pct, msg: self.progress.emit(pct, msg),
            )
            out_file = engine.generate(self.output_path)
            self.finished_path.emit(str(out_file))
        except Exception as exc:
            logger.exception("Contribution generation failed: %s", exc)
            self.failed.emit(str(exc))
