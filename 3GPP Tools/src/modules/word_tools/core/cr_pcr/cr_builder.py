"""Formal 3GPP Change Request (CR) OpenXML template population builders."""

from __future__ import annotations

import io
import re
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Sequence

from lxml import etree

from core.utils.paths import to_extended_path
from modules.word_tools.core.cr_pcr.models import (
    C_PROP_NS,
    ChangeBlock,
    ContributionMetadata,
    VT_NS,
    W_NS,
)
from modules.word_tools.core.styles_3gpp import (
    Style3GPP,
    ensure_3gpp_package_styles,
    get_template_cr_path,
)


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
    for idx, _block in enumerate(change_blocks):
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
