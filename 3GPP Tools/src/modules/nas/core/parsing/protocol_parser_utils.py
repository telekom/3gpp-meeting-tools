import html
import os
import zipfile
from core.utils.paths import to_extended_path
from pathlib import Path
from typing import Optional, List, Dict, Any
from xml.etree import ElementTree as ET

from lxml import etree as ET

from modules.nas.core.parsing.protocol_parser_constants import (
    TAG_T, TAG_TAB, TAG_BR, TAG_CR, TAG_HYPHEN, TAG_P,
    TAG_TR, TAG_TC, TAG_TCPR, TAG_GRIDSPAN, W_NS, TAG_VMERGE,
    TAG_R, TAG_RPR, TAG_VERTALIGN
)

_SUPERSCRIPT_TRANS = str.maketrans("0123456789+-=()", "⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾")
_SUBSCRIPT_TRANS = str.maketrans("0123456789+-=()", "₀₁₂₃₄₅₆₇₈₉₊₋₌₍₎")


def extract_document_root(docx_path: Path) -> Optional[ET.Element]:
    """Safely extracts and parses word/document.xml from a .docx zip archive."""
    ext_path = to_extended_path(docx_path)
    if not os.path.exists(ext_path):
        return None
    try:
        with zipfile.ZipFile(ext_path, "r") as zf:
            if "word/document.xml" not in zf.namelist():
                return None
            xml_bytes = zf.read("word/document.xml")
            return ET.fromstring(xml_bytes)
    except Exception:
        return None


def _extract_p_text(p_elem: ET.Element) -> str:
    """Extracts clean text from a <w:p> node, preserving spaces, tabs, line breaks, and numeric superscripts."""
    text_pieces: List[str] = []
    current_valign: Optional[str] = None

    for node in p_elem.iter():
        tag = node.tag
        if tag == TAG_R:
            current_valign = None
            rpr = node.find(TAG_RPR)
            if rpr is not None:
                va = rpr.find(TAG_VERTALIGN)
                if va is not None:
                    current_valign = va.get(f"{W_NS}val")
        elif tag == TAG_T and node.text:
            txt = node.text
            if current_valign == "superscript" and any(ch.isdigit() for ch in txt):
                txt = txt.translate(_SUPERSCRIPT_TRANS)
            elif current_valign == "subscript" and any(ch.isdigit() for ch in txt):
                txt = txt.translate(_SUBSCRIPT_TRANS)
            text_pieces.append(txt)
        elif tag == TAG_TAB:
            text_pieces.append(" ")
        elif tag in (TAG_BR, TAG_CR):
            text_pieces.append(" ")
        elif tag == TAG_HYPHEN:
            text_pieces.append("-")

    raw = "".join(text_pieces)
    raw = raw.replace("\u00a0", " ").replace("\xa0", " ")
    return " ".join(raw.split())


def _extract_tc_text(tc_elem: ET.Element) -> str:
    """Extracts text from a table cell <w:tc>, joining multiple paragraphs with spaces."""
    p_texts = [_extract_p_text(p) for p in tc_elem.findall(TAG_P)]
    return " ".join(pt for pt in p_texts if pt).strip()


def _convert_table_to_html(tbl_elem: ET.Element, is_figure_diagram: bool = False) -> str:
    """
    Converts a Word XML table into a styled HTML table with full rowspan and colspan support.
    Applies bit-grid styling if is_figure_diagram is True.
    """
    rows = tbl_elem.findall(TAG_TR)
    if not rows:
        return ""

    # Phase 1: Parse cell metadata grid with logical column offsets
    grid: List[List[Dict[str, Any]]] = []
    for row in rows:
        row_cells = []
        col_offset = 0
        for cell in row.findall(TAG_TC):
            tc_pr = cell.find(TAG_TCPR)
            colspan = 1
            vmerge_status = "none"  # "restart", "continue", or "none"

            if tc_pr is not None:
                gs = tc_pr.find(TAG_GRIDSPAN)
                if gs is not None:
                    val = gs.get(f"{W_NS}val")
                    if val and val.isdigit():
                        colspan = int(val)

                vm = tc_pr.find(TAG_VMERGE)
                if vm is not None:
                    val = vm.get(f"{W_NS}val")
                    vmerge_status = "restart" if val == "restart" else "continue"

            row_cells.append({
                "text": _extract_tc_text(cell),
                "col_offset": col_offset,
                "colspan": colspan,
                "vmerge": vmerge_status,
                "rowspan": 1,
                "skip": False,
            })
            col_offset += colspan
        grid.append(row_cells)

    # Phase 2: Resolve vertical merge rowspans by matching logical col_offset
    num_rows = len(grid)
    for r_idx in range(num_rows):
        for cell_data in grid[r_idx]:
            if cell_data["vmerge"] == "restart":
                target_offset = cell_data["col_offset"]
                span = 1
                for next_r in range(r_idx + 1, num_rows):
                    matched_next = None
                    for next_cell in grid[next_r]:
                        if next_cell["col_offset"] == target_offset and next_cell["vmerge"] == "continue":
                            matched_next = next_cell
                            break
                    if matched_next is not None:
                        span += 1
                        matched_next["skip"] = True
                    else:
                        break
                cell_data["rowspan"] = span
            elif cell_data["vmerge"] == "continue" and not cell_data["skip"]:
                cell_data["skip"] = True

    # Phase 3: Build HTML string
    table_font = "Consolas, 'Courier New', monospace" if is_figure_diagram else "'Segoe UI', Arial, sans-serif"
    table_style = (
        f"border-collapse: collapse; margin: 8px 0; border: 1px solid #CBD5E1; "
        f"font-family: {table_font}; font-size: 11px; width: 100%;"
    )
    html_parts = [f'<table border="1" cellspacing="0" cellpadding="4" style="{table_style}">']

    for r_idx, row_cells in enumerate(grid):
        html_parts.append("<tr>")
        for cell_data in row_cells:
            if cell_data["skip"]:
                continue

            tag = "th" if r_idx == 0 else "td"
            style_bits = ["border: 1px solid #CBD5E1;", "padding: 4px 6px;"]

            if is_figure_diagram:
                style_bits.append("text-align: center;")

            if r_idx == 0:
                style_bits.append("background-color: #F1F5F9; font-weight: bold; color: #1E293B;")
            elif is_figure_diagram and r_idx % 2 == 1:
                style_bits.append("background-color: #FAFAFA;")

            colspan_attr = f' colspan="{cell_data["colspan"]}"' if cell_data["colspan"] > 1 else ""
            rowspan_attr = f' rowspan="{cell_data["rowspan"]}"' if cell_data["rowspan"] > 1 else ""
            style_str = " ".join(style_bits)
            escaped_text = html.escape(cell_data["text"])

            html_parts.append(f'<{tag}{colspan_attr}{rowspan_attr} style="{style_str}">{escaped_text}</{tag}>')
        html_parts.append("</tr>")

    html_parts.append("</table>")
    return "".join(html_parts)
