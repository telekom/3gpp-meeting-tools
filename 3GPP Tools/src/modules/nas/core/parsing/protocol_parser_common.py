import html
import zipfile
from pathlib import Path
from typing import Optional, List, Dict, Any

try:
    from lxml import etree as ET
except ImportError:
    import xml.etree.ElementTree as ET

# WordprocessingML XML Namespaces
W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
TAG_BODY = f"{W_NS}body"
TAG_P = f"{W_NS}p"
TAG_TBL = f"{W_NS}tbl"
TAG_TR = f"{W_NS}tr"
TAG_TC = f"{W_NS}tc"
TAG_T = f"{W_NS}t"
TAG_TAB = f"{W_NS}tab"
TAG_BR = f"{W_NS}br"
TAG_CR = f"{W_NS}cr"
TAG_HYPHEN = f"{W_NS}noBreakHyphen"
TAG_TCPR = f"{W_NS}tcPr"
TAG_GRIDSPAN = f"{W_NS}gridSpan"
TAG_VMERGE = f"{W_NS}vMerge"
TAG_PPR = f"{W_NS}pPr"
TAG_JC = f"{W_NS}jc"


def extract_document_root(docx_path: Path) -> Optional[ET.Element]:
    """Safely extracts and parses word/document.xml from a .docx zip archive."""
    if not docx_path.exists():
        return None
    try:
        with zipfile.ZipFile(docx_path, "r") as zf:
            if "word/document.xml" not in zf.namelist():
                return None
            xml_bytes = zf.read("word/document.xml")
            return ET.fromstring(xml_bytes)
    except Exception:
        return None


def _extract_p_text(p_elem: ET.Element) -> str:
    """Extracts clean text from a <w:p> node, preserving spaces, tabs, and line breaks."""
    text_pieces: List[str] = []
    for node in p_elem.iter():
        tag = node.tag
        if tag == TAG_T and node.text:
            text_pieces.append(node.text)
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

    # Phase 1: Parse cell metadata grid
    grid: List[List[Dict[str, Any]]] = []
    for row in rows:
        row_cells = []
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
                "colspan": colspan,
                "vmerge": vmerge_status,
                "rowspan": 1,
                "skip": False,
            })
        grid.append(row_cells)

    # Phase 2: Resolve vertical merge rowspans
    num_rows = len(grid)
    for r_idx in range(num_rows):
        for c_idx, cell_data in enumerate(grid[r_idx]):
            if cell_data["vmerge"] == "restart":
                span = 1
                for next_r in range(r_idx + 1, num_rows):
                    if c_idx < len(grid[next_r]) and grid[next_r][c_idx]["vmerge"] == "continue":
                        span += 1
                        grid[next_r][c_idx]["skip"] = True
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