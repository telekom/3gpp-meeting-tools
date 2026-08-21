import html

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


def _extract_p_text(p_elem) -> str:
    """Extracts clean text from a <w:p> node, preserving spaces, tabs, and line breaks."""
    text_pieces = []
    for node in p_elem.iter():
        tag = node.tag
        if tag == TAG_T:
            if node.text:
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


def _extract_tc_text(tc_elem) -> str:
    """Extracts text from a table cell <w:tc>, joining multiple paragraphs with spaces."""
    p_texts = []
    for p in tc_elem.findall(TAG_P):
        pt = _extract_p_text(p)
        if pt:
            p_texts.append(pt)
    return " ".join(p_texts).strip()


def _convert_table_to_html(tbl_elem, is_figure_diagram: bool = False) -> str:
    """Converts a Word XML table into a styled HTML table supporting colspan and vertical alignment."""
    rows = tbl_elem.findall(TAG_TR)
    if not rows:
        return ""

    table_style = (
        "border-collapse: collapse; margin: 8px 0; border: 1px solid #CBD5E1; "
        "font-family: 'Segoe UI', Arial, sans-serif; font-size: 11px; width: 100%;"
    )
    html_parts = [f'<table border="1" cellspacing="0" cellpadding="4" style="{table_style}">']

    for r_idx, row in enumerate(rows):
        html_parts.append("<tr>")
        cells = row.findall(TAG_TC)

        for cell in cells:
            tcPr = cell.find(TAG_TCPR)
            colspan = 1
            is_vmerge_continue = False

            if tcPr is not None:
                gs = tcPr.find(TAG_GRIDSPAN)
                if gs is not None:
                    val = gs.get(f"{W_NS}val")
                    if val and val.isdigit():
                        colspan = int(val)

                vm = tcPr.find(TAG_VMERGE)
                if vm is not None:
                    val = vm.get(f"{W_NS}val")
                    if val != "restart":
                        is_vmerge_continue = True

            if is_vmerge_continue:
                continue

            cell_text = html.escape(_extract_tc_text(cell))
            tag = "th" if r_idx == 0 else "td"
            style_bits = ["border: 1px solid #E2E8F0;", "padding: 4px 6px;"]

            if r_idx == 0:
                style_bits.append("background-color: #F1F5F9; font-weight: bold; color: #1E293B;")

            colspan_attr = f' colspan="{colspan}"' if colspan > 1 else ""
            style_str = " ".join(style_bits)
            html_parts.append(f'<{tag}{colspan_attr} style="{style_str}">{cell_text}</{tag}>')

        html_parts.append("</tr>")

    html_parts.append("</table>")
    return "".join(html_parts)



