# --- File: src/modules/nas/core/nas_exporter.py ---
"""
High-fidelity clipboard and document exporter for 3GPP Protocol Evolution Matrix
and Information Element (IE) definitions.

Provides Word-compatible rich HTML (pasted as native styled Word tables and formatted
clause excerpts with Track Changes diff color preservation), GitHub-Flavored Markdown
tables and quote blocks, and pure TSV/plain text fallbacks.
"""

import html
import re
from typing import Any, List, Optional
from bs4 import BeautifulSoup, NavigableString, Tag

from PyQt5.QtCore import QMimeData
from PyQt5.QtWidgets import QApplication


# -----------------------------------------------------------------------------
# 1. Evolution Matrix Export Functions
# -----------------------------------------------------------------------------

def build_matrix_html(
    title: str,
    headers: List[str],
    rows: List[List[Any]],
    diff_colors: Optional[List[List[Optional[str]]]] = None,
) -> str:
    """
    Builds a Word/Outlook-compatible rich HTML table for the Evolution Matrix.
    Preserves cell padding, borders, headers, and diff background colors
    (soft green for Added, soft yellow for Modified, soft red for Removed).
    """
    clean_headers = [str(h or "").replace("\n", " ").strip() for h in headers]
    html_parts = [
        '<div style="font-family: Calibri, \'Segoe UI\', Arial, sans-serif; font-size: 10pt; color: #1E293B;">'
    ]

    if title and title.strip():
        html_parts.append(
            f'<div style="margin-bottom: 8px; font-weight: bold; color: #0369A1; font-size: 11pt;">'
            f'{html.escape(title.strip())}</div>'
        )

    html_parts.append(
        '<table style="border-collapse: collapse; width: 100%; border: 1px solid #CBD5E1; '
        'font-family: Calibri, \'Segoe UI\', Arial, sans-serif; font-size: 9.5pt;">'
    )

    # Table Header
    html_parts.append('<thead><tr style="background-color: #F1F5F9; color: #0F172A; font-weight: bold; border-bottom: 2px solid #94A3B8;">')
    for c_idx, h in enumerate(clean_headers):
        align = "text-align: center;" if c_idx >= 3 else "text-align: left;"
        html_parts.append(
            f'<th style="border: 1px solid #CBD5E1; padding: 6px 10px; {align} font-weight: bold; background-color: #F1F5F9;">'
            f'{html.escape(h)}</th>'
        )
    html_parts.append('</tr></thead><tbody>')

    # Table Body
    for r_idx, row in enumerate(rows):
        bg_row = "#FFFFFF" if r_idx % 2 == 0 else "#F8FAFC"
        html_parts.append(f'<tr style="background-color: {bg_row};">')
        for c_idx, val in enumerate(row):
            cell_bg = ""
            if diff_colors and r_idx < len(diff_colors) and c_idx < len(diff_colors[r_idx]):
                color = diff_colors[r_idx][c_idx]
                if color:
                    cell_bg = f"background-color: {color}; font-weight: bold;"

            align = "text-align: center;" if c_idx >= 3 else "text-align: left;"
            val_str = str(val or "")
            val_esc = html.escape(val_str)

            # Preserve tree indentation for child fields
            if c_idx == 1 and val_str.startswith(" "):
                leading_spaces = len(val_str) - len(val_str.lstrip(" "))
                val_esc = ("&nbsp;" * leading_spaces) + html.escape(val_str.lstrip(" "))

            html_parts.append(
                f'<td style="border: 1px solid #CBD5E1; padding: 4px 8px; {align} {cell_bg}">{val_esc}</td>'
            )
        html_parts.append('</tr>')

    html_parts.append('</tbody></table></div>')
    return "\n".join(html_parts)


def build_matrix_markdown(
    title: str,
    headers: List[str],
    rows: List[List[Any]],
) -> str:
    """
    Builds a clean GitHub-Flavored Markdown table for the Evolution Matrix.
    Escapes pipe characters and preserves hierarchy indentation.
    """
    clean_headers = [str(h or "").replace("\n", " ").replace("|", "\\|").strip() for h in headers]
    lines = []

    if title and title.strip():
        lines.append(f"### {title.strip()}\n")

    lines.append("| " + " | ".join(clean_headers) + " |")
    lines.append("| " + " | ".join(["---"] * len(clean_headers)) + " |")

    for row in rows:
        clean_row = []
        for c_idx, val in enumerate(row):
            v_str = str(val or "").replace("\n", " ").replace("|", "\\|").strip()
            if c_idx == 1 and "└─" in str(val):
                v_str = str(val).replace("    ", "&nbsp;&nbsp;&nbsp;&nbsp;").replace("|", "\\|")
            clean_row.append(v_str)
        lines.append("| " + " | ".join(clean_row) + " |")

    return "\n".join(lines)


def build_matrix_tsv(headers: List[str], rows: List[List[Any]]) -> str:
    """
    Exports the Evolution Matrix as clean tab-separated values (for Excel and scripting).
    """
    clean_headers = [str(h or "").replace("\n", " ").strip() for h in headers]
    lines = ["\t".join(clean_headers)]
    for row in rows:
        lines.append("\t".join(str(v or "").replace("\t", " ").replace("\n", " ") for v in row))
    return "\n".join(lines)


def copy_matrix_to_clipboard(
    title: str,
    headers: List[str],
    rows: List[List[Any]],
    diff_colors: Optional[List[List[Optional[str]]]] = None,
    mode: str = "word",
) -> str:
    """
    Copies the Evolution Matrix to the system clipboard in the requested format.
    Modes:
      - 'word': Dual-format rich HTML (native Word table) + TSV fallback.
      - 'markdown': GitHub-Flavored Markdown table.
      - 'tsv': Tab-separated values (for Excel).
    Returns user-facing status feedback.
    """
    clipboard = QApplication.clipboard()
    if not clipboard:
        return "Clipboard unavailable"

    tsv_text = build_matrix_tsv(headers, rows)

    if mode == "word":
        html_text = build_matrix_html(title, headers, rows, diff_colors)
        mime = QMimeData()
        mime.setHtml(html_text)
        mime.setText(tsv_text)
        clipboard.setMimeData(mime)
        return "Copied for Word / Outlook!"

    elif mode == "markdown":
        md_text = build_matrix_markdown(title, headers, rows)
        clipboard.setText(md_text)
        return "Copied as Markdown Table!"

    else:  # 'tsv'
        clipboard.setText(tsv_text)
        return "Copied TSV (Excel)!"


# -----------------------------------------------------------------------------
# 2. IE Definition Export Functions
# -----------------------------------------------------------------------------

def html_table_to_markdown(table_tag: Tag) -> str:
    """
    Converts an HTML <table> tag (including colspans and bit-layout headers)
    into a clean, aligned GitHub-Flavored Markdown table.
    """
    rows = []
    for tr in table_tag.find_all("tr", recursive=False):
        row_cells = []
        for cell in tr.find_all(["th", "td"], recursive=False):
            colspan = 1
            if cell.has_attr("colspan"):
                try:
                    colspan = int(cell["colspan"])
                except (ValueError, TypeError):
                    colspan = 1
            text = cell.get_text(strip=True).replace("\n", " ").replace("|", "\\|")
            row_cells.append(text)
            for _ in range(colspan - 1):
                row_cells.append("")
        if row_cells:
            rows.append(row_cells)

    if not rows:
        return ""

    max_cols = max(len(r) for r in rows)
    padded_rows = [r + [""] * (max_cols - len(r)) for r in rows]

    md_lines = [
        "| " + " | ".join(padded_rows[0]) + " |",
        "| " + " | ".join(["---"] * max_cols) + " |",
    ]
    for r in padded_rows[1:]:
        md_lines.append("| " + " | ".join(r) + " |")

    return "\n".join(md_lines)


def build_definition_word_html(
    html_content: str,
    spec_number: Optional[str] = None,
    version: Optional[str] = None,
    ie_name: Optional[str] = None,
    clause: Optional[str] = None,
) -> str:
    """
    Wraps the definition HTML into a Word-optimized document container with
    3GPP citation header banner, styled monospace code boxes, and crisp tables.
    """
    if not html_content or not html_content.strip():
        return ""

    spec_clean = f"TS {spec_number.strip()}" if spec_number and not spec_number.startswith("TS") else (spec_number or "")
    ver_clean = f"v{version.lstrip('vV').strip()}" if version else ""
    spec_ver = f"{spec_clean} {ver_clean}".strip()
    name_str = ie_name or "Information Element"
    cl_str = f" (Clause {clause})" if clause and clause != ie_name else ""

    citation_title = f"3GPP {spec_ver} — {name_str}{cl_str}".strip(" —") if spec_ver else f"{name_str}{cl_str}"

    # Clean action URLs for external pasting
    clean_body = re.sub(r'href="action:[^"]*"', 'style="color: #0369A1; text-decoration: underline;"', html_content)

    return f"""<div style="font-family: Calibri, 'Segoe UI', Arial, sans-serif; font-size: 11pt; color: #1E293B; max-width: 850px;">
  <div style="margin-bottom: 12px; padding: 8px 12px; background-color: #F1F5F9; border-left: 4px solid #0284C7; border-radius: 4px;">
    <strong style="color: #0369A1; font-size: 12pt;">{html.escape(citation_title)}</strong>
  </div>
  <div style="padding-left: 2px;">
    {clean_body}
  </div>
</div>"""


def build_definition_markdown(
    html_content: str,
    spec_number: Optional[str] = None,
    version: Optional[str] = None,
    ie_name: Optional[str] = None,
    clause: Optional[str] = None,
) -> str:
    """
    Converts an IE/protocol definition HTML structure into clean GitHub-Flavored Markdown:
    - ### Citation header
    - ```asn1 code blocks
    - #### Subsections
    - Markdown tables
    - > [!NOTE] Blockquotes for cross-references
    """
    if not html_content or not html_content.strip():
        return ""

    spec_clean = f"TS {spec_number.strip()}" if spec_number and not spec_number.startswith("TS") else (spec_number or "")
    ver_clean = f"v{version.lstrip('vV').strip()}" if version else ""
    spec_ver = f"{spec_clean} {ver_clean}".strip()
    name_str = ie_name or "Information Element"
    cl_str = f" (Clause {clause})" if clause and clause != ie_name else ""

    citation_title = f"### 3GPP {spec_ver} — {name_str}{cl_str}".strip(" —") if spec_ver else f"### {name_str}{cl_str}"

    if "<" not in html_content:
        return f"{citation_title}\n\n{html_content.strip()}"

    soup = BeautifulSoup(html_content, "html.parser")
    lines = [citation_title, ""]

    def process_node(node):
        if isinstance(node, NavigableString):
            txt = str(node).strip()
            if txt:
                lines.append(txt)
            return

        if not isinstance(node, Tag):
            return

        # Headings
        if node.name in ("h1", "h2"):
            txt = node.get_text(strip=True)
            if name_str not in txt:
                lines.append(f"\n#### {txt}\n")
            return

        if node.name in ("h3", "h4", "h5"):
            lines.append(f"\n#### {node.get_text(strip=True)}\n")
            return

        # Preformatted code (ASN.1, message syntax)
        if node.name == "pre":
            code = node.get_text()
            code_lines = code.split("\n")
            while code_lines and not code_lines[0].strip():
                code_lines.pop(0)
            while code_lines and not code_lines[-1].strip():
                code_lines.pop()
            clean_code = "\n".join(code_lines)
            lines.append("\n```asn1")
            lines.append(clean_code)
            lines.append("```\n")
            return

        # Specification tables
        if node.name == "table":
            tbl_md = html_table_to_markdown(node)
            if tbl_md:
                lines.append("\n" + tbl_md + "\n")
            return

        # Divs (Section titles, Cross-references, Callouts)
        if node.name == "div":
            style = node.get("style", "")
            is_bold_title = "font-weight: bold" in style or "font-weight:bold" in style

            if is_bold_title and not node.find(["pre", "table", "div"]):
                title = node.get_text(strip=True).rstrip(":")
                lines.append(f"\n#### {title}\n")
                return

            if (
                "border-left: 4px solid #0284C7" in style
                or "border: 1.5px solid #0284C7" in style
                or "Cross-Referenced" in node.get_text()
                or "Cross-Specification Reference" in node.get_text()
            ):
                callout_text = node.get_text("\n", strip=True)
                quote_lines = [f"> {l}" for l in callout_text.split("\n") if l.strip()]
                if quote_lines:
                    lines.append("\n> [!NOTE]\n" + "\n".join(quote_lines) + "\n")
                return

        # Paragraphs
        if node.name == "p":
            txt = node.get_text(strip=True)
            if txt:
                lines.append(f"\n{txt}\n")
            return

        # Recurse children
        for child in node.children:
            process_node(child)

    for top_child in soup.children:
        process_node(top_child)

    raw_res = "\n".join(lines)
    clean_res = re.sub(r"\n{3,}", "\n\n", raw_res).strip()
    return clean_res


def build_definition_plain_text(
    html_content: str,
    spec_number: Optional[str] = None,
    version: Optional[str] = None,
    ie_name: Optional[str] = None,
    clause: Optional[str] = None,
) -> str:
    """
    Returns clean plain text with full 3GPP citation header.
    """
    spec_clean = f"TS {spec_number.strip()}" if spec_number and not spec_number.startswith("TS") else (spec_number or "")
    ver_clean = f"v{version.lstrip('vV').strip()}" if version else ""
    spec_ver = f"{spec_clean} {ver_clean}".strip()
    name_str = ie_name or "Information Element"
    cl_str = f" (Clause {clause})" if clause and clause != ie_name else ""

    citation = f"3GPP {spec_ver} — {name_str}{cl_str}".strip(" —") if spec_ver else f"{name_str}{cl_str}"

    if "<" in html_content:
        soup = BeautifulSoup(html_content, "html.parser")
        text = soup.get_text("\n", strip=True)
    else:
        text = html_content.strip()

    return f"{citation}\n\n{text}".strip()


def copy_definition_to_clipboard(
    html_content: str,
    spec_number: Optional[str] = None,
    version: Optional[str] = None,
    ie_name: Optional[str] = None,
    clause: Optional[str] = None,
    mode: str = "word",
) -> str:
    """
    Copies the IE definition to clipboard in the requested format:
      - 'word': Rich HTML (formatted clause excerpt) + Markdown/plain fallback.
      - 'markdown': GitHub-Flavored Markdown.
      - 'text': Clean plain text.
      - 'citation': Formal citation header only.
    Returns user-facing status feedback.
    """
    clipboard = QApplication.clipboard()
    if not clipboard:
        return "Clipboard unavailable"

    spec_clean = f"TS {spec_number.strip()}" if spec_number and not spec_number.startswith("TS") else (spec_number or "")
    ver_clean = f"v{version.lstrip('vV').strip()}" if version else ""
    spec_ver = f"{spec_clean} {ver_clean}".strip()
    name_str = ie_name or "Information Element"
    cl_str = f" (Clause {clause})" if clause and clause != ie_name else ""
    citation = f"3GPP {spec_ver} — {name_str}{cl_str}".strip(" —") if spec_ver else f"{name_str}{cl_str}"

    if mode == "word":
        word_html = build_definition_word_html(html_content, spec_number, version, ie_name, clause)
        md_fallback = build_definition_markdown(html_content, spec_number, version, ie_name, clause)
        mime = QMimeData()
        mime.setHtml(word_html)
        mime.setText(md_fallback)
        clipboard.setMimeData(mime)
        return "Copied for Word / Outlook!"

    elif mode == "markdown":
        md_text = build_definition_markdown(html_content, spec_number, version, ie_name, clause)
        clipboard.setText(md_text)
        return "Copied as Markdown!"

    elif mode == "citation":
        clipboard.setText(citation)
        return "Copied Citation!"

    else:  # 'text'
        plain_text = build_definition_plain_text(html_content, spec_number, version, ie_name, clause)
        clipboard.setText(plain_text)
        return "Copied Plain Text!"
