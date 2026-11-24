# --- File: src/modules/emails/core/markdown_export/html_and_tables.py ---
"""HTML angle-bracket escaping, GFM table extraction from Outlook HTML, and list normalization."""

import re
from typing import Any, List

from .url_and_links import unwrap_outlook_safelinks

_ALLOWED_HTML_TAGS = {
    "a", "b", "i", "em", "strong", "code", "pre", "br", "hr", "p", "div", "span",
    "ul", "ol", "li", "h1", "h2", "h3", "h4", "h5", "h6", "sub", "sup", "mark", "blockquote",
}
_ANGLE_HTTP_RE = re.compile(r"\s*<https?://[^>]+>", re.IGNORECASE)


def escape_non_html_angles(text: str) -> str:
    """Escapes inline delegate tags like <Ashok> or <Jinguo> so Qt's Markdown renderer does not treat them as unclosed HTML."""
    if not text:
        return ""
    cleaned = re.sub(r"\s*<mailto:[^>\s]+>\s*", " ", text, flags=re.IGNORECASE)

    def _replace_tag(m: re.Match) -> str:
        inner = m.group(1)
        first = inner.strip().split()[0].lstrip("/").lower() if inner.strip() else ""
        if first in _ALLOWED_HTML_TAGS:
            return m.group(0)
        return f"&lt;{inner}&gt;"

    return re.sub(r"<([^<>\n]{1,120})>", _replace_tag, cleaned)


def _table_cell_to_plain_and_md(cell: Any) -> tuple:
    """Extracts (plain_text_for_matching, markdown_cell_text) preserving clickable hyperlinks inside HTML tables."""
    plain_text = " ".join(cell.get_text(" ", strip=True).replace("\xa0", " ").split())

    parts: List[str] = []

    def _walk(el: Any):
        name = getattr(el, "name", None)
        if name is None:
            parts.append(str(el))
            return
        if name.lower() == "a" and el.get("href"):
            href = unwrap_outlook_safelinks(str(el.get("href") or "").strip())
            label = " ".join(el.get_text(" ", strip=True).replace("\xa0", " ").split()) or href
            if href.lower().startswith(("http://", "https://")) and label:
                safe_label = label.replace("|", "\\|")
                parts.append(f"[{safe_label}]({href})")
                return
        if name.lower() == "br":
            parts.append(" ")
            return
        for ch in getattr(el, "children", ()):
            _walk(ch)

    _walk(cell)
    md_text = " ".join("".join(parts).replace("\xa0", " ").split()).replace("|", "\\|")
    if not md_text:
        md_text = plain_text.replace("|", "\\|")
    return plain_text, md_text


def inject_html_tables_into_body(body_text: str, html_body: str) -> str:
    """
    Replaces Outlook's flattened plain-text table cells in `body_text` with structured
    GFM Markdown pipe tables extracted from `html_body`. Preserves surrounding text intact.
    """
    if not body_text or not html_body or "<table" not in html_body.lower():
        return body_text
    try:
        from bs4 import BeautifulSoup
    except ImportError:
        return body_text

    soup = BeautifulSoup(html_body, "html.parser")
    tables = [t for t in soup.find_all("table") if not t.find("table")]
    if not tables:
        return body_text

    lines = body_text.splitlines()
    search_from = 0
    for tbl in tables:
        plain_grid: List[List[str]] = []
        md_grid: List[List[str]] = []
        for tr in tbl.find_all("tr"):
            row_pairs = [_table_cell_to_plain_and_md(c) for c in tr.find_all(["td", "th"])]
            plain_cells = [p for p, _m in row_pairs]
            md_cells = [m for _p, m in row_pairs]
            if any(plain_cells):
                plain_grid.append(plain_cells)
                md_grid.append(md_cells)
        if not plain_grid or max(len(r) for r in plain_grid) < 2:
            continue
        tbl_norm = " ".join(" ".join(cell for row in plain_grid for cell in row if cell).split())
        if len(tbl_norm) < 4:
            continue

        matched_span = None
        for i in range(search_from, len(lines)):
            first = " ".join(_ANGLE_HTTP_RE.sub("", lines[i]).replace("\xa0", " ").split())
            if not first or not tbl_norm.startswith(first):
                continue
            acc = first
            if acc == tbl_norm:
                matched_span = (i, i)
                break
            for j in range(i + 1, len(lines)):
                nxt = " ".join(_ANGLE_HTTP_RE.sub("", lines[j]).replace("\xa0", " ").split())
                if not nxt:
                    continue
                cand = acc + " " + nxt
                if not tbl_norm.startswith(cand):
                    break
                acc = cand
                if acc == tbl_norm:
                    matched_span = (i, j)
                    break
            if matched_span:
                break

        if matched_span:
            col_count = max(len(r) for r in md_grid)
            padded = [r + [""] * (col_count - len(r)) for r in md_grid]
            md_rows = [
                "| " + " | ".join(padded[0]) + " |",
                "| " + " | ".join(["---"] * col_count) + " |",
            ] + ["| " + " | ".join(br) + " |" for br in padded[1:]]
            replacement = [""] + md_rows + [""]
            s_i, s_j = matched_span
            lines[s_i : s_j + 1] = replacement
            search_from = s_i + len(replacement)

    return "\n".join(lines)


def _normalize_outlook_list_spacing(text: str) -> str:
    """Normalizes Word/Outlook list markers with 4+ spaces/tabs so CommonMark does not render them as code blocks."""
    if not text:
        return ""
    cleaned = text.replace("\xa0", " ")
    cleaned = re.sub(r"^([ \t]*(?:[-*+•]|\d+\.))[ \t]{2,}", r"\1 ", cleaned, flags=re.MULTILINE)
    cleaned = re.sub(r"^[ \t]{4,}(?=\S)", "", cleaned, flags=re.MULTILINE)
    return cleaned
