"""Compact Word/Outlook clipboard HTML: removes Word's implicit paragraph spacing inside table cells.

Word pastes HTML table cells that contain bare text using the target document's `Normal` style
(typically 8 pt space-after and 1.08 line spacing), which makes every row very tall. Wrapping each
cell's content in an explicit zero-margin `MsoNormal` paragraph and replacing cell padding with
tight `mso-padding-alt` values keeps pasted tables as compact as they look on screen.
"""

from __future__ import annotations

import re

_CELL_RE = re.compile(r"(<(td|th)\b[^>]*>)(.*?)(</\2>)", re.IGNORECASE | re.DOTALL)
_STYLE_ATTR_RE = re.compile(r'\bstyle="([^"]*)"', re.IGNORECASE)
_P_TAG_RE = re.compile(r"<p\b([^>]*)>", re.IGNORECASE)
_BLOCK_START_RE = re.compile(r"^\s*<(p|div|table|ul|ol|pre|h[1-6])\b", re.IGNORECASE)

_PARA_STYLE = (
    "margin: 0cm; margin-bottom: .0001pt; mso-margin-top-alt: 0cm; mso-margin-bottom-alt: 0cm; "
    "line-height: 105%;"
)
_CELL_PADDING = "padding: 1.5pt 4pt 1.5pt 4pt; mso-padding-alt: 1.5pt 4pt 1.5pt 4pt"

WORD_HEAD_STYLE = (
    "<style>"
    "p.MsoNormal, td p, th p { margin: 0cm; margin-bottom: .0001pt; line-height: normal; }"
    "table { border-collapse: collapse; mso-table-lspace: 0pt; mso-table-rspace: 0pt; }"
    "td, th { padding: 1.5pt 4pt 1.5pt 4pt; mso-padding-alt: 1.5pt 4pt 1.5pt 4pt; vertical-align: middle; }"
    "</style>"
)


def _set_style(attrs: str, declarations: str, *, drop: tuple = ("padding", "margin", "line-height")) -> str:
    """Return tag attribute string with `declarations` appended to (and conflicting props removed from) `style`."""
    match = _STYLE_ATTR_RE.search(attrs)
    if not match:
        return f'{attrs} style="{declarations}"'
    existing = match.group(1)
    kept = []
    for decl in existing.split(";"):
        decl = decl.strip()
        if not decl:
            continue
        prop = decl.split(":", 1)[0].strip().lower()
        if any(prop == d or prop.startswith(d + "-") for d in drop):
            continue
        kept.append(decl)
    new_style = "; ".join(kept + [declarations])
    return attrs[: match.start()] + f'style="{new_style}"' + attrs[match.end():]


def _fix_paragraph_tags(fragment: str) -> str:
    return _P_TAG_RE.sub(
        lambda m: "<p" + _set_style(m.group(1), _PARA_STYLE, drop=("margin", "line-height")) + ">",
        fragment,
    )


def _wrap_cell(match: "re.Match[str]") -> str:
    open_tag, tag_name, inner, close_tag = match.group(1), match.group(2), match.group(3), match.group(4)
    attrs = open_tag[len(tag_name) + 1 : -1]
    new_open = f"<{tag_name}" + _set_style(attrs, _CELL_PADDING, drop=("padding",)) + ">"
    if not inner.strip():
        return f"{new_open}<p class=MsoNormal style=\"{_PARA_STYLE}\">&nbsp;</p>{close_tag}"
    if _BLOCK_START_RE.match(inner):
        return f"{new_open}{_fix_paragraph_tags(inner)}{close_tag}"
    return f'{new_open}<p class=MsoNormal style="{_PARA_STYLE}">{inner}</p>{close_tag}'


def compact_html_for_word(html_text: str, *, full_document: bool = False) -> str:
    """Tighten table-cell spacing in `html_text` for pasting into Microsoft Word / Outlook.

    Wraps bare cell content in zero-margin `MsoNormal` paragraphs, shrinks cell padding, and optionally
    wraps the fragment in a full HTML document containing a compact `<style>` block. Idempotent: cells whose
    content already starts with a block element only get their `<p>` margins normalised. Nested tables inside
    cells are not expected in 3GPP exports and are left unwrapped on a best-effort basis.
    """
    if not html_text or "<t" not in html_text.lower():
        return html_text

    result = _CELL_RE.sub(_wrap_cell, html_text)
    if full_document:
        result = (
            "<html><head><meta charset=\"utf-8\">" + WORD_HEAD_STYLE + "</head><body>"
            + result + "</body></html>"
        )
    return result
