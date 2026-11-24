"""Shared slide models, OpenXML namespaces, filename sanitization, and Markdown table helpers."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Callable, List, Optional

CONVERTER_VERSION = "2.1.0"

TDOC_RE = re.compile(r"(?<![A-Za-z0-9])([A-Z][A-Z0-9]{0,3}-\d{6,8}(?:r\d{1,2})?)(?![0-9])", re.I)

OPENXML_NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def _clean_tdoc_prefix(tdoc_id: Optional[str]) -> str:
    """Produces a clean, filesystem-safe prefix from a TDoc ID."""
    if not tdoc_id:
        return ""
    clean = re.sub(r'[^\w\-]', '_', str(tdoc_id).strip())
    clean = re.sub(r'_+', '_', clean).strip("._ ")
    return f"{clean}_" if clean else ""


@dataclass
class ExtractedSlide:
    """Representation of an extracted PowerPoint slide."""
    index: int
    title: str = ""
    body_paragraphs: List[str] = field(default_factory=list)
    tables: List[List[List[str]]] = field(default_factory=list)
    speaker_notes: str = ""
    image_filename: Optional[str] = None


def _safe_log(msg: str, log_fn: Optional[Callable[[str], None]] = None) -> None:
    """Invokes logger safely handling character encoding mismatches on Windows."""
    if not log_fn:
        return
    try:
        log_fn(msg)
    except Exception:
        try:
            clean_msg = msg.encode("ascii", errors="replace").decode("ascii")
            log_fn(clean_msg)
        except Exception:
            pass


def _format_markdown_table(rows: List[List[str]]) -> str:
    """Renders a 2D matrix of text cells as a GitHub Flavored Markdown table."""
    if not rows:
        return ""
    col_count = max(len(r) for r in rows)
    if col_count == 0:
        return ""

    norm_rows = []
    for row in rows:
        cells = [c.replace("\r", " ").replace("\n", " ").replace("|", "\\|").strip() for c in row]
        if len(cells) < col_count:
            cells.extend([""] * (col_count - len(cells)))
        norm_rows.append(cells)

    header = norm_rows[0]
    separator = [":---"] * col_count
    body = norm_rows[1:]

    lines = [
        "| " + " | ".join(header) + " |",
        "| " + " | ".join(separator) + " |",
    ]
    for row in body:
        lines.append("| " + " | ".join(row) + " |")
    return "\n".join(lines)
