"""Sanitizer and HTML-to-Markdown converter for Microsoft 365 Copilot paste-backs."""

from __future__ import annotations

from html.parser import HTMLParser
import re
from typing import List, Optional, Tuple

_ZERO_WIDTH_RE = re.compile(r"[\u200b\u2060\ufeff]+")

# Leading leaked CSS rules from M365 Copilot HTML clipboard (e.g. `a { text-decoration: none; ... }`)
_LEADING_CSS_BLOCK_RE = re.compile(
    r"^(?:\s*[A-Za-z0-9_.*#,:>\s+-]+\s*\{[^{}]*:[^{}]*\}\s*)+"
)

# Bracketed SharePoint / Copilot citation chips appended to every paragraph
_SP_CITATION_BADGE_RE = re.compile(
    r"[ \t]*\[\\?\[[^\]]*(?:sharepoint\.com|epoint\.com|copilot)[^\]]*\\?\]\]\(https?://[^)\s]+\)",
    re.IGNORECASE,
)

# SharePoint links pointing to uploaded Copilot handoff packs or citation endpoints
_SP_FILE_LINK_RE = re.compile(
    r"\[([^\]\n]+)\]\(https?://[^)\s]*(?:sharepoint\.com|Microsoft%20Copilot%20Chat%20Files)[^)\s]*\)",
    re.IGNORECASE,
)

# Inline Markdown heading glued directly to prior sentence (e.g. `positions.### 1. 🧭 ...`)
_GLUED_HEADING_RE = re.compile(r"([^\n#])[ \t]*(#{1,6}[ \t]+\S)")

# Standalone numbered-emoji section header that lost its `### ` prefix on plain-text copy
_NUMBERED_EMOJI_SECTION_RE = re.compile(
    r"^[ \t]*(\d+\.\s+[\U0001F300-\U0001FAFF\u2600-\u27BF][\uFE0F\u200D\U0001F300-\U0001FAFF\u2600-\u27BF]*\s+\S[^\n]*)$"
)

# Bare protocol URIs left unlinked when Copilot strips Markdown links in plain text
_BARE_TDOC_URI_RE = re.compile(
    r"(?<!\()(?<!\[)\btdoc:([A-Za-z0-9]+-[0-9]{4,}[A-Za-z0-9._-]*)",
    re.IGNORECASE,
)
_BARE_JUMP_URI_RE = re.compile(r"(?<!\()(?<!\[)\bjump://msg/(\d+)\b", re.IGNORECASE)

_LIST_LINE_RE = re.compile(r"^[ \t]*(?:[-*+]\s+|\d+\.\s+)")
_HEADING_LINE_RE = re.compile(r"^[ \t]*#{1,6}\s+")
_FENCE_LINE_RE = re.compile(r"^[ \t]*```")

_COPILOT_ARTIFACT_HINT_RE = re.compile(
    r"(?:a\s*\{\s*text-decoration:|sharepoint\.com|Microsoft%20Copilot%20Chat%20Files"
    r"|[^\n#]#{2,4}\s+\d+\.|(?<!\()(?<!\[)\b(?:tdoc:[A-Za-z0-9]+-\d{4,}|jump://msg/\d+))",
    re.IGNORECASE,
)


def _is_sharepoint_citation_url(href: str) -> bool:
    low = (href or "").lower()
    return "sharepoint.com" in low or "microsoft%20copilot%20chat%20files" in low or "microsoft copilot chat files" in low


def _replace_sharepoint_link(match: re.Match) -> str:
    label = match.group(1).strip()
    unwrapped = label.replace("\\[", "[").replace("\\]", "]").strip("[] ").strip()
    low = unwrapped.lower()
    if low.endswith(".md"):
        return f"`{unwrapped}`"
    if not unwrapped or "epoint.com" in low or "sharepoint.com" in low or unwrapped.isdigit():
        return ""
    return unwrapped


class _CopilotHtmlToMarkdownParser(HTMLParser):
    """Lightweight stdlib HTML-to-GFM converter tailored to M365 Copilot clipboard HTML."""

    _SKIP_TAGS = frozenset({"style", "script", "head", "title", "meta", "link", "noscript"})

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self._out: List[str] = []
        self._skip_depth = 0
        self._list_stack: List[Tuple[str, int]] = []
        self._in_pre = False
        self._in_code = False
        self._anchor_href: Optional[str] = None
        self._anchor_buf: List[str] = []
        self._in_table = False
        self._table_rows: List[Tuple[bool, List[str]]] = []
        self._cur_row: List[str] = []
        self._cur_row_is_hdr = False
        self._in_cell = False
        self._cell_buf: List[str] = []

    def _emit(self, text: str) -> None:
        if self._anchor_href is not None:
            self._anchor_buf.append(text)
        elif self._in_cell:
            self._cell_buf.append(text)
        else:
            self._out.append(text)

    def handle_starttag(self, tag: str, attrs: List[Tuple[str, Optional[str]]]) -> None:
        t = tag.lower()
        if t in self._SKIP_TAGS:
            self._skip_depth += 1
            return
        if self._skip_depth > 0:
            return
        attr_map = {k.lower(): (v or "") for k, v in attrs}

        if t in ("h1", "h2", "h3", "h4", "h5", "h6"):
            level = max(2, min(4, int(t[1])))
            self._emit(f"\n\n{'#' * level} ")
        elif t == "p":
            if not self._in_cell:
                self._emit("\n\n")
        elif t == "br":
            self._emit(" " if self._in_cell else "\n")
        elif t == "hr":
            self._emit("\n\n---\n\n")
        elif t in ("strong", "b"):
            self._emit("**")
        elif t in ("em", "i"):
            self._emit("*")
        elif t == "pre":
            self._in_pre = True
            self._emit("\n\n```\n")
        elif t == "code":
            if not self._in_pre:
                self._in_code = True
                self._emit("`")
        elif t == "blockquote":
            self._emit("\n\n> ")
        elif t == "ul":
            self._list_stack.append(("ul", 0))
            if len(self._list_stack) == 1:
                self._emit("\n")
        elif t == "ol":
            self._list_stack.append(("ol", 0))
            if len(self._list_stack) == 1:
                self._emit("\n")
        elif t == "li":
            indent = "  " * max(0, len(self._list_stack) - 1)
            if self._list_stack and self._list_stack[-1][0] == "ol":
                kind, num = self._list_stack[-1]
                self._list_stack[-1] = (kind, num + 1)
                self._emit(f"\n{indent}{num + 1}. ")
            else:
                self._emit(f"\n{indent}- ")
        elif t == "a":
            self._anchor_href = attr_map.get("href", "").strip()
            self._anchor_buf = []
        elif t == "table":
            self._in_table = True
            self._table_rows = []
        elif t == "tr" and self._in_table:
            self._cur_row = []
            self._cur_row_is_hdr = False
        elif t in ("th", "td") and self._in_table:
            if t == "th":
                self._cur_row_is_hdr = True
            self._in_cell = True
            self._cell_buf = []

    def handle_endtag(self, tag: str) -> None:
        t = tag.lower()
        if t in self._SKIP_TAGS:
            self._skip_depth = max(0, self._skip_depth - 1)
            return
        if self._skip_depth > 0:
            return

        if t in ("h1", "h2", "h3", "h4", "h5", "h6"):
            self._emit("\n\n")
        elif t in ("p", "blockquote"):
            if not self._in_cell:
                self._emit("\n\n")
        elif t in ("strong", "b"):
            self._emit("**")
        elif t in ("em", "i"):
            self._emit("*")
        elif t == "pre":
            self._in_pre = False
            self._emit("\n```\n\n")
        elif t == "code":
            if self._in_code:
                self._in_code = False
                self._emit("`")
        elif t in ("ul", "ol"):
            if self._list_stack:
                self._list_stack.pop()
            if not self._list_stack:
                self._emit("\n\n")
        elif t == "a" and self._anchor_href is not None:
            href = self._anchor_href
            label = re.sub(r"\s+", " ", "".join(self._anchor_buf)).strip()
            self._anchor_href = None
            self._anchor_buf = []
            if _is_sharepoint_citation_url(href):
                unwrapped = label.strip("[] ").strip()
                if unwrapped.lower().endswith(".md"):
                    self._emit(f"`{unwrapped}`")
            elif label and href:
                self._emit(f"[{label}]({href})")
            elif label:
                self._emit(label)
        elif t in ("th", "td") and self._in_cell:
            cell_txt = re.sub(r"\s+", " ", "".join(self._cell_buf)).strip().replace("|", "\\|")
            self._cur_row.append(cell_txt)
            self._in_cell = False
            self._cell_buf = []
        elif t == "tr" and self._in_table:
            if self._cur_row:
                self._table_rows.append((self._cur_row_is_hdr, self._cur_row))
            self._cur_row = []
        elif t == "table" and self._in_table:
            self._in_table = False
            self._flush_table()

    def handle_data(self, data: str) -> None:
        if self._skip_depth > 0 or not data:
            return
        if self._in_pre:
            self._emit(data)
            return
        cleaned = re.sub(r"[ \t\r\n]+", " ", data)
        if cleaned == " " and (not self._out or self._out[-1].endswith(("\n", " "))):
            return
        self._emit(cleaned)

    def _flush_table(self) -> None:
        if not self._table_rows:
            return
        col_count = max(len(r) for _, r in self._table_rows)
        if col_count == 0:
            return
        lines: List[str] = ["\n"]
        for idx, (_is_hdr, row) in enumerate(self._table_rows):
            padded = row + [""] * (col_count - len(row))
            lines.append("| " + " | ".join(padded) + " |")
            if idx == 0:
                lines.append("| " + " | ".join(["---"] * col_count) + " |")
        lines.append("\n")
        self._out.append("\n".join(lines))

    def get_markdown(self) -> str:
        return "".join(self._out)


def copilot_html_to_markdown(html_text: str) -> str:
    """Convert rich HTML copied from Microsoft 365 Copilot into clean GFM Markdown."""
    raw = str(html_text or "").strip()
    if not raw:
        return ""
    parser = _CopilotHtmlToMarkdownParser()
    parser.feed(raw)
    parser.close()
    return sanitize_copilot_markdown(parser.get_markdown())


def _is_plain_prose_line(line: str) -> bool:
    s = line.strip()
    if not s:
        return False
    if _HEADING_LINE_RE.match(s) or _LIST_LINE_RE.match(s):
        return False
    if s.startswith(("|", ">", "---", "***", "```")):
        return False
    return True


def sanitize_copilot_markdown(raw_text: str) -> str:
    """Clean leaked CSS, SharePoint citation badges, glued headings, bare URIs, and cramped paragraphs."""
    text = _ZERO_WIDTH_RE.sub("", str(raw_text or ""))
    text = text.replace("\r\n", "\n").replace("\r", "\n").strip()
    if not text:
        return ""

    # 1. Strip leaked leading CSS rule blocks (e.g. `a { text-decoration: none; ... } tr th ...`)
    text = _LEADING_CSS_BLOCK_RE.sub("", text).lstrip()

    # 2. Strip SharePoint / Copilot citation chips and unwrap uploaded *_copilot.md links
    text = _SP_CITATION_BADGE_RE.sub("", text)
    text = _SP_FILE_LINK_RE.sub(_replace_sharepoint_link, text)

    # 3. Process line-by-line outside fenced code blocks
    raw_lines = text.split("\n")
    expanded_lines: List[str] = []
    in_fence = False
    for line in raw_lines:
        if _FENCE_LINE_RE.match(line):
            in_fence = not in_fence
            expanded_lines.append(line.rstrip())
            continue
        if in_fence:
            expanded_lines.append(line.rstrip())
            continue

        # Split inline glued headings (e.g. `positions.### 1. 🧭 ...`)
        unglued = _GLUED_HEADING_RE.sub(r"\1\n\n\2", line)
        for sub in unglued.split("\n"):
            sub_clean = re.sub(r"[ \t]{2,}", " ", sub).rstrip()
            m_sec = _NUMBERED_EMOJI_SECTION_RE.match(sub_clean)
            if m_sec and not _HEADING_LINE_RE.match(sub_clean):
                sub_clean = f"### {m_sec.group(1).strip()}"
            sub_clean = _BARE_TDOC_URI_RE.sub(lambda m: f"[{m.group(1).upper()}](tdoc:{m.group(1).upper()})", sub_clean)
            sub_clean = _BARE_JUMP_URI_RE.sub(lambda m: f"[Msg #{m.group(1)}](jump://msg/{m.group(1)})", sub_clean)
            expanded_lines.append(sub_clean)

    # 4. Ensure clean blank lines around headings and between single-newline prose paragraphs
    out_lines: List[str] = []
    in_fence = False
    for line in expanded_lines:
        if _FENCE_LINE_RE.match(line):
            in_fence = not in_fence
            out_lines.append(line)
            continue
        if in_fence:
            out_lines.append(line)
            continue

        s = line.strip()
        if not s:
            if out_lines and out_lines[-1] != "":
                out_lines.append("")
            continue

        prev = out_lines[-1] if out_lines else ""
        if _HEADING_LINE_RE.match(s):
            if prev:
                out_lines.append("")
            out_lines.append(s)
            out_lines.append("")
            continue

        if (
            prev
            and _is_plain_prose_line(prev)
            and _is_plain_prose_line(s)
            and not line.startswith(("  ", "\t"))
            and prev.rstrip()[-1] in ".!?)]`"
        ):
            out_lines.append("")

        out_lines.append(line)

    result = "\n".join(out_lines)
    result = re.sub(r"\n{3,}", "\n\n", result).strip()
    return result


def normalize_summary_for_display(summary_md: str, provider: str = "") -> str:
    """Normalize M365 Copilot summaries (including already-saved records) before Markdown rendering."""
    raw = str(summary_md or "")
    if not raw.strip():
        return raw
    prov_norm = str(provider or "").strip().lower().replace("_", "-")
    if prov_norm == "m365-copilot" or _COPILOT_ARTIFACT_HINT_RE.search(raw):
        return sanitize_copilot_markdown(raw)
    return raw
