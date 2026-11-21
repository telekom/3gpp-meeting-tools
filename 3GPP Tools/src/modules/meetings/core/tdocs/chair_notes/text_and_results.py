"""Regex constants, text normalization, paragraph color detection, and Result field parsing."""

from __future__ import annotations

import re
from typing import Dict, List, Tuple

from docx.document import Document as _Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table, _Cell
from docx.text.paragraph import Paragraph


# Private Use Area characters (Wingdings bullets like \uf0d8, \uf0b7, \uf0a7)
_PUA_RE = re.compile(r'[\uf000-\uf8ff]')

# Trivial 1-word TDoc disposition statuses that are NOT substantive technical agreements on their own
_TRIVIAL_STATUS_SET = {
    "noted",
    "approved",
    "agreed",
    "endorsed",
    "postponed",
    "withdrawn",
    "revised",
    "not treated",
    "not handled",
    "not pursued",
    "no agreement",
    "tbd",
    "ffs",
    "handled",
    "merged",
    "seen",
}

_TDOC_ID_RE = re.compile(
    r'\b([A-Z][A-Z0-9]{1,3}-\d{4,8}(?:[rR]\d{1,2}[a-zA-Z]?)?)\b'
)

# Matches RAN3 WID/SID bracket tags: e.g., `SID [FS_6G_Radio]: RP-261559` or `WID [NR_AIML_NGRAN_Ph3-Core]: RP-260773`
_RAN3_WID_BRACKET_RE = re.compile(
    r'(?:WID|SID|WI|SI)\s*\[\s*([A-Za-z0-9_\-\/]+)\s*\]',
    re.IGNORECASE,
)

# Matches RAN2 parenthetical WI metadata under headings:
# e.g. `(NR_AIML_air_Ph2, leading WG: RAN1; REL-20; WID: RP-260511)`
# or `(FS_NR_ish_redcap, leading WG: RAN2, REL-19, SID: RP-240899)`
_RAN2_PAREN_WI_RE = re.compile(
    r'^\(\s*([A-Za-z0-9]+_[A-Za-z0-9_\-]+)\s*[,;)]'
)

# Matches Release tags like REL-19, Rel-20, Release 19
_RELEASE_RE = re.compile(r'\b(?:REL|Release)[-_\s]*(\d{1,2})\b', re.IGNORECASE)

# Green hex colors used in RAN3 Chair Notes for Agreements / Working Assumptions
_GREEN_HEX_PREFIXES = ("008000", "00B050", "538135", "385723", "006600", "009900", "2E7D32")
# Blue hex colors used in RAN3 Chair Notes for FFS / To be continued
_BLUE_HEX_PREFIXES = ("0000FF", "3333FF", "0070C0", "002060", "2F5597")

# Cache format version for chair_notes_cache.json; bump when parser semantics change to force automatic re-scanning
CHAIR_NOTES_CACHE_VERSION = 2



def clean_chair_notes_text(text: str, preserve_newlines: bool = True) -> str:
    """Strips Wingdings Private Use Area bullets and normalizes whitespace."""
    if not text:
        return ""
    s = _PUA_RE.sub("", str(text))
    s = s.replace("\r\n", "\n").replace("\r", "\n")
    if preserve_newlines:
        lines = []
        for line in s.split("\n"):
            cl = re.sub(r'[ \t]+', ' ', line).strip()
            if cl:
                lines.append(cl)
        return "\n".join(lines).strip()
    return re.sub(r'\s+', ' ', s).strip()


def iter_block_items(parent):
    """
    Yields each paragraph and table child within `parent` (Document or _Cell)
    in true document order.
    """
    if isinstance(parent, _Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        raise ValueError(f"Unsupported parent type: {type(parent)}")

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


def _get_paragraph_dominant_color(p: Paragraph) -> str:
    """
    Inspects OpenXML run colors (`w:rPr/w:color/@w:val`) in paragraph `p`.
    Returns:
      - `"GREEN"` if green runs dominate the non-whitespace text,
      - `"BLUE"` if blue runs dominate,
      - `""` otherwise.
    """
    green_chars = 0
    blue_chars = 0
    other_chars = 0

    for run in p.runs:
        txt = _PUA_RE.sub("", run.text or "").strip()
        if not txt:
            continue
        n = len(txt)
        color_hex = ""
        try:
            if run.font and run.font.color and run.font.color.rgb:
                color_hex = str(run.font.color.rgb).upper()
        except Exception:
            color_hex = ""

        if not color_hex:
            # Check XML directly in case python-docx misses theme/val
            try:
                rPr = run._r.rPr
                if rPr is not None:
                    c_elms = rPr.findall("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}color")
                    if c_elms:
                        color_hex = (c_elms[0].get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val") or "").upper()
            except Exception:
                pass

        if color_hex in _GREEN_HEX_PREFIXES:
            green_chars += n
        elif color_hex in _BLUE_HEX_PREFIXES:
            blue_chars += n
        else:
            other_chars += n

    total = green_chars + blue_chars + other_chars
    if total == 0:
        return ""
    if green_chars > 0 and green_chars >= (blue_chars + other_chars):
        return "GREEN"
    if blue_chars > 0 and blue_chars >= (green_chars + other_chars):
        return "BLUE"
    # Also handle lines where a company prefix or bullet is black but the body is green
    if green_chars >= 15 and green_chars >= total * 0.45:
        return "GREEN"
    if blue_chars >= 15 and blue_chars >= total * 0.45:
        return "BLUE"
    return ""


def _classify_agreement_kind(text: str, default_kind: str = "Agreement") -> str:
    """Classifies whether a text block is an Agreement, Working Assumption, Conclusion, or FFS."""
    low = text.strip().lower()
    if low.startswith(("working assumption", "wa:", "wa ")):
        return "Working Assumption"
    if low.startswith(("conclusion", "conclusions")):
        return "Conclusion"
    if low.startswith(("ffs", "open issue", "to be continued", "for further study")):
        return "FFS / Open Issue"
    return default_kind


def _resolve_inherited_field(ai_num: str, mapping: Dict[str, str]) -> str:
    """
    Looks up `ai_num` in `mapping` (e.g. `10.2.1`), falling back up the parent hierarchy
    (`10.2` -> `10`) when child sub-items inherit their parent Agenda Item's WI/SI or Title.
    """
    if not ai_num or ai_num == "N/A":
        return ""
    if ai_num in mapping and mapping[ai_num]:
        return mapping[ai_num]
    parts = ai_num.split(".")
    for i in range(len(parts) - 1, 0, -1):
        prefix = ".".join(parts[:i])
        if prefix in mapping and mapping[prefix]:
            return mapping[prefix]
    return ""


_CHAIR_RESULT_PREFIXES: List[Tuple[re.Pattern, str]] = [
    (re.compile(r'^(?:not\s+handled|unhandled)\b', re.IGNORECASE), "Not Handled"),
    (re.compile(r'^not\s+treated\b', re.IGNORECASE), "Not Treated"),
    (re.compile(r'^not\s+pursued\b', re.IGNORECASE), "Not Pursued"),
    (re.compile(r'^not\s+agreed\b', re.IGNORECASE), "Not Agreed"),
    (re.compile(r'^not\s+approved\b', re.IGNORECASE), "Not Approved"),
    (re.compile(r'^not\s+noted\b', re.IGNORECASE), "Not Noted"),
    (re.compile(r'^not\s+endorsed\b', re.IGNORECASE), "Not Endorsed"),
    (re.compile(r'^(?:for\s+e-?mail\s+approval|e-?mail\s+approval)\b', re.IGNORECASE), "For e-mail approval"),
    (re.compile(r'^parallel\s+agreed\b', re.IGNORECASE), "Parallel Agreed"),
    (re.compile(r'^replied(?:\s+to)?\b', re.IGNORECASE), "Replied to"),
    (re.compile(r'^(?:block\s+approved?|approved)\b', re.IGNORECASE), "Approved"),
    (re.compile(r'^agreed\b', re.IGNORECASE), "Agreed"),
    (re.compile(r'^endorsed\b', re.IGNORECASE), "Endorsed"),
    (re.compile(r'^noted\b', re.IGNORECASE), "Noted"),
    (re.compile(r'^postponed\b', re.IGNORECASE), "Postponed"),
    (re.compile(r'^withdrawn\b', re.IGNORECASE), "Withdrawn"),
    (re.compile(r'^revised\b', re.IGNORECASE), "Revised"),
    (re.compile(r'^merged\b', re.IGNORECASE), "Merged"),
    (re.compile(r'^handled\b', re.IGNORECASE), "Handled"),
    (re.compile(r'^available\b', re.IGNORECASE), "Available"),
    (re.compile(r'^reserved\b', re.IGNORECASE), "Reserved"),
    (re.compile(r'^rejected\b', re.IGNORECASE), "Rejected"),
    (re.compile(r'^open(?:ed)?\b', re.IGNORECASE), "Open"),
]

_PREPOSITION_CONTINUATION_RE = re.compile(
    r'^(?:to|in|into|for|until|till|at|by|on|with|due|as|after|before|pending|subject)\b',
    re.IGNORECASE,
)


def parse_chair_result(raw_result: str) -> Tuple[str, str, List[str]]:
    """
    Normalizes a raw Chair's Notes / TdocsByAgenda `Result` cell into:
      `(clean_status, extra_remark, revised_to_tdocs)`

    Separates canonical 3GPP meeting statuses from embedded revision targets
    (`Revised to S2-...`), merge targets (`Merged into S2-...`), multi-line
    action notes (`OPEN\\nDiscuss if agree...`), tentative questions (`Block Approve?`),
    and free-text chair remarks (`Revision assigned after offline discussion, not opened yet.`).
    """
    raw = str(raw_result or "").strip()
    if not raw or raw.lower() in {"-", "—", "none", "nan", "n/a"}:
        return "", "", []

    # Preserve RAN Agreement / Working Assumption / FFS visual status pills
    if any(sym in raw for sym in ("🟢", "🟡", "🔵")):
        return raw, "", []

    lines = [
        re.sub(r'[ \t]+', ' ', ln).strip()
        for ln in raw.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    ]
    lines = [ln for ln in lines if ln]
    if not lines:
        return "", "", []

    first_line = lines[0]
    rest_lines = "\n".join(lines[1:]).strip()
    normalized_full = "\n".join(lines)

    # "Block Approve?" / "Block Approve" / "Block Approved" denotes bulk plenary approval
    if re.match(r'^block\s+approv(?:e|ed)\b', first_line, re.IGNORECASE):
        extra = normalized_full if "?" in first_line or rest_lines else ""
        return "Approved", extra, []

    # Tentative question/reminder in Result (e.g. "Agreed?")
    if first_line.endswith("?"):
        return "", normalized_full, []

    # 1. Revision pointer ("Revised to S2-2612689", "Rev in R3-264569 Agreed", etc.)
    m_rev = re.match(r'^rev(?:ised)?\.?\s+(?:to|in|into|as|by)\s+(.+)$', first_line, re.IGNORECASE)
    if m_rev:
        body = m_rev.group(1).strip()
        rev_matches = list(_TDOC_ID_RE.finditer(body))
        rev_targets = list(dict.fromkeys(m.group(1).upper() for m in rev_matches))
        after_tdoc = body[rev_matches[-1].end():].strip(" \t.,:;-–—") if rev_matches else body

        # RAN3 style: "Rev in R3-264569 Agreed" -> final outcome Agreed + Rev target
        if rev_matches and after_tdoc.lower() in {
            "agreed", "approved", "endorsed", "noted", "postponed", "withdrawn"
        }:
            return after_tdoc.capitalize(), normalized_full, rev_targets

        remark_parts = []
        if not rev_targets or after_tdoc:
            remark_parts.append(first_line)
        if rest_lines:
            remark_parts.append(rest_lines)
        return "Revised", "\n".join(remark_parts).strip(), rev_targets

    # 2. Merge pointer ("Merged into S2-2612833", "Merge into S2-...", etc.)
    if re.match(r'^merg(?:ed|e)\s+(?:into|with|in|to)\s+(.+)$', first_line, re.IGNORECASE):
        return "Merged", normalized_full, []

    # 3. Canonical status prefix (with optional trailing comment / continuation lines)
    for pat, canonical in _CHAIR_RESULT_PREFIXES:
        m = pat.match(first_line)
        if m:
            trailing_raw = first_line[m.end():]
            trailing_clean = trailing_raw.strip(" \t.:;-–—")
            if not trailing_clean:
                extra_remark = rest_lines
            elif trailing_raw.startswith((" ", "\t")) and _PREPOSITION_CONTINUATION_RE.match(trailing_clean):
                extra_remark = normalized_full
            else:
                extra_remark = f"{trailing_clean}\n{rest_lines}".strip() if rest_lines else trailing_clean
            return canonical, extra_remark, []

    # 4. RAN2 sentence-style statuses ("The CR is postponed", "The LS is approved")
    m_sent = re.match(
        r'^(?:the\s+)?(?:cr|p-?cr|ls|report|tp|tdoc|doc)\s+(?:is|was)\s+'
        r'(agreed|approved|endorsed|noted|postponed|withdrawn)\b',
        first_line,
        re.IGNORECASE,
    )
    if m_sent:
        return m_sent.group(1).capitalize(), normalized_full, []

    # 5. Free-text chair remarks containing an explicit unhandled / not-opened state
    low_full = " ".join(lines).lower()
    if any(p in low_full for p in ("not opened", "not handled", "unhandled")):
        return "Not Handled", normalized_full, []
    if "not treated" in low_full:
        return "Not Treated", normalized_full, []

    # 6. Unrecognized free-text remark: keep out of TDoc Status and move to Secretary Remarks
    return "", normalized_full, []
