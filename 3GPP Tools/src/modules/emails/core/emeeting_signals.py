# --- File: src/modules/emails/core/emeeting_signals.py ---
"""
Pure-Python e-Meeting signal engine:
- Extracts formal 3GPP e-Meeting `<<START>> ... <<END>>` template blocks.
- Detects and normalizes revision mentions (`r01`, `rev 2`, `S2-2601234r02`, `/Drafts/` links).
- Compares revision numbers for smart unparking when a new revision arrives after an objection.
- Detects 3GPP objection phrases in direct (non-quoted) email bodies.
"""

from __future__ import annotations

import html
import re
from typing import Any, Dict, List, Optional, Sequence, Tuple

from core.ui.theme import C

START_END_BLOCK_RE = re.compile(
    r"<<\s*START\s*>>(.*?)<<\s*END\s*>>",
    re.DOTALL | re.IGNORECASE,
)

FULL_TDOC_REV_RE = re.compile(
    r"\b((?:S[1-6]|R[1-5]|C[1-6]|SP|RP|CP)-\d{5,8})(?:r|rev\.?\s*)0?([1-9]\d*)\b",
    re.IGNORECASE,
)

SHORTHAND_REV_RE = re.compile(
    r"\b(?:r|rev\.?\s*)0?([1-9]\d*)\b",
    re.IGNORECASE,
)

DRAFTS_DOC_REV_RE = re.compile(
    r"(?:Drafts|Revisions|INBOX)[^\s\"'<>]*?\b((?:S[1-6]|R[1-5]|C[1-6]|SP|RP|CP)-\d{5,8})[ _-]*(?:r|rev)?0?([1-9]\d*)",
    re.IGNORECASE,
)

BASE_TDOC_EXTRACT_RE = re.compile(
    r"^((?:S[1-6]|R[1-5]|C[1-6]|SP|RP|CP)-\d{5,8})",
    re.IGNORECASE,
)

# High-precision 3GPP reflector objection / non-consensus phrases (evaluated on direct body only)
OBJECTION_PATTERNS: Tuple[re.Pattern, ...] = (
    re.compile(r"\b(?:we|i)\s+(?:do\s+not|don't|cannot|can't)\s+(?:support|agree|accept)\b", re.IGNORECASE),
    re.compile(r"\b(?:we|i)\s+(?:still\s+)?object(?:\s+to)?\b", re.IGNORECASE),
    re.compile(r"\bnot\s+ready\s+for\s+(?:agreement|approval|conclusion|endorsement)\b", re.IGNORECASE),
    re.compile(r"\bcannot\s+be\s+(?:agreed|approved|accepted|endorsed)\b", re.IGNORECASE),
    re.compile(r"\b(?:strong(?:ly)?\s+object|formal\s+objection|maintain\s+our\s+objection)\b", re.IGNORECASE),
    re.compile(r"\b(?:we|i)\s+have\s+(?:strong\s+|serious\s+|major\s+)?concerns?\s+(?:with|on|about|regarding)\b", re.IGNORECASE),
    re.compile(r"\bprefer\s+not\s+to\s+(?:agree|approve|proceed)\b", re.IGNORECASE),
)


def extract_base_tdoc_id(raw_tdoc: Any) -> str:
    """Returns uppercase base TDoc ID (`S2-2601234` from `S2-2601234r02`)."""
    txt = str(raw_tdoc or "").strip().upper()
    if not txt:
        return ""
    m = BASE_TDOC_EXTRACT_RE.match(txt)
    return m.group(1).upper() if m else txt


def extract_start_end_blocks(body_text: str) -> Tuple[List[str], str]:
    """
    Splits `body_text` into `(formal_blocks, remaining_free_text)`.
    Supports one or multiple `<<START>> ... <<END>>` blocks in a single email.
    """
    text = str(body_text or "")
    if not text or "<<" not in text:
        return [], text.strip()
    blocks = [m.group(1).strip() for m in START_END_BLOCK_RE.finditer(text) if m.group(1).strip()]
    if not blocks:
        return [], text.strip()
    remaining = START_END_BLOCK_RE.sub("", text).strip()
    return blocks, remaining


def parse_rev_number(rev_token: Any) -> int:
    """Converts `'r02'`, `'S2-2601234r02'`, `'2'`, or `''` into an integer revision (`0` if none)."""
    txt = str(rev_token or "").strip()
    if not txt:
        return 0
    m_full = FULL_TDOC_REV_RE.search(txt)
    if m_full:
        return int(m_full.group(2))
    m_short = SHORTHAND_REV_RE.search(txt)
    if m_short:
        return int(m_short.group(1))
    if txt.isdigit():
        return int(txt)
    return 0


def format_rev_tag(rev_num: int) -> str:
    """Formats a positive revision integer as `'r01'`, `'r02'`, etc. Returns `''` for `<= 0`."""
    return f"r{int(rev_num):02d}" if int(rev_num or 0) > 0 else ""


def extract_revisions_from_text(
    text: str,
    primary_tdoc: str = "",
    attachments: Optional[Sequence[str]] = None,
) -> Tuple[int, List[str]]:
    """
    Scans `text` (and attachment names) for explicit TDoc revisions (`S2-2601234r02`)
    and shorthand revisions (`r01`, `rev 2`).
    Returns `(max_rev_num, normalized_rev_tokens)`.
    """
    body = str(text or "")
    base_primary = extract_base_tdoc_id(primary_tdoc)
    rev_nums: List[int] = []
    tokens: List[str] = []

    def _add(base: str, num: int):
        if num <= 0 or num > 99:
            return
        rev_nums.append(num)
        tok = f"{base}r{num:02d}" if base else f"r{num:02d}"
        if tok not in tokens:
            tokens.append(tok)

    for m in FULL_TDOC_REV_RE.finditer(body):
        b = m.group(1).upper()
        n = int(m.group(2))
        if not base_primary or b == base_primary:
            _add(b, n)

    for m in DRAFTS_DOC_REV_RE.finditer(body):
        b = m.group(1).upper()
        n = int(m.group(2))
        if not base_primary or b == base_primary:
            _add(b, n)

    for att in (attachments or ()):
        att_s = str(att or "")
        for m in FULL_TDOC_REV_RE.finditer(att_s):
            b = m.group(1).upper()
            n = int(m.group(2))
            if not base_primary or b == base_primary:
                _add(b, n)
        for m in SHORTHAND_REV_RE.finditer(att_s):
            _add(base_primary, int(m.group(1)))

    # Shorthand revisions (`r01`, `rev 02`) inside <<START>>..<<END>> or direct body when primary_tdoc is known
    blocks, _ = extract_start_end_blocks(body)
    scan_targets = blocks if blocks else ([body] if base_primary else [])
    for chunk in scan_targets:
        for m in SHORTHAND_REV_RE.finditer(chunk):
            _add(base_primary, int(m.group(1)))

    max_rev = max(rev_nums) if rev_nums else 0
    return max_rev, tokens


def detect_objection_in_text(direct_body: str) -> Tuple[bool, str]:
    """
    Checks direct (non-quoted) message body for 3GPP objection / non-agreement phrases.
    Returns `(has_objection, matched_snippet)`.
    """
    text = str(direct_body or "").strip()
    if not text:
        return False, ""
    # Strip trailing quote lines starting with '>'
    clean_lines = [
        ln.strip()
        for ln in text.splitlines()
        if ln.strip() and not ln.lstrip().startswith(">")
    ]
    clean_text = "\n".join(clean_lines)
    for pat in OBJECTION_PATTERNS:
        m = pat.search(clean_text)
        if m:
            # Extract surrounding line as tooltip context
            start = clean_text.rfind("\n", 0, m.start()) + 1
            end = clean_text.find("\n", m.end())
            if end == -1:
                end = len(clean_text)
            snippet = clean_text[start:end].strip()
            if len(snippet) > 140:
                snippet = snippet[:137] + "..."
            return True, snippet
    return False, ""


def analyze_message_signals(
    msg: Dict[str, Any],
    primary_tdoc: str = "",
) -> Dict[str, Any]:
    """
    Computes e-Meeting signals for a single message dict:
    - `has_template`: True if `<<START>>...<<END>>` block is present
    - `template_blocks`: list of extracted formal blocks
    - `max_rev_num`: highest revision integer mentioned (0 if none)
    - `max_rev_tag`: `'r01'`, `'r02'`, etc.
    - `rev_tokens`: full list of revision tokens mentioned
    - `has_objection`: True if direct body matches objection keywords
    - `objection_snippet`: matched line preview
    """
    raw_body = str(msg.get("clean_body") or msg.get("body_text") or msg.get("raw_body_text") or "")
    subj = str(msg.get("subject") or "")
    linked = msg.get("linked_tdocs") or []
    base_tdoc = extract_base_tdoc_id(primary_tdoc or (linked[0] if linked else ""))

    blocks, _ = extract_start_end_blocks(raw_body)
    max_rev, rev_tokens = extract_revisions_from_text(
        f"{subj}\n{raw_body}",
        primary_tdoc=base_tdoc,
        attachments=msg.get("attachments") or [],
    )
    has_obj, obj_snip = detect_objection_in_text(raw_body)
    return {
        "has_template": bool(blocks),
        "template_blocks": blocks,
        "max_rev_num": max_rev,
        "max_rev_tag": format_rev_tag(max_rev),
        "rev_tokens": rev_tokens,
        "has_objection": has_obj,
        "objection_snippet": obj_snip,
    }


def compute_group_signals(
    messages: Sequence[Dict[str, Any]],
    *,
    primary_tdoc: str = "",
    is_objected: bool = False,
    objected_at_rev: str = "",
) -> Dict[str, Any]:
    """
    Rolls up e-Meeting signals across a group of messages (a Thread or a Base TDoc)
    and determines smart-parking / unparking status:
    - If `is_objected=True` at `objected_at_rev='r01'` and `max_rev_num >= 2` (`r02`),
      `unparked_by_new_rev=True`!
    """
    max_rev = 0
    unread_max_rev = 0
    all_rev_tokens: List[str] = []
    has_template = False
    unread_has_template = False
    has_objection = False
    unread_has_objection = False
    objection_snippets: List[str] = []

    for m in messages:
        sig = m.get("_signals") or analyze_message_signals(m, primary_tdoc=primary_tdoc)
        m["_signals"] = sig
        is_unread = not bool(m.get("is_read", True))
        if sig["has_template"]:
            has_template = True
            if is_unread:
                unread_has_template = True
        if sig["max_rev_num"] > max_rev:
            max_rev = sig["max_rev_num"]
        if is_unread and sig["max_rev_num"] > unread_max_rev:
            unread_max_rev = sig["max_rev_num"]
        for tok in sig["rev_tokens"]:
            if tok not in all_rev_tokens:
                all_rev_tokens.append(tok)
        if sig["has_objection"]:
            has_objection = True
            if is_unread:
                unread_has_objection = True
            sender = str(m.get("company") or m.get("sender_name") or "Sender").strip()
            if sig["objection_snippet"]:
                objection_snippets.append(f"{sender}: \"{sig['objection_snippet']}\"")

    obj_rev_num = parse_rev_number(objected_at_rev)
    unparked_by_new_rev = bool(is_objected and max_rev > obj_rev_num)
    effectively_parked = bool(is_objected and not unparked_by_new_rev)

    return {
        "max_rev_num": max_rev,
        "max_rev_tag": format_rev_tag(max_rev),
        "unread_max_rev_num": unread_max_rev,
        "unread_max_rev_tag": format_rev_tag(unread_max_rev),
        "rev_tokens": all_rev_tokens,
        "has_template": has_template,
        "unread_has_template": unread_has_template,
        "has_objection": has_objection,
        "unread_has_objection": unread_has_objection,
        "objection_snippets": objection_snippets,
        "is_objected": bool(is_objected),
        "objected_at_rev": format_rev_tag(obj_rev_num) if obj_rev_num > 0 else (str(objected_at_rev or "").strip()),
        "unparked_by_new_rev": unparked_by_new_rev,
        "effectively_parked": effectively_parked,
    }


def format_signal_badges(sig: Dict[str, Any]) -> str:
    """Formats compact table cell badges from a message or group signal dictionary."""
    badges: List[str] = []
    if sig.get("is_objected"):
        if sig.get("unparked_by_new_rev"):
            badges.append(f"🛑→🆕{sig.get('max_rev_tag')}")
        else:
            badges.append("🛑")
    elif sig.get("max_rev_tag"):
        badges.append(f"🆕{sig['max_rev_tag']}")

    if sig.get("unread_has_objection") or (sig.get("has_objection") and not sig.get("is_objected")):
        badges.append("⚠️")
    if sig.get("has_template"):
        badges.append("📋")
    return " ".join(badges)


_SHORTHAND_LINK_OR_MD_RE = re.compile(
    r"(\[[^\]]*\]\([^\)]*\))|\b((?:r|rev\.?\s*)0?([1-9]\d*))\b",
    re.IGNORECASE,
)


def linkify_shorthand_revisions(text: str, primary_tdoc: str = "") -> str:
    """
    Turns shorthand revision mentions (`r01`, `rev 2`) into clickable `[r01](tdoc:S2-2601234R01)`
    links when a `primary_tdoc` (`S2-2601234`) is known, leaving existing Markdown links untouched.
    """
    base = extract_base_tdoc_id(primary_tdoc)
    if not text or not base:
        return text

    def _repl(m: re.Match) -> str:
        if m.group(1):
            return m.group(1)
        raw_tok = m.group(2)
        rev_num = int(m.group(3))
        if rev_num <= 0 or rev_num > 99:
            return raw_tok
        target = f"{base}R{rev_num:02d}"
        return f"[{raw_tok}](tdoc:{target})"

    return _SHORTHAND_LINK_OR_MD_RE.sub(_repl, text)


def build_message_jump_strip_html(
    messages: Sequence[Dict[str, Any]],
    *,
    primary_tdoc: str = "",
    newest_first: bool = True,
) -> str:
    """
    Builds a compact HTML bar of clickable message chips (`jump://msg/<idx>`) for multi-message
    threads and TDocs so the user can jump directly to any company's reply, `🆕 rXX`, `📋`, or `⚠️`.
    Returns empty string if `len(messages) <= 1`.
    """
    if not messages or len(messages) <= 1:
        return ""
    indexed = list(enumerate(messages, 1))
    if newest_first:
        indexed.reverse()

    chips: List[str] = []
    for idx, m in indexed:
        sig = m.get("_signals") or analyze_message_signals(m, primary_tdoc=primary_tdoc)
        comp = str(m.get("company") or m.get("sender_name") or "Unknown").strip()
        if len(comp) > 18:
            comp = comp[:16] + "…"
        is_unread = int(m.get("is_read", 1) or 0) == 0 if isinstance(m.get("is_read"), int) else (not bool(m.get("is_read", True)))
        badges: List[str] = []
        if is_unread:
            badges.append("🔵")
        if sig.get("max_rev_tag"):
            badges.append(f"🆕{sig['max_rev_tag']}")
        if sig.get("has_template"):
            badges.append("📋")
        if sig.get("has_objection"):
            badges.append("⚠️")
        badge_str = (" " + "".join(badges)) if badges else ""
        fg = C.INFO_DARK if is_unread else (C.DANGER if sig.get("has_objection") else C.TEXT_PRIMARY)
        weight = "bold" if (is_unread or sig.get("max_rev_tag") or sig.get("has_objection")) else "normal"
        chips.append(
            f'<a href="jump://msg/{idx}" '
            f'style="text-decoration:none; color:{fg}; font-weight:{weight};">'
            f'[{idx}.&nbsp;{html.escape(comp)}{html.escape(badge_str)}]</a>'
        )
    return (
        f'<span style="color:{C.TEXT_MUTED}; font-size:11px;"><b>Jump:</b> </span>'
        + f' <span style="color:{C.BORDER_STRONG};">·</span> '.join(chips)
    )

