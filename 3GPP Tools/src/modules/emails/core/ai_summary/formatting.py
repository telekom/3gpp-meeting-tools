"""Deterministic facts extraction, LLM prompt formatting, and fault-tolerant output parsing."""

from __future__ import annotations

import json
import re
from typing import Any, Dict, List, Optional, Sequence, Tuple

from modules.emails.core.emeeting_signals import compute_group_signals, extract_base_tdoc_id
from modules.emails.core.markdown_export import (
    split_direct_and_quoted_body,
    unwrap_outlook_safelinks,
)

_META_BLOCK_RE = re.compile(
    r"```\s*summary-meta\s*\n(.*?)\n```",
    re.DOTALL | re.IGNORECASE,
)

_DELTA_HEADING_RE = re.compile(
    r"(?im)^#{2,4}\s*🆕?\s*What\s+Changed[^\n]*\n(.*?)(?=^#{1,3}\s+1\.|\Z)",
    re.DOTALL | re.MULTILINE,
)


def sort_messages_chronologically(messages: Sequence[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Return messages sorted oldest-to-newest by `date_sent` / `date_received` / `date_str`."""
    return sorted(
        list(messages or []),
        key=lambda m: str(m.get("date_sent") or m.get("date_received") or m.get("date_str") or ""),
    )


def get_message_timestamp_str(msg: Dict[str, Any]) -> str:
    return str(msg.get("date_sent") or msg.get("date_received") or msg.get("date_str") or "").strip()


def extract_clean_direct_body(msg: Dict[str, Any]) -> str:
    """Return direct message body (preserving `<<START>>..<<END>>` blocks) without trailing quote noise."""
    precomputed = str(msg.get("clean_body") or "").strip()
    if precomputed:
        return unwrap_outlook_safelinks(precomputed)
    raw = str(msg.get("raw_body_text") or msg.get("body_text") or "").strip()
    if not raw:
        return ""
    direct, _quoted = split_direct_and_quoted_body(raw)
    return unwrap_outlook_safelinks(direct or raw)


def extract_deterministic_facts(
    messages: Sequence[Dict[str, Any]],
    *,
    primary_tdoc: str = "",
    tdoc_meta: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """Extract 100% deterministic metrics and e-Meeting signals directly from DB records."""
    sorted_msgs = sort_messages_chronologically(messages)
    companies = list(
        dict.fromkeys(
            str(m.get("company") or "").strip()
            for m in sorted_msgs
            if str(m.get("company") or "").strip()
            and str(m.get("company") or "").strip().lower() not in ("unknown", "list")
        )
    )
    is_obj = bool((tdoc_meta or {}).get("is_objected")) or any(bool(m.get("is_objected")) for m in sorted_msgs)
    obj_rev = str((tdoc_meta or {}).get("objected_at_rev") or "").strip()
    sig = compute_group_signals(
        sorted_msgs,
        primary_tdoc=primary_tdoc,
        is_objected=is_obj,
        objected_at_rev=obj_rev,
    )
    facts: Dict[str, Any] = {
        "message_count": len(sorted_msgs),
        "companies": companies,
        "latest_date": get_message_timestamp_str(sorted_msgs[-1]) if sorted_msgs else "",
        "max_rev_tag": sig.get("max_rev_tag") or "",
        "rev_tokens": list(sig.get("rev_tokens") or []),
        "has_template": bool(sig.get("has_template")),
        "has_objection": bool(sig.get("has_objection")),
        "is_objected": bool(sig.get("is_objected")),
        "unparked_by_new_rev": bool(sig.get("unparked_by_new_rev")),
    }
    if tdoc_meta:
        for src_k, dst_k in (
            ("title", "tdoc_title"),
            ("source", "tdoc_author"),
            ("ai", "ai"),
            ("tdoc_status", "tdoc_status"),
            ("my_status", "my_status"),
        ):
            val = str(tdoc_meta.get(src_k) or "").strip()
            if val:
                facts[dst_k] = val
    return facts


def format_messages_for_llm(
    messages: Sequence[Dict[str, Any]],
    *,
    index_offset: int = 0,
    max_chars: int = 85000,
    per_msg_char_cap: int = 8000,
) -> Tuple[str, List[str]]:
    """
    Format messages chronologically as `[Msg #N]` blocks for LLM prompts.
    Returns `(formatted_text, included_message_ids)`.
    """
    sorted_msgs = sort_messages_chronologically(messages)
    blocks: List[str] = []
    included_ids: List[str] = []
    total_chars = 0

    for pos, msg in enumerate(sorted_msgs, start=1 + index_offset):
        i = int(msg.get("_msg_idx") or pos)
        mid = str(msg.get("id") or msg.get("email_id") or "").strip()
        sender = str(msg.get("sender_name") or msg.get("sender_email") or "Unknown").strip()
        company = str(msg.get("company") or "Unknown").strip()
        ts = get_message_timestamp_str(msg)[:16] or "Unknown Date"
        subj = str(msg.get("subject") or "").strip()
        tdocs = [extract_base_tdoc_id(t) for t in (msg.get("linked_tdocs") or []) if extract_base_tdoc_id(t)]
        atts = [str(a).strip() for a in (msg.get("attachments") or []) if str(a).strip()]

        body = extract_clean_direct_body(msg)
        if not body:
            body = "(No direct text; message may consist solely of quoted history or attachments.)"
        elif len(body) > per_msg_char_cap:
            body = body[:per_msg_char_cap] + "\n...[truncated]..."

        header_lines = [
            f"--- [Msg #{i}] ---",
            f"From: {sender} ({company})",
            f"Date: {ts}",
            f"Subject: {subj}",
        ]
        if tdocs:
            header_lines.append(f"TDocs: {', '.join(dict.fromkeys(tdocs))}")
        if atts:
            header_lines.append(f"Attachments: {', '.join(atts[:8])}")
        block = "\n".join(header_lines) + "\n\n" + body

        if blocks and total_chars + len(block) > max_chars:
            blocks.append(f"... ({len(sorted_msgs) - len(blocks)} additional message(s) truncated for context budget) ...")
            break
        blocks.append(block)
        total_chars += len(block) + 2
        if mid:
            included_ids.append(mid)

    return "\n\n".join(blocks), included_ids


def parse_summary_output(raw_text: str) -> Tuple[str, str, Dict[str, Any]]:
    """
    Parse raw LLM Markdown output into `(clean_summary_md, delta_md, structured_meta)`.
    Fault-tolerant: malformed or missing `summary-meta` blocks never raise or fail the run.
    """
    text = str(raw_text or "").strip()
    if not text:
        return "", "", {}

    structured_meta: Dict[str, Any] = {}
    meta_match = _META_BLOCK_RE.search(text)
    if meta_match:
        json_candidate = meta_match.group(1).strip()
        try:
            parsed = json.loads(json_candidate)
            if isinstance(parsed, dict):
                structured_meta = parsed
        except Exception:
            structured_meta = {}
        text = (_META_BLOCK_RE.sub("", text)).strip()

    delta_md = ""
    delta_match = _DELTA_HEADING_RE.search(text)
    if delta_match:
        delta_md = delta_match.group(1).strip()

    return text, delta_md, structured_meta
