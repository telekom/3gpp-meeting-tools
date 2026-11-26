"""Pure-Python Microsoft 365 Copilot send & paste-back handoff engine."""

from __future__ import annotations

from dataclasses import dataclass, field
import os
from pathlib import Path
import re
import tempfile
from typing import Any, Dict, List, Optional, Sequence, Tuple
import webbrowser

from core.ai.cli_launch import load_standing_cli_context

from .copilot_sanitizer import sanitize_copilot_markdown
from .formatting import (
    extract_deterministic_facts,
    format_messages_for_llm,
    get_message_timestamp_str,
    parse_summary_output,
    sort_messages_chronologically,
)
from .models import EmailSummaryRecord, SummaryScope
from .orchestrator import load_summary_prompt
from .store import EmailSummaryStore, utc_now_iso

M365_COPILOT_WEB_URL = "https://copilot.cloud.microsoft/"
M365_COPILOT_APP_URI = "ms-copilot:"
M365_CLIPBOARD_CHAR_THRESHOLD = 14000


@dataclass
class M365HandoffPacket:
    """Prepared M365 Copilot prompt + provenance metadata for paste-back saving."""

    scope: SummaryScope
    mode: str
    prompt_text: str
    clipboard_prompt: str
    attached_file: Optional[Path]
    cutoff_utc: str
    covered_message_ids: List[str] = field(default_factory=list)
    message_count: int = 0
    new_message_count: int = 0
    incremental_hops: int = 0
    deterministic_facts: Dict[str, Any] = field(default_factory=dict)
    leaf_summary_reused_count: int = 0
    leaf_raw_count: int = 0
    pack_mode: str = "full_uncut"
    was_truncated: bool = False


def _sanitize_filename_part(value: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", str(value or "").strip())
    return cleaned.strip("_")[:60] or "scope"


def stage_handoff_markdown(
    meeting_dir: Optional[Path],
    scope: SummaryScope,
    content: str,
) -> Path:
    """Write a self-contained UTF-8 Markdown dossier for M365 Copilot upload."""
    if meeting_dir is not None:
        base_dir = Path(meeting_dir) / "Agenda" / "CopilotHandoff"
    else:
        base_dir = Path(tempfile.gettempdir()) / "3gpp_copilot_handoff"
    base_dir.mkdir(parents=True, exist_ok=True)
    safe_key = _sanitize_filename_part(scope.scope_key)
    out_path = base_dir / f"{scope.scope_type}_{safe_key}_copilot.md"
    out_path.write_text(str(content or ""), encoding="utf-8")
    return out_path


def _max_msg_timestamp(messages: Sequence[Dict[str, Any]]) -> str:
    stamps = [get_message_timestamp_str(m) for m in messages if get_message_timestamp_str(m)]
    return max(stamps) if stamps else utc_now_iso()


def _finalize_clipboard_and_file(
    *,
    scope: SummaryScope,
    mode_desc: str,
    sys_prompt: str,
    body_payload: str,
    meeting_dir: Optional[Path],
    clipboard_threshold: int,
    force_file: bool,
) -> Tuple[str, str, Optional[Path]]:
    full_prompt = f"=== INSTRUCTIONS ===\n{sys_prompt}\n\n=== DISCUSSION CONTEXT ===\n{body_payload}"
    attached_file: Optional[Path] = None
    if force_file or len(full_prompt) > clipboard_threshold:
        file_doc = (
            f"# 3GPP Email Discussion Pack ({scope.scope_type.upper()}: {scope.scope_key})\n"
            f"Title: {scope.title}\n"
            f"Mode: {mode_desc}\n\n"
            f"{body_payload}\n"
        )
        attached_file = stage_handoff_markdown(meeting_dir, scope, file_doc)
        clip_prompt = (
            f"=== INSTRUCTIONS ===\n{sys_prompt}\n\n"
            f"=== ATTACHED FILE ===\n"
            f"Analyze the attached Markdown file (`{attached_file.name}`) containing `{scope.scope_key}` — {scope.title} ({mode_desc}).\n"
            f"Preserve exact `[Msg #N](jump://msg/N)` and `[TDOC](tdoc:TDOC)` citation links in clean Markdown with blank lines between headings and paragraphs (omit SharePoint file citation links) so I can paste your response back into 3GPP Tools."
        )
    else:
        clip_prompt = full_prompt
    return full_prompt, clip_prompt, attached_file


def build_leaf_handoff_packet(
    store: EmailSummaryStore,
    scope: SummaryScope,
    messages: Sequence[Dict[str, Any]],
    *,
    meeting_dir: Optional[Path] = None,
    wg: str = "",
    meeting_id: str = "",
    mode: str = "auto",
    tdoc_meta: Optional[Dict[str, Any]] = None,
    clipboard_threshold: int = M365_CLIPBOARD_CHAR_THRESHOLD,
    force_file: bool = False,
) -> M365HandoffPacket:
    """Build a full or incremental M365 Copilot handoff packet for Thread, TDoc, or Selection."""
    norm_scope = scope.normalized()
    sorted_msgs = sort_messages_chronologically(messages)
    existing = store.get_summary(norm_scope.scope_type, norm_scope.scope_key)
    fresh = store.compute_freshness(norm_scope.scope_type, norm_scope.scope_key, sorted_msgs)

    req_mode = (mode or "auto").strip().lower()
    use_incremental = (
        existing is not None
        and norm_scope.scope_type in ("thread", "tdoc")
        and fresh.new_count > 0
        and req_mode != "full"
        and (req_mode == "incremental" or not fresh.recommend_full_rebuild)
    )

    primary_tdoc = norm_scope.scope_key if norm_scope.scope_type == "tdoc" else ""
    facts = extract_deterministic_facts(sorted_msgs, primary_tdoc=primary_tdoc, tdoc_meta=tdoc_meta)

    if use_incremental and existing is not None:
        new_id_set = set(fresh.new_message_ids)
        delta_msgs = []
        for idx, msg in enumerate(sorted_msgs, start=1):
            mid = str(msg.get("id") or msg.get("email_id") or "").strip()
            if mid in new_id_set:
                mc = dict(msg)
                mc["_msg_idx"] = idx
                delta_msgs.append(mc)
        sys_prompt = load_summary_prompt(
            "email_summary_incremental_system.txt",
            "Update the previous 3GPP email summary using the new messages. Output ### 🆕 What Changed Since Last Cutoff followed by the full updated 4-section summary.",
        )
        delta_block, included_ids = format_messages_for_llm(delta_msgs, max_chars=500000, per_msg_char_cap=40000)
        was_trunc = len(included_ids) < len(delta_msgs) or "...[truncated]..." in delta_block
        body_payload = (
            f"Scope: {norm_scope.scope_type.upper()} `{norm_scope.scope_key}` — {norm_scope.title}\n"
            f"Previous Cutoff: {existing.cutoff_utc or 'unknown'}\n"
            f"Total Messages Now: {len(sorted_msgs)} (New messages: {len(delta_msgs)})\n\n"
            f"=== PREVIOUS SUMMARY ===\n{existing.summary_md}\n\n"
            f"=== NEW MESSAGES ===\n{delta_block}"
        )
        mode_tag = "incremental"
        mode_desc = f"Incremental +{len(delta_msgs)} new message(s)"
        new_cnt = len(delta_msgs)
        next_hops = existing.incremental_hops + 1
        covered_ids = list(
            dict.fromkeys(
                list(existing.covered_message_ids)
                + [str(m.get("id") or m.get("email_id") or "").strip() for m in sorted_msgs]
            )
        )
    else:
        if norm_scope.scope_type == "selection":
            sys_prompt = load_summary_prompt(
                "email_summary_selection_system.txt",
                "Synthesize across the selected 3GPP email discussions and address the focus prompt if provided.",
            )
        else:
            sys_prompt = load_summary_prompt(
                "email_summary_thread_system.txt",
                "Summarize the 3GPP email discussion into Core Contention, Company Positions, Revisions/Open Issues, and Consensus/Next Steps.",
            )
        msgs_block, included_ids = format_messages_for_llm(sorted_msgs, max_chars=500000, per_msg_char_cap=40000)
        was_trunc = len(included_ids) < len(sorted_msgs) or "...[truncated]..." in msgs_block
        meta_lines = [
            f"Scope: {norm_scope.scope_type.upper()} `{norm_scope.scope_key}`",
            f"Title: {norm_scope.title}",
            f"Meeting: {meeting_id or 'N/A'} ({wg or 'N/A'})",
            f"Total Messages: {len(sorted_msgs)}",
            f"Participating Companies: {', '.join(facts.get('companies') or []) or 'Unknown'}",
        ]
        if facts.get("tdoc_title"):
            meta_lines.append(
                f"TDoc Title: {facts['tdoc_title']} (Author: {facts.get('tdoc_author', 'N/A')}, Status: {facts.get('tdoc_status', 'N/A')})"
            )
        if norm_scope.focus_prompt:
            meta_lines.append(f"Focus Prompt: {norm_scope.focus_prompt}")
        body_payload = "\n".join(meta_lines) + "\n\nCleaned Chronological Messages:\n\n" + msgs_block
        mode_tag = "full"
        mode_desc = f"Full ({len(sorted_msgs)} messages)"
        new_cnt = len(sorted_msgs)
        next_hops = 0
        covered_ids = [str(m.get("id") or m.get("email_id") or "").strip() for m in sorted_msgs]

    full_prompt, clip_prompt, attached_file = _finalize_clipboard_and_file(
        scope=norm_scope,
        mode_desc=mode_desc,
        sys_prompt=sys_prompt,
        body_payload=body_payload,
        meeting_dir=meeting_dir,
        clipboard_threshold=clipboard_threshold,
        force_file=force_file,
    )
    return M365HandoffPacket(
        scope=norm_scope,
        mode=mode_tag,
        prompt_text=full_prompt,
        clipboard_prompt=clip_prompt,
        attached_file=attached_file,
        cutoff_utc=_max_msg_timestamp(sorted_msgs),
        covered_message_ids=[x for x in covered_ids if x],
        message_count=len(sorted_msgs),
        new_message_count=new_cnt,
        incremental_hops=next_hops,
        deterministic_facts=facts,
        pack_mode="full_uncut",
        was_truncated=was_trunc,
    )


def build_ai_handoff_packet(
    store: EmailSummaryStore,
    ai_code: str,
    ai_label: str,
    leaf_buckets: Dict[str, Dict[str, Any]],
    *,
    meeting_dir: Optional[Path] = None,
    wg: str = "",
    meeting_id: str = "",
    tdoc_meta_map: Optional[Dict[str, Dict[str, Any]]] = None,
    focus_prompt: str = "",
    clipboard_threshold: int = M365_CLIPBOARD_CHAR_THRESHOLD,
    force_file: bool = False,
    pack_mode: str = "compact",
) -> M365HandoffPacket:
    """
    Build a Per-AI cross-thinking M365 Copilot packet.
    - `pack_mode="full_uncut"`: embeds full uncut deduplicated emails (ideal for Copilot file uploads).
    - `pack_mode="compact"`: reuses fresh saved leaf summaries and embeds uncut raw/delta messages for the rest.
    """
    clean_ai = str(ai_code or "").strip()
    clean_pack_mode = "full_uncut" if str(pack_mode or "").strip().lower() == "full_uncut" else "compact"
    norm_scope = SummaryScope(
        scope_type="ai",
        scope_key=clean_ai,
        title=ai_label or f"AI {clean_ai}",
        focus_prompt=focus_prompt.strip(),
    ).normalized()

    standing_ctx = load_standing_cli_context().strip()
    sys_prompt = load_summary_prompt(
        "email_summary_ai_reduce_system.txt",
        "Perform cross-TDoc strategic synthesis across the provided leaf summaries and discussions under this Agenda Item.",
    )

    sections: List[str] = []
    all_msgs_for_ai: List[Dict[str, Any]] = []
    reused_cnt = 0
    raw_cnt = 0
    was_truncated = False
    leaf_keys: List[str] = []
    per_bucket_cap = 500000 if clean_pack_mode == "full_uncut" else 250000
    per_msg_cap = 40000

    seen_sigs: set = set()
    dedup_buckets: List[Tuple[str, Dict[str, Any]]] = []
    for bkey, bucket in leaf_buckets.items():
        msgs = list(bucket.get("messages") or [])
        if not msgs:
            continue
        sig = tuple(sorted(str(m.get("id") or m.get("email_id") or "") for m in msgs if (m.get("id") or m.get("email_id"))))
        if sig and sig in seen_sigs:
            continue
        if sig:
            seen_sigs.add(sig)
        dedup_buckets.append((bkey, bucket))

    for idx, (bkey, bucket) in enumerate(dedup_buckets, 1):
        msgs = list(bucket.get("messages") or [])
        all_msgs_for_ai.extend(msgs)
        stype = str(bucket.get("scope_type") or "thread").lower()
        skey = str(bucket.get("scope_key") or bkey)
        btitle = str(bucket.get("title") or skey)
        alias_tdocs = [str(t).strip().upper() for t in (bucket.get("alias_tdocs") or []) if str(t).strip()]
        tmeta = bucket.get("tdoc_meta") or ((tdoc_meta_map or {}).get(skey.upper()) if stype == "tdoc" else None)
        leaf_keys.append(f"{stype}:{skey}")

        fresh = store.compute_freshness(stype, skey, msgs)
        existing_leaf = store.get_summary(stype, skey)
        link_tag = f"[{skey}](tdoc:{skey})" if stype == "tdoc" else f"[{btitle}](thread:{skey})"

        item_lines = [f"## Item {idx}: {link_tag} — {btitle}", f"- **Scope:** `{stype}:{skey}` | **Messages:** {len(msgs)}"]
        if alias_tdocs:
            td_links = ", ".join(f"[{td}](tdoc:{td})" for td in alias_tdocs)
            item_lines.append(f"- **Associated TDocs:** {td_links}")
        if tmeta and (tmeta.get("tdoc_status") or tmeta.get("my_status")):
            item_lines.append(
                f"- **TDoc Status:** {tmeta.get('tdoc_status') or 'N/A'} | **My Status:** {tmeta.get('my_status') or 'Neutral'}"
            )

        if clean_pack_mode == "compact" and existing_leaf is not None and not fresh.is_stale:
            reused_cnt += 1
            item_lines.append("- **Source:** Saved up-to-date AI Leaf Summary")
            item_lines.append("")
            item_lines.append(existing_leaf.summary_md)
        else:
            raw_cnt += 1
            if clean_pack_mode == "compact" and existing_leaf is not None and fresh.new_count > 0:
                new_id_set = set(fresh.new_message_ids)
                sorted_msgs = sort_messages_chronologically(msgs)
                delta_msgs = []
                for pos, msg in enumerate(sorted_msgs, 1):
                    mid = str(msg.get("id") or msg.get("email_id") or "").strip()
                    if mid in new_id_set:
                        mc = dict(msg)
                        mc["_msg_idx"] = pos
                        delta_msgs.append(mc)
                delta_txt, inc_ids = format_messages_for_llm(delta_msgs, max_chars=per_bucket_cap, per_msg_char_cap=per_msg_cap)
                if len(inc_ids) < len(delta_msgs) or "...[truncated]..." in delta_txt:
                    was_truncated = True
                item_lines.append(f"- **Source:** Saved Leaf Summary + {len(delta_msgs)} New Message(s)")
                item_lines.append("")
                item_lines.append(existing_leaf.summary_md)
                item_lines.append(f"\n### New Messages (+{len(delta_msgs)})\n{delta_txt}")
            else:
                raw_txt, inc_ids = format_messages_for_llm(msgs, max_chars=per_bucket_cap, per_msg_char_cap=per_msg_cap)
                if len(inc_ids) < len(msgs) or "...[truncated]..." in raw_txt:
                    was_truncated = True
                item_lines.append("- **Source:** Full Cleaned Chronological Messages")
                item_lines.append("")
                item_lines.append(raw_txt)
        sections.append("\n".join(item_lines))

    dedup_msgs: Dict[str, Dict[str, Any]] = {}
    for msg in all_msgs_for_ai:
        mid = str(msg.get("id") or msg.get("email_id") or "").strip()
        if mid:
            dedup_msgs[mid] = msg
    ai_msgs = list(dedup_msgs.values()) if dedup_msgs else all_msgs_for_ai
    ai_facts = extract_deterministic_facts(ai_msgs)
    ai_facts["leaf_count"] = len(leaf_keys)
    ai_facts["leaf_keys"] = leaf_keys

    mode_label = "Full Uncut Emails" if clean_pack_mode == "full_uncut" else f"Compact ({reused_cnt} pre-summarized, {raw_cnt} raw/delta)"
    header_parts = [
        f"Agenda Item: {ai_label or clean_ai}",
        f"Meeting: {meeting_id or 'N/A'} ({wg or 'N/A'})",
        f"Topics Included: {len(sections)} ({mode_label})",
        f"Total Messages Covered: {len(ai_msgs)}",
    ]
    if norm_scope.focus_prompt:
        header_parts.append(f"Delegate Focus Prompt: {norm_scope.focus_prompt}")
    if standing_ctx:
        header_parts.append(f"\n=== STANDING COMPANY / DELEGATE CONTEXT ===\n{standing_ctx[:3000]}")
    header_parts.append("\n=== TOPIC SUMMARIES & DISCUSSIONS ===\n\n" + "\n\n---\n\n".join(sections))
    body_payload = "\n".join(header_parts)

    full_prompt, clip_prompt, attached_file = _finalize_clipboard_and_file(
        scope=norm_scope,
        mode_desc=f"Per-AI Cross-Thinking ({len(sections)} topics, {len(ai_msgs)} msgs)",
        sys_prompt=sys_prompt,
        body_payload=body_payload,
        meeting_dir=meeting_dir,
        clipboard_threshold=clipboard_threshold,
        force_file=force_file,
    )
    return M365HandoffPacket(
        scope=norm_scope,
        mode="ai_briefing",
        prompt_text=full_prompt,
        clipboard_prompt=clip_prompt,
        attached_file=attached_file,
        cutoff_utc=_max_msg_timestamp(ai_msgs),
        covered_message_ids=list(dedup_msgs.keys()),
        message_count=len(ai_msgs),
        new_message_count=len(ai_msgs),
        incremental_hops=0,
        deterministic_facts=ai_facts,
        leaf_summary_reused_count=reused_cnt,
        leaf_raw_count=raw_cnt,
        pack_mode=clean_pack_mode,
        was_truncated=was_truncated,
    )


def save_external_summary_from_handoff(
    store: EmailSummaryStore,
    packet: M365HandoffPacket,
    raw_response_text: str,
    *,
    provider: str = "m365-copilot",
    model: str = "M365 Copilot",
) -> Tuple[Optional[EmailSummaryRecord], str]:
    """Parse and persist an externally generated M365 Copilot summary into `EmailSummaryStore`."""
    cleaned = sanitize_copilot_markdown(raw_response_text)
    if not cleaned:
        return None, "Pasted summary text is empty."

    summary_md, delta_md, structured_meta = parse_summary_output(cleaned)
    if not summary_md:
        summary_md = cleaned

    rec = EmailSummaryRecord(
        scope_type=packet.scope.scope_type,
        scope_key=packet.scope.scope_key,
        title=packet.scope.title,
        summary_md=summary_md,
        delta_md=delta_md,
        structured_meta=structured_meta,
        deterministic_facts=dict(packet.deterministic_facts),
        cutoff_utc=packet.cutoff_utc,
        covered_message_ids=list(packet.covered_message_ids),
        message_count=packet.message_count,
        incremental_hops=packet.incremental_hops,
        provider=str(provider or "m365-copilot").strip() or "m365-copilot",
        model=str(model or "M365 Copilot").strip() or "M365 Copilot",
        reasoning_effort="",
        focus_prompt=packet.scope.focus_prompt,
    )
    saved = store.upsert_summary(rec, save_history=True)
    return saved, ""


def launch_m365_copilot(target: str = "web") -> Tuple[bool, str]:
    """Open Microsoft 365 Copilot Web (`copilot.cloud.microsoft`) or desktop app (`ms-copilot:`)."""
    clean = str(target or "web").strip().lower()
    if clean == "clipboard_only":
        return True, "Prompt copied to clipboard (browser launch skipped)."
    if clean == "app" and os.name == "nt" and hasattr(os, "startfile"):
        try:
            os.startfile(M365_COPILOT_APP_URI)
            return True, f"Opened Microsoft Copilot app ({M365_COPILOT_APP_URI})."
        except Exception:
            pass
    try:
        webbrowser.open(M365_COPILOT_WEB_URL)
        return True, f"Opened M365 Copilot in browser ({M365_COPILOT_WEB_URL})."
    except Exception as exc:
        return False, f"Could not launch M365 Copilot: {exc}"
