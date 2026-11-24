# --- File: src/modules/emails/core/markdown_export/exporter.py ---
"""Option A interactive HTML-banner & LLM-clean Markdown thread exporter."""

import html
import urllib.parse
from typing import Any, Dict, List

from core.ui.theme import C

from .formatting_helpers import (
    _TDOC_RE,
    _normalize_subject_for_comparison,
    clean_tdoc_id,
    format_compact_timestamp,
    format_sender_display,
    partition_email_attachments,
    split_sender_company,
)
from .reader_ui import format_emeeting_body_with_callouts
from .url_and_links import extract_thread_document_urls, linkify_tdoc_tokens


def export_thread_to_markdown(
    thread: Dict[str, Any],
    messages: List[Dict[str, Any]],
    deduplicate_quotes: bool = True,
    interactive_links: bool = False,
    newest_first: bool = False,
    include_metadata: bool = True,
    include_ignore_action: bool = False,
    enable_tdoc_links: bool = False,
    include_discussion_heading: bool = True,
) -> str:
    """
    Exports a tracked email thread to a lossless, human- and LLM-readable Markdown document.
    Preserves inline replies, metadata, participants/companies, and TDoc references,
    unwraps Outlook Safelinks/Proofpoint, and formats URLs as clickable Markdown links.
    """
    participants = []
    seen_p = set()
    tdocs_mentioned = {clean_tdoc_id(t) for t in (thread.get("linked_tdocs") or []) if clean_tdoc_id(t)}

    for msg in messages:
        name = (msg.get("sender_name") or msg.get("sender_email") or "Unknown").strip()
        comp = (msg.get("company") or "").strip()
        label = format_sender_display(name, comp)
        if label not in seen_p:
            seen_p.add(label)
            participants.append(label)
        for m in _TDOC_RE.findall(f"{msg.get('subject', '')} {msg.get('body_text', '')}"):
            cleaned_m = clean_tdoc_id(m)
            if cleaned_m:
                tdocs_mentioned.add(cleaned_m)

    dates = [m.get("date_sent", "") for m in messages if m.get("date_sent")]
    date_range = f"{min(dates)} to {max(dates)}" if dates else "N/A"
    doc_urls = extract_thread_document_urls(messages)

    disc_heading = "## Discussion (Newest First)" if newest_first else "## Chronological Discussion"
    lines: List[str] = []
    if include_metadata:
        lines.extend([
            f"# Email Thread: {thread.get('title', 'Untitled Thread')}",
            "",
            "## Thread Metadata",
            f"- **Thread ID:** `{thread.get('thread_id', '')}`",
            f"- **Meeting:** `{thread.get('meeting_id', '') or 'N/A'}` ({thread.get('wg', '') or 'N/A'})",
            f"- **Linked AIs:** {', '.join(thread.get('linked_ais') or []) or 'None'}",
            f"- **Linked Sub-Topics:** {', '.join(thread.get('linked_sub_items') or []) or 'None'}",
            f"- **Linked WIs:** {', '.join(thread.get('linked_wis') or []) or 'None'}",
            f"- **Linked / Mentioned TDocs:** {', '.join(sorted(tdocs_mentioned)) or 'None'}",
            f"- **Date Range:** {date_range}",
            f"- **Messages:** {len(messages)}",
            f"- **Participants:** {', '.join(participants) or 'None'}",
        ])
        user_status = str(thread.get("user_status") or "").strip()
        user_tags = [str(x).strip() for x in (thread.get("user_tags") or []) if str(x).strip()]
        user_notes = str(thread.get("user_notes") or "").strip()
        if user_status and user_status != "⚪ Neutral":
            lines.append(f"- **My Status:** {user_status}")
        if user_tags:
            lines.append(f"- **My Tags:** {', '.join(user_tags)}")
        if doc_urls:
            lines.append(f"- **Linked Documents ({len(doc_urls)}):**")
            for u in doc_urls:
                fname = urllib.parse.unquote(urllib.parse.urlsplit(u).path.rsplit("/", 1)[-1]) or u
                lines.append(f"  - [📄 {fname}]({u})")
        if user_notes:
            lines.append("")
            lines.append("> [!NOTE] Delegate Notes")
            for nline in user_notes.splitlines():
                lines.append(f"> {nline}" if nline.strip() else ">")
        lines.extend(["", "---", ""])
    if include_discussion_heading:
        lines.extend([disc_heading, ""])

    base_subject_norm = _normalize_subject_for_comparison(
        thread.get("normalized_subject") or thread.get("title") or ""
    )
    linked_tdocs_list = [clean_tdoc_id(t) for t in (thread.get("linked_tdocs") or []) if clean_tdoc_id(t)]
    primary_tdoc = linked_tdocs_list[0] if linked_tdocs_list else ""
    from modules.emails.core.emeeting_signals import analyze_message_signals
    indexed_msgs = list(enumerate(messages, 1))
    if newest_first:
        indexed_msgs.reverse()

    unread_seq = 0
    for idx, msg in indexed_msgs:
        sender = (msg.get("sender_name") or msg.get("sender_email") or "Unknown").strip()
        comp = (msg.get("company") or "").strip()
        sender_disp = format_sender_display(sender, comp)
        date_raw = msg.get("date_sent", "Unknown Date")
        date_disp = format_compact_timestamp(date_raw) if interactive_links else date_raw
        subj = (msg.get("subject") or "").strip()
        msg_id = str(msg.get("id") or "").strip()
        is_unread = int(msg.get("is_read", 1) or 0) == 0
        is_flagged = int(msg.get("flag_status", 0) or 0) == 2
        is_ignored = int(msg.get("is_ignored", 0) or 0) == 1
        can_open_outlook = interactive_links and msg_id and not msg_id.startswith("MSG-")
        msg_sig = msg.get("_signals") or analyze_message_signals(msg, primary_tdoc=primary_tdoc)

        if interactive_links:
            anchors = f'<a name="msg-{idx}"></a>'
            if is_unread:
                unread_seq += 1
                anchors += f'<a name="unread-{unread_seq}"></a>'
            enc_id = urllib.parse.quote(msg_id, safe="")
            if is_unread:
                card_bg, card_accent, card_border = C.INFO_SOFT, C.ACCENT_ALT, C.INFO_BORDER
            elif is_flagged:
                card_bg, card_accent, card_border = C.WARNING_SOFT, C.WARNING_ACCENT, C.WARNING_BORDER
            elif is_ignored:
                card_bg, card_accent, card_border = C.BG_SURFACE, C.BORDER_STRONG, C.BORDER_SUBTLE
            else:
                card_bg, card_accent, card_border = C.BG_CANVAS, C.BORDER_HOVER, C.BORDER_SUBTLE

            name_part, comp_part = split_sender_company(sender, comp)
            sender_html = f'<b style="color:{C.TEXT_PRIMARY};">{html.escape(name_part)}</b>'
            if comp_part:
                sender_html += f' <span style="color:{C.TEXT_SECONDARY};">({html.escape(comp_part)})</span>'
            badge_parts = []
            if is_unread:
                badge_parts.append(f'<span style="color:{C.INFO}; font-weight:bold;">🔵 NEW</span>')
            if is_flagged:
                badge_parts.append(f'<span style="color:{C.WARNING}; font-weight:bold;">🚩 FLAGGED</span>')
            if msg_sig.get("max_rev_tag"):
                rev_t = str(msg_sig["max_rev_tag"])
                badge_parts.append(f'<span style="color:{C.SUCCESS_DARK}; font-weight:bold;">🆕 {html.escape(rev_t)}</span>')
            if msg_sig.get("has_template"):
                badge_parts.append(f'<span style="color:{C.WARNING_DARK}; font-weight:bold;">📋 e-Comment</span>')
            if msg_sig.get("has_objection"):
                badge_parts.append(f'<span style="color:{C.DANGER}; font-weight:bold;">⚠️ Objection?</span>')
            if is_ignored:
                badge_parts.append(f'<span style="color:{C.TEXT_MUTED}; font-weight:bold;">🚫 IGNORED</span>')
            badges_html = (" · " + " · ".join(badge_parts)) if badge_parts else ""
            left_html = (
                f'{anchors}<span style="color:{C.TEXT_MUTED};">{idx}.</span> {sender_html}'
                f' <span style="color:{C.TEXT_MUTED};">· {html.escape(str(date_disp))}</span>{badges_html}'
            )

            action_links = []
            if can_open_outlook:
                action_links.append(
                    f'<a href="outlook://reply/{enc_id}" style="text-decoration:none; color:{C.ACCENT_ALT};">↩️&nbsp;Reply</a>'
                )
                action_links.append(
                    f'<a href="outlook://msg/{enc_id}" style="text-decoration:none; color:{C.ACCENT_ALT};">📧&nbsp;Outlook</a>'
                )
                action_links.append(
                    f'<a href="outlook://reload/{enc_id}" style="text-decoration:none; color:{C.ACCENT_ALT};">🔄</a>'
                )
            if msg_id:
                action_links.append(
                    f'<a href="outlook://read/{enc_id}" style="text-decoration:none; color:{C.ACCENT_ALT};">✓&nbsp;Read</a>'
                    if is_unread
                    else f'<a href="outlook://unread/{enc_id}" style="text-decoration:none; color:{C.ACCENT_ALT};">🔵&nbsp;Unread</a>'
                )
                action_links.append(
                    f'<a href="outlook://unflag/{enc_id}" style="text-decoration:none; color:{C.ACCENT_ALT};">🏳️&nbsp;Unflag</a>'
                    if is_flagged
                    else f'<a href="outlook://flag/{enc_id}" style="text-decoration:none; color:{C.ACCENT_ALT};">🚩</a>'
                )
                if include_ignore_action:
                    action_links.append(
                        f'<a href="outlook://unignore/{enc_id}" style="text-decoration:none; color:{C.ACCENT_ALT};">↩️&nbsp;Un-ignore</a>'
                        if is_ignored
                        else f'<a href="outlook://ignore/{enc_id}" style="text-decoration:none; color:{C.ACCENT_ALT};">🚫</a>'
                    )
            action_links.append(
                f'<a href="outlook://copy/{enc_id}" style="text-decoration:none; color:{C.ACCENT_ALT};">📋</a>'
            )
            right_html = f' <span style="color:{C.BORDER_STRONG};">·</span> '.join(action_links)
            lines.append(
                f'<table width="100%" cellspacing="0" cellpadding="4" '
                f'style="background-color:{card_bg}; border-top:1px solid {card_border}; '
                f'border-bottom:1px solid {card_border}; margin-top:6px; margin-bottom:4px;">'
                f'<tr>'
                f'<td width="4" style="background-color:{card_accent};"></td>'
                f'<td valign="middle"><h3 style="margin:0; font-size:medium; font-weight:normal;">{left_html}</h3></td>'
                f'<td align="right" valign="middle" style="font-size:90%;">{right_html}</td>'
                f'</tr></table>'
            )
            lines.append("")
        else:
            lines.append(f"### {idx}. [{date_disp}] {sender_disp}")

        subj_norm = _normalize_subject_for_comparison(subj)
        if subj and base_subject_norm and subj_norm and subj_norm != base_subject_norm:
            subj_disp = linkify_tdoc_tokens(subj) if (interactive_links and enable_tdoc_links) else subj
            lines.append(f"**Subject:** {subj_disp}  ")
        match_loc = str(msg.get("match_location") or "").strip()
        folder_tag = str(msg.get("folder_tag") or "").strip()
        matched_tdoc = clean_tdoc_id(msg.get("matched_tdoc") or "")
        rev_matched = str(msg.get("rev_matched") or "").strip()
        if match_loc or folder_tag or matched_tdoc:
            meta_bits = []
            if folder_tag:
                meta_bits.append(f"**Tag:** `[{folder_tag}]`")
            if match_loc:
                td_detail = f" ({matched_tdoc}" + (f" · {rev_matched}" if rev_matched else "") + ")" if matched_tdoc else ""
                meta_bits.append(f"**Match In:** {match_loc}{td_detail}")
            lines.append(" · ".join(meta_bits) + "  ")
        raw_atts = msg.get("attachments") or []
        if raw_atts:
            real_atts, inline_imgs = partition_email_attachments(raw_atts) if interactive_links else (raw_atts, [])
            if real_atts:
                if can_open_outlook:
                    enc_id = urllib.parse.quote(msg_id, safe="")
                    att_links = [
                        f"[📎 {a}](outlook://att/{enc_id}/{urllib.parse.quote(str(a), safe='')})"
                        for a in real_atts
                    ]
                    att_line = f"**Attachments:** {' · '.join(att_links)}"
                else:
                    att_line = f"**Attachments:** {', '.join(real_atts)}"
                if inline_imgs:
                    att_line += f" *(+{len(inline_imgs)} inline)*"
                lines.append(att_line + "  ")
        lines.append("")

        raw_body = (msg.get("clean_body") if deduplicate_quotes and msg.get("clean_body") else msg.get("body_text")) or ""
        msg_primary_tdoc = matched_tdoc or primary_tdoc
        body = format_emeeting_body_with_callouts(
            raw_body,
            primary_tdoc=msg_primary_tdoc,
            interactive_links=interactive_links,
            enable_tdoc_links=enable_tdoc_links,
            deduplicate_quotes=deduplicate_quotes,
        )
        if body:
            lines.append(body)
        else:
            lines.append("*(Empty message body)*")
        lines.append("")
        if not interactive_links:
            lines.append("---")
            lines.append("")

    return "\n".join(lines).strip() + "\n"
