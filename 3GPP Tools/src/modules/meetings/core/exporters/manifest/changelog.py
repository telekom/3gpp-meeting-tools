# --- File: src/modules/meetings/core/exporters/manifest/changelog.py ---
"""Markdown changelog formatters for master and local Work Item index files."""

import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

from .models import ExportManifest


def format_master_changelog_markdown(
    manifest: ExportManifest,
    current_session: Dict[str, Any],
    new_doc_entries: List[Dict[str, Any]],
    updated_doc_entries: Optional[List[Dict[str, Any]]] = None,
    base_dir: Optional[Path] = None,
) -> str:
    """
    Renders an unobtrusive HTML <details> collapsible changelog block at the bottom
    of the master index.md file, optimized for human reading and LLM incremental analysis.
    """
    updated_doc_entries = updated_doc_entries or []
    sessions = manifest.sessions or [current_session]
    total_docs = manifest.get_total_tdocs()

    is_update = current_session.get("session_id", 1) > 1
    added_count = len(new_doc_entries)
    updated_count = len(updated_doc_entries)
    timestamp = current_session.get("timestamp", datetime.datetime.now().strftime("%Y-%m-%d %H:%M"))

    if is_update:
        delta_str = f"+{added_count} TDoc(s) added"
        if updated_count:
            delta_str += f", {updated_count} updated"
        summary_title = (
            f"🕒 <b>Export Changelog & LLM Incremental Update Guide</b> — "
            f"<i>Latest update: {timestamp} ({delta_str}, {total_docs} total)</i>"
        )
    else:
        summary_title = (
            f"🕒 <b>Export Changelog & LLM Incremental Update Guide</b> — "
            f"<i>Initial Export ({total_docs} TDocs)</i>"
        )

    lines = [
        "\n---\n",
        "## 🔄 Export History & Incremental LLM Guide\n",
        "<details>",
        f"<summary>{summary_title}</summary>\n",
        "### 📅 Export Sessions History\n",
    ]

    for s in sessions:
        sid = s.get("session_id", 1)
        st_time = s.get("timestamp", "")
        s_lbl = s.get("label", f"Session #{sid}")
        s_added = len(s.get("added_tdocs", []))
        s_tot = s.get("total_tdocs", 0)
        s_scope = s.get("scope_mode", "")
        scope_suffix = f" *(scope: `{s_scope}`)*" if s_scope else ""
        if sid == 1:
            lines.append(f"- **Session #1 ({s_lbl}):** `{st_time}` — {s_added} initial TDoc(s){scope_suffix}")
        else:
            lines.append(
                f"- **Session #{sid} ({s_lbl}):** `{st_time}` — **+{s_added} added** ({s_tot} total in dossier){scope_suffix}"
            )

    lines.append("")

    # Delta table for newly added TDocs in this session
    if new_doc_entries:
        lines.append(f"### 🆕 Newly Added TDocs in Latest Update (+{len(new_doc_entries)})\n")
        lines.append("| TDoc | Type | Work Item | AI | Title | Source | Dossier Link |")
        lines.append("|:---|:---|:---|:---:|:---|:---|:---:|")

        for d in sorted(new_doc_entries, key=lambda x: str(x.get("tdoc_id", ""))):
            tid = d.get("tdoc_id", "")
            doc_type = d.get("type", "")
            wi = d.get("wi", "General")
            ai = d.get("ai", "")
            title = str(d.get("title", "")).replace("\n", " ").replace("|", "\\|")
            source = str(d.get("source", "")).replace("\n", " ").replace("|", "\\|")
            rel_link = d.get("rel_path") or d.get("target_md_rel_path") or ""

            link_cell = f"[📄 Open]({rel_link})" if rel_link else tid
            lines.append(f"| `{tid}` | {doc_type} | `{wi}` | {ai} | {title} | {source} | {link_cell} |")

        lines.append("")

    # Updated / Revised TDocs table (if any)
    if updated_doc_entries:
        lines.append(f"### 🔄 Updated / Re-converted TDocs ({len(updated_doc_entries)})\n")
        lines.append("| TDoc | Type | Work Item | Title | Source | Link |")
        lines.append("|:---|:---|:---|:---|:---|:---:|")
        for d in sorted(updated_doc_entries, key=lambda x: str(x.get("tdoc_id", ""))):
            tid = d.get("tdoc_id", "")
            doc_type = d.get("type", "")
            wi = d.get("wi", "General")
            title = str(d.get("title", "")).replace("\n", " ").replace("|", "\\|")
            source = str(d.get("source", "")).replace("\n", " ").replace("|", "\\|")
            rel_link = d.get("rel_path") or d.get("target_md_rel_path") or ""
            link_cell = f"[📄 Open]({rel_link})" if rel_link else tid
            lines.append(f"| `{tid}` | {doc_type} | `{wi}` | {title} | {source} | {link_cell} |")
        lines.append("")

    # Actionable LLM Instruction Guide Prompt
    lines.append("#### 💡 Suggested Instruction for LLM Analysis:\n")
    if is_update and new_doc_entries:
        new_ids_str = ", ".join(f"`{d['tdoc_id']}`" for d in new_doc_entries[:15])
        if len(new_doc_entries) > 15:
            new_ids_str += f", and {len(new_doc_entries) - 15} more"
        lines.append(
            "> *Copy/paste this prompt to your LLM to perform an incremental analysis without re-analyzing previous documents:*\n"
            ">\n"
            f"> \"I have extended this meeting briefing with **{len(new_doc_entries)} additional TDocs** ({new_ids_str}) "
            "as detailed in the Changelog table above. Without repeating or regenerating the analysis of previously summarized "
            "contributions, please perform an incremental analysis focusing strictly on:\n"
            "> 1. **New Proposals:** Summarize the core proposals and architectural objectives in the newly added TDocs.\n"
            "> 2. **Cross-Impact & Consensus Shifts:** Identify if these new contributions alter, resolve, or contradict "
            "discussions/stances from the earlier documents.\n"
            "> 3. **Action Items & Stances:** Note any specific operator/vendor consensus changes or recommendations for "
            "our delegation based on these additions.\""
        )
    else:
        lines.append(
            "> *Copy/paste this prompt when initiating analysis on this briefing dossier:*\n"
            ">\n"
            f"> \"Analyze this 3GPP meeting briefing dossier ({total_docs} contributions). "
            "Synthesize the key debate clusters per Work Item, identify competing proposals, highlight vendor alignments, "
            "and suggest priority items for delegate preparation.\""
        )

    lines.append("</details>\n")
    return "\n".join(lines)


def format_local_wi_changelog_markdown(
    wi_name: str,
    added_docs_in_wi: List[Dict[str, Any]],
    timestamp: str,
) -> str:
    """Renders a compact collapsible changelog block inside a local Work Item index.md."""
    if not added_docs_in_wi:
        return ""

    n = len(added_docs_in_wi)
    lines = [
        "\n---\n",
        "<details>",
        f"<summary>🕒 <b>Work Item Incremental Update (+{n} TDoc(s) added on {timestamp})</b></summary>\n",
        f"The following {n} contribution(s) were added to this Work Item briefing a posteriori:\n",
    ]

    for d in sorted(added_docs_in_wi, key=lambda x: str(x.get("tdoc_id", ""))):
        tid = d.get("tdoc_id", "")
        title = d.get("title", "")
        source = d.get("source", "")
        cat = d.get("category", "")
        src_str = f" ({source})" if source else ""
        lines.append(f"- **`{tid}`** [{cat}]: *{title}*{src_str}")

    lines.append("\n</details>\n")
    return "\n".join(lines)
