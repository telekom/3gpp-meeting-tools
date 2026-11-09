# --- File: src/modules/meetings/core/cross_wg_ls_tracker.py ---
"""
Cross-WG Liaison Statement (LS) Dialogue & Thread Tracker.

Traces inter-group questions, guidance, and reply chains between
Working Groups (SA2, RAN2, RAN3, CT1, SA3, etc.) across meeting cycles.
"""

import logging
import re
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Set, Tuple

logger = logging.getLogger(__name__)

_TDOC_EXTRACT_RE = re.compile(r'\b([A-Z][A-Z0-9]{1,3}-\d{4,8})\b')


@dataclass
class LiaisonThreadNode:
    tdoc_id: str
    wg: str
    meeting_number: str
    cycle: str
    direction: str             # "IN" or "OUT"
    title: str
    source: str = ""
    status: str = ""
    target_groups: List[str] = field(default_factory=list)
    reply_to_ids: List[str] = field(default_factory=list)
    reply_in_ids: List[str] = field(default_factory=list)
    dossier_rel_path: str = ""
    replies: List["LiaisonThreadNode"] = field(default_factory=list)


class CrossWgLiaisonTracker:
    """Tracks and correlates Liaison Statements across Working Groups."""

    def __init__(self):
        self.nodes: Dict[str, LiaisonThreadNode] = {}
        self.roots: List[LiaisonThreadNode] = []

    def ingest_ls_tdoc(
        self,
        tdoc_id: str,
        wg: str,
        meeting_number: str,
        cycle: str,
        tdoc_type: str,
        title: str,
        source: str = "",
        status: str = "",
        reply_to: str = "",
        reply_in: str = "",
        dossier_rel_path: str = "",
    ) -> LiaisonThreadNode:
        """Registers a Liaison Statement TDoc into the tracker."""
        tid = str(tdoc_id or "").strip().upper()
        norm_type = "OUT" if "out" in str(tdoc_type).lower() else "IN"

        # Extract explicit reply-to TDocs
        raw_replies = []
        if reply_to:
            raw_replies.extend(_TDOC_EXTRACT_RE.findall(str(reply_to).upper()))
        # Check title for reply patterns like "(reply to S2-2501234)"
        title_replies = _TDOC_EXTRACT_RE.findall(str(title).upper())
        for tr in title_replies:
            if tr != tid and tr not in raw_replies:
                raw_replies.append(tr)

        # Extract explicit reply-in TDocs
        raw_reply_in = []
        if reply_in:
            raw_reply_in.extend(_TDOC_EXTRACT_RE.findall(str(reply_in).upper()))

        # Target WGs extraction from title (e.g. "to RAN2, CT1")
        targets = []
        to_match = re.search(r'\bto\s+([A-Za-z0-9,\s/&]+)', str(title), re.IGNORECASE)
        if to_match:
            candidate_str = to_match.group(1).upper()
            for cand in ["SA1", "SA2", "SA3", "SA4", "SA5", "RAN1", "RAN2", "RAN3", "RAN4", "CT1", "CT3", "CT4"]:
                if cand in candidate_str:
                    targets.append(cand)

        node = LiaisonThreadNode(
            tdoc_id=tid,
            wg=str(wg or "").strip().upper(),
            meeting_number=str(meeting_number or "").strip(),
            cycle=str(cycle or "").strip(),
            direction=norm_type,
            title=str(title or "").strip(),
            source=str(source or "").strip(),
            status=str(status or "").strip(),
            target_groups=targets,
            reply_to_ids=raw_replies,
            reply_in_ids=raw_reply_in,
            dossier_rel_path=str(dossier_rel_path or "").replace("\\", "/"),
        )
        self.nodes[tid] = node
        return node

    def build_threads(self) -> None:
        """Resolves parent-child reply relationships across the ingested nodes."""
        self.roots = []
        claimed_as_child = set()

        # Connect children via reply_to_ids
        for tid, node in self.nodes.items():
            for parent_id in node.reply_to_ids:
                if parent_id in self.nodes:
                    parent_node = self.nodes[parent_id]
                    if node not in parent_node.replies:
                        parent_node.replies.append(node)
                    claimed_as_child.add(tid)

        # Roots are nodes that were not claimed as replies to another known node
        for tid, node in self.nodes.items():
            if tid not in claimed_as_child:
                self.roots.append(node)

        # Sort roots chronologically by cycle and meeting
        self.roots.sort(key=lambda n: (n.cycle, n.meeting_number, n.tdoc_id))

    def generate_markdown_report(self) -> str:
        """Generates a structured, navigable Markdown summary of the LS dialogue."""
        self.build_threads()

        lines = [
            "# 📬 Cross-WG Liaison Statement (LS) Dialogue Tracker",
            "",
            "> [!NOTE]",
            "> This report traces inter-group technical correspondence, questions, and replies exchanged across Working Groups.",
            "",
            "## 🧵 Dialogue Threads",
            "",
        ]

        if not self.roots:
            lines.append("*No cross-WG Liaison Statements detected in this scope.*")
            lines.append("")
            return "\n".join(lines)

        for root in self.roots:
            lines.extend(self._render_node_tree(root, depth=0))
            lines.append("")

        lines.extend([
            "---",
            "",
            "## 📋 Comprehensive Liaison Statement Inventory",
            "",
            "| Cycle | WG | Meeting | TDoc | Dir | Title | Target Groups | Status | Dossier |",
            "| :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- |",
        ])

        all_sorted = sorted(self.nodes.values(), key=lambda n: (n.cycle, n.wg, n.tdoc_id))
        for n in all_sorted:
            doc_link = f"[{n.tdoc_id}]({n.dossier_rel_path})" if n.dossier_rel_path else n.tdoc_id
            dir_badge = "📤 **OUT**" if n.direction == "OUT" else "📥 **IN**"
            targets_str = ", ".join(n.target_groups) if n.target_groups else "—"
            clean_title = re.sub(r'[\r\n\t]+', ' ', n.title).replace("|", "\\|")
            if len(clean_title) > 90:
                clean_title = clean_title[:87] + "..."
            lines.append(
                f"| {n.cycle} | **{n.wg}** | {n.meeting_number} | {doc_link} | {dir_badge} | {clean_title} | {targets_str} | {n.status or '—'} | [Dossier]({n.dossier_rel_path}) |"
                if n.dossier_rel_path else
                f"| {n.cycle} | **{n.wg}** | {n.meeting_number} | {doc_link} | {dir_badge} | {clean_title} | {targets_str} | {n.status or '—'} | — |"
            )

        lines.append("")
        return "\n".join(lines)

    def _render_node_tree(self, node: LiaisonThreadNode, depth: int = 0) -> List[str]:
        prefix = "  " * depth + ("└─ " if depth > 0 else "### ")
        dir_icon = "📤" if node.direction == "OUT" else "📥"
        link_str = f"[{node.tdoc_id}]({node.dossier_rel_path})" if node.dossier_rel_path else node.tdoc_id
        targets_str = f" *(to {', '.join(node.target_groups)})*" if node.target_groups else ""
        
        line = f"{prefix}{dir_icon} **{node.wg}** {node.meeting_number} ({node.cycle}): {link_str} — *\"{node.title}\"*{targets_str}"
        lines = [line]

        for child in node.replies:
            lines.extend(self._render_node_tree(child, depth=depth + 1))

        return lines
