"""Agreement Markdown export, LLM delta feed, master index, and consolidated briefing mixin."""

from __future__ import annotations

import logging
import os
import re
from pathlib import Path
from typing import Any, Dict, List, Optional

from core.utils.paths import to_extended_path
from modules.meetings.core.cohort_resolver import format_round_label
from modules.meetings.core.cross_wg_ls_tracker import CrossWgLiaisonTracker
from modules.meetings.core.exporters.export_manifest import ExportManifest
from modules.meetings.core.exporters.topic_digest_builder import (
    write_consolidated_briefing_markdown,
    write_master_index_markdown,
)
from modules.meetings.core.topic_decision import TopicDecision
from modules.meetings.core.topic_manager import MetaTopicDefinition
from ..checkup import check_and_refresh_dossier
from ..path_utils import _resolve_doc_server_url, sanitize_folder_name

logger = logging.getLogger(__name__)


class TopicDossierArtifactsMixin:
    """Writes RAN agreement files, delta feeds, master indexes, and consolidated topic briefings."""

    def _write_agreement_markdown(
        self,
        target_dir: Path,
        dec: TopicDecision,
        candidate_docs: List[Dict[str, Any]],
        seq: int = 1,
    ) -> str:
        """Exports a standalone Markdown document for a RAN Chairman's Notes text agreement."""
        wg_clean = sanitize_folder_name(dec.wg)
        mtg_clean = sanitize_folder_name(dec.meeting_number)
        meeting_folder_id = dec.meeting_folder_id or f"{wg_clean}_{mtg_clean}"

        agreements_dir = target_dir / wg_clean / meeting_folder_id / "Agreements"
        os.makedirs(to_extended_path(agreements_dir), exist_ok=True)

        if dec.tdoc_id:
            tid_clean = sanitize_folder_name(dec.tdoc_id)
            if seq > 1:
                filename = f"Agreement_{tid_clean}_{seq}.md"
            else:
                filename = f"Agreement_{tid_clean}.md"
        else:
            ai_clean = sanitize_folder_name(dec.ai or "General")
            filename = f"Agreement_AI_{ai_clean}_{seq:02d}.md"

        out_file = agreements_dir / filename
        rel_path = f"{wg_clean}/{meeting_folder_id}/Agreements/{filename}"

        bg_link = "—"
        bg_doc = None
        if dec.tdoc_id:
            bg_doc = next((doc for doc in candidate_docs if doc.get("tdoc_id") == dec.tdoc_id), None)
            if bg_doc and bg_doc.get("rel_path"):
                doc_rel = bg_doc["rel_path"]
                try:
                    rel_to_doc = os.path.relpath(target_dir / doc_rel, agreements_dir).replace("\\", "/")
                    bg_link = f"[{dec.tdoc_id}]({rel_to_doc}) *(Noted Proposal)*"
                except Exception:
                    bg_link = f"[{dec.tdoc_id}](../../{doc_rel}) *(Noted Proposal)*"
            else:
                bg_link = f"`{dec.tdoc_id}` *(Noted / Discussed)*"

        portal_url = dec.source_url or (
            _resolve_doc_server_url(dec.tdoc_id, bg_doc.get("meeting_dict") if bg_doc else None, bg_doc.get("row_data") if bg_doc else None)
            if dec.tdoc_id
            else ""
        )
        portal_link = f"[{dec.tdoc_id}]({portal_url})" if portal_url else "—"

        lines = [
            f"# 📜 3GPP Text Agreement: {dec.wg} #{dec.meeting_number} (AI {dec.ai or 'N/A'})",
            "",
            "| Metadata Attribute | Value |",
            "| :--- | :--- |",
            f"| **Working Group** | **{dec.wg}** |",
            f"| **Meeting** | #{dec.meeting_number} (`{meeting_folder_id}`, {dec.date or 'N/A'}) |",
            f"| **Agenda Item (AI)** | `{dec.ai or 'N/A'}` |",
            f"| **Decision Type** | **{dec.badge_icon} {dec.decision_type}** |",
            f"| **Status** | {dec.status or 'Agreed'} |",
            f"| **Target Specification** | {dec.target_spec or '—'} |",
            f"| **Related Work Item** | {dec.wi or '—'} |",
            f"| **Chairman's Notes Source** | {dec.source or 'Chairman Notes'} |",
            f"| **Background Contribution** | {bg_link} |",
            f"| **Official 3GPP Portal** | {portal_link} |",
            "",
            "---",
            "",
            "## 📝 Verbatim Consensus Text (Chairman's Notes)",
            "",
            "> [!NOTE] 3GPP RAN Decision Nature",
            "> In 3GPP RAN working groups (RAN1, RAN2, RAN3), the **verbatim text below represents the formal, binding consensus** reached by the delegation.",
            "> The cited contribution was \"Noted\" in the meeting report and serves as technical provenance/context.",
            "",
            "```text",
            dec.title_or_summary.strip(),
            "```",
            "",
            "### Formatted Excerpt",
            "",
        ]

        for paragraph in dec.title_or_summary.strip().split("\n"):
            p_clean = paragraph.strip()
            if p_clean:
                lines.append(f"> {p_clean}")
            else:
                lines.append(">")

        lines.extend([
            "",
            "---",
            "",
            "## 🔍 Context & Topic Traceability",
            f"- **Meta-Topic:** {self.topic.name}",
            f"- **Synchronized Round:** {format_round_label(dec.cohort_key)}",
            f"- **Meeting Record:** Documented under Agenda Item **{dec.ai or 'N/A'}** in the Chairman's Notes of 3GPP **{dec.wg}** Meeting **#{dec.meeting_number}**.",
            "",
        ])

        try:
            with open(to_extended_path(out_file), "w", encoding="utf-8") as f:
                f.write("\n".join(lines))
        except Exception as e:
            logger.warning(f"Failed to write agreement markdown {out_file}: {e}")

        return rel_path

    def _write_agreements_timeline(self, target_dir: Path, decisions: List[TopicDecision]):
        """Generates agreements_timeline.md unifying RAN text agreements and SA/CT pCRs."""
        lines = [
            f"# ⚖️ Unified Technical Decisions & Baseline: {self.topic.name}",
            "",
            "> [!NOTE]",
            "> This timeline unites technical decisions across Working Groups at an equal level:",
            "> - **RAN Working Groups (RAN1/2/3):** Substantive text agreements, working assumptions, and endorsed Way Forwards from Chairman's Notes.",
            "> - **SA & CT Groups (SA2, CT1, etc.):** Approved Pseudo-CRs (pCRs) to TRs and normative Change Requests (CRs) to Technical Specifications.",
            "",
            "| Meeting Date | WG | Type | Substantive Decision / Scope | Contribution | Target Spec |",
            "| :--- | :--- | :--- | :--- | :--- | :--- |",
        ]

        sorted_decisions = sorted(decisions, key=lambda d: (d.date or d.cohort_key, d.wg, d.decision_type))
        for dec in sorted_decisions:
            lines.append(dec.to_markdown_row())

        lines.append("")
        with open(to_extended_path(target_dir / "agreements_timeline.md"), "w", encoding="utf-8") as f:
            f.write("\n".join(lines))

    def _write_llm_delta_feed(
        self,
        target_dir: Path,
        session_id: int,
        session_timestamp: str,
        new_decisions: List[TopicDecision],
        new_tdocs: List[Dict[str, Any]],
        ls_tracker: CrossWgLiaisonTracker,
    ):
        """Writes an incremental delta feed optimized for LLM / AI digestion."""
        delta_dir = target_dir / "deltas"
        os.makedirs(to_extended_path(delta_dir), exist_ok=True)
        delta_file = delta_dir / f"delta_session_{session_id}.md"

        lines = [
            f"# 🧠 Incremental AI Delta Feed — Session #{session_id}",
            f"**Generated:** {session_timestamp} | **Topic:** {self.topic.name}",
            "",
            "> [!IMPORTANT] LLM Prompting Guidance",
            "> This document contains strictly the **new decisions and contributions** added during this update session.",
            "> Feed this file alongside `company_positions_e2e_prompt.md` and per-meeting `_company_positions_input.md` files for token-efficient analysis.",
            "",
            "## 🆕 Newly Added Technical Decisions",
            "",
        ]

        if new_decisions:
            lines.extend([
                "| Date | WG | Meeting | Decision Type | Summary / Scope | Contribution |",
                "| :--- | :--- | :--- | :--- | :--- | :--- |",
            ])
            for d in new_decisions:
                doc_link = f"[{d.tdoc_id or 'Agreement'}](../{d.dossier_rel_path})" if d.dossier_rel_path else (d.tdoc_id or "—")
                summary_clean = re.sub(r'[\r\n\t]+', ' ', d.title_or_summary).replace("|", "\\|")
                m_lbl = d.meeting_folder_id or f"{d.wg}_{d.meeting_number}"
                lines.append(f"| {d.date or d.cohort_key} | **{d.wg}** | {m_lbl} | {d.badge_icon} {d.decision_type} | {summary_clean} | {doc_link} |")
        else:
            lines.append("*No new decisions recorded in this session.*")

        lines.extend([
            "",
            "## 📄 Newly Exported / Restored Document Dossiers",
            "",
        ])
        if new_tdocs:
            lines.extend([
                "| TDoc | WG | Meeting | Type | Title | Relative Dossier Path |",
                "| :--- | :--- | :--- | :--- | :--- | :--- |",
            ])
            for doc in new_tdocs:
                tid = doc["tdoc_id"]
                rel_p = doc["rel_path"]
                m_lbl = doc.get("meeting_folder_id") or doc.get("meeting_number", "")
                title_c = re.sub(r'[\r\n\t]+', ' ', doc.get("title", "")).replace("|", "\\|")
                lines.append(f"| **{tid}** | {doc.get('wg')} | {m_lbl} | {doc.get('type')} | {title_c} | [{rel_p}](../{rel_p}) |")
        else:
            lines.append("*All relevant documents were already exported in earlier sessions.*")

        lines.append("")
        with open(to_extended_path(delta_file), "w", encoding="utf-8") as f:
            f.write("\n".join(lines))

    def _write_master_index(
        self,
        target_dir: Path,
        manifest: ExportManifest,
        all_decisions: List[TopicDecision],
        all_candidates: List[Dict[str, Any]],
        session_id: int,
        meeting_records: Optional[List[Dict[str, Any]]] = None,
        synthesis_status: Optional[str] = None,
    ):
        """Delegates master `index.md` generation to `write_master_index_markdown`."""
        recs = meeting_records or []
        total_stubs = sum(int(r.get("stub_count", 0) or 0) for r in recs)
        write_master_index_markdown(
            target_dir=target_dir,
            topic=self.topic,
            manifest=manifest,
            total_decisions=len(all_decisions),
            total_active_docs=len(all_candidates),
            stub_count=total_stubs,
            round_count=len(self.cohorts),
            session_id=session_id,
            scope_mode=self.export_scope_mode,
            meeting_records=recs,
            synthesis_status=synthesis_status,
            synthesis_state=manifest.synthesis_state,
        )

    def _write_consolidated_briefing(
        self,
        target_dir: Path,
        all_decisions: List[TopicDecision],
        all_candidates: List[Dict[str, Any]],
        meeting_records: Optional[List[Dict[str, Any]]] = None,
    ):
        """Delegates consolidated single-file briefing generation to `write_consolidated_briefing_markdown`."""
        write_consolidated_briefing_markdown(
            target_dir=target_dir,
            topic=self.topic,
            all_decisions=all_decisions,
            meeting_records=meeting_records or [],
            round_count=len(self.cohorts),
        )

    @staticmethod
    def check_and_refresh_dossier(
        target_dir: Path,
        topic: Optional[MetaTopicDefinition] = None,
    ) -> Dict[str, Any]:
        """Static alias for `check_and_refresh_dossier(target_dir, topic=topic)`."""
        return check_and_refresh_dossier(target_dir, topic=topic)
