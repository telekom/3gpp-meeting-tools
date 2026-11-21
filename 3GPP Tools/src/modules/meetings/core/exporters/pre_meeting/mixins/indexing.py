# --- File: src/modules/meetings/core/exporters/pre_meeting/mixins/indexing.py ---
"""Contributor matrix, local Work Item covers, master cover, and table rendering mixin."""

import datetime
import logging
from pathlib import Path
from typing import Dict, List, Optional

from core.utils.company_sanitizer import CompanySanitizer
from core.utils.paths import to_extended_path
from modules.meetings.core.exporters.post_meeting_exporter import (
    quote_markdown_path,
    safe_markdown_rel_link,
    sanitize_folder_name,
    slugify_anchor,
)
from modules.meetings.core.tdocs.tdoc_url_resolver import resolve_tdoc_download_url

logger = logging.getLogger("modules.meetings.core.exporters.pre_meeting_exporter")


class PreMeetingIndexingMixin:
    """Renders contributor matrices, WI briefings, master index.md, and summary tables."""

    def _render_contributor_matrix(self, docs_list: List[dict], max_companies: int = 25) -> str:
        """Computes and renders a Markdown table of top contributing companies."""
        comp_stats: Dict[str, Dict[str, int]] = {}

        for d in docs_list:
            cat = d.get("category", "")
            companies = d.get("sanitized_companies")
            if not companies:
                companies = ["Other"]
            for comp in companies:
                if comp not in comp_stats:
                    comp_stats[comp] = {"pCR": 0, "CR": 0, "DISC": 0, "LS_OUT": 0, "total": 0}
                if cat in comp_stats[comp]:
                    comp_stats[comp][cat] += 1
                comp_stats[comp]["total"] += 1

        if not comp_stats:
            return "*No contributor data available.*\n"

        sorted_comps = sorted(comp_stats.items(), key=lambda x: (x[1]["total"], x[0].lower()), reverse=True)
        if max_companies:
            sorted_comps = sorted_comps[:max_companies]

        headers = ["Contributor", "Total Docs", "pCR", "CR", "Discussion", "LS out"]
        lines = [
            "| " + " | ".join(headers) + " |",
            "|" + "|".join([":---"] + [":---:"] * (len(headers) - 1)) + "|",
        ]

        for comp, stats in sorted_comps:
            lines.append(
                f"| **{comp}** | **{stats['total']}** | {stats['pCR']} | {stats['CR']} | {stats['DISC']} | {stats['LS_OUT']} |"
            )

        return "\n".join(lines) + "\n"

    def _generate_local_wi_indices(self, filtered_entries: dict):
        """Writes an index.md briefing cover inside each Work Item directory."""
        for rel_name, wis in filtered_entries.items():
            rel_slug = sanitize_folder_name(rel_name)
            for wi_name, wi_data in wis.items():
                wi_slug = sanitize_folder_name(wi_name)
                wi_dir = self.export_dir / rel_slug / wi_slug
                wi_index_file = wi_dir / "index.md"

                info = wi_data["info"]
                wi_title = info.get("name") or wi_name
                wi_code = info.get("code")

                lines = []
                master_rel = safe_markdown_rel_link(self.export_dir / "index.md", wi_dir)
                lines.append(f"[⬅ Back to Master Briefing]({master_rel})\n")
                lines.append(f"# Work Item Pre-Meeting Briefing: {wi_name}\n")
                if wi_title and wi_title != wi_name:
                    lines.append(f"**Name:** {wi_title}  \n")
                if wi_code:
                    lines.append(f"**WI Code:** `{wi_code}`  \n")
                lines.append(f"**Release:** {rel_name}  \n")

                # Covered Agenda Items
                if wi_data["ais"]:
                    sorted_ais = sorted(list(wi_data["ais"]))
                    lines.append("\n### 📋 Covered Agenda Items\n")
                    for ai in sorted_ais:
                        desc = wi_data["ai_titles"].get(ai, "")
                        if desc:
                            lines.append(f"* **AI {ai}**: {desc}")
                        else:
                            lines.append(f"* **AI {ai}**")
                    lines.append("")

                # Contributor activity matrix for this WI
                all_wi_docs = [
                    d for cat_docs in wi_data["categories"].values() for d in cat_docs if d.get("category") != "LS_IN"
                ]
                if all_wi_docs:
                    lines.append("### 🏢 Top Contributors in this Work Item\n")
                    lines.append(self._render_contributor_matrix(all_wi_docs, max_companies=12))
                    lines.append("")

                lines.append("---\n")

                # Document Tables
                category_titles = [
                    ("LS_IN", "📥 Incoming Liaison Statements (LS in)"),
                    ("LS_OUT", "📤 Proposed Outgoing Liaison Statements (LS out)"),
                    ("pCR", "📝 Proposed Pseudo-CRs (pCR)"),
                    ("CR", "📄 Proposed Change Requests (CR)"),
                    ("DISC", "💬 Discussion, Evaluation & Status Papers (Disc / WID)"),
                ]

                has_docs = False
                for cat_key, cat_title in category_titles:
                    docs = wi_data["categories"].get(cat_key, [])
                    if docs:
                        has_docs = True
                        lines.append(f"### {cat_title}\n")
                        lines.append(self._render_markdown_table(docs, is_local_wi=True, base_dir=wi_dir))
                        lines.append("")

                if not has_docs:
                    lines.append("*No documents exported under this Work Item.*\n")

                # Local incremental changelog if documents were added to this WI
                if hasattr(self, "session_plan") and self.session_plan.is_incremental:
                    from modules.meetings.core.exporters.export_manifest import format_local_wi_changelog_markdown
                    added_in_wi = [
                        d for d in all_wi_docs
                        if str(d.get("tdoc_id", "")).strip().upper() in self.session_plan.new_tdocs
                    ]
                    if added_in_wi:
                        lines.append(format_local_wi_changelog_markdown(
                            wi_name, added_in_wi, self.session_plan.session_timestamp
                        ))

                try:
                    with open(to_extended_path(wi_index_file), "w", encoding="utf-8") as f:
                        f.write("\n".join(lines))
                except Exception as e:
                    logger.error(f"[PreMeetingExporter] Error writing WI briefing index for {wi_name}: {e}")

    def _generate_master_index(self, filtered_entries: dict):
        """Generates the master index.md cover page with executive summary and contributor landscape."""
        master_file = self.export_dir / "index.md"
        wg = str(self.mtg_info.get("wg_name", "3GPP")).strip()
        mtg = str(self.mtg_info.get("meeting_number", "")).strip()
        mtg_title = f"{wg} #{mtg}" if (wg and mtg) else "3GPP Meeting"

        cat_counts = {"LS_IN": 0, "LS_OUT": 0, "pCR": 0, "CR": 0, "DISC": 0}
        total_unique_tdocs = set()
        total_wis = set()
        all_company_docs = []

        for rel in filtered_entries.values():
            for wi_name, wi_data in rel.items():
                if wi_name != "General":
                    total_wis.add(wi_name)
                for cat, docs in wi_data["categories"].items():
                    cat_counts[cat] = cat_counts.get(cat, 0) + len(docs)
                    for d in docs:
                        total_unique_tdocs.add(d["tdoc_id"])
                        if cat != "LS_IN":
                            all_company_docs.append(d)

        now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
        scope_labels = {
            "followed": "🎯 Followed Work Items only",
            "starred": "⭐ Starred Work Items only",
            "tracked": "🎯 Followed & ⭐ Starred Work Items",
            "visible": "🔍 Current Filtered Table View",
            "whole": "🌐 Whole Meeting (All Work Items & Contributions)",
            "custom": "📋 Custom Work Item & Agenda Item Selection",
        }
        scope_str = scope_labels.get(self.scope_mode, self.scope_mode)

        lines = []
        lines.append(f"# {mtg_title} — Pre-Meeting Briefing & Dossier Report\n")
        lines.append(f"**Meeting:** {mtg_title}  \n")
        lines.append(f"**Generated:** {now_str}  \n")
        lines.append(f"**Export Scope:** {scope_str}  \n")
        lines.append(f"**Work Items Covered:** {len(total_wis)}  \n")
        lines.append(f"**Total Active Contribution Proposals:** {len(total_unique_tdocs)}  \n")
        if self.deduplicate_revisions:
            lines.append("**Revision Mode:** 🔄 Deduplicated (latest active revision exported per chain)  \n")
        if self.company_filter:
            lines.append(f"**Company Filter Active:** {', '.join(sorted(self.company_filter))}  \n")
        lines.append("\n---\n")

        # Executive Summary Table
        lines.append("## 📊 Executive Summary\n")
        lines.append("| Category | Total Active Proposals | Scope Description |")
        lines.append("|:---|:---:|:---|")
        lines.append(f"| 📥 Incoming LS (`LS in`) | **{cat_counts['LS_IN']}** | Received incoming Liaison Statements |")
        lines.append(f"| 📤 Proposed Outgoing LS (`LS out`) | **{cat_counts['LS_OUT']}** | Proposed candidate reply & new outgoing LS drafts |")
        lines.append(f"| 📝 Proposed Pseudo-CRs (`pCR`) | **{cat_counts['pCR']}** | Proposed specification architecture drafts |")
        lines.append(f"| 📄 Proposed Change Requests (`CR`) | **{cat_counts['CR']}** | Proposed specification change proposals |")
        lines.append(f"| 💬 Discussion & Evaluation (`Disc`) | **{cat_counts['DISC']}** | Study reports, Key Issue solutions & discussion papers |")
        lines.append(f"| **Total Document Proposals** | **{sum(cat_counts.values())}** | *(Unique Active TDocs: {len(total_unique_tdocs)})* |")
        lines.append("\n---\n")

        # Contributor Landscape
        if all_company_docs:
            lines.append("## 🏢 Contributor Activity Landscape\n")
            lines.append("Summary of top submitting companies across all proposed CRs, pCRs, Discussion papers, and draft LS replies:\n")
            lines.append(self._render_contributor_matrix(all_company_docs, max_companies=20))
            lines.append("\n---\n")

        # Table of Contents
        lines.append("## 📑 Table of Contents\n")
        for rel_name in sorted(filtered_entries.keys()):
            rel_slug = sanitize_folder_name(rel_name)
            lines.append(f"* **{rel_name}**")
            for wi_name in sorted(filtered_entries[rel_name].keys()):
                wi_slug = sanitize_folder_name(wi_name)
                wi_anchor = f"wi-{slugify_anchor(wi_slug)}"
                wi_index_file = self.export_dir / rel_slug / wi_slug / "index.md"
                if self.generate_local_indices:
                    wi_link = safe_markdown_rel_link(wi_index_file, self.export_dir)
                    lines.append(f"  * [{wi_name}](#{wi_anchor}) ([Briefing Cover]({wi_link}))")
                else:
                    lines.append(f"  * [{wi_name}](#{wi_anchor})")
        lines.append("\n---\n")

        # Per-Release / Per-WI Detailed Tables
        for rel_name in sorted(filtered_entries.keys()):
            rel_slug = sanitize_folder_name(rel_name)
            lines.append(f"## Release: {rel_name}\n")

            for wi_name in sorted(filtered_entries[rel_name].keys()):
                wi_data = filtered_entries[rel_name][wi_name]
                wi_slug = sanitize_folder_name(wi_name)
                wi_anchor = f"wi-{slugify_anchor(wi_slug)}"

                info = wi_data["info"]
                wi_title = info.get("name") or wi_name
                wi_code = info.get("code")

                lines.append(f'<a id="{wi_anchor}" name="{wi_anchor}"></a>\n')
                lines.append(f"### 🎯 Work Item: `{wi_name}`\n")
                if wi_title and wi_title != wi_name:
                    lines.append(f"**Title:** {wi_title}  \n")
                if wi_code:
                    lines.append(f"**WI Code:** `{wi_code}`  \n")

                wi_index_file = self.export_dir / rel_slug / wi_slug / "index.md"
                if self.generate_local_indices:
                    wi_link = safe_markdown_rel_link(wi_index_file, self.export_dir)
                    lines.append(f"**Briefing Subfolder:** [`{rel_slug}/{wi_slug}/`]({wi_link})  \n")
                else:
                    lines.append(f"**Briefing Subfolder:** `{rel_slug}/{wi_slug}/`  \n")

                if wi_data["ais"]:
                    ai_bullets = []
                    for ai in sorted(list(wi_data["ais"])):
                        desc = wi_data["ai_titles"].get(ai, "")
                        ai_bullets.append(f"`AI {ai}`: {desc}" if desc else f"`AI {ai}`")
                    lines.append(f"**Covered Agenda Items:** {'; '.join(ai_bullets)}  \n")

                lines.append("")

                category_titles = [
                    ("LS_IN", "📥 Incoming Liaison Statements (LS in)"),
                    ("LS_OUT", "📤 Proposed Outgoing Liaison Statements (LS out)"),
                    ("pCR", "📝 Proposed Pseudo-CRs (pCR)"),
                    ("CR", "📄 Proposed Change Requests (CR)"),
                    ("DISC", "💬 Discussion, Evaluation & Status Papers (Disc / WID)"),
                ]

                has_docs = False
                for cat_key, cat_title in category_titles:
                    docs = wi_data["categories"].get(cat_key, [])
                    if docs:
                        has_docs = True
                        lines.append(f"#### {cat_title}\n")
                        lines.append(self._render_markdown_table(docs, is_local_wi=False, base_dir=self.export_dir))
                        lines.append("")

                if not has_docs:
                    lines.append("*No documents exported in this category.*\n")

                lines.append("")

        # Append Changelog & Incremental LLM Guide and update persistent manifest
        if hasattr(self, "session_plan"):
            from modules.meetings.core.exporters.export_manifest import (
                format_master_changelog_markdown,
                save_manifest,
                ExportManifest,
            )
            all_added = []
            all_updated = []
            seen_added = set()
            seen_updated = set()

            for rel in filtered_entries.values():
                for wi in rel.values():
                    for cat_docs in wi["categories"].values():
                        for d in cat_docs:
                            tid = str(d.get("tdoc_id", "")).strip().upper()
                            if tid in self.session_plan.new_tdocs and tid not in seen_added:
                                seen_added.add(tid)
                                all_added.append(d)
                            elif tid in self.session_plan.tdocs_to_update and tid not in seen_updated:
                                seen_updated.add(tid)
                                all_updated.append(d)

            all_added_entries = []
            for rel_name, wis in filtered_entries.items():
                for wi_name, wi_data in wis.items():
                    for cat_name, cat_docs in wi_data["categories"].items():
                        for d in cat_docs:
                            tid = str(d.get("tdoc_id", "")).strip().upper()
                            if tid in self.session_plan.new_tdocs or tid in self.session_plan.tdocs_to_update:
                                entry_copy = dict(d)
                                entry_copy["release"] = rel_name
                                entry_copy["wi"] = wi_name
                                entry_copy["category"] = cat_name
                                entry_copy["rel_path"] = d.get("target_md_rel_path", "")
                                all_added_entries.append(entry_copy)

            if getattr(self, "manifest", None) is None:
                self.manifest = ExportManifest(
                    meeting=mtg_title,
                    export_type="pre_meeting",
                    created_at=self.session_plan.session_timestamp,
                )

            current_session = {
                "session_id": self.session_plan.session_id,
                "timestamp": self.session_plan.session_timestamp,
                "label": "Initial Export" if self.session_plan.session_id == 1 else f"Incremental Update #{self.session_plan.session_id - 1}",
                "scope_mode": self.scope_mode,
                "added_tdocs": sorted(list(self.session_plan.new_tdocs)),
                "updated_tdocs": sorted(list(self.session_plan.tdocs_to_update)),
                "total_tdocs": len(self.manifest.tdocs) + len(self.session_plan.new_tdocs),
            }

            changelog_block = format_master_changelog_markdown(
                self.manifest,
                current_session,
                new_doc_entries=all_added,
                updated_doc_entries=all_updated,
                base_dir=self.export_dir,
            )
            lines.append(changelog_block)

            self.manifest.record_session(
                session_id=self.session_plan.session_id,
                timestamp=self.session_plan.session_timestamp,
                scope_mode=self.scope_mode,
                added_tdocs=list(self.session_plan.new_tdocs),
                updated_tdocs=list(self.session_plan.tdocs_to_update),
                added_entries=all_added_entries,
            )
            save_manifest(self.export_dir, self.manifest)

        try:
            with open(to_extended_path(master_file), "w", encoding="utf-8") as f:
                f.write("\n".join(lines))
        except Exception as e:
            logger.error(f"[PreMeetingExporter] Error writing master index.md: {e}", exc_info=True)

    def _render_markdown_table(self, docs: List[dict], is_local_wi: bool = False, base_dir: Optional[Path] = None) -> str:
        """Renders standard markdown table with Pre-Meeting status, stance, and multi-reply indicators."""
        if not docs:
            return "*None*\n"

        if base_dir is None:
            base_dir = self.export_dir

        has_stance = any(
            (d.get("my_status") and d.get("my_status") != "⚪ Neutral") or d.get("my_notes")
            for d in docs
        )

        headers = ["TDoc", "Type", "Status", "AI", "Title", "Source"]
        if has_stance and self.include_user_notes:
            headers.append("Stance")
        headers.extend(["Cross-WI", "Links"])

        lines = [
            "| " + " | ".join(headers) + " |",
            "|" + "|".join([":---"] * len(headers)) + "|",
        ]

        for d in docs:
            tdoc_id = d["tdoc_id"]
            target_abs = d.get("target_md_abs_path")
            if target_abs and base_dir:
                link_path = safe_markdown_rel_link(target_abs, base_dir)
            elif is_local_wi:
                link_path = quote_markdown_path(d.get("wi_local_rel_path", ""))
            else:
                link_path = quote_markdown_path(d.get("target_md_rel_path", ""))

            link_str = f"[{tdoc_id}]({link_path})" if link_path else tdoc_id

            title_safe = str(d.get("title", "")).replace("\n", " ").replace("\r", "").replace("|", "\\|")
            source_safe = CompanySanitizer.format_source_markdown_table_cell(d.get("source", ""))
            cross_wi_str = ", ".join(f"`{w}`" for w in d.get("cross_wis", [])) if d.get("cross_wis") else "—"

            raw_status = str(d.get("status", "")).strip()
            cat = d.get("category", "")

            # Format status with predecessor or reply notes
            if d.get("is_reserved"):
                status_cell = "⏳ Reserved"
            elif cat == "LS_IN" and d.get("replies"):
                rep_items = []
                for rep_id in d["replies"]:
                    rep_abs = getattr(self, "exported_tdoc_paths", {}).get(rep_id)
                    if rep_abs and base_dir:
                        r_rel = safe_markdown_rel_link(rep_abs, base_dir)
                        rep_items.append(f"[{rep_id}]({r_rel})")
                    else:
                        rep_items.append(rep_id)
                rep_str = ", ".join(rep_items)
                status_cell = f"replies: {rep_str}" if not raw_status else f"{raw_status} (replies: {rep_str})"
            elif cat == "LS_OUT" and d.get("reply_to_ids"):
                in_items = []
                for in_id in d["reply_to_ids"]:
                    in_abs = getattr(self, "exported_tdoc_paths", {}).get(in_id)
                    if in_abs and base_dir:
                        i_rel = safe_markdown_rel_link(in_abs, base_dir)
                        in_items.append(f"[{in_id}]({i_rel})")
                    else:
                        in_items.append(in_id)
                in_str = ", ".join(in_items)
                status_cell = f"reply to {in_str}" if not raw_status else f"{raw_status} (reply to {in_str})"
            elif d.get("supersedes"):
                preds_str = ", ".join(d["supersedes"])
                status_cell = f"{raw_status} (rev of {preds_str})" if raw_status else f"rev of {preds_str}"
            else:
                status_cell = raw_status or "Available"

            tdoc_url = d.get("tdoc_url") or resolve_tdoc_download_url(
                tdoc_id,
                docs_url=self.docs_ftp_url,
                revisions_url=self.revisions_url,
                meta=d.get("row_data"),
            )
            if link_path and tdoc_url and not d.get("is_reserved"):
                links_cell = f"[Dossier]({link_path}) · [📥 3GPP]({tdoc_url})"
            elif link_path:
                links_cell = f"[Dossier]({link_path})"
            elif tdoc_url:
                links_cell = f"[📥 3GPP]({tdoc_url})"
            else:
                links_cell = "—"

            row_cells = [
                link_str,
                str(d.get("type", "")),
                status_cell,
                str(d.get("ai", "")),
                title_safe,
                source_safe,
            ]

            if has_stance and self.include_user_notes:
                st = d.get("my_status", "")
                nt = d.get("my_notes", "")
                if st and st != "⚪ Neutral":
                    stance_cell = f"**{st}**" + (f" (*{nt}*)" if nt else "")
                elif nt:
                    stance_cell = f"📝 *{nt}*"
                else:
                    stance_cell = "—"
                row_cells.append(stance_cell.replace("\n", " ").replace("|", "\\|"))

            row_cells.extend([
                cross_wi_str,
                links_cell,
            ])
            lines.append("| " + " | ".join(row_cells) + " |")

        return "\n".join(lines)
