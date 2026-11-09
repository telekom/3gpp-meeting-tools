# --- File: src/modules/meetings/core/exporters/pre_meeting_exporter.py ---
"""
pre_meeting_exporter.py - Pre-Meeting Report & Markdown Dossier Exporter.

Generates an organized, executive-grade Pre-Meeting Markdown Dossier tailored for
delegates preparing for an upcoming 3GPP meeting or analyzing the technical input landscape:
- Captures active company contributions (Discussion papers, WID/SIDs, proposed CRs, proposed pCRs)
  and incoming liaison statements (LS in).
- Handles multi-proposal LS replies: associates incoming LS in documents with all candidate
  outgoing LS reply proposals (1-to-many mapping) and cross-links competing proposals.
- Intelligent revision deduplication: traces pre-meeting revision chains (e.g. S2-260100 -> S2-260500)
  to export only the latest active non-withdrawn draft, annotating predecessor provenance.
- Contributor activity matrix: computes company submission volume across WIs and categories.
- Personal triage embedding: integrates delegate stances ('🟢 Support', '🔴 Object', '🔵 My TDoc')
  and notes from user_tdocs.db into dossier callouts and index tables.
- Graceful handling of unuploaded 'Reserved' TDocs without stalling parallel downloads.
"""

from concurrent.futures import ThreadPoolExecutor, as_completed
import datetime
import logging
import os
import re
import shutil
from pathlib import Path
from typing import Dict, List, Optional, Set, Any, Tuple

from core.utils.company_sanitizer import CompanySanitizer
from core.utils.paths import to_extended_path
from modules.meetings.core.tdocs.tdoc_file_handler import TDocFileHandler
from modules.meetings.core.tdocs.tdoc_url_resolver import resolve_tdoc_download_url
from modules.meetings.core.tdocs.tdoc_markdown_converter import (
    sanitize_filename,
    is_conversion_stale,
    resolve_revision_mode_for_type,
    CONVERTER_VERSION,
)
from modules.meetings.ui.tdocs_models import TDocsTableModel
from modules.meetings.core.exporters.post_meeting_exporter import (
    PostMeetingExporterThread,
    sanitize_folder_name,
    safe_relpath,
    quote_markdown_path,
    safe_markdown_rel_link,
    slugify_anchor,
)

logger = logging.getLogger(__name__)


class PreMeetingExporterThread(PostMeetingExporterThread):
    """
    Asynchronous worker to generate an organized, executive-grade Pre-Meeting Markdown Briefing & Dossier.
    Structure:
      Export/PreMeeting_<WG>_<Meeting>_<Timestamp>/
        ├── index.md (Master Cover, Contributor Matrix, LS Clusters & Summary Tables)
        ├── <Release>/
        │   └── <WorkItem>/
        │       ├── index.md (WI Specific Briefing Cover & Contributor Breakdown)
        │       ├── LS_IN/      (Incoming LS + media/)
        │       ├── LS_OUT/     (Candidate Draft LS Replies + media/)
        │       ├── pCR/        (Proposed Pseudo-CRs + media/)
        │       ├── CR/         (Proposed Change Requests + media/)
        │       └── Discussion/ (Discussion, Evaluation & Status Papers + media/)
        └── General/
            ├── LS_IN/          (Overarching incoming LS without a specific WI)
            └── LS_OUT/         (Candidate LS Replies answering General LS)
    """

    def __init__(
        self,
        meeting_dir: Path,
        all_rows: list,
        visible_rows: list,
        agenda_map: dict,
        mtg_info: dict,
        docs_ftp_url: str,
        revisions_url: str,
        scope_mode: str = "followed",  # "followed", "starred", "tracked", "visible", "whole", "custom"
        custom_wis: Optional[Set[str]] = None,
        custom_ais: Optional[Set[str]] = None,
        custom_wi_ais: Optional[Dict[str, Set[str]]] = None,
        followed_tokens: Optional[Set[str]] = None,
        starred_tokens: Optional[Set[str]] = None,
        include_ls_in: bool = True,
        include_pcr: bool = True,
        include_cr: bool = True,
        include_ls_out: bool = True,
        include_disc: bool = True,  # Default True for Pre-Meeting!
        include_general_ls: bool = True,
        deduplicate_revisions: bool = True,  # Export latest active revision per chain
        include_reserved_in_index: bool = True,  # List 'Reserved' TDocs in tables without downloading
        include_postponed_or_treated: bool = True,  # Mid-meeting tolerance (keep full input corpus)
        company_filter: Optional[Set[str]] = None,  # Filter by specific companies
        include_user_notes: bool = True,  # Embed My Status & My Notes into dossiers & index
        generate_local_indices: bool = True,
        download_missing: bool = True,
        force_reconvert: bool = False,
        compile_dossier: bool = False,
        output_dir: Optional[Path] = None,
        parent=None,
    ):
        super().__init__(
            meeting_dir=meeting_dir,
            all_rows=all_rows,
            visible_rows=visible_rows,
            agenda_map=agenda_map,
            mtg_info=mtg_info,
            docs_ftp_url=docs_ftp_url,
            revisions_url=revisions_url,
            scope_mode=scope_mode,
            custom_wis=custom_wis,
            custom_ais=custom_ais,
            custom_wi_ais=custom_wi_ais,
            followed_tokens=followed_tokens,
            starred_tokens=starred_tokens,
            include_ls_in=include_ls_in,
            include_pcr=include_pcr,
            include_cr=include_cr,
            include_ls_out=include_ls_out,
            include_disc=include_disc,
            include_general_ls=include_general_ls,
            generate_local_indices=generate_local_indices,
            download_missing=download_missing,
            force_reconvert=force_reconvert,
            compile_dossier=compile_dossier,
            output_dir=output_dir,
            parent=parent,
        )
        self.deduplicate_revisions = deduplicate_revisions
        self.include_reserved_in_index = include_reserved_in_index
        self.include_postponed_or_treated = include_postponed_or_treated
        self.company_filter = {str(c).strip().lower() for c in (company_filter or set()) if str(c).strip()}
        self.include_user_notes = include_user_notes

        # Pre-Meeting specific default output directory if not explicitly overridden
        if not output_dir:
            timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S")
            wg = str(self.mtg_info.get("wg_name", "3GPP")).strip().replace(" ", "_")
            mtg = str(self.mtg_info.get("meeting_number", "")).strip()
            folder_name = f"PreMeeting_{wg}_{mtg}_{timestamp}" if (wg or mtg) else f"PreMeeting_{timestamp}"
            self.export_dir = self.meeting_dir / "Export" / folder_name

    def run(self):
        try:
            self.progress.emit("Initializing Work Items knowledge base...")
            wi_db_info = self._load_work_items_info()

            self.progress.emit("Analyzing pre-meeting revision lineages and contributions...")
            filtered_entries = self._filter_and_group_tdocs(wi_db_info)

            if not filtered_entries:
                self.finished.emit(False, "No TDocs match the selected pre-meeting export criteria.")
                return

            total_entries = sum(
                len(docs)
                for rel in filtered_entries.values()
                for wi in rel.values()
                for docs in wi["categories"].values()
            )
            self.progress.emit(f"Found {total_entries} matching pre-meeting contribution(s). Exporting dossiers...")

            os.makedirs(to_extended_path(self.export_dir), exist_ok=True)

            # Incremental Manifest & Session Planning
            from modules.meetings.core.exporters.export_manifest import (
                load_or_reconstruct_manifest,
                plan_incremental_session,
                merge_manifest_into_tree,
                save_manifest,
                format_master_changelog_markdown,
                format_local_wi_changelog_markdown,
                ExportManifest,
            )

            self.manifest = load_or_reconstruct_manifest(
                self.export_dir, self.mtg_info, self.all_rows, export_type="pre_meeting"
            )
            self.session_plan = plan_incremental_session(
                self.manifest, filtered_entries, self.export_dir, force_reconvert=self.force_reconvert
            )

            if self.session_plan.is_incremental:
                logger.info(
                    f"[PreMeetingExporter] Incremental export session #{self.session_plan.session_id}: "
                    f"{len(self.session_plan.new_tdocs)} new, {len(self.session_plan.tdocs_to_skip)} skipped, "
                    f"{len(self.session_plan.preserved_tdocs)} preserved."
                )
                filtered_entries = merge_manifest_into_tree(self.manifest, filtered_entries, self.export_dir)

            # Convert and assemble all dossiers
            self._export_all_dossiers(filtered_entries)

            if self.isInterruptionRequested():
                self.finished.emit(False, "Export cancelled by user.")
                return

            # Generate index.md files
            self.progress.emit("Generating Pre-Meeting Markdown briefing reports...")
            if self.generate_local_indices:
                self._generate_local_wi_indices(filtered_entries)

            self._generate_master_index(filtered_entries)

            # Compile consolidated Navigable Dossier (HTML & PDF)
            if self.compile_dossier and not self.isInterruptionRequested():
                try:
                    self.progress.emit("Compiling consolidated Navigable Dossier (HTML & PDF)...")
                    self._compile_consolidated_dossier(filtered_entries)
                except Exception as doss_err:
                    logger.warning(f"[PreMeetingExporter] Consolidated dossier generation warning: {doss_err}")

            if hasattr(self, "session_plan") and self.session_plan.is_incremental:
                preserved_count = len(self.session_plan.tdocs_to_skip) + len(self.session_plan.preserved_tdocs)
                msg = (
                    f"Successfully updated Pre-Meeting Briefing & Dossier (Update #{self.session_plan.session_id - 1}):\n"
                    f"• {len(self.session_plan.new_tdocs)} TDoc(s) added\n"
                    f"• {preserved_count} TDoc(s) preserved\n\n"
                    f"Location:\n{self.export_dir}"
                )
            else:
                msg = f"Successfully generated Pre-Meeting Briefing & Dossier in:\n{self.export_dir}"
            self.finished.emit(True, msg)

        except Exception as e:
            logger.error(f"[PreMeetingExporter] Critical failure during export: {e}", exc_info=True)
            self.finished.emit(False, str(e))

    def _compile_consolidated_dossier(self, filtered_entries: dict):
        """Compiles non-reserved deliverables into master navigable HTML and PDF dossier."""
        unique_docs = {}
        for rel in filtered_entries.values():
            for wi in rel.values():
                for cat_docs in wi["categories"].values():
                    for doc in cat_docs:
                        tid = doc["tdoc_id"]
                        # Skip unuploaded reserved documents from PDF/HTML compilation
                        if doc.get("is_reserved") and not self._find_word_doc(self.meeting_dir / tid, tid):
                            continue
                        if tid not in unique_docs:
                            unique_docs[tid] = doc["row_data"]

        if not unique_docs:
            return

        from modules.meetings.core.exporters.dossier_exporter import DossierExporterThread

        dossier_worker = DossierExporterThread(
            meeting_dir=self.meeting_dir,
            tdocs_list=list(unique_docs.values()),
            docs_ftp_url=self.docs_ftp_url,
            revisions_url=self.revisions_url,
            mtg_info=self.mtg_info,
            export_format="both",
            include_prompt=True,
            embed_base64_images=True,
            download_missing=False,
            force_reconvert=False,
            output_dir=self.export_dir,
        )
        dossier_worker.run()

    def _build_revision_lineage_map(self) -> Tuple[Set[str], Dict[str, List[str]], Dict[str, str]]:
        """
        Builds complete revision lineage relationships across all meeting rows:
        - superseded_set: Set of TDoc IDs that have a newer active (non-withdrawn) revision in all_rows.
        - predecessors_map: latest_leaf_tdoc_id -> List of superseded predecessor TDoc IDs in chronological order.
        - leaf_map: tdoc_id -> latest_active_leaf_tdoc_id
        """
        tdoc_map = {str(r.get("TDoc", "")).strip().upper(): r for r in self.all_rows if r.get("TDoc")}

        # Build forward edges: parent -> list of child revisions
        forward_edges: Dict[str, Set[str]] = {}
        for tdoc_id, row in tdoc_map.items():
            # 1. Inspect 'Revised to'
            rev_to = str(row.get("Revised to", "")).strip()
            if rev_to and rev_to.lower() != "none":
                for cand in re.findall(r'[A-Za-z0-9]+-\d+(?:r\d+[a-zA-Z]?)?', rev_to):
                    cand_upper = cand.upper()
                    if cand_upper in tdoc_map and cand_upper != tdoc_id:
                        forward_edges.setdefault(tdoc_id, set()).add(cand_upper)

            # 2. Inspect 'Is revision of'
            rev_of = str(row.get("Is revision of", "")).strip()
            if rev_of and rev_of.lower() != "none":
                for cand in re.findall(r'[A-Za-z0-9]+-\d+(?:r\d+[a-zA-Z]?)?', rev_of):
                    cand_upper = cand.upper()
                    if cand_upper in tdoc_map and cand_upper != tdoc_id:
                        forward_edges.setdefault(cand_upper, set()).add(tdoc_id)

            # 3. Base TDoc regex match (e.g. S2-260100r01 is revision of S2-260100)
            base_match = re.search(r'^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$', tdoc_id, re.IGNORECASE)
            if base_match:
                base_tdoc = base_match.group(1).upper()
                if base_tdoc in tdoc_map and base_tdoc != tdoc_id:
                    forward_edges.setdefault(base_tdoc, set()).add(tdoc_id)

        def _is_active_tdoc(tid: str) -> bool:
            r = tdoc_map.get(tid)
            if not r:
                return False
            st = str(r.get("TDoc Status", "")).strip().lower()
            return "withdrawn" not in st

        # Trace each node to its latest active leaf
        leaf_map: Dict[str, str] = {}
        superseded_set: Set[str] = set()
        predecessors_map: Dict[str, List[str]] = {}

        for tdoc_id in tdoc_map:
            if not _is_active_tdoc(tdoc_id):
                continue

            curr = tdoc_id
            visited = {curr}
            chain = [curr]

            while True:
                children = forward_edges.get(curr, set())
                active_children = [c for c in children if c not in visited and _is_active_tdoc(c)]
                if not active_children:
                    break
                # Sort naturally to pick newest revision (e.g. r02 over r01)
                active_children.sort(key=lambda x: str(x))
                next_node = active_children[-1]
                visited.add(next_node)
                chain.append(next_node)
                curr = next_node

            leaf_tdoc = curr
            leaf_map[tdoc_id] = leaf_tdoc

            if leaf_tdoc != tdoc_id:
                superseded_set.add(tdoc_id)
                preds = predecessors_map.setdefault(leaf_tdoc, [])
                if tdoc_id not in preds:
                    preds.append(tdoc_id)

        # Sort predecessor lists for clean display
        for leaf, preds in predecessors_map.items():
            preds.sort()

        return superseded_set, predecessors_map, leaf_map

    def _build_ls_reply_map(self) -> Tuple[Dict[str, List[str]], Dict[str, List[str]], Dict[str, dict]]:
        """
        Builds multi-candidate reply mappings for Pre-Meeting preparation:
        - in_to_replies: incoming_tdoc_id -> List of candidate draft LS out tdoc_ids (sorted by TDoc ID)
        - out_to_in: outgoing_tdoc_id -> List of incoming LS tdoc_ids it answers
        - tdoc_map: tdoc_id -> row_dict
        Under deduplicate_revisions, resolves candidate reply LS out documents to their latest active revision.
        """
        tdoc_map = {str(r.get("TDoc", "")).strip().upper(): r for r in self.all_rows if r.get("TDoc")}
        superseded_set, predecessors_map, leaf_map = self._build_revision_lineage_map()

        def _resolve_candidate(tid: str) -> str:
            upper = tid.upper()
            if self.deduplicate_revisions and upper in leaf_map:
                return leaf_map[upper]
            return upper

        in_to_replies_set: Dict[str, Set[str]] = {}
        out_to_in_set: Dict[str, Set[str]] = {}

        for r in self.all_rows:
            tdoc = str(r.get("TDoc", "")).strip().upper()
            if not tdoc:
                continue

            # Check if tdoc itself is withdrawn
            status = str(r.get("TDoc Status", "")).strip().lower()
            if "withdrawn" in status:
                continue

            rep_in = str(r.get("Reply in", "")).strip()
            rep_to = str(r.get("Reply to", "")).strip()
            orig_ls = str(r.get("Original LS", "")).strip()

            resolved_out = _resolve_candidate(tdoc)

            # 1. Incoming LS pointing forward to draft replies via 'Reply in'
            if rep_in:
                targets = [_resolve_candidate(x) for x in re.findall(r'[A-Za-z0-9]+-\d+(?:r\d+[a-zA-Z]?)?', rep_in)]
                for tgt in targets:
                    if tgt in tdoc_map:
                        tgt_status = str(tdoc_map[tgt].get("TDoc Status", "")).strip().lower()
                        if "withdrawn" not in tgt_status:
                            in_to_replies_set.setdefault(tdoc, set()).add(tgt)
                            out_to_in_set.setdefault(tgt, set()).add(tdoc)

            # 2. Outgoing proposed LS pointing back to incoming LS via 'Reply to' or 'Original LS'
            back_sources = []
            if rep_to:
                back_sources.extend(re.findall(r'[A-Za-z0-9]+-\d+(?:r\d+[a-zA-Z]?)?', rep_to))
            if orig_ls:
                back_sources.extend(re.findall(r'[A-Za-z0-9]+-\d+(?:r\d+[a-zA-Z]?)?', orig_ls))

            for inc in back_sources:
                inc_upper = inc.upper()
                in_to_replies_set.setdefault(inc_upper, set()).add(resolved_out)
                out_to_in_set.setdefault(resolved_out, set()).add(inc_upper)

        in_to_replies = {k: sorted(list(v)) for k, v in in_to_replies_set.items()}
        out_to_in = {k: sorted(list(v)) for k, v in out_to_in_set.items()}

        return in_to_replies, out_to_in, tdoc_map

    def _filter_and_group_tdocs(self, wi_db_info: dict) -> dict:
        """
        Filters meeting TDocs according to Pre-Meeting rules:
        - Captures active company proposals (Discussion papers, WID/SIDs, proposed CRs, proposed pCRs)
          and incoming liaison statements (LS in).
        - Connects 1-to-many candidate LS reply proposals and allows candidate LS out to inherit parent WI.
        - Deduplicates earlier pre-meeting revisions when deduplicate_revisions=True.
        - Groups by Release -> Work Item -> Category.
        """
        tree: Dict[str, Dict[str, Any]] = {}

        superseded_set, predecessors_map, leaf_map = self._build_revision_lineage_map()
        in_to_replies, out_to_in, tdoc_map = self._build_ls_reply_map()
        visible_tdoc_ids = {str(r.get("TDoc", "")).strip().upper() for r in self.visible_rows if r.get("TDoc")}

        for row in self.all_rows:
            tdoc_id = str(row.get("TDoc", "")).strip().upper()
            if not tdoc_id:
                continue

            status = str(row.get("TDoc Status", "")).strip().lower()

            # Always skip explicitly withdrawn contributions
            if "withdrawn" in status:
                continue

            # Skip superseded revisions if deduplication is enabled
            if self.deduplicate_revisions and tdoc_id in superseded_set:
                continue

            is_reserved = any(res in status for res in ["reserved", "not uploaded", "not yet uploaded"])
            if is_reserved and not self.include_reserved_in_index:
                continue

            # Mid-meeting tolerance check (if user chose to exclude already-treated docs)
            if not self.include_postponed_or_treated:
                if any(tr in status for tr in ["agreed", "approved", "noted", "rejected", "postponed", "not agreed"]):
                    continue

            doc_type = str(row.get("Type", "")).strip().lower()

            category = None
            if "ls in" in doc_type or doc_type == "ls":
                if self.include_ls_in:
                    category = "LS_IN"
            elif "ls out" in doc_type or "ls_out" in doc_type:
                if self.include_ls_out:
                    category = "LS_OUT"
            elif "pcr" in doc_type or "p-cr" in doc_type:
                if self.include_pcr:
                    category = "pCR"
            elif "cr" in doc_type:
                if self.include_cr:
                    category = "CR"
            elif any(d in doc_type for d in ["disc", "discussion", "report", "wid", "sid", "other", "draft", "evaluation"]):
                if self.include_disc:
                    category = "DISC"

            if not category:
                continue

            if self.scope_mode == "visible" and tdoc_id not in visible_tdoc_ids:
                continue

            # Contributor company filter (if active)
            row_companies = set(row.get("_All_Sanitized_Companies", row.get("_Sanitized_Companies", [])))
            row_companies_lower = {str(c).strip().lower() for c in row_companies}

            if self.company_filter:
                company_matched = bool(row_companies_lower & self.company_filter)
                # Keep LS_IN if it has candidate replies from a matching company
                if not company_matched and category == "LS_IN":
                    for cand_rep in in_to_replies.get(tdoc_id, []):
                        rep_row = tdoc_map.get(cand_rep, {})
                        rep_comps = {
                            str(c).strip().lower()
                            for c in rep_row.get("_All_Sanitized_Companies", rep_row.get("_Sanitized_Companies", []))
                        }
                        if rep_comps & self.company_filter:
                            company_matched = True
                            break
                if not company_matched:
                    continue

            # Resolve associated Work Items
            associated_wis = self._resolve_tdoc_wis(row, wi_db_info)

            # Work Item Inheritance for candidate LS OUT:
            # If an LS out has no official WI, inherit the WI from the parent LS in it answers
            if category == "LS_OUT" and (not associated_wis or associated_wis == ["General"]):
                inherited_wis = []
                for inc_id in out_to_in.get(tdoc_id, []):
                    inc_row = tdoc_map.get(inc_id)
                    if inc_row:
                        inc_wis = self._resolve_tdoc_wis(inc_row, wi_db_info)
                        for iw in inc_wis:
                            if iw != "General" and iw not in inherited_wis:
                                inherited_wis.append(iw)
                if inherited_wis:
                    associated_wis = inherited_wis

            raw_wis = TDocsTableModel._split_related_wis(row.get("Related WIs", ""))
            clean_wis = [w for w in raw_wis if w and w.lower() not in {"(no wi)", "none", "nan", "-"}]
            has_official_wi = bool(clean_wis)

            ai_num = str(row.get("Agenda Item", "")).strip() or "Unassigned"
            ai_desc = str(row.get("Agenda item description", "")).strip()

            if not associated_wis:
                if (category in ("LS_IN", "LS_OUT")) and self.include_general_ls:
                    associated_wis = ["General"]
                elif self.scope_mode in ["whole", "custom"]:
                    associated_wis = ["General"]
                else:
                    continue

            # Scope filtering
            matching_wis = []
            for wi in associated_wis:
                if wi == "General":
                    if self.scope_mode in ["whole", "visible"] or ((category in ("LS_IN", "LS_OUT")) and self.include_general_ls):
                        matching_wis.append(wi)
                    elif self.scope_mode == "custom":
                        if "general" in self.custom_wis or (self.custom_ais and ai_num in self.custom_ais):
                            matching_wis.append(wi)
                        elif self.custom_wi_ais and "general" in self.custom_wi_ais and ai_num in self.custom_wi_ais["general"]:
                            matching_wis.append(wi)
                    continue

                wi_tokens = TDocsTableModel._extract_base_tokens(wi)
                if self.scope_mode == "followed":
                    if wi_tokens & wi_db_info["followed_tokens"]:
                        matching_wis.append(wi)
                elif self.scope_mode == "starred":
                    if wi_tokens & wi_db_info["starred_tokens"]:
                        matching_wis.append(wi)
                elif self.scope_mode == "tracked":
                    if wi_tokens & wi_db_info["tracked_tokens"]:
                        matching_wis.append(wi)
                elif self.scope_mode == "custom":
                    wi_match = any(t in self.custom_wis for t in wi_tokens) or (wi.lower() in self.custom_wis)
                    if wi_match:
                        matched_ai_set = None
                        if self.custom_wi_ais:
                            if wi.lower() in self.custom_wi_ais:
                                matched_ai_set = self.custom_wi_ais[wi.lower()]
                            else:
                                for t in wi_tokens:
                                    if t in self.custom_wi_ais:
                                        matched_ai_set = self.custom_wi_ais[t]
                                        break
                        if matched_ai_set is not None:
                            if ai_num in matched_ai_set:
                                matching_wis.append(wi)
                        else:
                            matching_wis.append(wi)
                    elif self.custom_ais and ai_num in self.custom_ais:
                        if wi.lower() == "general" or not self.custom_wis:
                            matching_wis.append(wi)
                else:  # "whole" or "visible"
                    matching_wis.append(wi)

            # Fallback to General for unassigned LS
            if (category in ("LS_IN", "LS_OUT")) and self.include_general_ls and not has_official_wi and not matching_wis:
                matching_wis.append("General")
                if "General" not in associated_wis:
                    associated_wis.append("General")

            if not matching_wis:
                continue

            # Resolve release
            release = self._resolve_tdoc_release(row, matching_wis, wi_db_info)

            # Sibling replies for LS_OUT (competing proposals answering the same incoming LS)
            sibling_replies = set()
            if category == "LS_OUT":
                for parent_in in out_to_in.get(tdoc_id, []):
                    for other_out in in_to_replies.get(parent_in, []):
                        if other_out != tdoc_id:
                            sibling_replies.add(other_out)

            # Insert into tree (handle cross-WI duplication)
            for target_wi in matching_wis:
                other_wis = [w for w in associated_wis if w != target_wi]

                if release not in tree:
                    tree[release] = {}
                if target_wi not in tree[release]:
                    tree[release][target_wi] = {
                        "info": self._get_wi_display_info(target_wi, wi_db_info, release),
                        "ais": set(),
                        "ai_titles": {},
                        "categories": {"LS_IN": [], "LS_OUT": [], "pCR": [], "CR": [], "DISC": []},
                    }

                tree[release][target_wi]["ais"].add(ai_num)
                if ai_num and ai_desc and ai_num not in tree[release][target_wi]["ai_titles"]:
                    tree[release][target_wi]["ai_titles"][ai_num] = ai_desc

                tdoc_url = resolve_tdoc_download_url(
                    tdoc_id, docs_url=self.docs_ftp_url, revisions_url=self.revisions_url, meta=row
                )
                entry = {
                    "tdoc_id": tdoc_id,
                    "title": str(row.get("Title", "")).strip(),
                    "source": str(row.get("Source", "")).strip(),
                    "type": str(row.get("Type", "")).strip(),
                    "status": str(row.get("TDoc Status", "")).strip(),
                    "ai": ai_num,
                    "spec": str(row.get("Spec", "")).strip(),
                    "category": category,
                    "cross_wis": other_wis,
                    "replies": in_to_replies.get(tdoc_id, []),
                    "reply_to_ids": out_to_in.get(tdoc_id, []),
                    "sibling_replies": sorted(list(sibling_replies)),
                    "supersedes": predecessors_map.get(tdoc_id, []),
                    "is_reserved": is_reserved,
                    "my_status": str(row.get("My Status", "⚪ Neutral")).replace("🔄 ", "").strip(),
                    "my_notes": str(row.get("My Notes", "")).replace("🔄 [From Base]: ", "").strip(),
                    "for_action": str(row.get("For", "")).strip(),
                    "sanitized_companies": row.get("_Sanitized_Companies", []),
                    "row_data": row,
                    "tdoc_url": tdoc_url,
                }
                tree[release][target_wi]["categories"][category].append(entry)

        # Post-pass: Guarantee that any candidate LS OUT replying to an included LS IN is present
        all_included_tdocs = {
            doc["tdoc_id"]
            for rel_data in tree.values()
            for wi_data in rel_data.values()
            for doc_list in wi_data["categories"].values()
            for doc in doc_list
        }

        needed_replies = set()
        for rel_data in tree.values():
            for wi_data in rel_data.values():
                for ls_in_doc in wi_data["categories"].get("LS_IN", []):
                    for rep_id in ls_in_doc.get("replies", []):
                        if rep_id not in all_included_tdocs:
                            needed_replies.add(rep_id)

        for rep_id in sorted(needed_replies):
            rep_row = tdoc_map.get(rep_id)
            if not rep_row:
                continue

            rep_status = str(rep_row.get("TDoc Status", "")).strip().lower()
            if "withdrawn" in rep_status:
                continue

            rep_rel = self._resolve_tdoc_release(rep_row, ["General"], wi_db_info)
            target_wi = "General"

            if rep_rel not in tree:
                tree[rep_rel] = {}
            if target_wi not in tree[rep_rel]:
                tree[rep_rel][target_wi] = {
                    "info": self._get_wi_display_info(target_wi, wi_db_info, rep_rel),
                    "ais": set(),
                    "ai_titles": {},
                    "categories": {"LS_IN": [], "LS_OUT": [], "pCR": [], "CR": [], "DISC": []},
                }

            rep_ai = str(rep_row.get("Agenda Item", "")).strip() or "Unassigned"
            rep_ai_desc = str(rep_row.get("Agenda item description", "")).strip()
            tree[rep_rel][target_wi]["ais"].add(rep_ai)
            if rep_ai and rep_ai_desc and rep_ai not in tree[rep_rel][target_wi]["ai_titles"]:
                tree[rep_rel][target_wi]["ai_titles"][rep_ai] = rep_ai_desc

            raw_wis = TDocsTableModel._split_related_wis(rep_row.get("Related WIs", ""))
            clean_wis = [w for w in raw_wis if w and w.lower() not in {"(no wi)", "none", "nan", "-"}]

            rep_url = resolve_tdoc_download_url(
                rep_id, docs_url=self.docs_ftp_url, revisions_url=self.revisions_url, meta=rep_row
            )
            rep_entry = {
                "tdoc_id": rep_id,
                "title": str(rep_row.get("Title", "")).strip(),
                "source": str(rep_row.get("Source", "")).strip(),
                "type": str(rep_row.get("Type", "")).strip(),
                "status": str(rep_row.get("TDoc Status", "")).strip(),
                "ai": rep_ai,
                "spec": str(rep_row.get("Spec", "")).strip(),
                "category": "LS_OUT",
                "cross_wis": clean_wis,
                "replies": in_to_replies.get(rep_id, []),
                "reply_to_ids": out_to_in.get(rep_id, []),
                "sibling_replies": [],
                "supersedes": predecessors_map.get(rep_id, []),
                "is_reserved": any(res in rep_status for res in ["reserved", "not uploaded"]),
                "my_status": str(rep_row.get("My Status", "⚪ Neutral")).replace("🔄 ", "").strip(),
                "my_notes": str(rep_row.get("My Notes", "")).replace("🔄 [From Base]: ", "").strip(),
                "for_action": str(rep_row.get("For", "")).strip(),
                "sanitized_companies": rep_row.get("_Sanitized_Companies", []),
                "row_data": rep_row,
                "tdoc_url": rep_url,
            }
            tree[rep_rel][target_wi]["categories"]["LS_OUT"].append(rep_entry)
            all_included_tdocs.add(rep_id)

        # Sort entries within each category
        for rel in tree.values():
            for wi in rel.values():
                for cat in wi["categories"]:
                    wi["categories"][cat].sort(key=lambda x: str(x.get("tdoc_id", "")))

        return tree

    def _export_all_dossiers(self, filtered_entries: dict):
        """Converts and copies documents into their respective type subfolders with Pre-Meeting callouts."""
        unique_tdocs = {}
        for rel in filtered_entries.values():
            for wi in rel.values():
                for cat_docs in wi["categories"].values():
                    for doc in cat_docs:
                        tid = doc["tdoc_id"]
                        if tid not in unique_tdocs:
                            unique_tdocs[tid] = doc

        total_unique = len(unique_tdocs)
        converted_cache = {}  # tdoc_id -> (local_md_file, local_media_dir)

        # Phase 0: Parallel pre-download of missing non-reserved documents
        if self.download_missing:
            missing_dl_tasks: List[Tuple[str, Path]] = []
            for tdoc_id, doc_entry in unique_tdocs.items():
                # Skip unuploaded reserved documents from network download
                if doc_entry.get("is_reserved"):
                    continue

                # Skip documents preserved from previous export or already existing on disk
                if doc_entry.get("is_preserved_from_previous"):
                    continue
                if hasattr(self, "session_plan") and tdoc_id in self.session_plan.tdocs_to_skip:
                    continue

                row_data = doc_entry["row_data"]
                base_match = re.search(r'^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$', tdoc_id, re.IGNORECASE)
                base_tdoc = base_match.group(1).upper() if base_match else tdoc_id.upper()
                tdoc_folder = self.meeting_dir / base_tdoc
                local_md_file = tdoc_folder / f"{sanitize_filename(tdoc_id)}.md"
                legacy_cache_file = tdoc_folder / f"{tdoc_id}_LLM_v{CONVERTER_VERSION}.md"
                doc_path = self._find_word_doc(tdoc_folder, tdoc_id)
                doc_type = row_data.get("Type", "") if row_data else ""
                expected_mode = resolve_revision_mode_for_type(doc_type, default_mode="markup")

                needs_conversion = self.force_reconvert or is_conversion_stale(
                    doc_path, local_md_file, expected_version=CONVERTER_VERSION, expected_mode=expected_mode
                )
                if needs_conversion and not self.force_reconvert and os.path.exists(to_extended_path(legacy_cache_file)):
                    if not is_conversion_stale(doc_path, legacy_cache_file, expected_version=CONVERTER_VERSION, expected_mode=expected_mode):
                        needs_conversion = False

                if needs_conversion and not doc_path:
                    missing_dl_tasks.append((tdoc_id, tdoc_folder))

            if missing_dl_tasks:
                total_dl = len(missing_dl_tasks)
                dl_workers = min(6, total_dl)
                self.progress.emit(f"Downloading {total_dl} missing pre-meeting TDoc(s) in parallel ({dl_workers} workers)...")
                logger.info(f"[PreMeetingExporter] Concurrently downloading {total_dl} missing TDocs using {dl_workers} workers...")
                dl_done = 0
                with ThreadPoolExecutor(max_workers=dl_workers) as executor:
                    future_to_tid = {
                        executor.submit(self._download_and_extract_tdoc, tid, folder): tid
                        for tid, folder in missing_dl_tasks
                    }
                    for future in as_completed(future_to_tid):
                        if self.isInterruptionRequested():
                            try:
                                executor.shutdown(wait=False, cancel_futures=True)
                            except TypeError:
                                executor.shutdown(wait=False)
                            return
                        tid = future_to_tid[future]
                        dl_done += 1
                        try:
                            res_path = future.result()
                            if res_path:
                                self.progress.emit(f"Downloaded TDoc {tid} ({dl_done}/{total_dl})")
                            else:
                                self.progress.emit(f"Failed to download {tid} ({dl_done}/{total_dl})")
                        except Exception as dl_err:
                            logger.error(f"[PreMeetingExporter] Error downloading {tid}: {dl_err}")

        # Phase 1: Ensure all canonical Markdown files exist in meeting cache
        for idx, (tdoc_id, doc_entry) in enumerate(unique_tdocs.items(), 1):
            if self.isInterruptionRequested():
                return

            if doc_entry.get("is_preserved_from_previous"):
                continue
            if hasattr(self, "session_plan") and tdoc_id in self.session_plan.tdocs_to_skip:
                continue

            row_data = doc_entry["row_data"]
            # If reserved and not present locally, we skip conversion and will write a clean reserved stub
            if doc_entry.get("is_reserved") and not self._find_word_doc(self.meeting_dir / tdoc_id, tdoc_id):
                converted_cache[tdoc_id] = (None, None)
                continue

            self.progress.emit(f"Converting {tdoc_id} to Markdown ({idx}/{total_unique})...")
            local_md_file, local_media_dir = self._ensure_tdoc_markdown(tdoc_id, row_data)
            converted_cache[tdoc_id] = (local_md_file, local_media_dir)

        # Pre-register target paths for all documents across the dossier structure
        self.exported_tdoc_paths: Dict[str, Path] = {}
        self.tdoc_titles: Dict[str, str] = {}
        self.tdoc_sources: Dict[str, str] = {}

        for rel_name, wis in filtered_entries.items():
            rel_slug = sanitize_folder_name(rel_name)
            for wi_name, wi_data in wis.items():
                wi_slug = sanitize_folder_name(wi_name)
                for cat_name, cat_docs in wi_data["categories"].items():
                    cat_folder = "Discussion" if cat_name == "DISC" else cat_name
                    type_dir = self.export_dir / rel_slug / wi_slug / cat_folder
                    for doc in cat_docs:
                        tdoc_id = doc["tdoc_id"]
                        target_md_file = type_dir / f"{sanitize_filename(tdoc_id)}.md"
                        doc["target_md_abs_path"] = target_md_file
                        doc["target_md_rel_path"] = f"{rel_slug}/{wi_slug}/{cat_folder}/{sanitize_filename(tdoc_id)}.md"
                        doc["wi_local_rel_path"] = f"{cat_folder}/{sanitize_filename(tdoc_id)}.md"
                        if tdoc_id not in self.exported_tdoc_paths:
                            self.exported_tdoc_paths[tdoc_id] = target_md_file
                            self.tdoc_titles[tdoc_id] = doc.get("title", "")
                            self.tdoc_sources[tdoc_id] = doc.get("source", "")

        # Phase 2: Copy dossiers into target directory structure with rich pre-meeting callouts
        for rel_name, wis in filtered_entries.items():
            rel_slug = sanitize_folder_name(rel_name)
            for wi_name, wi_data in wis.items():
                wi_slug = sanitize_folder_name(wi_name)

                for cat_name, cat_docs in wi_data["categories"].items():
                    if not cat_docs:
                        continue

                    cat_folder = "Discussion" if cat_name == "DISC" else cat_name
                    type_dir = self.export_dir / rel_slug / wi_slug / cat_folder
                    type_media_dir = type_dir / "media"
                    os.makedirs(to_extended_path(type_dir), exist_ok=True)

                    for doc in cat_docs:
                        tdoc_id = doc["tdoc_id"]
                        if doc.get("is_preserved_from_previous"):
                            continue
                        if hasattr(self, "session_plan") and tdoc_id in self.session_plan.tdocs_to_skip:
                            continue

                        local_md_file, local_media_dir = converted_cache.get(tdoc_id, (None, None))
                        target_md_file = doc.get("target_md_abs_path") or (type_dir / f"{sanitize_filename(tdoc_id)}.md")

                        header_notices = []

                        # Notice 0: Incremental Addition Header
                        if hasattr(self, "session_plan") and self.session_plan.is_incremental and tdoc_id in self.session_plan.new_tdocs:
                            update_lbl = f"Incremental Update #{self.session_plan.session_id - 1}"
                            header_notices.append(
                                f"> [!NOTE]\n"
                                f"> **Incremental Addition:** Added to this briefing on {self.session_plan.session_timestamp} ({update_lbl})."
                            )

                        # Notice 1: Revision Lineage Notice (Supersedes older revisions)
                        if doc.get("supersedes"):
                            preds_str = ", ".join(f"`{p}`" for p in doc["supersedes"])
                            header_notices.append(
                                f"> [!NOTE]\n"
                                f"> **Revision History:** This active pre-meeting contribution supersedes earlier draft(s): {preds_str}."
                            )

                        # Notice 2: Cross-WI Contribution
                        if doc.get("cross_wis"):
                            other_wis_str = ", ".join(f"`{w}`" for w in doc["cross_wis"])
                            header_notices.append(
                                f"> [!NOTE]\n"
                                f"> **Cross-WI Contribution:** This TDoc is also associated with {other_wis_str}."
                            )

                        # Notice 3: Personal Pre-Meeting Stance & Notes
                        if self.include_user_notes:
                            my_st = doc.get("my_status", "")
                            my_nt = doc.get("my_notes", "")
                            if (my_st and my_st != "⚪ Neutral") or my_nt:
                                stance_header = f"Delegate Pre-Meeting Stance ({my_st}):" if my_st else "Delegate Pre-Meeting Notes:"
                                notes_body = f"\n> {my_nt}" if my_nt else ""
                                header_notices.append(
                                    f"> [!IMPORTANT]\n"
                                    f"> **{stance_header}**{notes_body}"
                                )

                        # Notice 4: Candidate LS Replies (For LS_IN)
                        if doc.get("category") == "LS_IN" and doc.get("replies"):
                            rep_bullets = []
                            for rep_id in doc["replies"]:
                                rep_abs = self.exported_tdoc_paths.get(rep_id)
                                rep_title = self.tdoc_titles.get(rep_id, "")
                                rep_src = self.tdoc_sources.get(rep_id, "")
                                src_str = f" ({rep_src})" if rep_src else ""
                                title_str = f" — *{rep_title}*" if rep_title else ""
                                if rep_abs:
                                    rel = safe_markdown_rel_link(rep_abs, target_md_file.parent)
                                    rep_bullets.append(f"> - **[{rep_id}]({rel})**{title_str}{src_str}")
                                else:
                                    rep_bullets.append(f"> - **{rep_id}**{title_str}{src_str}")
                            n_cand = len(rep_bullets)
                            cand_lbl = f"{n_cand} candidate proposal(s)" if n_cand > 1 else "1 candidate proposal"
                            header_notices.append(
                                f"> [!TIP]\n"
                                f"> **Proposed Reply Liaison Statements ({cand_lbl}):**\n" + "\n".join(rep_bullets)
                            )

                        # Notice 5: Draft LS Reply provenance & Competing Sibling Replies (For LS_OUT)
                        if doc.get("category") == "LS_OUT" and doc.get("reply_to_ids"):
                            in_links = []
                            for in_id in doc["reply_to_ids"]:
                                in_abs = self.exported_tdoc_paths.get(in_id)
                                in_title = self.tdoc_titles.get(in_id, "")
                                title_str = f" (*{in_title}*)" if in_title else ""
                                if in_abs:
                                    rel = safe_markdown_rel_link(in_abs, target_md_file.parent)
                                    in_links.append(f"**[{in_id}]({rel})**{title_str}")
                                else:
                                    in_links.append(f"**{in_id}**{title_str}")

                            reply_notice = f"> [!NOTE]\n> **Draft Reply Liaison Statement:**\n> - **Replies to Incoming LS:** {', '.join(in_links)}"

                            if doc.get("sibling_replies"):
                                sib_links = []
                                for sib_id in doc["sibling_replies"]:
                                    sib_abs = self.exported_tdoc_paths.get(sib_id)
                                    sib_src = self.tdoc_sources.get(sib_id, "")
                                    src_str = f" ({sib_src})" if sib_src else ""
                                    if sib_abs:
                                        rel = safe_markdown_rel_link(sib_abs, target_md_file.parent)
                                        sib_links.append(f"**[{sib_id}]({rel})**{src_str}")
                                    else:
                                        sib_links.append(f"**{sib_id}**{src_str}")
                                reply_notice += f"\n> - **Competing Reply Proposals ({len(sib_links)} other(s)):** {', '.join(sib_links)}"

                            header_notices.append(reply_notice)

                        notice_block = "\n\n".join(header_notices) + ("\n\n" if header_notices else "")

                        # Navigation breadcrumbs
                        master_index_file = self.export_dir / "index.md"
                        master_rel = safe_markdown_rel_link(master_index_file, target_md_file.parent)
                        wi_index_file = self.export_dir / rel_slug / wi_slug / "index.md"
                        if self.generate_local_indices:
                            wi_rel = safe_markdown_rel_link(wi_index_file, target_md_file.parent)
                            nav_bar = f"[📁 Master Briefing]({master_rel}) · [🎯 {wi_name} Briefing]({wi_rel})\n\n---\n\n"
                        else:
                            nav_bar = f"[📁 Master Briefing]({master_rel})\n\n---\n\n"

                        full_header_block = nav_bar + notice_block

                        # If reserved and uncached, write clean informative reserved stub
                        if doc.get("is_reserved") and not local_md_file:
                            try:
                                source_lines = "\n".join(CompanySanitizer.format_source_markdown_header_lines(doc.get('source', '')))
                                with open(to_extended_path(target_md_file), "w", encoding="utf-8") as f:
                                    f.write(
                                        f"{full_header_block}"
                                        f"# {tdoc_id}: {doc['title']}\n\n"
                                        f"> ⏳ **Status: Reserved / Not Uploaded**\n"
                                        f"> *This document number has been allocated or reserved on the meeting agenda but has not yet been uploaded to 3GPP servers.*\n\n"
                                        f"{source_lines}\n"
                                        f"- **Type:** {doc['type']}\n"
                                        f"- **Agenda Item:** {doc['ai']}\n"
                                        f"- **Status:** {doc['status'] or 'Reserved'}\n"
                                    )
                            except Exception as err:
                                logger.error(f"[PreMeetingExporter] Error writing reserved stub for {tdoc_id}: {err}")
                            continue

                        if local_md_file and os.path.exists(to_extended_path(local_md_file)):
                            try:
                                with open(to_extended_path(local_md_file), "r", encoding="utf-8") as f:
                                    md_content = f.read()

                                tdoc_url = doc.get("tdoc_url") or resolve_tdoc_download_url(
                                    tdoc_id, docs_url=self.docs_ftp_url, revisions_url=self.revisions_url, meta=doc.get("row_data")
                                )
                                if tdoc_url and "- **3GPP Server URL:**" not in md_content:
                                    url_bullet = f"- **3GPP Server URL:** [{tdoc_url}]({tdoc_url})\n"
                                    if "- **Original file:**" in md_content:
                                        md_content = md_content.replace("- **Original file:**", f"{url_bullet}- **Original file:**", 1)
                                    elif "- **Exporter:**" in md_content:
                                        md_content = md_content.replace("- **Exporter:**", f"{url_bullet}- **Exporter:**", 1)
                                    elif "\n---\n" in md_content:
                                        md_content = md_content.replace("\n---\n", f"\n{url_bullet}\n---\n", 1)

                                if full_header_block:
                                    if md_content.startswith("<!---"):
                                        first_break = md_content.find("\n\n")
                                        if first_break != -1:
                                            md_content = md_content[:first_break + 2] + full_header_block + md_content[first_break + 2:]
                                        else:
                                            md_content = full_header_block + md_content
                                    else:
                                        md_content = full_header_block + md_content

                                with open(to_extended_path(target_md_file), "w", encoding="utf-8") as f:
                                    f.write(md_content)

                                # Copy media files
                                if local_media_dir and local_media_dir.is_dir():
                                    media_files = list(local_media_dir.glob(f"{tdoc_id}_*"))
                                    if not media_files:
                                        media_files = list(local_media_dir.glob(f"{sanitize_filename(tdoc_id)}_*"))
                                    if media_files:
                                        os.makedirs(to_extended_path(type_media_dir), exist_ok=True)
                                        for mf in media_files:
                                            dst_img = type_media_dir / mf.name
                                            if not dst_img.exists():
                                                shutil.copy2(to_extended_path(mf), to_extended_path(dst_img))
                            except Exception as err:
                                logger.error(f"[PreMeetingExporter] Error writing dossier for {tdoc_id}: {err}")
                        else:
                            # Fallback stub
                            try:
                                source_lines = "\n".join(CompanySanitizer.format_source_markdown_header_lines(doc.get('source', '')))
                                with open(to_extended_path(target_md_file), "w", encoding="utf-8") as f:
                                    f.write(
                                        f"{full_header_block}"
                                        f"# {tdoc_id}: {doc['title']}\n\n"
                                        f"{source_lines}\n"
                                        f"- **Type:** {doc['type']}\n"
                                        f"- **Agenda Item:** {doc['ai']}\n"
                                        f"- **Status:** {doc['status']}\n\n"
                                        f"> ⚠️ *Note: Document conversion was not available in local cache.*\n"
                                    )
                            except Exception:
                                pass

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
                            if tid in self.session_plan.new_tdocs:
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
