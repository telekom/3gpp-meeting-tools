# --- File: src/modules/meetings/core/exporters/pre_meeting/exporter_thread.py ---
"""
PreMeetingExporterThread worker shell: QThread configuration, run() pipeline, and consolidated dossier.
Feature logic lives in mixins/ (see pre_meeting/README.md for a 'where to edit what' map).
"""

import datetime
import logging
import os
from pathlib import Path
from typing import Dict, Optional, Set

from core.utils.paths import to_extended_path
from modules.meetings.core.exporters.post_meeting_exporter import PostMeetingExporterThread
from modules.meetings.core.exporters.pre_meeting.mixins import (
    PreMeetingExportMixin,
    PreMeetingGroupingMixin,
    PreMeetingIndexingMixin,
)

logger = logging.getLogger("modules.meetings.core.exporters.pre_meeting_exporter")


class PreMeetingExporterThread(
    PreMeetingGroupingMixin,
    PreMeetingExportMixin,
    PreMeetingIndexingMixin,
    PostMeetingExporterThread,
):
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

