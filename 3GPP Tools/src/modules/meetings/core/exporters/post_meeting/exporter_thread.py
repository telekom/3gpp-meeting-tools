# --- File: src/modules/meetings/core/exporters/post_meeting/exporter_thread.py ---
"""
PostMeetingExporterThread worker shell: QThread signals, __init__, run() pipeline, and dossier compilation.
Feature logic lives in the mixins/ package (see README.md for a 'where to edit what' map).
"""
import datetime
import os
from pathlib import Path
from typing import (
    Dict,
    Optional,
    Set,
)
from PyQt5.QtCore import QThread, pyqtSignal
from core.utils.paths import to_extended_path
from modules.meetings.core.exporters.post_meeting.path_utils import logger
from modules.meetings.core.exporters.post_meeting.mixins import (
    GroupingMixin,
    ConversionMixin,
    IndexingMixin,
    AgreementsMixin,
)


class PostMeetingExporterThread(
    GroupingMixin,
    ConversionMixin,
    IndexingMixin,
    AgreementsMixin,
    QThread,
):
    """
    Asynchronous worker to generate an organized, executive-grade Post-Meeting Markdown Dossier.
    Structure:
      Export/PostMeeting_<WG>_<Meeting>_<Timestamp>/
        ├── index.md (Master Cover & Summary Tables)
        ├── <Release>/
        │   └── <WorkItem>/
        │       ├── index.md (WI Specific Dossier Cover)
        │       ├── LS_IN/ (S2-xxxxxx.md + media/)
        │       ├── pCR/   (S2-xxxxxx.md + media/)
        │       ├── CR/    (S2-xxxxxx.md + media/)
        │       └── LS_OUT/(S2-xxxxxx.md + media/)
        └── General/
            └── LS_IN/ (Overarching incoming LS without a specific WI)
    """
    progress = pyqtSignal(str)
    finished = pyqtSignal(bool, str)

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
        include_disc: bool = False,
        include_chair_agreements: bool = True,
        include_general_ls: bool = True,
        generate_local_indices: bool = True,
        download_missing: bool = True,
        force_reconvert: bool = False,
        compile_dossier: bool = False,
        output_dir: Optional[Path] = None,
        parent=None,
    ):
        super().__init__(parent)
        self.meeting_dir = Path(meeting_dir)
        self.all_rows = all_rows or []
        self.visible_rows = visible_rows or []
        self.agenda_map = agenda_map or {}
        self.mtg_info = mtg_info or {}
        self.docs_ftp_url = docs_ftp_url or ""
        self.revisions_url = revisions_url or ""
        self.scope_mode = scope_mode
        self.custom_wis = {str(w).strip().lower() for w in (custom_wis or set()) if str(w).strip()}
        self.custom_ais = {str(a).strip() for a in (custom_ais or set()) if str(a).strip()}
        self.custom_wi_ais: Dict[str, Set[str]] = {}
        if custom_wi_ais:
            for k, v in custom_wi_ais.items():
                self.custom_wi_ais[str(k).strip().lower()] = {str(a).strip() for a in (v or set()) if str(a).strip()}
        self.followed_tokens = {str(t).strip().lower() for t in (followed_tokens or set()) if str(t).strip()}
        self.starred_tokens = {str(t).strip().lower() for t in (starred_tokens or set()) if str(t).strip()}
        self.include_ls_in = include_ls_in
        self.include_pcr = include_pcr
        self.include_cr = include_cr
        self.include_ls_out = include_ls_out
        self.include_disc = include_disc
        self.include_chair_agreements = include_chair_agreements
        self.include_general_ls = include_general_ls
        self.generate_local_indices = generate_local_indices
        self.download_missing = download_missing
        self.force_reconvert = force_reconvert
        self.compile_dossier = compile_dossier

        # Resolve output directory
        if output_dir:
            self.export_dir = Path(output_dir)
        else:
            timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S")
            wg = str(self.mtg_info.get("wg_name", "3GPP")).strip().replace(" ", "_")
            mtg = str(self.mtg_info.get("meeting_number", "")).strip()
            folder_name = f"PostMeeting_{wg}_{mtg}_{timestamp}" if (wg or mtg) else f"PostMeeting_{timestamp}"
            self.export_dir = self.meeting_dir / "Export" / folder_name

    def run(self):
        try:
            self.progress.emit("Initializing Work Items knowledge base...")
            wi_db_info = self._load_work_items_info()

            self.progress.emit("Categorizing and filtering post-meeting TDocs...")
            filtered_entries = self._filter_and_group_tdocs(wi_db_info)

            if not filtered_entries:
                self.finished.emit(False, "No TDocs match the selected export criteria.")
                return

            total_entries = sum(len(docs) for rel in filtered_entries.values() for wi in rel.values() for docs in wi["categories"].values())
            self.progress.emit(f"Found {total_entries} matching contribution instance(s). Exporting dossiers...")

            os.makedirs(to_extended_path(self.export_dir), exist_ok=True)

            # Incremental Manifest & Session Planning
            from modules.meetings.core.exporters.export_manifest import (
                load_or_reconstruct_manifest,
                plan_incremental_session,
                merge_manifest_into_tree,
            )

            self.manifest = load_or_reconstruct_manifest(
                self.export_dir, self.mtg_info, self.all_rows, export_type="post_meeting"
            )
            self.session_plan = plan_incremental_session(
                self.manifest, filtered_entries, self.export_dir, force_reconvert=self.force_reconvert
            )

            if self.session_plan.is_incremental:
                logger.info(
                    f"[PostMeetingExporter] Incremental export session #{self.session_plan.session_id}: "
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
            self.progress.emit("Generating Markdown index reports...")
            if self.generate_local_indices:
                self._generate_local_wi_indices(filtered_entries)

            self._generate_master_index(filtered_entries)

            # Compile consolidated Navigable Dossier (HTML & PDF)
            if self.compile_dossier and not self.isInterruptionRequested():
                try:
                    self.progress.emit("Compiling consolidated Navigable Dossier (HTML & PDF)...")
                    self._compile_consolidated_dossier(filtered_entries)
                except Exception as doss_err:
                    logger.warning(f"[PostMeetingExporter] Consolidated dossier generation warning: {doss_err}")

            if hasattr(self, "session_plan") and self.session_plan.is_incremental:
                preserved_count = len(self.session_plan.tdocs_to_skip) + len(self.session_plan.preserved_tdocs)
                msg = (
                    f"Successfully updated Post-Meeting Dossier (Update #{self.session_plan.session_id - 1}):\n"
                    f"• {len(self.session_plan.new_tdocs)} TDoc(s) added\n"
                    f"• {preserved_count} TDoc(s) preserved\n\n"
                    f"Location:\n{self.export_dir}"
                )
            else:
                msg = f"Successfully generated Post-Meeting Dossier in:\n{self.export_dir}"
            self.finished.emit(True, msg)

        except Exception as e:
            logger.error(f"[PostMeetingExporter] Critical failure during export: {e}", exc_info=True)
            self.finished.emit(False, str(e))

    def _compile_consolidated_dossier(self, filtered_entries: dict):
        """Compiles all exported deliverables into a master navigable HTML and PDF dossier."""
        unique_docs = {}
        for rel in filtered_entries.values():
            for wi in rel.values():
                for cat_docs in wi["categories"].values():
                    for doc in cat_docs:
                        tid = doc["tdoc_id"]
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
