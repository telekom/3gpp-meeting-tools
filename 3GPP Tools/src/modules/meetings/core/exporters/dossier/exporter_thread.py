"""Asynchronous QThread worker orchestrating unified 3GPP Dossier compilation."""

from __future__ import annotations

import datetime
import logging
import os
import re
import shutil
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from PyQt5.QtCore import QThread, pyqtSignal

from core.utils.paths import to_extended_path
from modules.meetings.core.tdocs.tdoc_file_handler import TDocFileHandler
from modules.meetings.core.tdocs.tdoc_markdown_converter import (
    CONVERTER_VERSION,
    is_conversion_stale,
    resolve_revision_mode_for_type,
    sanitize_filename,
)
from modules.meetings.core.tdocs.tdoc_url_resolver import resolve_tdoc_download_url

from .html_builder import build_html_dossier, convert_markdown_to_html
from .pdf_renderer import render_bookmarked_pdf
from .prompt_builder import generate_ai_prompt

logger = logging.getLogger(__name__)


class DossierExporterThread(QThread):
    """
    Asynchronous worker to compile selected or post-meeting 3GPP Technical Documents
    into a unified Navigable HTML Dossier, Bookmarked Vector PDF, and AI Citation Prompt.
    """

    progress = pyqtSignal(str)
    finished = pyqtSignal(bool, str, dict)  # success, message, file_paths

    def __init__(
        self,
        meeting_dir: Path,
        tdocs_list: List[dict],
        docs_ftp_url: str,
        revisions_url: str,
        mtg_info: dict,
        export_format: str = "both",  # "both", "html", "pdf"
        include_prompt: bool = True,
        embed_base64_images: bool = True,
        download_missing: bool = True,
        force_reconvert: bool = False,
        output_dir: Optional[Path] = None,
        parent=None,
    ):
        super().__init__(parent)
        self.meeting_dir = Path(meeting_dir)
        self.tdocs_list = tdocs_list or []
        self.docs_ftp_url = docs_ftp_url or ""
        self.revisions_url = revisions_url or ""
        self.mtg_info = mtg_info or {}
        self.export_format = export_format.lower()
        self.include_prompt = include_prompt
        self.embed_base64_images = embed_base64_images
        self.download_missing = download_missing
        self.force_reconvert = force_reconvert

        # Resolve output directory
        if output_dir:
            self.export_dir = Path(output_dir)
        else:
            timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S")
            wg = str(self.mtg_info.get("wg_name", "3GPP")).strip().replace(" ", "_")
            mtg = str(self.mtg_info.get("meeting_number", "")).strip()
            folder_name = f"Dossier_{wg}_{mtg}_{timestamp}" if (wg or mtg) else f"Dossier_{timestamp}"
            self.export_dir = self.meeting_dir / "Export" / folder_name

    def run(self):
        try:
            if not self.tdocs_list:
                self.finished.emit(False, "No TDocs selected for dossier generation.", {})
                return

            os.makedirs(to_extended_path(self.export_dir), exist_ok=True)
            media_export_dir = self.export_dir / "media"
            os.makedirs(to_extended_path(media_export_dir), exist_ok=True)

            total_docs = len(self.tdocs_list)
            self.progress.emit(f"Preparing {total_docs} document(s) for dossier...")

            # --- Phase 1: Parallel Pre-download of Missing Documents ---
            if self.download_missing:
                missing_download_tasks: List[Tuple[str, Path]] = []
                seen_tdocs = set()

                for tdoc_data in self.tdocs_list:
                    tdoc_id = str(tdoc_data.get("TDoc", "")).strip()
                    if not tdoc_id or tdoc_id in seen_tdocs:
                        continue
                    seen_tdocs.add(tdoc_id)

                    base_match = re.search(
                        r"^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$", tdoc_id, re.IGNORECASE
                    )
                    base_tdoc = base_match.group(1).upper() if base_match else tdoc_id.upper()
                    doc_type = str(tdoc_data.get("Type", "Other"))
                    expected_mode = resolve_revision_mode_for_type(doc_type, default_mode="markup")

                    tdoc_folder = self.meeting_dir / base_tdoc
                    local_md_file = tdoc_folder / f"{sanitize_filename(tdoc_id)}.md"
                    doc_path = self._find_word_doc(tdoc_folder, tdoc_id)

                    needs_conversion = self.force_reconvert or is_conversion_stale(
                        doc_path,
                        local_md_file,
                        expected_version=CONVERTER_VERSION,
                        expected_mode=expected_mode,
                    )
                    if needs_conversion and not doc_path:
                        missing_download_tasks.append((tdoc_id, tdoc_folder))

                if missing_download_tasks:
                    total_missing = len(missing_download_tasks)
                    dl_workers = min(6, total_missing)
                    self.progress.emit(f"Downloading {total_missing} missing TDoc(s) in parallel...")
                    logger.info(
                        f"[DossierExporter] Pre-downloading {total_missing} TDocs using {dl_workers} workers..."
                    )
                    dl_done = 0
                    with ThreadPoolExecutor(max_workers=dl_workers) as executor:
                        future_to_tid = {
                            executor.submit(self._download_and_extract_tdoc, tid, folder): tid
                            for (tid, folder) in missing_download_tasks
                        }
                        for future in as_completed(future_to_tid):
                            if self.isInterruptionRequested():
                                try:
                                    executor.shutdown(wait=False, cancel_futures=True)
                                except TypeError:
                                    executor.shutdown(wait=False)
                                self.finished.emit(False, "Export cancelled by user.", {})
                                return
                            tid = future_to_tid[future]
                            dl_done += 1
                            try:
                                future.result()
                                self.progress.emit(f"Downloaded TDoc {tid} ({dl_done}/{total_missing})")
                            except Exception as dl_err:
                                logger.error(f"[DossierExporter] Error downloading {tid}: {dl_err}")

            if self.isInterruptionRequested():
                self.finished.emit(False, "Export cancelled by user.", {})
                return

            # --- Phase 2: Convert / Load Markdown for Each Document ---
            converted_docs: List[Dict[str, Any]] = []
            for idx, tdoc_data in enumerate(self.tdocs_list, 1):
                if self.isInterruptionRequested():
                    self.finished.emit(False, "Export cancelled by user.", {})
                    return

                tdoc_id = str(tdoc_data.get("TDoc", "")).strip()
                if not tdoc_id:
                    continue

                base_match = re.search(
                    r"^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$", tdoc_id, re.IGNORECASE
                )
                base_tdoc = base_match.group(1).upper() if base_match else tdoc_id.upper()
                doc_type = str(tdoc_data.get("Type", "Other"))
                raw_ai = str(tdoc_data.get("Agenda Item", "Unassigned")).strip()
                title = str(tdoc_data.get("Title", "")).strip()
                source = str(tdoc_data.get("Source", "")).strip()
                status = str(tdoc_data.get("TDoc Status", "")).strip()
                expected_mode = resolve_revision_mode_for_type(doc_type, default_mode="markup")

                self.progress.emit(f"Processing Markdown for {tdoc_id} ({idx}/{total_docs})...")

                tdoc_folder = self.meeting_dir / base_tdoc
                local_md_file = tdoc_folder / f"{sanitize_filename(tdoc_id)}.md"
                local_media_dir = tdoc_folder / "media"
                doc_path = self._find_word_doc(tdoc_folder, tdoc_id)

                needs_conversion = self.force_reconvert or is_conversion_stale(
                    doc_path,
                    local_md_file,
                    expected_version=CONVERTER_VERSION,
                    expected_mode=expected_mode,
                )

                if needs_conversion and doc_path:
                    try:
                        os.makedirs(to_extended_path(tdoc_folder), exist_ok=True)
                        os.makedirs(to_extended_path(local_media_dir), exist_ok=True)
                        from modules.meetings.core.tdocs.tdoc_markdown_converter import (
                            convert_document_to_markdown,
                        )

                        convert_document_to_markdown(
                            doc_path,
                            out_dir=tdoc_folder,
                            tdoc_id=tdoc_id,
                            meta=tdoc_data,
                            revisions=expected_mode,
                            media_dir=local_media_dir,
                            media_rel="media",
                            cancellation_check=self.isInterruptionRequested,
                        )
                    except Exception as conv_err:
                        logger.warning(
                            f"[DossierExporter] Direct conversion failed for {tdoc_id}: {conv_err}"
                        )

                md_text = ""
                if os.path.exists(to_extended_path(local_md_file)):
                    try:
                        with open(
                            to_extended_path(local_md_file),
                            "r",
                            encoding="utf-8",
                            errors="ignore",
                        ) as f:
                            md_text = f.read()
                    except Exception as read_err:
                        logger.error(
                            f"[DossierExporter] Failed reading markdown for {tdoc_id}: {read_err}"
                        )

                # Copy media files to export/media/
                if local_media_dir.is_dir():
                    for img in local_media_dir.iterdir():
                        if img.is_file() and (
                            tdoc_id.upper() in img.name.upper() or not img.name.startswith("S")
                        ):
                            dst = media_export_dir / img.name
                            try:
                                if not dst.exists():
                                    shutil.copy2(to_extended_path(img), to_extended_path(dst))
                            except Exception:
                                pass

                tdoc_download_url = resolve_tdoc_download_url(
                    tdoc_id,
                    docs_url=self.docs_ftp_url,
                    revisions_url=self.revisions_url,
                    meta=tdoc_data,
                )

                converted_docs.append({
                    "tdoc_id": tdoc_id,
                    "title": title,
                    "source": source,
                    "type": doc_type,
                    "status": status,
                    "ai": raw_ai,
                    "row_data": tdoc_data,
                    "md_text": md_text,
                    "media_dir": local_media_dir,
                    "tdoc_url": tdoc_download_url,
                })

            if not converted_docs:
                self.finished.emit(False, "Failed to compile any documents into the dossier.", {})
                return

            # --- Phase 3: Synthesize Interactive HTML Dossier ---
            self.progress.emit("Synthesizing interactive HTML dossier...")
            html_content, _tdoc_html_map = self._build_html_dossier(converted_docs, media_export_dir)

            html_path = self.export_dir / "dossier.html"
            with open(to_extended_path(html_path), "w", encoding="utf-8") as f:
                f.write(html_content)
            logger.info(f"[DossierExporter] Saved HTML dossier: {html_path}")

            result_files = {"html": str(html_path)}

            # --- Phase 4: Generate Navigable PDF with Bookmarks ---
            page_spans: Dict[str, Tuple[int, int]] = {}
            pdf_path = self.export_dir / "dossier.pdf"

            if self.export_format in ("both", "pdf"):
                self.progress.emit("Rendering vector PDF with hierarchical bookmarks...")
                page_spans = self._render_bookmarked_pdf(html_content, pdf_path, converted_docs)
                result_files["pdf"] = str(pdf_path)

            # --- Phase 5: Generate Companion AI Citation Prompt ---
            if self.include_prompt:
                self.progress.emit("Synthesizing companion AI citation map...")
                prompt_path = self.export_dir / "dossier_prompt.txt"
                self._generate_ai_prompt(converted_docs, page_spans, prompt_path)
                result_files["prompt"] = str(prompt_path)

            msg = f"Successfully generated Navigable Dossier in:\n{self.export_dir}"
            self.finished.emit(True, msg, result_files)

        except Exception as e:
            logger.error(f"[DossierExporter] Critical error: {e}", exc_info=True)
            self.finished.emit(False, str(e), {})

    # -------------------------------------------------------------------------
    # Delegated Pipeline Methods (Preserved Instance Hooks)
    # -------------------------------------------------------------------------

    def _build_html_dossier(
        self, docs: List[Dict[str, Any]], media_export_dir: Path
    ) -> Tuple[str, Dict[str, str]]:
        return build_html_dossier(
            docs,
            media_export_dir,
            self.mtg_info,
            meeting_dir=self.meeting_dir,
            embed_base64_images=self.embed_base64_images,
            convert_md_fn=self._convert_markdown_to_html,
        )

    def _convert_markdown_to_html(
        self, md_text: str, local_media_dir: Path, media_export_dir: Path, tdoc_id: str
    ) -> str:
        return convert_markdown_to_html(
            md_text,
            local_media_dir,
            media_export_dir,
            tdoc_id,
            meeting_dir=self.meeting_dir,
            embed_base64_images=self.embed_base64_images,
        )

    def _render_bookmarked_pdf(
        self,
        arg1: str,
        arg2: Any,
        arg3: Any = None,
        arg4: Optional[List[Dict[str, Any]]] = None,
    ) -> Dict[str, Tuple[int, int]]:
        return render_bookmarked_pdf(
            arg1,
            arg2,
            arg3,
            arg4,
            is_cancelled=self.isInterruptionRequested,
            progress_cb=self.progress.emit,
        )

    def _generate_ai_prompt(
        self,
        docs: List[Dict[str, Any]],
        page_spans: Dict[str, Tuple[int, int]],
        prompt_path: Path,
    ):
        return generate_ai_prompt(docs, page_spans, prompt_path, self.mtg_info)

    # -------------------------------------------------------------------------
    # Helper Methods
    # -------------------------------------------------------------------------

    def _find_word_doc(self, folder: Path, tdoc_id: str) -> Optional[Path]:
        if not folder.exists():
            return None
        for ext in [".docx", ".docm", ".pptx", ".doc", ".ppt", ".pdf"]:
            files = list(folder.glob(f"*{ext}"))
            if files:
                for f in files:
                    if tdoc_id.lower() in f.name.lower():
                        return f
                return files[0]
        return None

    def _download_and_extract_tdoc(self, tdoc_id: str, tdoc_folder: Path) -> Optional[Path]:
        base_match = re.search(r"^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$", tdoc_id, re.IGNORECASE)
        is_revision = bool(base_match)

        candidate_urls: List[str] = []
        if is_revision:
            if self.revisions_url:
                candidate_urls.append(self.revisions_url)
            if self.docs_ftp_url and self.docs_ftp_url not in candidate_urls:
                candidate_urls.append(self.docs_ftp_url)
        else:
            if self.docs_ftp_url:
                candidate_urls.append(self.docs_ftp_url)
            if self.revisions_url and self.revisions_url not in candidate_urls:
                candidate_urls.append(self.revisions_url)

        if not candidate_urls:
            return None

        for base_url in candidate_urls:
            if self.isInterruptionRequested():
                return None
            try:
                TDocFileHandler.download_and_extract_tdoc(
                    tdoc_id,
                    base_url,
                    tdoc_folder,
                    timeout=15,
                    cancel_check=self.isInterruptionRequested,
                )
                doc_path = self._find_word_doc(tdoc_folder, tdoc_id)
                if doc_path:
                    return doc_path
            except Exception as e:
                logger.debug(
                    "[DossierExporter] Download attempt failed for %s from %s: %s",
                    tdoc_id,
                    base_url,
                    e,
                )

        return None
