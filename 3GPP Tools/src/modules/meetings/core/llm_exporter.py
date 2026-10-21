import logging
import os
import re
import shutil
from pathlib import Path
from typing import List, Optional

import pythoncom
import win32com.client
from PyQt5.QtCore import QThread, pyqtSignal

from core.utils.paths import to_extended_path
from modules.meetings.core.tdoc_file_handler import TDocFileHandler
from modules.meetings.core.tdoc_markdown_converter import (
    convert_document_to_markdown,
    sanitize_filename,
    write_index,
    ExportedDoc,
)

TDOC_MD_EXTRACTOR_VERSION = "2.0.0"
LLM_EXTRACTOR_VERSION = TDOC_MD_EXTRACTOR_VERSION  # Backwards-compatible alias


class TDocMarkdownExporterThread(QThread):
    """
    Asynchronous worker to convert and export 3GPP Technical Documents (.docx/.doc)
    into high-fidelity, standalone Markdown (.md) documents and an index.md manifest.
    """
    finished = pyqtSignal(bool, str)
    progress = pyqtSignal(str)

    def __init__(
        self,
        meeting_dir: Path,
        tdocs_list: list,
        docs_ftp_url: str,
        revisions_url: str,
        is_bulk: bool = True,
        output_dir: Optional[Path] = None,
        *args,
        **kwargs,
    ):
        super().__init__()
        self.meeting_dir = Path(meeting_dir)
        self.tdocs_list = tdocs_list
        self.docs_ftp_url = docs_ftp_url
        self.revisions_url = revisions_url
        self.is_bulk = is_bulk
        self.export_dir = Path(output_dir) if output_dir else (self.meeting_dir / "Export" / "Markdown")
        self.media_dir = self.export_dir / "media"

    def run(self):
        word_app = None
        try:
            pythoncom.CoInitialize()
            os.makedirs(to_extended_path(self.export_dir), exist_ok=True)
            os.makedirs(to_extended_path(self.media_dir), exist_ok=True)

            exported_docs: List[ExportedDoc] = []
            failed_tdocs: List[str] = []
            total = len(self.tdocs_list)

            for idx, tdoc_data in enumerate(self.tdocs_list, 1):
                if self.isInterruptionRequested():
                    logging.info("[Markdown Exporter] Interruption requested by user. Halting.")
                    break

                tdoc_id = str(tdoc_data.get("TDoc", "")).strip()
                if not tdoc_id:
                    continue

                base_match = re.search(r'^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$', tdoc_id, re.IGNORECASE)
                base_tdoc = base_match.group(1).upper() if base_match else tdoc_id.upper()
                doc_type = str(tdoc_data.get("Type", "Other"))
                raw_ai = str(tdoc_data.get("Agenda Item", "Unknown")).strip()
                title = str(tdoc_data.get("Title", "")).strip()
                status = str(tdoc_data.get("TDoc Status", "")).strip()

                msg = f"Processing {tdoc_id} ({idx}/{total})..."
                self.progress.emit(msg)
                logging.info(f"[Markdown Exporter] {msg}")

                tdoc_folder = self.meeting_dir / base_tdoc
                cache_file = tdoc_folder / f"{tdoc_id}_LLM_v{TDOC_MD_EXTRACTOR_VERSION}.md"
                ext_cache_file = to_extended_path(cache_file)
                target_md_file = self.export_dir / f"{sanitize_filename(tdoc_id)}.md"

                doc_obj: Optional[ExportedDoc] = None

                # 1. Use local cached conversion if available
                if os.path.exists(ext_cache_file):
                    logging.info(f"[Markdown Exporter] Using local conversion cache for {tdoc_id}")
                    try:
                        with open(ext_cache_file, "r", encoding="utf-8") as f:
                            md_content = f.read()

                        # Write to export folder
                        with open(to_extended_path(target_md_file), "w", encoding="utf-8") as f:
                            f.write(md_content)

                        images_count = 0
                        cached_media = tdoc_folder / "media"
                        if cached_media.is_dir():
                            for img in cached_media.glob(f"{tdoc_id}_*"):
                                images_count += 1
                                dst = self.media_dir / img.name
                                if not dst.exists():
                                    try:
                                        shutil.copy2(to_extended_path(img), to_extended_path(dst))
                                    except Exception:
                                        pass

                        doc_obj = ExportedDoc(
                            tdoc_id=tdoc_id,
                            name=tdoc_id,
                            md_path=target_md_file,
                            text=md_content,
                            meta=tdoc_data,
                            doc_type=doc_type,
                            agenda_item=raw_ai,
                            title=title,
                            status=status,
                            images=images_count,
                            warnings=[],
                        )
                    except Exception as cache_err:
                        logging.warning(f"[Markdown Exporter] Failed reading cache for {tdoc_id}: {cache_err}")

                # 2. Convert from source document if not cached
                if not doc_obj:
                    doc_path = self._find_word_doc(tdoc_folder, tdoc_id)
                    if not doc_path:
                        dl_msg = f"Downloading missing TDoc: {tdoc_id}..."
                        self.progress.emit(dl_msg)
                        logging.info(f"[Markdown Exporter] {dl_msg}")
                        doc_path = self._download_and_extract_tdoc(tdoc_id, tdoc_folder)

                    if doc_path:
                        ext_msg = f"Converting {tdoc_id} to Markdown..."
                        self.progress.emit(ext_msg)
                        logging.info(f"[Markdown Exporter] {ext_msg}")
                        try:
                            # Primary engine: Direct OpenXML extraction with tables, diagrams & track changes
                            doc_obj = convert_document_to_markdown(
                                doc_path,
                                out_dir=self.export_dir,
                                tdoc_id=tdoc_id,
                                meta=tdoc_data,
                                revisions="markup",
                                media_dir=self.media_dir,
                                media_rel="media",
                                cancellation_check=self.isInterruptionRequested,
                            )
                            # Update local cache in tdoc folder for future AI triage / re-exports
                            if doc_obj and doc_obj.text:
                                try:
                                    os.makedirs(to_extended_path(cache_file.parent), exist_ok=True)
                                    with open(ext_cache_file, "w", encoding="utf-8") as f:
                                        f.write(doc_obj.text)
                                except Exception:
                                    pass

                        except Exception as conv_err:
                            logging.warning(
                                f"[Markdown Exporter] Direct OpenXML conversion failed for {tdoc_id} ({conv_err}). "
                                f"Attempting Word COM fallback..."
                            )
                            if not word_app:
                                word_app = win32com.client.DispatchEx("Word.Application")
                                word_app.Visible = False
                                word_app.DisplayAlerts = 0

                            md_text = self._extract_from_word(word_app, doc_path, doc_type)
                            if md_text and not md_text.startswith("Error"):
                                if not md_text.strip().startswith(f"# {tdoc_id}"):
                                    header = f"# {tdoc_id}: {title}\n**Source:** {tdoc_data.get('Source', '')}\n**Agenda Item:** {raw_ai}\n\n---\n\n"
                                    md_text = header + md_text

                                with open(to_extended_path(target_md_file), "w", encoding="utf-8") as f:
                                    f.write(md_text)

                                try:
                                    os.makedirs(to_extended_path(cache_file.parent), exist_ok=True)
                                    with open(ext_cache_file, "w", encoding="utf-8") as f:
                                        f.write(md_text)
                                except Exception:
                                    pass

                                doc_obj = ExportedDoc(
                                    tdoc_id=tdoc_id,
                                    name=tdoc_id,
                                    md_path=target_md_file,
                                    text=md_text,
                                    meta=tdoc_data,
                                    doc_type=doc_type,
                                    agenda_item=raw_ai,
                                    title=title,
                                    status=status,
                                    images=0,
                                    warnings=[],
                                )

                    if not doc_obj:
                        warn_msg = f"Could not convert {tdoc_id} (missing or invalid Word document)."
                        logging.warning(f"[Markdown Exporter] {warn_msg}")
                        failed_tdocs.append(tdoc_id)

                if doc_obj:
                    exported_docs.append(doc_obj)

            # 3. Create index.md manifest for multi-file exports
            if exported_docs:
                try:
                    write_index(self.export_dir, exported_docs)
                    logging.info(f"[Markdown Exporter] Wrote index manifest in {self.export_dir}")
                except Exception as idx_err:
                    logging.warning(f"[Markdown Exporter] Could not write index.md: {idx_err}")

            if self.is_bulk:
                msg = f"Successfully exported {len(exported_docs)} Markdown document(s) and index manifest to:\n{self.export_dir}"
            else:
                names = ", ".join(d.name for d in exported_docs)
                msg = f"Successfully exported {names} to:\n{self.export_dir}"

            if failed_tdocs:
                sample_failed = ", ".join(failed_tdocs[:5])
                msg += f"\n\n⚠️ Failed to convert ({len(failed_tdocs)} document(s)): {sample_failed}{'...' if len(failed_tdocs) > 5 else ''}"

            self.finished.emit(True, msg)

        except Exception as e:
            logging.error(f"[Markdown Exporter] Critical thread failure: {e}", exc_info=True)
            self.finished.emit(False, str(e))
        finally:
            if word_app:
                try:
                    word_app.DisplayAlerts = 0
                    word_app.Quit(SaveChanges=0)
                except Exception:
                    pass
                del word_app
            pythoncom.CoUninitialize()

    def _find_word_doc(self, folder: Path, tdoc_id: str) -> Optional[Path]:
        if not folder.exists():
            return None
        for ext in [".docx", ".doc"]:
            files = list(folder.glob(f"*{ext}"))
            if files:
                for f in files:
                    if tdoc_id.lower() in f.name.lower():
                        return f
                return files[0]
        return None

    def _download_and_extract_tdoc(self, tdoc_id: str, tdoc_folder: Path) -> Optional[Path]:
        base_match = re.search(r'^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$', tdoc_id, re.IGNORECASE)
        is_revision = bool(base_match)

        base_url = self.revisions_url if (is_revision and self.revisions_url) else self.docs_ftp_url
        if not base_url:
            logging.error(f"[Markdown Exporter] No base URL configured for downloading {tdoc_id}")
            return None

        try:
            logging.info(f"[Markdown Exporter] Triggering TDocFileHandler for {tdoc_id} against {base_url}")
            TDocFileHandler.download_and_extract_tdoc(tdoc_id, base_url, tdoc_folder)
            return self._find_word_doc(tdoc_folder, tdoc_id)
        except Exception as e:
            err_msg = f"Failed to download {tdoc_id}: {str(e)}"
            self.progress.emit(err_msg)
            logging.error(f"[Markdown Exporter] {err_msg}", exc_info=True)
            return None

    def _extract_from_word(self, word_app, doc_path: Path, doc_type: str) -> str:
        doc = None
        try:
            doc = word_app.Documents.Open(str(doc_path), False, True, False)
            md_lines = []
            in_new_block = False

            all_new_trigger = re.compile(r'(?i)all (?:new )?text (?:is )?(?:new|added)')
            placeholder_clause_trigger = re.compile(r'^(\d+\.)+[a-zA-Z]$')
            numeric_clause_trigger = re.compile(r'^(\d+\.)+\d+$')
            boundary_trigger = re.compile(r'(?i)<[-\s]*(next|end of)\s*change')

            for para in doc.Paragraphs:
                text = para.Range.Text.strip('\r\x07\x0b\xa0 ').strip()
                if not text:
                    continue

                style_name = ""
                try:
                    style_name = para.Style.NameLocal
                except Exception:
                    pass

                if boundary_trigger.search(text):
                    in_new_block = False
                    md_lines.append(f"\n*[{text.strip()}]*\n")
                    continue

                words = text.split()
                if in_new_block and "Heading" in style_name and words and numeric_clause_trigger.match(words[0]):
                    in_new_block = False

                if all_new_trigger.search(text):
                    in_new_block = True
                    md_lines.append("\n> **Note:** Entering 'All Text New' block.\n")
                    continue

                if "Heading" in style_name and words and placeholder_clause_trigger.match(words[0]):
                    in_new_block = True
                    md_lines.append(f"\n> **Note:** Entering placeholder clause '{words[0]}'.\n")

                if in_new_block:
                    md_lines.append(f"[ADDED BLOCK]: {text}")
                else:
                    if "CR" in doc_type or "pCR" in doc_type:
                        revs = para.Range.Revisions
                        if revs.Count > 0:
                            inserted, deleted = [], []
                            for rev in revs:
                                if rev.Type == 1:
                                    inserted.append(rev.Range.Text.strip('\r\x07\x0b '))
                                elif rev.Type == 2:
                                    deleted.append(rev.Range.Text.strip('\r\x07\x0b '))

                            prefix = ""
                            if inserted:
                                prefix += f"[INSERTED: {', '.join(inserted)}] "
                            if deleted:
                                prefix += f"[DELETED: {', '.join(deleted)}] "
                            md_lines.append(f"{prefix}{text}")
                        else:
                            md_lines.append(text)
                    else:
                        if "Heading" in style_name:
                            depth = ''.join(filter(str.isdigit, style_name))
                            prefix = "#" * int(depth) if depth else "##"
                            md_lines.append(f"\n{prefix} {text}\n")
                        else:
                            md_lines.append(text)

            return "\n\n".join(md_lines)

        except Exception as e:
            return f"Error parsing document: {str(e)}"
        finally:
            if doc:
                try:
                    doc.Close(SaveChanges=False)
                except Exception:
                    pass


# Backwards compatibility alias
LLMExporterThread = TDocMarkdownExporterThread
