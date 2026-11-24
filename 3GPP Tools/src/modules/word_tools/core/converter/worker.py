"""Background QThread converter supporting Word COM, PDF media repair, LibreOffice, and Markdown."""

from __future__ import annotations

import logging
import os
import shutil
import tempfile
import time
from pathlib import Path
from typing import Optional, Union

import pythoncom
import win32com.client
from PyQt5.QtCore import QThread, pyqtSignal

from core.network.session import NetworkSession
from core.utils.office_printer_guard import temporary_virtual_default_printer
from core.utils.paths import to_extended_path
from modules.word_tools.core.converter.com_utils import (
    WD_FORMAT_DOC,
    WD_FORMAT_HTML,
    WD_FORMAT_PDF,
    WD_FORMAT_RTF,
    WD_FORMAT_TXT,
    WD_FORMAT_XML_DOCX,
    WD_FORMAT_XPS,
    _configure_unattended_word,
    _is_rpc_disconnect_error,
    _is_valid_pdf,
    _safe_com_cleanup,
    _sanitize_file_attributes,
    _suppress_faulthandler_for_word_com,
)
from modules.word_tools.core.converter.doc_to_docx import _export_pdf_with_word_once
from modules.word_tools.core.libreoffice_converter import (
    convert_document_libreoffice,
    get_libreoffice_missing_msg,
    is_libreoffice_available,
)
from modules.word_tools.core.sensitivity_label import set_sensitivity_label
from modules.word_tools.core.word_media_repair import (
    create_pdf_safe_docx,
    list_legacy_metafiles,
)

logger = logging.getLogger(__name__)


class WordConverterThread(QThread):
    ui_log_msg = pyqtSignal(str, int)
    finished_path = pyqtSignal(str)
    conversion_failed = pyqtSignal(str)

    FORMAT_MAP = {
        "pdf": WD_FORMAT_PDF,
        "html": WD_FORMAT_HTML,
        "xps": WD_FORMAT_XPS,
        "rtf": WD_FORMAT_RTF,
        "txt": WD_FORMAT_TXT,
        "docx": WD_FORMAT_XML_DOCX,
        "doc": WD_FORMAT_DOC,
    }

    def __init__(
        self,
        doc_source: str,
        target_format: str,
        engine: str = "auto",
        revisions: str = "markup",
        extract_media: bool = True,
        output_dir: Optional[Union[str, Path]] = None,
        include_header: bool = True,
    ):
        super().__init__()
        self.doc_source = doc_source
        self.target_format = target_format.lower().replace(".", "").strip()
        self.engine = engine.lower().strip()
        self.revisions = revisions
        self.extract_media = extract_media
        self.output_dir = str(output_dir) if output_dir else None
        self.include_header = include_header
        self._downloaded_temp: Optional[Path] = None

    def _resolve_path(self, input_str: str) -> str:
        if not input_str:
            raise ValueError("Input document is empty. Please select a valid file or URL.")

        if input_str.startswith("http://") or input_str.startswith("https://"):
            if "sharepoint.com" in input_str.lower() or "onedrive" in input_str.lower():
                self._log("🔗 Corporate link detected. Delegating authentication to MS Word...", logging.INFO)
                return input_str.split("?")[0] if "?web=" in input_str else input_str

            self._log("⏳ Downloading document via NetworkSession...", logging.INFO)
            session = NetworkSession.get_instance()
            r = session.get(input_str, allow_redirects=True, timeout=30)
            r.raise_for_status()

            url_path = input_str.split("?")[0]
            ext = Path(url_path).suffix if Path(url_path).suffix else ".docx"
            fd, tmp_str = tempfile.mkstemp(prefix="puml2visio_conv_", suffix=ext)
            os.close(fd)
            tmp_path = Path(tmp_str)
            with open(tmp_path, "wb") as f:
                f.write(r.content)
            self._downloaded_temp = tmp_path
            return str(tmp_path)

        return input_str

    def run(self):
        self._downloaded_temp = None
        temp_snapshot_to_cleanup = None
        try:
            source_path = self._resolve_path(self.doc_source)
            is_remote_url = source_path.startswith("http://") or source_path.startswith("https://")
            if is_remote_url:
                source = Path(source_path.split("?")[0].split("/")[-1] or "remote_document.docx")
                original_source = source
            else:
                original_source = Path(source_path).resolve()
                _sanitize_file_attributes(original_source)

                if original_source.suffix.lower() in ('.pptx', '.ppt'):
                    source = original_source
                else:
                    from modules.word_tools.core.document_opener import ensure_readable_word_document
                    readable_path, temp_snapshot_to_cleanup = ensure_readable_word_document(
                        original_source,
                        logger_fn=lambda msg: self._log(msg, logging.INFO),
                    )
                    source = readable_path

            # 1. Markdown Export Pipeline (OpenXML / tdoc_markdown_converter)
            if self.target_format in ("md", "markdown"):
                self._run_markdown_export(
                    source,
                    raw_source=source_path if is_remote_url else None,
                    original_source=original_source,
                )
                return

            # 2. Explicit LibreOffice Engine Request
            if not is_remote_url and (
                self.engine == "libreoffice" or (source.suffix.lower() == ".doc" and self.target_format == "docx_libreoffice")
            ):
                if not is_libreoffice_available():
                    self._log(get_libreoffice_missing_msg(), logging.ERROR)
                    return
                self._log(f"⏳ Converting '{source.name}' to .{self.target_format} using Headless LibreOffice...", logging.INFO)
                out_path = convert_document_libreoffice(source, target_format=self.target_format)
                self._log(f"✅ Conversion complete: {out_path.name}", logging.INFO)
                self.finished_path.emit(str(out_path))
                return

            # 3. Automated Word-first pipeline.
            try:
                self._run_word_export(source, raw_source=source_path if is_remote_url else None)
            except Exception as word_err:
                # PDF-specific recovery: we proved that malformed/problematic
                # EMF/WMF graphics can make Word's PDF renderer abort.  Repair
                # only a temporary DOCX and retry Word before using LibreOffice.
                if (
                    not is_remote_url
                    and self.target_format == "pdf"
                    and source.suffix.lower() == ".docx"
                ):
                    try:
                        repaired = self._try_pdf_media_repair(source, word_err)
                        if repaired:
                            return
                    except Exception as repair_err:
                        self._log(
                            f"⚠️ Legacy-graphic PDF repair did not recover the export "
                            f"({repair_err}).",
                            logging.WARNING,
                        )

                if (
                    not is_remote_url
                    and self.engine != "word"
                    and is_libreoffice_available()
                    and self.target_format in ("docx", "pdf", "html", "rtf", "txt")
                ):
                    self._log(
                        f"⚠️ Word COM conversion failed ({word_err}). "
                        f"Initiating LibreOffice fallback...",
                        logging.WARNING,
                    )
                    out_path = convert_document_libreoffice(
                        source, target_format=self.target_format
                    )
                    self._log(
                        f"✅ Conversion complete (via LibreOffice): {out_path.name}",
                        logging.INFO,
                    )
                    self.finished_path.emit(str(out_path))
                else:
                    raise word_err

        except Exception as e:
            self._log(f"❌ Conversion Error: {str(e)}", logging.ERROR)
            self.conversion_failed.emit(str(e))
        finally:
            if temp_snapshot_to_cleanup and temp_snapshot_to_cleanup.exists():
                try:
                    temp_snapshot_to_cleanup.unlink()
                except Exception:
                    pass
            if self._downloaded_temp and self._downloaded_temp.exists():
                try:
                    self._downloaded_temp.unlink()
                except Exception:
                    pass
            self._downloaded_temp = None

    def _try_pdf_media_repair(self, source: Path, original_error: Exception) -> bool:
        """
        Retry a failed Word PDF export once after rasterizing all EMF/WMF media
        in a temporary DOCX.

        This favors predictable conversion time over forensic identification of
        the individual bad metafile. The original DOCX is never modified.
        """
        media = list_legacy_metafiles(source)
        if not media:
            self._log(
                "ℹ️ Word PDF export failed, but the DOCX contains no EMF/WMF "
                "graphics to repair.",
                logging.INFO,
            )
            return False

        self._log(
            f"⚠️ Word PDF export failed ({original_error}). "
            f"Found {len(media)} legacy EMF/WMF graphic(s).",
            logging.WARNING,
        )

        work_dir = Path(tempfile.mkdtemp(prefix="3gpp_word_pdf_repair_"))
        repaired_docx = work_dir / f"{source.stem}_pdf_safe.docx"
        repaired_pdf = work_dir / f"{source.stem}_pdf_safe.pdf"
        final_target = source.with_suffix(".pdf")

        try:
            def ui_log(message: str) -> None:
                self._log(message, logging.INFO)

            replacements = create_pdf_safe_docx(
                source_docx=source,
                output_docx=repaired_docx,
                log=ui_log,
            )

            if not replacements:
                return False

            self._log(
                "⏳ Retrying Word PDF export once using the temporary "
                "PDF-safe document...",
                logging.INFO,
            )

            _export_pdf_with_word_once(repaired_docx, repaired_pdf)

            if not _is_valid_pdf(repaired_pdf):
                raise RuntimeError(
                    "Word did not produce a valid PDF from the repaired document."
                )

            if final_target.exists():
                try:
                    _sanitize_file_attributes(final_target)
                    final_target.unlink()
                except OSError:
                    pass

            shutil.copy2(to_extended_path(repaired_pdf), to_extended_path(final_target))
            _sanitize_file_attributes(final_target)

            self._log(
                f"✅ PDF conversion recovered after rasterizing "
                f"{len(replacements)} legacy EMF/WMF graphic(s) in the "
                f"temporary conversion copy. Original DOCX was not modified.",
                logging.INFO,
            )
            self.finished_path.emit(str(final_target))
            return True

        finally:
            shutil.rmtree(work_dir, ignore_errors=True)

    def _run_word_export(self, source: Path, raw_source: Optional[str] = None):
        word = None
        doc = None
        success = False
        out_path = ""
        out_name = ""
        com_server_dead = False

        try:
            pythoncom.CoInitialize()
            if raw_source:
                from core.utils.paths import get_project_root
                out_dir = get_project_root().parent / "output"
                out_dir.mkdir(parents=True, exist_ok=True)
            else:
                out_dir = source.parent
            out_name = f"{source.stem}.{self.target_format}"
            out_path = str(out_dir / out_name)

            if self.target_format not in self.FORMAT_MAP:
                raise ValueError(f"Unsupported conversion format: {self.target_format}")

            self._log(f"⏳ Spawning Word Converter Engine for {out_name}...", logging.INFO)
            with temporary_virtual_default_printer(log=logger):
                word = win32com.client.DispatchEx("Word.Application")
                _configure_unattended_word(word, log=logger)

                open_target = raw_source if raw_source else str(source)
                try:
                    doc = word.Documents.Open(
                        FileName=open_target,
                        ConfirmConversions=False,
                        ReadOnly=True,
                        AddToRecentFiles=False,
                        PasswordDocument="__3gpp_no_pw__",
                        WritePasswordDocument="__3gpp_no_pw__",
                        NoEncodingDialog=True,
                    )
                except Exception as open_err:
                    try:
                        if hasattr(word, "ProtectedViewWindows") and word.ProtectedViewWindows.Count > 0:
                            pv = word.ProtectedViewWindows.Item(1)
                            doc = pv.Edit()
                    except Exception:
                        pass
                    if doc is None:
                        raise open_err

                try:
                    set_sensitivity_label(doc)
                except Exception:
                    pass

                self._log(f"⏳ Converting and saving {out_name} to {self.target_format}...", logging.INFO)

                if self.target_format in ("pdf", "xps"):
                    export_format = WD_FORMAT_PDF if self.target_format == "pdf" else WD_FORMAT_XPS
                    word.Options.UpdateFieldsAtPrint = False
                    word.Options.UpdateLinksAtPrint = False
                    word.Options.PrintBackground = False

                    try:
                        with _suppress_faulthandler_for_word_com():
                            doc.ExportAsFixedFormat(
                                OutputFileName=out_path,
                                ExportFormat=export_format,
                                OpenAfterExport=False,
                                OptimizeFor=0,
                                IncludeDocProps=True,
                                CreateBookmarks=1,
                            )
                    except Exception as export_err:
                        com_server_dead = _is_rpc_disconnect_error(export_err)
                        if com_server_dead:
                            raise RuntimeError(
                                f"Microsoft Word export lost its COM/RPC connection ({export_err})."
                            ) from export_err
                        raise

                    # Ensure Word background printer thread has completely flushed
                    try:
                        while word.BackgroundPrintingStatus > 0:
                            time.sleep(0.05)
                    except Exception:
                        pass
                else:
                    doc.SaveAs2(out_path, FileFormat=self.FORMAT_MAP[self.target_format])

            if self.target_format == "pdf" and not _is_valid_pdf(Path(out_path)):
                raise RuntimeError(
                    "Microsoft Word returned from PDF export without producing a valid PDF."
                )

            success = True

        except Exception as exc:
            if _is_rpc_disconnect_error(exc):
                com_server_dead = True
            raise
        finally:
            # Do not call back into WINWORD after its RPC server has disconnected.
            _safe_com_cleanup(
                word=word,
                doc=doc,
                com_server_dead=com_server_dead,
            )
            word = None
            doc = None

        if success:
            self._log(f"✅ Conversion complete: {out_name}", logging.INFO)
            self.finished_path.emit(out_path)

    def _run_markdown_export(
        self,
        source: Path,
        raw_source: Optional[str] = None,
        original_source: Optional[Path] = None,
    ) -> None:
        """Export document to Markdown with figure extraction and revision tracking."""
        from core.utils.paths import get_project_root, strip_extended_path
        from modules.meetings.core.tdocs.tdoc_markdown_converter import convert_document_to_markdown

        orig = original_source or source
        if self.output_dir:
            out_dir = Path(self.output_dir).resolve()
        elif raw_source:
            out_dir = get_project_root().parent / "output"
        else:
            out_dir = orig.parent
        out_dir.mkdir(parents=True, exist_ok=True)

        self._log(f"⏳ Converting '{orig.name}' to Markdown (revisions: {self.revisions})...", logging.INFO)

        media_folder = out_dir / f"{orig.stem}_media" if self.extract_media else None
        media_rel = f"{orig.stem}_media"

        exported = convert_document_to_markdown(
            doc_path=source,
            out_dir=out_dir,
            tdoc_id=orig.stem,
            revisions=self.revisions,
            media_dir=media_folder,
            media_rel=media_rel,
            convert_images=self.extract_media,
            include_header=self.include_header,
        )

        clean_md_path = strip_extended_path(str(exported.md_path))
        img_info = f" (extracted {exported.images} image/drawing(s))" if exported.images else ""
        self._log(f"✅ Markdown conversion complete: {Path(clean_md_path).name}{img_info}", logging.INFO)
        self.finished_path.emit(clean_md_path)

    def _log(self, message: str, level: int = logging.INFO) -> None:
        """
        Send worker output through the application's central logging pipeline.

        The root logging handlers send the same LogRecord to:
          - CLI
          - 3gpp_tools.log
          - GuiLogHandler / application console
        """
        logger.log(level, message)
