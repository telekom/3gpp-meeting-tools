import ctypes
import faulthandler
import gc
import logging
import os
from pathlib import Path
import shutil
import stat
import subprocess
import tempfile
import time
from contextlib import contextmanager
from typing import Optional, Union

import pythoncom
import win32com.client
from PyQt5.QtCore import QThread, pyqtSignal

from core.utils.utils import get_proxies
from modules.word_tools.core.sensitivity_label import set_sensitivity_label
from modules.word_tools.core.libreoffice_converter import (
    convert_doc_to_docx_libreoffice,
    convert_document_libreoffice,
    is_libreoffice_available,
    get_libreoffice_missing_msg,
)
from modules.word_tools.core.word_media_repair import (
    create_pdf_safe_docx,
    list_legacy_metafiles,
)

logger = logging.getLogger(__name__)

# Word WdSaveFormat Constants
WD_FORMAT_DOC = 0
WD_FORMAT_XML_DOCX = 12      # wdFormatXMLDocument (.docx)
WD_FORMAT_DOC_DEFAULT = 16   # wdFormatDocumentDefault
WD_FORMAT_PDF = 17           # wdFormatPDF
WD_FORMAT_XPS = 18           # wdFormatXPS
WD_FORMAT_RTF = 6            # wdFormatRTF
WD_FORMAT_HTML = 8           # wdFormatHTML
WD_FORMAT_TXT = 2            # wdFormatText


def _sanitize_file_attributes(file_path: Path) -> None:
    """Removes read-only flags and NTFS Zone.Identifier streams using native Win32 APIs."""
    p = Path(file_path).resolve()
    if not p.exists():
        return

    try:
        os.chmod(p, stat.S_IWRITE | stat.S_IREAD)
    except Exception:
        pass

    if os.name == "nt":
        # Reset file attributes (removes Read-Only / Hidden flags)
        try:
            ctypes.windll.kernel32.SetFileAttributesW(str(p), 0x80)  # FILE_ATTRIBUTE_NORMAL
        except Exception:
            pass

        # Unlink the NTFS Zone.Identifier alternate data stream directly
        try:
            zone_stream = f"{str(p)}:Zone.Identifier"
            ctypes.windll.kernel32.DeleteFileW(zone_stream)
        except Exception:
            pass

        # Fallback via PowerShell Unblock-File
        try:
            subprocess.run(
                ["powershell", "-NoProfile", "-Command", f'Unblock-File -LiteralPath "{str(p)}"'],
                creationflags=0x08000000,  # CREATE_NO_WINDOW
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                timeout=2,
            )
        except Exception:
            pass



def _is_rpc_disconnect_error(exc: BaseException) -> bool:
    """Detect Word COM/RPC disconnects after WINWORD becomes unavailable."""
    rpc_hresults = {
        -2147023170,  # 0x800706BE RPC_S_CALL_FAILED
        -2147023174,  # 0x800706BA RPC_S_SERVER_UNAVAILABLE
        -2147417848,  # 0x80010108 RPC_E_DISCONNECTED
        -2147418111,  # 0x80010001 RPC_E_CALL_REJECTED
    }
    hresult = getattr(exc, "hresult", None)
    if isinstance(hresult, int) and hresult in rpc_hresults:
        return True
    msg = str(exc).lower()
    return any(marker in msg for marker in (
        "0x800706be", "0x800706ba", "0x80010108", "0x80010001",
        "remote procedure call failed", "rpc server is unavailable",
        "object invoked has disconnected", "call was rejected by callee",
    ))


@contextmanager
def _suppress_faulthandler_for_word_com():
    """
    Suppress only faulthandler's process-wide native dump during fragile Word
    export calls. Python/pywin32 exceptions still propagate normally.
    """
    try:
        was_enabled = faulthandler.is_enabled()
    except Exception:
        was_enabled = False
    try:
        if was_enabled:
            try:
                faulthandler.disable()
            except Exception:
                pass
        yield
    finally:
        if was_enabled:
            try:
                faulthandler.enable()
            except Exception:
                pass


def _safe_com_cleanup(word=None, doc=None, new_doc=None, com_server_dead: bool = False) -> None:
    """
    Safely closes Word documents, quits Word, releases COM pointers,
    and uninitializes the COM apartment without triggering RPC SEH faults.
    """
    # Temporarily shield faulthandler so harmless first-chance RPC server
    # disconnects during WINWORD exit are not dumped to stderr as fatal errors.
    was_faulthandler_enabled = False
    try:
        if faulthandler.is_enabled():
            was_faulthandler_enabled = True
            faulthandler.disable()
    except Exception:
        pass

    try:
        # A disconnected WINWORD RPC server must not receive further COM calls.
        if com_server_dead:
            new_doc = None
            doc = None
            word = None

        # 1. Wait for any background printing/exporting to complete
        if word is not None and not com_server_dead:
            try:
                while word.BackgroundPrintingStatus > 0:
                    time.sleep(0.05)
            except Exception:
                pass

        # 2. Safely close documents
        for d in (() if com_server_dead else (new_doc, doc)):
            if d is not None:
                try:
                    d.Close(SaveChanges=False)
                except Exception:
                    pass

        # 3. Safely close any lingering Protected View windows
        if word is not None and not com_server_dead:
            try:
                if hasattr(word, "ProtectedViewWindows"):
                    while word.ProtectedViewWindows.Count > 0:
                        word.ProtectedViewWindows.Item(1).Close()
            except Exception:
                pass

            # 4. Quit Word Application
            try:
                word.Quit(SaveChanges=0)  # 0 = wdDoNotSaveChanges
            except Exception:
                pass

    finally:
        # 5. Clear all COM references in local scope
        new_doc = None
        doc = None
        word = None

        # 6. Force garbage collection so win32com wrapper destructors execute
        try:
            gc.collect()
        except Exception:
            pass

        # 7. Pump any remaining COM messages
        try:
            pythoncom.PumpWaitingMessages()
        except Exception:
            pass

        # 8. Uninitialize COM apartment
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

        # 9. Restore faulthandler if it was active
        if was_faulthandler_enabled:
            try:
                faulthandler.enable()
            except Exception:
                pass



def _is_valid_pdf(path: Path) -> bool:
    try:
        if not path.exists() or path.stat().st_size < 5:
            return False
        with path.open("rb") as fh:
            return fh.read(5) == b"%PDF-"
    except OSError:
        return False


def _export_pdf_with_word_once(source: Path, target: Path) -> None:
    """
    Perform one isolated Word PDF export.

    Used by the media-repair diagnostic path.  Each test gets a fresh Word COM
    instance so a failed PDF renderer cannot contaminate the next candidate.
    """
    word = None
    doc = None
    com_server_dead = False
    try:
        pythoncom.CoInitialize()
        if target.exists():
            try:
                target.unlink()
            except OSError:
                pass

        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0

        try:
            word.AutomationSecurity = 1
            word.Options.ConfirmConversions = False
            word.Options.DoNotPromptForConvert = True
            word.Options.WarnBeforeSavingPrintingSendingMarkup = False
            word.Options.SaveInterval = 0
            word.Options.UpdateFieldsAtPrint = False
            word.Options.UpdateLinksAtPrint = False
            word.Options.PrintBackground = False
        except Exception:
            pass

        doc = word.Documents.Open(
            FileName=os.path.normpath(str(source)),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
        )

        try:
            set_sensitivity_label(doc)
        except Exception:
            pass

        with _suppress_faulthandler_for_word_com():
            doc.ExportAsFixedFormat(
                OutputFileName=os.path.normpath(str(target)),
                ExportFormat=WD_FORMAT_PDF,
                OpenAfterExport=False,
                OptimizeFor=0,
                IncludeDocProps=True,
                CreateBookmarks=1,
            )

        try:
            while word.BackgroundPrintingStatus > 0:
                time.sleep(0.05)
        except Exception:
            pass

        if not _is_valid_pdf(target):
            raise RuntimeError("Word returned without producing a valid PDF.")
    except Exception as exc:
        com_server_dead = _is_rpc_disconnect_error(exc)
        raise
    finally:
        _safe_com_cleanup(word=word, doc=doc, com_server_dead=com_server_dead)


def convert_doc_to_docx_word(
    doc_path: Union[str, Path],
    output_path: Optional[Union[str, Path]] = None,
    logger: Optional[logging.Logger] = None,
) -> Path:
    """Synchronously converts a legacy .doc to .docx using Microsoft Word COM automation."""
    log = logger or logging.getLogger(__name__)
    source = Path(doc_path).resolve()

    if not source.exists():
        raise FileNotFoundError(f"Source document not found: {source}")

    if source.suffix.lower() == ".docx":
        return source

    target = Path(output_path).resolve() if output_path else source.with_suffix(".docx")
    if target.exists() and target.stat().st_size > 0:
        return target

    _sanitize_file_attributes(source)
    target.parent.mkdir(parents=True, exist_ok=True)
    if target.exists():
        try:
            _sanitize_file_attributes(target)
            target.unlink()
        except Exception:
            pass

    temp_dir = Path(tempfile.gettempdir())
    temp_target = temp_dir / f"stage_{source.stem}.docx"
    if temp_target.exists():
        try:
            temp_target.unlink()
        except Exception:
            pass

    source_str = os.path.normpath(str(source))
    temp_target_str = os.path.normpath(str(temp_target))

    word = None
    doc = None
    new_doc = None
    saved = False

    try:
        pythoncom.CoInitialize()

        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0

        try:
            word.AutomationSecurity = 1  # msoAutomationSecurityLow
            word.Options.ConfirmConversions = False
            word.Options.DoNotPromptForConvert = True
            word.Options.WarnBeforeSavingPrintingSendingMarkup = False
            word.Options.SaveInterval = 0
            word.Options.PrintBackground = False
        except Exception:
            pass

        try:
            doc = word.Documents.Open(
                FileName=source_str,
                ConfirmConversions=False,
                ReadOnly=False,
                AddToRecentFiles=False,
            )
        except Exception:
            try:
                if hasattr(word, "ProtectedViewWindows") and word.ProtectedViewWindows.Count > 0:
                    pv = word.ProtectedViewWindows.Item(1)
                    doc = pv.Edit()
                else:
                    doc = word.Documents.Open(
                        FileName=source_str,
                        ConfirmConversions=False,
                        ReadOnly=True,
                        AddToRecentFiles=False,
                    )
            except Exception:
                raise RuntimeError("MS Word failed to open source document.")

        if doc is None:
            raise RuntimeError(f"Word failed to acquire a valid document handle for {source.name}")

        try:
            doc.Convert()
        except Exception:
            pass

        try:
            set_sensitivity_label(doc)
        except Exception:
            pass

        try:
            doc.SaveAs2(FileName=temp_target_str, FileFormat=WD_FORMAT_XML_DOCX)
            if temp_target.exists() and temp_target.stat().st_size > 0:
                saved = True
        except Exception:
            pass

        if not saved:
            try:
                doc.SaveAs2(FileName=temp_target_str, FileFormat=WD_FORMAT_DOC_DEFAULT)
                if temp_target.exists() and temp_target.stat().st_size > 0:
                    saved = True
            except Exception:
                pass

        if not saved or not temp_target.exists() or temp_target.stat().st_size == 0:
            new_doc = word.Documents.Add()
            new_doc.Content.FormattedText = doc.Content.FormattedText
            new_doc.SaveAs2(FileName=temp_target_str, FileFormat=WD_FORMAT_XML_DOCX)
            new_doc.Close(SaveChanges=False)
            new_doc = None
            if temp_target.exists() and temp_target.stat().st_size > 0:
                saved = True

        if not temp_target.exists() or temp_target.stat().st_size == 0:
            raise RuntimeError(f"MS Word failed to write target .docx for {source.name}")

        shutil.copy2(temp_target, target)
        _sanitize_file_attributes(target)
        try:
            temp_target.unlink()
        except Exception:
            pass

        log.info(f"Successfully converted via MS Word: {source.name} -> {target.name}")
        return target

    finally:
        _safe_com_cleanup(word=word, doc=doc, new_doc=new_doc)
        word = None
        doc = None
        new_doc = None


def convert_doc_to_docx(
    doc_path: Union[str, Path],
    output_path: Optional[Union[str, Path]] = None,
    logger: Optional[logging.Logger] = None,
) -> Path:
    """Primary Entry Point: Attempts Word COM conversion first, falling back to LibreOffice."""
    log = logger or logging.getLogger(__name__)
    source = Path(doc_path).resolve()

    try:
        return convert_doc_to_docx_word(source, output_path=output_path, logger=log)
    except Exception as word_err:
        log.warning(f"Word COM conversion failed for '{source.name}' ({word_err}). Falling back to LibreOffice...")
        try:
            return convert_doc_to_docx_libreoffice(source, output_path=output_path, logger=log)
        except Exception as lo_err:
            log.error(f"Both Word and LibreOffice conversions failed for '{source.name}'. LibreOffice error: {lo_err}")
            raise


class WordConverterThread(QThread):
    ui_log_msg = pyqtSignal(str, int)
    finished_path = pyqtSignal(str)
    finished = pyqtSignal()

    FORMAT_MAP = {
        "pdf": WD_FORMAT_PDF,
        "html": WD_FORMAT_HTML,
        "xps": WD_FORMAT_XPS,
        "rtf": WD_FORMAT_RTF,
        "txt": WD_FORMAT_TXT,
        "docx": WD_FORMAT_XML_DOCX,
        "doc": WD_FORMAT_DOC,
    }

    def __init__(self, doc_source: str, target_format: str, engine: str = "auto"):
        super().__init__()
        self.doc_source = doc_source
        self.target_format = target_format.lower().replace(".", "").strip()
        self.engine = engine.lower().strip()

    def _resolve_path(self, input_str: str) -> str:
        if not input_str:
            raise ValueError("Input document is empty. Please select a valid file or URL.")

        if input_str.startswith("http://") or input_str.startswith("https://"):
            if "sharepoint.com" in input_str.lower() or "onedrive" in input_str.lower():
                self._log("🔗 Corporate link detected. Delegating authentication to MS Word...", logging.INFO)
                return input_str.split("?")[0] if "?web=" in input_str else input_str

            self._log("⏳ Downloading document via proxy...", logging.INFO)
            import requests
            r = requests.get(input_str, allow_redirects=True, proxies=get_proxies(), timeout=30)
            r.raise_for_status()

            tmp_path = Path(tempfile.gettempdir()) / "puml2visio_conv_temp.docx"
            with open(tmp_path, "wb") as f:
                f.write(r.content)
            return str(tmp_path)

        return input_str

    def run(self):
        try:
            source_path = self._resolve_path(self.doc_source)
            source = Path(source_path).resolve()
            _sanitize_file_attributes(source)

            # 1. Explicit LibreOffice Engine Request
            if self.engine == "libreoffice" or (source.suffix.lower() == ".doc" and self.target_format == "docx_libreoffice"):
                if not is_libreoffice_available():
                    self._log(get_libreoffice_missing_msg(), logging.ERROR)
                    return
                self._log(f"⏳ Converting '{source.name}' to .{self.target_format} using Headless LibreOffice...", logging.INFO)
                out_path = convert_document_libreoffice(source, target_format=self.target_format)
                self._log(f"✅ Conversion complete: {out_path.name}", logging.INFO)
                self.finished_path.emit(str(out_path))
                return

            # 2. Automated Word-first pipeline.
            try:
                self._run_word_export(source)
            except Exception as word_err:
                # PDF-specific recovery: we proved that malformed/problematic
                # EMF/WMF graphics can make Word's PDF renderer abort.  Repair
                # only a temporary DOCX and retry Word before using LibreOffice.
                if (
                    self.target_format == "pdf"
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
                    self.engine != "word"
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
        finally:
            self.finished.emit()

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

            shutil.copy2(repaired_pdf, final_target)
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

    def _run_word_export(self, source: Path):
        word = None
        doc = None
        success = False
        out_path = ""
        out_name = ""
        com_server_dead = False

        try:
            pythoncom.CoInitialize()
            out_dir = source.parent
            out_name = f"{source.stem}.{self.target_format}"
            out_path = str(out_dir / out_name)

            if self.target_format not in self.FORMAT_MAP:
                raise ValueError(f"Unsupported conversion format: {self.target_format}")

            self._log(f"⏳ Spawning Word Converter Engine for {out_name}...", logging.INFO)
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0

            try:
                word.AutomationSecurity = 1  # msoAutomationSecurityLow
                word.Options.ConfirmConversions = False
                word.Options.DoNotPromptForConvert = True
                word.Options.WarnBeforeSavingPrintingSendingMarkup = False
                word.Options.SaveInterval = 0
                word.Options.PrintBackground = False
            except Exception:
                pass

            try:
                doc = word.Documents.Open(
                    FileName=str(source),
                    ConfirmConversions=False,
                    ReadOnly=True,
                    AddToRecentFiles=False,
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

    def _log(self, message: str, level: int = logging.INFO) -> None:
        """
        Send worker output through the application's central logging pipeline.

        The root logging handlers send the same LogRecord to:
          - CLI
          - 3gpp_tools.log
          - GuiLogHandler / application console
        """
        logger.log(level, message)