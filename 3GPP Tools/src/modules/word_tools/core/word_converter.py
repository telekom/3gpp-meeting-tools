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


def _safe_com_cleanup(word=None, doc=None, new_doc=None) -> None:
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
        # 1. Wait for any background printing/exporting to complete
        if word is not None:
            try:
                while word.BackgroundPrintingStatus > 0:
                    time.sleep(0.05)
            except Exception:
                pass

        # 2. Safely close documents
        for d in (new_doc, doc):
            if d is not None:
                try:
                    d.Close(SaveChanges=False)
                except Exception:
                    pass

        # 3. Safely close any lingering Protected View windows
        if word is not None:
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
                self.ui_log_msg.emit("🔗 Corporate link detected. Delegating authentication to MS Word...", logging.INFO)
                return input_str.split("?")[0] if "?web=" in input_str else input_str

            self.ui_log_msg.emit("⏳ Downloading document via proxy...", logging.INFO)
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
                    self.ui_log_msg.emit(get_libreoffice_missing_msg(), logging.ERROR)
                    return
                self.ui_log_msg.emit(f"⏳ Converting '{source.name}' to .{self.target_format} using Headless LibreOffice...", logging.INFO)
                out_path = convert_document_libreoffice(source, target_format=self.target_format)
                self.ui_log_msg.emit(f"✅ Conversion complete: {out_path.name}", logging.INFO)
                self.finished_path.emit(str(out_path))
                return

            # 2. Automated Word-First Pipeline with LibreOffice Fallback
            try:
                self._run_word_export(source)
            except Exception as word_err:
                if self.engine != "word" and is_libreoffice_available() and self.target_format in ("docx", "pdf", "html", "rtf", "txt"):
                    self.ui_log_msg.emit(
                        f"⚠️ Word COM conversion failed ({word_err}). Initiating LibreOffice fallback...",
                        logging.WARNING,
                    )
                    out_path = convert_document_libreoffice(source, target_format=self.target_format)
                    self.ui_log_msg.emit(f"✅ Conversion complete (via LibreOffice): {out_path.name}", logging.INFO)
                    self.finished_path.emit(str(out_path))
                else:
                    raise word_err

        except Exception as e:
            self.ui_log_msg.emit(f"❌ Conversion Error: {str(e)}", logging.ERROR)
        finally:
            self.finished.emit()

    def _run_word_export(self, source: Path):
        word = None
        doc = None
        success = False
        out_path = ""
        out_name = ""

        try:
            pythoncom.CoInitialize()
            out_dir = source.parent
            out_name = f"{source.stem}.{self.target_format}"
            out_path = str(out_dir / out_name)

            if self.target_format not in self.FORMAT_MAP:
                raise ValueError(f"Unsupported conversion format: {self.target_format}")

            self.ui_log_msg.emit(f"⏳ Spawning Word Converter Engine for {out_name}...", logging.INFO)
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

            self.ui_log_msg.emit(f"⏳ Converting and saving {out_name} to {self.target_format}...", logging.INFO)

            if self.target_format in ("pdf", "xps"):
                export_format = WD_FORMAT_PDF if self.target_format == "pdf" else WD_FORMAT_XPS
                word.Options.UpdateFieldsAtPrint = False
                word.Options.UpdateLinksAtPrint = False
                word.Options.PrintBackground = False

                doc.ExportAsFixedFormat(
                    OutputFileName=out_path,
                    ExportFormat=export_format,
                    OpenAfterExport=False,
                    OptimizeFor=0,
                    IncludeDocProps=True,
                    CreateBookmarks=1,
                )

                # Ensure Word background printer thread has completely flushed
                try:
                    while word.BackgroundPrintingStatus > 0:
                        time.sleep(0.05)
                except Exception:
                    pass
            else:
                doc.SaveAs2(out_path, FileFormat=self.FORMAT_MAP[self.target_format])

            success = True

        finally:
            # Clean up Word and uninitialize COM BEFORE emitting finished signals
            _safe_com_cleanup(word=word, doc=doc)
            word = None
            doc = None

        if success:
            self.ui_log_msg.emit(f"✅ Conversion complete: {out_name}", logging.INFO)
            self.finished_path.emit(out_path)