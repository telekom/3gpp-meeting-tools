import logging
import os
from pathlib import Path
import shutil
import stat
import tempfile
from typing import Optional, Union

import pythoncom
import win32com.client
from PyQt5.QtCore import QThread, pyqtSignal

from core.utils.utils import get_proxies

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
    """Removes read-only flags and NTFS Zone.Identifier streams from the file."""
    try:
        if file_path.exists():
            os.chmod(file_path, stat.S_IWRITE | stat.S_IREAD)
    except Exception:
        pass

    try:
        zone_stream = Path(f"{file_path.resolve()}:Zone.Identifier")
        if zone_stream.exists():
            zone_stream.unlink()
    except Exception:
        pass


def convert_doc_to_docx(
    doc_path: Union[str, Path],
    output_path: Optional[Union[str, Path]] = None,
    logger: Optional[logging.Logger] = None,
) -> Path:
    """
    Synchronously converts a legacy Word 97-2003 (.doc) binary file to OpenXML (.docx).
    Uses isolated local %TEMP% staging to completely bypass OneDrive sync interceptors,
    followed by multi-stage SaveAs2 and clipboard-free FormattedText RAM cloning.
    """
    log = logger or logging.getLogger(__name__)
    source = Path(doc_path).resolve()

    if not source.exists():
        raise FileNotFoundError(f"Source document not found: {source}")

    if source.suffix.lower() == ".docx":
        return source

    target = Path(output_path).resolve() if output_path else source.with_suffix(".docx")
    if target.exists() and target.stat().st_size > 0:
        return target

    # 1. Clean attributes and prepare paths
    _sanitize_file_attributes(source)
    target.parent.mkdir(parents=True, exist_ok=True)
    if target.exists():
        try:
            _sanitize_file_attributes(target)
            target.unlink()
        except Exception:
            pass

    # Stage conversion in local %TEMP% to eliminate OneDrive co-authoring COM locks
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
        word.DisplayAlerts = 0  # wdAlertsNone

        try:
            word.AutomationSecurity = 1  # msoAutomationSecurityLow
            word.Options.ConfirmConversions = False
            word.Options.DoNotPromptForConvert = True
            word.Options.WarnBeforeSavingPrintingSendingMarkup = False
            word.Options.SaveInterval = 0  # Disable AutoRecover during batch automation
        except Exception:
            pass

        # 2. Open source document
        try:
            doc = word.Documents.Open(
                FileName=source_str,
                ConfirmConversions=False,
                ReadOnly=True,
                AddToRecentFiles=False,
            )
        except Exception:
            if hasattr(word, "ProtectedViewWindows") and word.ProtectedViewWindows.Count > 0:
                log.info(f"Unlocking Protected View window for {source.name}...")
                pv = word.ProtectedViewWindows.Item(1)
                doc = pv.Edit()
            else:
                doc = word.Documents.Open(
                    FileName=source_str,
                    ConfirmConversions=False,
                    ReadOnly=False,
                    AddToRecentFiles=False,
                )

        if doc is None:
            raise RuntimeError(f"Word failed to acquire a valid document handle for {source.name}")

        # 3. Apply Sensitivity Label if configured
        try:
            from modules.word_tools.core.sensitivity_label import set_sensitivity_label
            set_sensitivity_label(doc)
        except Exception:
            pass

        # 4. Primary Conversion: Direct SaveAs2 to isolated local temporary file
        try:
            doc.SaveAs2(FileName=temp_target_str, FileFormat=WD_FORMAT_XML_DOCX)
            if temp_target.exists() and temp_target.stat().st_size > 0:
                saved = True
        except Exception as err_save2:
            log.debug(f"Direct SaveAs2 failed ({err_save2}). Trying default format...")

        if not saved:
            try:
                doc.SaveAs2(FileName=temp_target_str, FileFormat=WD_FORMAT_DOC_DEFAULT)
                if temp_target.exists() and temp_target.stat().st_size > 0:
                    saved = True
            except Exception as err_def:
                log.debug(f"SaveAs2 default failed ({err_def}). Trying FormattedText clone...")

        # 5. Secondary Conversion: Direct RAM FormattedText clone (clipboard-free)
        if not saved or not temp_target.exists() or temp_target.stat().st_size == 0:
            log.info(f"Cloning formatted content stream for {source.name}...")
            new_doc = word.Documents.Add()

            try:
                from modules.word_tools.core.sensitivity_label import set_sensitivity_label
                set_sensitivity_label(new_doc)
            except Exception:
                pass

            # Duplicates complete formatting, tables, and Visio drawings directly in memory
            new_doc.Content.FormattedText = doc.Content.FormattedText
            new_doc.SaveAs2(FileName=temp_target_str, FileFormat=WD_FORMAT_XML_DOCX)
            new_doc.Close(SaveChanges=False)
            new_doc = None

            if temp_target.exists() and temp_target.stat().st_size > 0:
                saved = True

        # 6. Tertiary Fallback: Fresh InsertFile stream transfer (ensuring source handle is closed)
        if not saved or not temp_target.exists() or temp_target.stat().st_size == 0:
            log.info(f"Attempting detached InsertFile stream for {source.name}...")
            try:
                doc.Close(SaveChanges=False)
                doc = None
            except Exception:
                pass

            new_doc = word.Documents.Add()
            new_doc.Range(0, 0).InsertFile(FileName=source_str)
            new_doc.SaveAs2(FileName=temp_target_str, FileFormat=WD_FORMAT_XML_DOCX)
            new_doc.Close(SaveChanges=False)
            new_doc = None

            if temp_target.exists() and temp_target.stat().st_size > 0:
                saved = True

        # 7. Finalize: Copy staging file to actual destination
        if not temp_target.exists() or temp_target.stat().st_size == 0:
            raise RuntimeError(f"Target .docx file was not generated for {source.name}")

        shutil.copy2(temp_target, target)
        _sanitize_file_attributes(target)

        try:
            temp_target.unlink()
        except Exception:
            pass

        log.info(f"Successfully converted {source.name} -> {target.name}")
        return target

    except Exception as e:
        log.error(f"COM conversion failed for {source.name}: {e}")
        raise
    finally:
        if new_doc:
            try:
                new_doc.Close(SaveChanges=False)
            except Exception:
                pass

        if doc:
            try:
                doc.Close(SaveChanges=False)
            except Exception:
                pass

        if word:
            try:
                if hasattr(word, "ProtectedViewWindows"):
                    while word.ProtectedViewWindows.Count > 0:
                        word.ProtectedViewWindows.Item(1).Close()
                word.Quit()
            except Exception:
                pass

        pythoncom.CoUninitialize()


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

    def __init__(self, doc_source: str, target_format: str):
        super().__init__()
        self.doc_source = doc_source
        self.target_format = target_format.lower().replace(".", "")

    def _resolve_path(self, input_str: str) -> str:
        if not input_str:
            raise ValueError("Input document is empty. Please select a valid file, open document, or URL.")

        if input_str.startswith("http://") or input_str.startswith("https://"):
            if "sharepoint.com" in input_str.lower() or "onedrive" in input_str.lower():
                self.ui_log_msg.emit(
                    "🔗 Corporate link detected. Delegating secure authentication to MS Word...",
                    logging.INFO,
                )
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
        word = None
        doc = None
        try:
            pythoncom.CoInitialize()

            self.ui_log_msg.emit(f"⏳ Preparing document for {self.target_format.upper()} conversion...", logging.INFO)
            source_path = self._resolve_path(self.doc_source)
            source = Path(source_path).resolve()
            _sanitize_file_attributes(source)

            out_dir = source.parent
            out_name = source.stem + f".{self.target_format}"
            out_path = str(out_dir / out_name)

            if self.target_format not in self.FORMAT_MAP:
                raise ValueError(f"Unsupported conversion format: {self.target_format}")

            self.ui_log_msg.emit(f"⏳ Spawning detached Word Converter Engine for {out_name}...", logging.INFO)
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0

            try:
                word.AutomationSecurity = 1
            except Exception:
                pass

            try:
                doc = word.Documents.Open(str(source), False, True, False)
            except Exception:
                if hasattr(word, "ProtectedViewWindows") and word.ProtectedViewWindows.Count > 0:
                    doc = word.ProtectedViewWindows.Item(1).Edit()
                else:
                    raise

            self.ui_log_msg.emit(f"⏳ Converting and saving {out_name} to {self.target_format}...", logging.INFO)

            if self.target_format in ("pdf", "xps"):
                export_format = WD_FORMAT_PDF if self.target_format == "pdf" else WD_FORMAT_XPS
                word.Options.UpdateFieldsAtPrint = False
                word.Options.UpdateLinksAtPrint = False

                doc.ExportAsFixedFormat(
                    OutputFileName=out_path,
                    ExportFormat=export_format,
                    OpenAfterExport=False,
                    OptimizeFor=0,
                    IncludeDocProps=True,
                    CreateBookmarks=1,
                )
            else:
                doc.SaveAs2(out_path, FileFormat=self.FORMAT_MAP[self.target_format])

            self.ui_log_msg.emit(f"✅ Conversion complete: {out_name}", logging.INFO)
            self.finished_path.emit(out_path)

        except Exception as e:
            self.ui_log_msg.emit(f"❌ Conversion Error: {str(e)}", logging.ERROR)
        finally:
            try:
                if doc:
                    doc.Close(SaveChanges=False)
                if word:
                    word.Quit()
            except Exception:
                pass

            pythoncom.CoUninitialize()
            self.finished.emit()