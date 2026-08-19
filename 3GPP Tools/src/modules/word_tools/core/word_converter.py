import logging
from pathlib import Path
import tempfile
from typing import Optional, Union

import pythoncom
import win32com.client
from PyQt5.QtCore import QThread, pyqtSignal

from core.utils.utils import get_proxies


def _unblock_file(file_path: Path) -> None:
    """Removes the Windows Zone.Identifier NTFS stream to prevent Word Protected View."""
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
    Synchronously converts a legacy Word 97-2003 (.doc) binary file to OpenXML (.docx)
    using Microsoft Word COM automation with Protected View bypass and error recovery.
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

    # 1. Strip Mark-of-the-Web to prevent Protected View locking
    _unblock_file(source)

    # Ensure target directory exists and stale target is removed
    target.parent.mkdir(parents=True, exist_ok=True)
    if target.exists():
        try:
            target.unlink()
        except Exception:
            pass

    word = None
    doc = None
    try:
        pythoncom.CoInitialize()

        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0  # wdAlertsNone

        # Disable macro warning dialogs
        try:
            word.AutomationSecurity = 1  # msoAutomationSecurityLow
        except Exception:
            pass

        # 2. Open Document with fallback for Protected View
        try:
            doc = word.Documents.Open(
                FileName=str(source),
                ConfirmConversions=False,
                ReadOnly=True,
                AddToRecentFiles=False,
            )
        except Exception as open_err:
            if hasattr(word, "ProtectedViewWindows") and word.ProtectedViewWindows.Count > 0:
                log.info(f"Unlocking Protected View window for {source.name}...")
                pv = word.ProtectedViewWindows.Item(1)
                doc = pv.Edit()
            else:
                raise open_err

        if doc is None:
            raise RuntimeError(f"Word failed to acquire a valid document handle for {source.name}")

        # 3. Apply corporate Sensitivity Label if configured in word_tools
        try:
            from modules.word_tools.core.sensitivity_label import apply_configured_sensitivity_label
            apply_configured_sensitivity_label(doc)
        except Exception:
            pass

        # 4. Save as .docx (wdFormatXMLDocument = 16) with tiered fallbacks
        save_success = False
        save_errors = []

        try:
            doc.SaveAs2(str(target), 16)
            save_success = True
        except Exception as e1:
            save_errors.append(str(e1))

        if not save_success:
            try:
                doc.SaveAs2(FileName=str(target), FileFormat=16)
                save_success = True
            except Exception as e2:
                save_errors.append(str(e2))

        if not save_success:
            try:
                doc.SaveAs(str(target), 16)
                save_success = True
            except Exception as e3:
                save_errors.append(str(e3))

        if not save_success:
            raise RuntimeError(f"Failed to SaveAs2/SaveAs .docx: {'; '.join(save_errors)}")

        log.info(f"Successfully converted {source.name} -> {target.name}")
        return target

    except Exception as e:
        log.error(f"COM conversion failed for {source.name}: {e}")
        raise
    finally:
        try:
            if doc:
                doc.Close(SaveChanges=False)
        except Exception:
            pass

        try:
            if word and hasattr(word, "ProtectedViewWindows"):
                while word.ProtectedViewWindows.Count > 0:
                    word.ProtectedViewWindows.Item(1).Close()
        except Exception:
            pass

        try:
            if word:
                word.Quit()
        except Exception:
            pass

        pythoncom.CoUninitialize()


class WordConverterThread(QThread):
    ui_log_msg = pyqtSignal(str, int)
    finished_path = pyqtSignal(str)
    finished = pyqtSignal()

    FORMAT_MAP = {
        "pdf": 17,   # wdFormatPDF
        "html": 8,   # wdFormatHTML
        "xps": 18,   # wdFormatXPS
        "rtf": 6,    # wdFormatRTF
        "txt": 2,    # wdFormatText
        "docx": 16,  # wdFormatXMLDocument
        "doc": 0,    # wdFormatDocument
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
                self.ui_log_msg.emit("🔗 Corporate link detected. Delegating secure authentication to MS Word...",
                                     logging.INFO)
                return input_str.split("?")[0] if "?web=" in input_str else input_str

            self.ui_log_msg.emit("⏳ Downloading document via proxy...", logging.INFO)
            import requests
            r = requests.get(input_str, allow_redirects=True, proxies=get_proxies(), timeout=30)
            r.raise_for_status()

            tmp_path = Path(tempfile.gettempdir()) / "puml2visio_conv_temp.docx"
            with open(tmp_path, 'wb') as f:
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
            _unblock_file(source)

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
                export_format = 17 if self.target_format == "pdf" else 18
                word.Options.UpdateFieldsAtPrint = False
                word.Options.UpdateLinksAtPrint = False

                doc.ExportAsFixedFormat(
                    OutputFileName=out_path,
                    ExportFormat=export_format,
                    OpenAfterExport=False,
                    OptimizeFor=0,
                    IncludeDocProps=True,
                    CreateBookmarks=1
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