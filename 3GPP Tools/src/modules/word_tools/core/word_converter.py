import logging
from pathlib import Path
import tempfile
from typing import Optional, Union

import pythoncom
import win32com.client
from PyQt5.QtCore import QThread, pyqtSignal

from core.utils.utils import get_proxies


def convert_doc_to_docx(
    doc_path: Union[str, Path],
    output_path: Optional[Union[str, Path]] = None,
    logger: Optional[logging.Logger] = None,
) -> Path:
    """
    Synchronously converts a legacy Word 97-2003 (.doc) binary file to OpenXML (.docx)
    using Microsoft Word COM automation.
    """
    log = logger or logging.getLogger(__name__)
    source = Path(doc_path)

    if not source.exists():
        raise FileNotFoundError(f"Source document not found: {source}")

    if source.suffix.lower() == ".docx":
        return source

    target = Path(output_path) if output_path else source.with_suffix(".docx")
    if target.exists() and target.stat().st_size > 0:
        return target

    word = None
    doc = None
    try:
        pythoncom.CoInitialize()
        # Spawn an isolated, headless Word instance
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0  # wdAlertsNone

        # Signature: Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles)
        doc = word.Documents.Open(str(source.resolve()), False, True, False)

        # wdFormatXMLDocument = 16 (.docx)
        doc.SaveAs2(str(target.resolve()), FileFormat=16)
        log.info(f"Successfully converted {source.name} -> {target.name}")
        return target
    except Exception as e:
        log.error(f"COM conversion failed for {source.name}: {e}")
        raise
    finally:
        try:
            if doc:
                doc.Close(SaveChanges=False)
            if word:
                word.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()


class WordConverterThread(QThread):
    ui_log_msg = pyqtSignal(str, int)
    finished_path = pyqtSignal(str)  # Used by QueueManager for success notifications
    finished = pyqtSignal()

    # Microsoft Word WdSaveFormat Enumerations
    # https://learn.microsoft.com/en-us/office/vba/api/word.wdsaveformat
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

            out_dir = Path(source_path).parent
            out_name = Path(source_path).stem + f".{self.target_format}"
            out_path = str(out_dir / out_name)

            if self.target_format not in self.FORMAT_MAP:
                raise ValueError(f"Unsupported conversion format: {self.target_format}")

            self.ui_log_msg.emit(f"⏳ Spawning detached Word Converter Engine for {out_name}...", logging.INFO)
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0

            doc = word.Documents.Open(source_path, False, True, False)

            self.ui_log_msg.emit(f"⏳ Converting and saving {out_name} to {self.target_format}...", logging.INFO)

            if self.target_format in ("pdf", "xps"):
                export_format = 17 if self.target_format == "pdf" else 18
                word.Options.UpdateFieldsAtPrint = False
                word.Options.UpdateLinksAtPrint = False

                doc.ExportAsFixedFormat(
                    OutputFileName=out_path,
                    ExportFormat=export_format,
                    OpenAfterExport=False,
                    OptimizeFor=0,  # wdExportOptimizeForPrint
                    IncludeDocProps=True,
                    CreateBookmarks=1  # wdExportCreateHeadingBookmarks
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