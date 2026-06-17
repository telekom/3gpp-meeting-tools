# src/modules/word_tools/core/word_converter.py

import os
import tempfile
import logging
from pathlib import Path
import win32com.client
import pythoncom
from PyQt5.QtCore import QThread, pyqtSignal


class WordConverterThread(QThread):
    ui_log_msg = pyqtSignal(str, int)
    finished_path = pyqtSignal(str)  # Used by QueueManager for success notifications
    finished = pyqtSignal()

    # Microsoft Word WdSaveFormat Enumerations
    # https://learn.microsoft.com/en-us/office/vba/api/word.wdsaveformat
    FORMAT_MAP = {
        "pdf": 17,  # wdFormatPDF
        "html": 8,  # wdFormatHTML
        "xps": 18,  # wdFormatXPS
        "rtf": 6,  # wdFormatRTF
        "txt": 2,  # wdFormatText
    }

    def __init__(self, doc_source: str, target_format: str):
        super().__init__()
        self.doc_source = doc_source
        self.target_format = target_format.lower().replace(".", "")

    def _get_proxies(self):
        return {
            "http": os.environ.get("HTTP_PROXY") or os.environ.get("http_proxy"),
            "https": os.environ.get("HTTPS_PROXY") or os.environ.get("https_proxy"),
        }

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
            r = requests.get(input_str, allow_redirects=True, proxies=self._get_proxies(), timeout=30)
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

            # Determine output path (next to the source file if local, or temp if URL)
            out_dir = Path(source_path).parent
            out_name = Path(source_path).stem + f".{self.target_format}"
            out_path = str(out_dir / out_name)

            if self.target_format not in self.FORMAT_MAP:
                raise ValueError(f"Unsupported conversion format: {self.target_format}")

            self.ui_log_msg.emit("⏳ Spawning detached Word Converter Engine...", logging.INFO)
            # DispatchEx forces a completely new, invisible instance of Word
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False

            # 0 is the native Word constant for wdAlertsNone
            word.DisplayAlerts = 0

            # We must pass the arguments positionally to bypass the file-lock popup.
            # Signature: Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles)
            doc = word.Documents.Open(source_path, False, True, False)

            self.ui_log_msg.emit("⏳ Converting and saving...", logging.INFO)
            doc.SaveAs2(out_path, FileFormat=self.FORMAT_MAP[self.target_format])

            self.ui_log_msg.emit(f"✅ Conversion complete: {out_name}", logging.INFO)
            self.finished_path.emit(out_path)

        except Exception as e:
            self.ui_log_msg.emit(f"❌ Conversion Error: {str(e)}", logging.ERROR)
        finally:
            # Rigorous cleanup to prevent ghost processes
            try:
                if doc: doc.Close(SaveChanges=False)
                if word: word.Quit()
            except Exception:
                pass

            pythoncom.CoUninitialize()
            self.finished.emit()