import tempfile
import logging
from pathlib import Path
import win32com.client
import pythoncom
from PyQt5.QtCore import QThread, pyqtSignal

from core.utils.utils import get_proxies


class WordComparatorThread(QThread):
    ui_log_msg = pyqtSignal(str, int)
    finished = pyqtSignal()

    def __init__(self, doc_a: str, doc_b: str, keep_open: bool = False):
        super().__init__()
        self.doc_a = doc_a
        self.doc_b = doc_b
        self.keep_open = keep_open

    def _resolve_path(self, input_str: str, doc_label: str) -> str:
        """Determines if the input is a local path or a URL. Downloads URLs via Proxy."""
        if not input_str:
            raise ValueError(f"Document {doc_label} input is empty. Please select a valid file, open document, or URL.")

        if input_str.startswith("http://") or input_str.startswith("https://"):

            # --- NEW: SharePoint & Office 365 Authentication Bypass ---
            if "sharepoint.com" in input_str.lower() or "onedrive" in input_str.lower():
                self.ui_log_msg.emit(
                    f"🔗 Corporate link detected for Document {doc_label}. Delegating secure authentication to MS Word...",
                    logging.INFO)

                # Strip web-viewer parameters (like ?web=1) so Word doesn't get confused
                clean_url = input_str.split("?")[0] if "?web=" in input_str else input_str

                # We return the URL directly. word.Documents.Open(clean_url) handles the SSO!
                return clean_url

            # --- Standard Public URL Download ---
            self.ui_log_msg.emit(f"⏳ Downloading Document {doc_label} via proxy...", logging.INFO)
            import requests

            proxies = get_proxies()
            r = requests.get(input_str, allow_redirects=True, proxies=proxies, timeout=30)
            r.raise_for_status()

            tmp_path = Path(tempfile.gettempdir()) / f"puml2visio_cmp_{doc_label}.docx"
            with open(tmp_path, 'wb') as f:
                f.write(r.content)
            return str(tmp_path)

        return input_str
    def run(self):
        doc_original = None
        doc_revised = None
        try:
            pythoncom.CoInitialize()  # Thread-safe COM initialization

            self.ui_log_msg.emit("⏳ Preparing documents for comparison...", logging.INFO)
            path_a = self._resolve_path(self.doc_a, "A")
            path_b = self._resolve_path(self.doc_b, "B")

            self.ui_log_msg.emit("⏳ Spawning Native Word Diff Engine...", logging.INFO)
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = True

            # Open both documents in the background
            doc_original = word.Documents.Open(path_a, ReadOnly=True)
            doc_revised = word.Documents.Open(path_b, ReadOnly=True)

            # Fire the native COM diff engine
            word.CompareDocuments(OriginalDocument=doc_original, RevisedDocument=doc_revised)

            self.ui_log_msg.emit(
                "✅ Comparison generated successfully! Please check the newly opened Microsoft Word window.",
                logging.INFO)

        except Exception as e:
            self.ui_log_msg.emit(f"❌ Comparison Error: {str(e)}", logging.ERROR)
        finally:
            # Clean up conditionally based on user UI preference
            if not self.keep_open:
                try:
                    if doc_original: doc_original.Close(SaveChanges=False)
                    if doc_revised: doc_revised.Close(SaveChanges=False)
                except Exception:
                    pass

            pythoncom.CoUninitialize()
            self.finished.emit()