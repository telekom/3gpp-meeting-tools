import tempfile
import logging
import traceback  # <--- NEW: Crucial for getting exact error lines
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
        # ... (Keep your existing _resolve_path code here) ...
        if not input_str:
            raise ValueError(f"Document {doc_label} input is empty. Please select a valid file, open document, or URL.")

        if input_str.startswith("http://") or input_str.startswith("https://"):
            if "sharepoint.com" in input_str.lower() or "onedrive" in input_str.lower():
                self.ui_log_msg.emit(
                    f"🔗 Corporate link detected for Document {doc_label}. Delegating secure authentication to MS Word...",
                    logging.INFO)
                clean_url = input_str.split("?")[0] if "?web=" in input_str else input_str
                return clean_url

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
        word = None
        original_security = None
        temp_dir = None

        try:
            pythoncom.CoInitialize()

            self.ui_log_msg.emit("⏳ Step 1: Initializing paths...", logging.INFO)
            path_a = self._resolve_path(self.doc_a, "A")
            path_b = self._resolve_path(self.doc_b, "B")

            file_a_name = Path(path_a).name
            file_b_name = Path(path_b).name
            self.ui_log_msg.emit(f"   ➔ Doc A (Base): {file_a_name}", logging.INFO)
            self.ui_log_msg.emit(f"   ➔ Doc B (Rev) : {file_b_name}", logging.INFO)

            if Path(path_a).resolve() == Path(path_b).resolve():
                self.ui_log_msg.emit("⚠️ WARNING: Document A and Document B are the exact same file!", logging.WARNING)

            self.ui_log_msg.emit("⏳ Step 2: Creating local sandbox copies...", logging.INFO)
            import shutil
            import os
            import stat
            temp_dir = Path(tempfile.gettempdir()) / "3gpp_compare_tmp"
            temp_dir.mkdir(parents=True, exist_ok=True)

            temp_path_a = temp_dir / f"A_{file_a_name}"
            temp_path_b = temp_dir / f"B_{file_b_name}"

            # Copy content only, dropping strict Windows metadata
            shutil.copyfile(path_a, temp_path_a)
            shutil.copyfile(path_b, temp_path_b)

            # Explicitly strip OS-level Read-Only locks
            os.chmod(temp_path_a, stat.S_IWRITE)
            os.chmod(temp_path_b, stat.S_IWRITE)

            self.ui_log_msg.emit("⏳ Step 3: Spawning Native Word Diff Engine...", logging.INFO)
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = True

            try:
                original_security = word.AutomationSecurity
                word.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
            except Exception:
                pass

            # Disable alerts so InsertFile doesn't throw hidden "conversion" popups
            word.DisplayAlerts = 0

            self.ui_log_msg.emit("⏳ Step 4: Bypassing corporate locks via Skeleton Key...", logging.INFO)

            def process_sandbox_doc(filepath, label, original_filename):
                self.ui_log_msg.emit(f"   ➔ Extracting Doc {label} into unlocked container...", logging.INFO)

                # 1. Create a pristine, completely unlocked blank document
                doc = word.Documents.Add()

                # 2. Insert the locked file's contents natively.
                doc.Content.InsertFile(FileName=str(filepath))

                # 3. Accept all revisions safely
                if doc.Revisions.Count > 0:
                    doc.Revisions.AcceptAll()

                # ---> NEW: Apply Sensitivity Label BEFORE saving to prevent IT popup blockers!
                from modules.word_tools.core.sensitivity_label import set_sensitivity_label
                set_sensitivity_label(doc, self.ui_log_msg)

                # 4. Save the document using the exact original filename
                save_dir = temp_dir / f"unlocked_{label}"
                save_dir.mkdir(parents=True, exist_ok=True)
                out_path = save_dir / original_filename

                doc.SaveAs(FileName=str(out_path))
                return doc

            # Process both documents safely, passing the original filenames down
            doc_original = process_sandbox_doc(temp_path_a, "A", file_a_name)
            doc_revised = process_sandbox_doc(temp_path_b, "B", file_b_name)

            self.ui_log_msg.emit("⏳ Step 5: Executing CompareDocuments engine...", logging.INFO)

            # Restore alerts for the user
            word.DisplayAlerts = -1

            # Execute comparison using the exact kwargs you verified
            cmp_doc = word.CompareDocuments(
                OriginalDocument=doc_original,
                RevisedDocument=doc_revised,
                Destination=2,
                IgnoreAllComparisonWarnings=True,
                CompareFormatting=True,
                CompareCaseChanges=True,
                CompareWhitespace=True,
                CompareFields=True
            )

            self.ui_log_msg.emit("⏳ Step 6: Closing source documents...", logging.INFO)
            try:
                doc_original.Close(SaveChanges=False)
                doc_revised.Close(SaveChanges=False)
                doc_original = None
                doc_revised = None
            except Exception:
                pass

            if cmp_doc:
                cmp_doc.Activate()
            word.Activate()

            self.ui_log_msg.emit("✅ Comparison generated successfully!", logging.INFO)

        except Exception as e:
            import traceback
            err_trace = traceback.format_exc()
            self.ui_log_msg.emit(f"❌ Comparison Error: {str(e)}\n\nTraceback:\n{err_trace}", logging.ERROR)
            print(f"Detailed Comparison Traceback:\n{err_trace}")

        finally:
            self.ui_log_msg.emit("🧹 Step 7: Cleaning up thread...", logging.INFO)

            if word and original_security is not None:
                try:
                    word.AutomationSecurity = original_security
                except Exception:
                    pass

            try:
                if doc_original: doc_original.Close(SaveChanges=False)
                if doc_revised: doc_revised.Close(SaveChanges=False)
            except Exception:
                pass

            # -> NEW: Wipe the entire temp directory in one shot, destroying the locked copies and the saved unlocked copies
            try:
                if temp_dir and temp_dir.exists():
                    import shutil
                    shutil.rmtree(temp_dir, ignore_errors=True)
            except Exception as e:
                print(f"Cleanup warning: {e}")

            pythoncom.CoUninitialize()
            self.finished.emit()