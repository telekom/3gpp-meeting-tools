# --- File: src/modules/meetings/core/llm_exporter.py ---
import os
import re
import pythoncom
import win32com.client
import logging
from pathlib import Path
from PyQt5.QtCore import QThread, pyqtSignal

from modules.meetings.core.tdoc_file_handler import TDocFileHandler

LLM_EXTRACTOR_VERSION = "1.0.0"


class LLMExporterThread(QThread):
    finished = pyqtSignal(bool, str)
    progress = pyqtSignal(str)

    def __init__(self, meeting_dir: Path, tdocs_list: list, docs_ftp_url: str, revisions_url: str,
                 is_bulk: bool = True, max_chars: int = 200000):
        super().__init__()
        self.meeting_dir = meeting_dir
        self.tdocs_list = tdocs_list
        self.docs_ftp_url = docs_ftp_url
        self.revisions_url = revisions_url
        self.is_bulk = is_bulk
        self.max_chars = max_chars
        self.export_dir = self.meeting_dir / "Export" / "LLM_Corpus"

    def run(self):
        word_app = None
        try:
            pythoncom.CoInitialize()
            self.export_dir.mkdir(parents=True, exist_ok=True)

            corpus = {}
            saved_files = []

            for tdoc_data in self.tdocs_list:
                tdoc_id = str(tdoc_data.get("TDoc", "")).strip()
                if not tdoc_id: continue

                base_match = re.search(r'^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$', tdoc_id, re.IGNORECASE)
                base_tdoc = base_match.group(1).upper() if base_match else tdoc_id.upper()

                ai = str(tdoc_data.get("Agenda Item", "Unknown")).replace(" ", "_")
                doc_type = str(tdoc_data.get("Type", "Other"))

                if "CR" in doc_type or "pCR" in doc_type:
                    category = "CRs"
                elif "LS" in doc_type:
                    category = "LSs"
                else:
                    category = "Discussion_Papers"

                msg = f"Processing {tdoc_id}..."
                self.progress.emit(msg)
                logging.info(f"[LLM Exporter] {msg}")

                tdoc_folder = self.meeting_dir / base_tdoc
                cache_file = tdoc_folder / f"{tdoc_id}_LLM_v{LLM_EXTRACTOR_VERSION}.md"

                md_content = ""
                if cache_file.exists():
                    logging.info(f"[LLM Exporter] Found local cache for {tdoc_id}")
                    with open(cache_file, "r", encoding="utf-8") as f:
                        md_content = f.read()
                else:
                    doc_path = self._find_word_doc(tdoc_folder, tdoc_id)
                    if not doc_path:
                        dl_msg = f"Downloading missing TDoc: {tdoc_id}..."
                        self.progress.emit(dl_msg)
                        logging.info(f"[LLM Exporter] {dl_msg}")
                        doc_path = self._download_and_extract_tdoc(tdoc_id, tdoc_folder)

                    if doc_path:
                        if not word_app:
                            word_app = win32com.client.DispatchEx("Word.Application")
                            word_app.Visible = False
                            word_app.DisplayAlerts = 0

                        ext_msg = f"Extracting {tdoc_id}..."
                        self.progress.emit(ext_msg)
                        logging.info(f"[LLM Exporter] {ext_msg}")
                        md_content = self._extract_from_word(word_app, doc_path, doc_type)

                        if md_content:
                            with open(cache_file, "w", encoding="utf-8") as f:
                                f.write(md_content)
                    else:
                        warn_msg = f"> ⚠️ Could not locate or download an unzipped Word document for {tdoc_id}.\n"
                        self.progress.emit(f"Failed to find Word doc for {tdoc_id}")
                        logging.warning(f"[LLM Exporter] {warn_msg}")
                        md_content = warn_msg

                if ai not in corpus: corpus[ai] = {}
                if category not in corpus[ai]: corpus[ai][category] = []

                header = f"# TDoc: {tdoc_id}\n**Title:** {tdoc_data.get('Title', '')}\n**Source:** {tdoc_data.get('Source', '')}\n\n"

                # ---> THE FIX: Store as a tuple (text_content, tdoc_id) to accurately report skipped files later
                full_tdoc_text = header + md_content + "\n\n---\n"
                corpus[ai][category].append((full_tdoc_text, tdoc_id))

            # --- SMART CHUNKING LOGIC ---
            if self.is_bulk:
                for ai, categories in corpus.items():
                    for category, contents in categories.items():

                        context_header = (
                            f"# 3GPP LLM Corpus\n"
                            f"**Agenda Item:** {ai}\n"
                            f"**Document Category:** {category}\n\n"
                            f"## Context Guide for LLM\n"
                            f"This file contains a programmatic compilation of 3GPP Technical Documents (TDocs). "
                            f"These documents represent telecommunications standards proposals, revisions, and working group agreements.\n\n"
                            f"**Structural Rules for parsing this text:**\n"
                            f"- `[ADDED BLOCK]:` Denotes entirely new text inserted into the specification where tracking wasn't explicitly isolated.\n"
                            f"- `[INSERTED: <text>]`: Denotes specific inline text additions explicitly marked via Word Track Changes.\n"
                            f"- `[DELETED: <text>]`: Denotes specific inline text removals explicitly marked via Word Track Changes.\n\n"
                            f"**Your Task:** Please use this corpus to analyze technical agreements, architectural changes, or contradictions within this specific Agenda Item.\n\n"
                            f"---\n\n"
                        )

                        chunk_idx = 1
                        current_text = context_header
                        has_content = False

                        for tdoc_text, tdoc_id in contents:
                            # 1. Skip documents that individually exceed the limit
                            if len(tdoc_text) > self.max_chars:
                                warn_msg = f"⚠️ Skipped {tdoc_id}: Size ({len(tdoc_text)} chars) exceeds configured limit of {self.max_chars}."
                                logging.warning(f"[LLM Exporter] {warn_msg}")
                                self.progress.emit(warn_msg)
                                continue

                            # 2. If appending this document pushes us over the limit, save the chunk and start a new one
                            if len(current_text) + len(tdoc_text) > self.max_chars and has_content:
                                suffix = f"_Part{chunk_idx}"
                                mega_file = self.export_dir / f"AI_{ai}_Agreed_{category}{suffix}.md"
                                with open(mega_file, "w", encoding="utf-8") as f:
                                    f.write(current_text)
                                saved_files.append(mega_file.name)

                                # Reset variables for the next chunk
                                chunk_idx += 1
                                current_text = context_header + tdoc_text
                                has_content = True
                            else:
                                # Safe to append
                                current_text += tdoc_text
                                has_content = True

                        # 3. Write whatever is remaining in the buffer
                        if has_content:
                            suffix = f"_Part{chunk_idx}" if chunk_idx > 1 else ""
                            mega_file = self.export_dir / f"AI_{ai}_Agreed_{category}{suffix}.md"
                            with open(mega_file, "w", encoding="utf-8") as f:
                                f.write(current_text)
                            saved_files.append(mega_file.name)

                self.finished.emit(True, f"Generated {len(saved_files)} Mega-Files (Chunked) in:\n{self.export_dir}")
            else:
                self.finished.emit(True, f"Exported single TDoc to local cache:\n{cache_file}")

        except Exception as e:
            logging.error(f"[LLM Exporter] Critical thread failure: {e}", exc_info=True)
            self.finished.emit(False, str(e))
        finally:
            if word_app:
                try:
                    word_app.Quit()
                except:
                    pass
            pythoncom.CoUninitialize()

    def _find_word_doc(self, folder: Path, tdoc_id: str):
        if not folder.exists(): return None
        for ext in [".docx", ".doc"]:
            files = list(folder.glob(f"*{ext}"))
            if files:
                for f in files:
                    if tdoc_id.lower() in f.name.lower(): return f
                return files[0]
        return None

    def _download_and_extract_tdoc(self, tdoc_id: str, tdoc_folder: Path):
        base_match = re.search(r'^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$', tdoc_id, re.IGNORECASE)
        is_revision = bool(base_match)

        base_url = self.revisions_url if (is_revision and self.revisions_url) else self.docs_ftp_url
        if not base_url:
            logging.error(f"[LLM Exporter] No base URL configured for downloading {tdoc_id}")
            return None

        try:
            logging.info(f"[LLM Exporter] Triggering TDocFileHandler for {tdoc_id} against {base_url}")
            TDocFileHandler.download_and_extract_tdoc(tdoc_id, base_url, tdoc_folder)
            return self._find_word_doc(tdoc_folder, tdoc_id)
        except Exception as e:
            err_msg = f"Failed to download {tdoc_id}: {str(e)}"
            self.progress.emit(err_msg)
            logging.error(f"[LLM Exporter] {err_msg}", exc_info=True)
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
                text = para.Range.Text.strip('\r\x07\x0b ')
                if not text: continue

                style_name = ""
                try:
                    style_name = para.Style.NameLocal
                except:
                    pass

                if boundary_trigger.search(text):
                    in_new_block = False
                    md_lines.append(f"\n*[{text.strip()}]*\n")
                    continue

                if in_new_block and "Heading" in style_name and numeric_clause_trigger.match(text.split()[0]):
                    in_new_block = False

                if all_new_trigger.search(text):
                    in_new_block = True
                    md_lines.append(f"\n> **Note to LLM:** Entering 'All Text New' block.\n")
                    continue

                if "Heading" in style_name and placeholder_clause_trigger.match(text.split()[0]):
                    in_new_block = True
                    md_lines.append(f"\n> **Note to LLM:** Entering placeholder clause '{text.split()[0]}'.\n")

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
                            if inserted: prefix += f"[INSERTED: {', '.join(inserted)}] "
                            if deleted: prefix += f"[DELETED: {', '.join(deleted)}] "
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
                except:
                    pass