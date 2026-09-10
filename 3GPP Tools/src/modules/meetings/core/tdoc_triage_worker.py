# --- File: src/modules/meetings/core/tdoc_triage_worker.py ---
import logging
import re
from pathlib import Path
from typing import Optional

import pythoncom
import win32com.client
from PyQt5.QtCore import QThread, pyqtSignal

from core.ai.ollama_client import OllamaClient
from core.utils.paths import get_project_root
from modules.meetings.core.llm_exporter import LLM_EXTRACTOR_VERSION
from modules.meetings.core.tdoc_file_handler import TDocFileHandler

PROMPTS_DIR = get_project_root() / "config" / "prompts"


def load_prompt_file(filename: str, fallback_default: str) -> str:
    """Loads a prompt file from config/prompts/ with timestamp hot-reloading."""
    prompt_path = PROMPTS_DIR / filename
    if prompt_path.exists():
        try:
            with open(prompt_path, "r", encoding="utf-8") as f:
                return f.read().strip()
        except Exception as e:
            logging.warning(f"[AI Triage] Failed to read {filename}: {e}")
    return fallback_default.strip()


def prepare_focused_text(raw_text: str, doc_type: str, max_chars: int = 16000) -> str:
    """
    Optimizes 3GPP document text for 7B parameter models (e.g., qwen2.5:7b)
    by prioritizing the CR cover sheet table and modified clauses with track changes.
    """
    if not raw_text or len(raw_text) <= max_chars:
        return raw_text

    is_cr = (
        "CR" in doc_type.upper()
        or "[INSERTED:" in raw_text
        or "[ADDED BLOCK]:" in raw_text
    )

    if is_cr:
        # 1. Preamble & Cover Sheet: The first ~3,500 characters contain the CR table
        # (Reason for change, Summary of change, Consequences if not approved, Clauses affected)
        preamble_limit = min(3500, len(raw_text))
        preamble = raw_text[:preamble_limit]
        remaining = raw_text[preamble_limit:]

        # 2. Filter remaining text for lines containing tracked changes or headings
        paragraphs = remaining.split("\n\n")
        focused_paras = []
        budget = max_chars - len(preamble) - 300

        change_keywords = (
            "[INSERTED:",
            "[DELETED:",
            "[ADDED BLOCK]:",
            "*[<",
            "#",
            "##",
            "###",
        )

        truncated = False
        for p in paragraphs:
            p_strip = p.strip()
            if not p_strip:
                continue

            if any(kw in p_strip for kw in change_keywords):
                if len(p_strip) + 2 > budget:
                    truncated = True
                    break
                focused_paras.append(p_strip)
                budget -= len(p_strip) + 2
            elif len(p_strip) < 120 and (
                p_strip.startswith("Clause ") or p_strip.startswith("Section ")
            ):
                if len(p_strip) + 2 <= budget:
                    focused_paras.append(p_strip)
                    budget -= len(p_strip) + 2

        result = preamble + "\n\n" + "\n\n".join(focused_paras)
        if truncated:
            result += "\n\n> ⚠️ *[Remaining unchanged/modified clauses truncated for 7B context window]*"
        return result
    else:
        # Discussion Paper / LS: Take front ~10,000 chars and end ~5,000 chars
        front = raw_text[:10000]
        tail = raw_text[-5000:]
        return (
            front
            + "\n\n> ⚠️ *[... Middle sections omitted for focused executive triage ...]*\n\n"
            + tail
        )


class TDocTriageWorker(QThread):
    """Background worker that prepares focused TDoc text and streams LLM triage output."""

    progress_status = pyqtSignal(str)
    token_received = pyqtSignal(str)
    generation_completed = pyqtSignal(str)
    generation_failed = pyqtSignal(str)

    def __init__(
        self,
        client: OllamaClient,
        model: str,
        tdoc_data: dict,
        meeting_dir: Path,
        docs_ftp_url: str = "",
        revisions_url: str = "",
        parent=None,
    ):
        super().__init__(parent)
        self.client = client
        self.model = model
        self.tdoc_data = tdoc_data
        self.meeting_dir = Path(meeting_dir)
        self.docs_ftp_url = docs_ftp_url
        self.revisions_url = revisions_url
        self._is_cancelled = False

    def cancel(self):
        """Flag the thread to abort processing."""
        self._is_cancelled = True

    def run(self):
        word_app = None
        try:
            tdoc_id = str(self.tdoc_data.get("TDoc", "")).strip()
            if not tdoc_id:
                self.generation_failed.emit("Invalid or empty TDoc identifier.")
                return

            base_match = re.search(
                r"^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$", tdoc_id, re.IGNORECASE
            )
            base_tdoc = base_match.group(1).upper() if base_match else tdoc_id.upper()
            doc_type = str(self.tdoc_data.get("Type", "CR"))

            tdoc_folder = self.meeting_dir / base_tdoc
            cache_file = tdoc_folder / f"{tdoc_id}_LLM_v{LLM_EXTRACTOR_VERSION}.md"

            raw_markdown = ""

            # 1. Check existing LLM markdown cache
            if cache_file.exists():
                self.progress_status.emit("📖 Reading cached document...")
                with open(cache_file, "r", encoding="utf-8") as f:
                    raw_markdown = f.read()
            else:
                # 2. Find local Word file or download and extract
                self.progress_status.emit(f"🔍 Locating Word document for {tdoc_id}...")
                doc_path = self._find_word_doc(tdoc_folder, tdoc_id)

                if not doc_path:
                    # Also check meeting TDocs directory
                    alt_folder = self.meeting_dir / "TDocs" / base_tdoc
                    doc_path = self._find_word_doc(alt_folder, tdoc_id)

                if not doc_path:
                    self.progress_status.emit(f"⬇️ Downloading {tdoc_id}...")
                    doc_path = self._download_and_extract_tdoc(tdoc_id, tdoc_folder)

                if self._is_cancelled:
                    self.generation_failed.emit("Triage cancelled by user.")
                    return

                if not doc_path or not doc_path.exists():
                    self.generation_failed.emit(
                        f"Could not locate or extract a Word document for {tdoc_id}."
                    )
                    return

                # 3. Extract text from Word document using COM
                self.progress_status.emit("📄 Extracting document text...")
                pythoncom.CoInitialize()
                word_app = win32com.client.DispatchEx("Word.Application")
                word_app.Visible = False
                word_app.DisplayAlerts = 0

                raw_markdown = self._extract_from_word(word_app, doc_path, doc_type)

                if raw_markdown:
                    tdoc_folder.mkdir(parents=True, exist_ok=True)
                    with open(cache_file, "w", encoding="utf-8") as f:
                        f.write(raw_markdown)

            if self._is_cancelled:
                self.generation_failed.emit("Triage cancelled by user.")
                return

            # 4. Prepare focused text context
            self.progress_status.emit("⚡ Optimizing context for 7B model...")
            focused_content = prepare_focused_text(raw_markdown, doc_type, max_chars=16000)

            # 5. Load prompts and format payload
            sys_prompt = load_prompt_file(
                "triage_system.txt",
                "You are an expert 3GPP standards delegate. Provide a concise technical triage in 3 parts: "
                "1. Objective & Problem Statement, 2. Key Technical Changes, 3. Potential Impacts & Contentious Points.",
            )
            usr_template = load_prompt_file(
                "triage_user.txt",
                "TDoc: {tdoc_id}\nTitle: {title}\nSource: {source}\nAgenda Item: {agenda_item}\nType: {doc_type}\n\nContent:\n{content}",
            )

            formatted_user = usr_template.format(
                tdoc_id=tdoc_id,
                title=self.tdoc_data.get("Title", "No Title"),
                source=self.tdoc_data.get("Source", "Unknown"),
                agenda_item=self.tdoc_data.get("Agenda Item", "N/A"),
                doc_type=doc_type,
                content=focused_content,
            )

            messages = [
                {"role": "system", "content": sys_prompt},
                {"role": "user", "content": formatted_user},
            ]

            # 6. Stream tokens from Ollama
            self.progress_status.emit("🦙 Streaming analysis tokens...")
            accumulated = []
            stream_options = {"temperature": 0.1, "num_ctx": 8192}

            for delta in self.client.stream_chat(
                model=self.model, messages=messages, options=stream_options
            ):
                if self._is_cancelled:
                    self.generation_failed.emit("Triage cancelled by user.")
                    return
                accumulated.append(delta)
                self.token_received.emit(delta)

            full_triage = "".join(accumulated).strip()
            self.generation_completed.emit(full_triage)

        except Exception as e:
            if not self._is_cancelled:
                logging.error(f"[AI Triage] Error during triage: {e}", exc_info=True)
                self.generation_failed.emit(str(e))
        finally:
            if word_app:
                try:
                    word_app.DisplayAlerts = 0
                    word_app.Quit(SaveChanges=0)
                except Exception:
                    pass
                del word_app
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    def _find_word_doc(self, folder: Path, tdoc_id: str) -> Optional[Path]:
        if not folder or not folder.exists():
            return None
        for ext in [".docx", ".doc"]:
            files = list(folder.glob(f"*{ext}"))
            if files:
                for f in files:
                    if tdoc_id.lower() in f.name.lower():
                        return f
                return files[0]
        return None

    def _download_and_extract_tdoc(self, tdoc_id: str, target_folder: Path) -> Optional[Path]:
        base_match = re.search(
            r"^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$", tdoc_id, re.IGNORECASE
        )
        is_revision = bool(base_match)
        base_url = (
            self.revisions_url
            if (is_revision and self.revisions_url)
            else self.docs_ftp_url
        )

        if not base_url:
            return None

        try:
            TDocFileHandler.download_and_extract_tdoc(tdoc_id, base_url, target_folder)
            return self._find_word_doc(target_folder, tdoc_id)
        except Exception as e:
            logging.error(f"[AI Triage] Download failed for {tdoc_id}: {e}")
            return None

    def _extract_from_word(self, word_app, doc_path: Path, doc_type: str) -> str:
        doc = None
        try:
            doc = word_app.Documents.Open(str(doc_path), False, True, False)
            md_lines = []
            in_new_block = False

            all_new_trigger = re.compile(r"(?i)all (?:new )?text (?:is )?(?:new|added)")
            placeholder_clause_trigger = re.compile(r"^(\d+\.)+[a-zA-Z]$")
            numeric_clause_trigger = re.compile(r"^(\d+\.)+\d+$")
            boundary_trigger = re.compile(r"(?i)<[-\s]*(next|end of)\s*change")

            for para in doc.Paragraphs:
                text = para.Range.Text.strip("\r\x07\x0b ")
                if not text:
                    continue

                style_name = ""
                try:
                    style_name = para.Style.NameLocal
                except Exception:
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
                                    inserted.append(rev.Range.Text.strip("\r\x07\x0b "))
                                elif rev.Type == 2:
                                    deleted.append(rev.Range.Text.strip("\r\x07\x0b "))

                            prefix = ""
                            if inserted:
                                prefix += f"[INSERTED: {', '.join(inserted)}] "
                            if deleted:
                                prefix += f"[DELETED: {', '.join(deleted)}] "
                            md_lines.append(f"{prefix}{text}")
                        else:
                            md_lines.append(text)
                    else:
                        if "Heading" in style_name:
                            depth = "".join(filter(str.isdigit, style_name))
                            prefix = "#" * int(depth) if depth else "##"
                            md_lines.append(f"\n{prefix} {text}\n")
                        else:
                            md_lines.append(text)

            return "\n\n".join(md_lines)
        except Exception as e:
            return f"Error extracting Word content: {e}"
        finally:
            if doc:
                try:
                    doc.Close(SaveChanges=False)
                except Exception:
                    pass