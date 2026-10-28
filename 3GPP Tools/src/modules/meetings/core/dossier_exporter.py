# --- File: src/modules/meetings/core/dossier_exporter.py ---
"""
dossier_exporter.py - Unified Navigable Dossier Exporter (Interactive HTML + Bookmarked PDF + AI Prompt)

Modern implementation of F-04:
Takes a set of 3GPP Technical Documents (selected, visible, or post-meeting scope)
and compiles them into:
1. dossier.html: A self-contained, responsive, interactive HTML dossier with sticky/floating
   navigation, master index, search, syntax-highlighted revision marks, and embedded images.
2. dossier.pdf: A vector-quality PDF with native hierarchical PDF outline bookmarks
   (Agenda Item > TDoc ID > Sections) and internal hyperlinks.
3. dossier_prompt.txt: A companion AI citation map with exact 1-indexed page spans and
   tailored evaluation prompts for NotebookLM, Claude 3.5 Sonnet, Gemini, and Ollama.
"""

from __future__ import annotations

import base64
from concurrent.futures import ThreadPoolExecutor, as_completed
import datetime
import html
import logging
import os
import re
import shutil
from pathlib import Path
from typing import Callable, Dict, List, Optional, Set, Tuple, Any

import markdown
import pypdf
from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtGui import QTextDocument
from PyQt5.QtPrintSupport import QPrinter

from core.utils.paths import to_extended_path
from modules.meetings.core.tdoc_file_handler import TDocFileHandler
from modules.meetings.core.tdoc_url_resolver import resolve_tdoc_download_url
from modules.meetings.core.tdoc_markdown_converter import (
    convert_document_to_markdown,
    sanitize_filename,
    is_conversion_stale,
    resolve_revision_mode_for_type,
    CONVERTER_VERSION,
    ExportedDoc,
)

logger = logging.getLogger(__name__)


# -------------------------------------------------------------------------
# Modern CSS Stylesheet for Interactive HTML and Vector PDF Print
# -------------------------------------------------------------------------
DOSSIER_CSS = """
:root {
  --primary: #005A9E;
  --primary-hover: #004578;
  --primary-light: #EFF6FF;
  --bg-page: #F8FAFC;
  --bg-card: #FFFFFF;
  --text-main: #1E293B;
  --text-muted: #64748B;
  --border-color: #E2E8F0;
  --border-subtle: #CBD5E1;
  --diff-ins-bg: #DCFCE7;
  --diff-ins-text: #166534;
  --diff-del-bg: #FEE2E2;
  --diff-del-text: #991B1B;
}

* {
  box-sizing: border-box;
}

body {
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
  font-size: 14px;
  line-height: 1.6;
  color: var(--text-main);
  background-color: var(--bg-page);
  margin: 0;
  padding: 0;
}

/* --- Container & Layout --- */
.dossier-wrapper {
  max-width: 1100px;
  margin: 0 auto;
  padding: 24px 20px 80px 20px;
}

/* --- Floating Action Navigation Bar --- */
.floating-nav {
  position: sticky;
  top: 0;
  z-index: 1000;
  background: rgba(255, 255, 255, 0.95);
  backdrop-filter: blur(8px);
  border-bottom: 1px solid var(--border-color);
  padding: 10px 20px;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  box-shadow: 0 2px 6px rgba(0,0,0,0.04);
}

.floating-nav .title-pill {
  font-weight: 700;
  color: var(--primary);
  font-size: 13px;
  display: flex;
  align-items: center;
  gap: 6px;
}

.floating-nav .nav-controls {
  display: flex;
  align-items: center;
  gap: 8px;
}

.floating-nav select {
  font-size: 12px;
  padding: 5px 10px;
  border: 1px solid var(--border-subtle);
  border-radius: 4px;
  background-color: #fff;
  color: var(--text-main);
  outline: none;
  max-width: 320px;
}

.floating-nav button, .floating-nav a.btn {
  font-size: 12px;
  font-weight: 600;
  padding: 5px 12px;
  border-radius: 4px;
  border: 1px solid var(--border-subtle);
  background-color: #fff;
  color: var(--text-main);
  text-decoration: none;
  cursor: pointer;
  display: inline-flex;
  align-items: center;
  gap: 4px;
  transition: all 0.15s ease-in-out;
}

.floating-nav button:hover, .floating-nav a.btn:hover {
  background-color: var(--primary-light);
  border-color: var(--primary);
  color: var(--primary);
}

.floating-nav button.primary {
  background-color: var(--primary);
  color: #fff;
  border-color: var(--primary);
}

.floating-nav button.primary:hover {
  background-color: var(--primary-hover);
}

/* --- Header & Summary Cover Card --- */
.cover-card {
  background: var(--bg-card);
  border: 1px solid var(--border-color);
  border-radius: 8px;
  padding: 28px 32px;
  margin-bottom: 28px;
  box-shadow: 0 1px 3px rgba(0,0,0,0.03);
}

.cover-card h1 {
  margin: 0 0 8px 0;
  color: var(--primary);
  font-size: 24px;
  font-weight: 700;
}

.cover-card .meta-subtitle {
  color: var(--text-muted);
  font-size: 13px;
  margin-bottom: 20px;
}

.stats-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(140px, 1fr));
  gap: 12px;
  margin: 16px 0;
}

.stat-chip {
  background: var(--primary-light);
  border: 1px solid #BFDBFE;
  border-radius: 6px;
  padding: 10px 14px;
  text-align: center;
}

.stat-chip .val {
  font-size: 20px;
  font-weight: 700;
  color: var(--primary);
}

.stat-chip .lbl {
  font-size: 11px;
  text-transform: uppercase;
  color: var(--text-muted);
  font-weight: 600;
}

/* --- Master Index Table --- */
.index-card {
  background: var(--bg-card);
  border: 1px solid var(--border-color);
  border-radius: 8px;
  padding: 24px;
  margin-bottom: 32px;
  box-shadow: 0 1px 3px rgba(0,0,0,0.03);
}

.index-card h2 {
  margin-top: 0;
  color: var(--text-main);
  font-size: 18px;
  border-bottom: 2px solid var(--primary-light);
  padding-bottom: 8px;
}

table.index-table {
  width: 100%;
  border-collapse: collapse;
  margin-top: 14px;
  font-size: 12px;
}

table.index-table th {
  background: #F1F5F9;
  color: #334155;
  text-align: left;
  padding: 8px 10px;
  border: 1px solid var(--border-color);
  font-weight: 600;
}

table.index-table td {
  padding: 8px 10px;
  border: 1px solid var(--border-color);
  vertical-align: top;
}

table.index-table tr:nth-child(even) {
  background-color: #F8FAFC;
}

table.index-table tr:hover {
  background-color: var(--primary-light);
}

.badge {
  display: inline-block;
  padding: 2px 6px;
  border-radius: 4px;
  font-size: 11px;
  font-weight: 600;
  text-align: center;
}

.badge-cr { background: #FEF3C7; color: #92400E; }
.badge-pcr { background: #E0E7FF; color: #3730A3; }
.badge-ls { background: #E0F2FE; color: #075985; }
.badge-disc { background: #F1F5F9; color: #475569; }
.badge-agreed { background: #DCFCE7; color: #166534; }
.badge-noted { background: #F3F4F6; color: #374151; }

/* --- TDoc Contribution Sections --- */
.tdoc-section {
  background: var(--bg-card);
  border: 1px solid var(--border-color);
  border-radius: 8px;
  padding: 28px 32px;
  margin-bottom: 36px;
  box-shadow: 0 1px 3px rgba(0,0,0,0.03);
}

.tdoc-header {
  border-bottom: 2px solid var(--border-color);
  padding-bottom: 14px;
  margin-bottom: 20px;
}

.tdoc-header-top {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  margin-bottom: 8px;
}

.tdoc-id-title {
  font-size: 20px;
  font-weight: 700;
  color: var(--primary);
  margin: 0;
}

.tdoc-actions {
  display: flex;
  gap: 8px;
}

.tdoc-actions a {
  font-size: 11px;
  font-weight: 600;
  padding: 4px 8px;
  border-radius: 4px;
  border: 1px solid var(--border-subtle);
  background: #F8FAFC;
  color: var(--text-main);
  text-decoration: none;
}

.tdoc-actions a:hover {
  background: var(--primary-light);
  color: var(--primary);
  border-color: var(--primary);
}

.tdoc-meta-row {
  display: flex;
  flex-wrap: wrap;
  gap: 16px;
  font-size: 12px;
  color: var(--text-muted);
  margin-top: 6px;
}

.tdoc-meta-row strong {
  color: var(--text-main);
}

/* --- Document Body & CommonMark Rendering --- */
.tdoc-body {
  font-size: 13.5px;
  line-height: 1.6;
}

.tdoc-body h1, .tdoc-body h2, .tdoc-body h3, .tdoc-body h4 {
  color: #0F172A;
  margin-top: 24px;
  margin-bottom: 10px;
  font-weight: 600;
}

.tdoc-body h1 { font-size: 18px; border-bottom: 1px solid var(--border-color); padding-bottom: 4px; }
.tdoc-body h2 { font-size: 16px; }
.tdoc-body h3 { font-size: 14.5px; }

.tdoc-body table {
  width: 100%;
  border-collapse: collapse;
  margin: 16px 0;
  font-size: 12px;
}

.tdoc-body th, .tdoc-body td {
  border: 1px solid var(--border-color);
  padding: 6px 10px;
  vertical-align: top;
}

.tdoc-body th {
  background-color: #F8FAFC;
  font-weight: 600;
  color: #334155;
}

.tdoc-body img {
  max-width: 100%;
  height: auto;
  border: 1px solid var(--border-color);
  border-radius: 4px;
  margin: 14px 0;
  display: block;
}

.tdoc-body pre, .tdoc-body code {
  font-family: Consolas, Monaco, "Courier New", monospace;
  font-size: 12px;
  background-color: #F1F5F9;
  border-radius: 4px;
}

.tdoc-body code {
  padding: 2px 4px;
}

.tdoc-body pre {
  padding: 12px;
  overflow-x: auto;
  border: 1px solid var(--border-color);
}

/* --- Revision Marks & Diff Highlighting --- */
.diff-ins, ins {
  background-color: var(--diff-ins-bg);
  color: var(--diff-ins-text);
  text-decoration: underline;
  padding: 1px 3px;
  border-radius: 2px;
}

.diff-del, del {
  background-color: var(--diff-del-bg);
  color: var(--diff-del-text);
  text-decoration: line-through;
  padding: 1px 3px;
  border-radius: 2px;
}

.change-marker {
  display: block;
  background-color: var(--primary-light);
  color: var(--primary);
  border-left: 4px solid var(--primary);
  font-weight: 700;
  font-size: 12px;
  padding: 6px 12px;
  margin: 16px 0;
  border-radius: 0 4px 4px 0;
}

/* --- Print Styling for Vector PDF Export --- */
@media print {
  body {
    background-color: #FFFFFF !important;
    font-size: 10pt !important;
    color: #000000 !important;
  }
  .dossier-wrapper {
    max-width: 100% !important;
    padding: 0 !important;
  }
  .no-print, .floating-nav, .tdoc-actions {
    display: none !important;
  }
  .cover-card, .index-card, .tdoc-section {
    border: none !important;
    box-shadow: none !important;
    padding: 0 !important;
    margin: 0 !important;
  }
  .tdoc-section {
    page-break-before: always !important;
  }
  .tdoc-section:first-of-type {
    page-break-before: avoid !important;
  }
  table.index-table th {
    background-color: #EEEEEE !important;
    color: #000000 !important;
  }
  .change-marker {
    background-color: #F0F4F8 !important;
    border-left: 3px solid #005A9E !important;
    color: #005A9E !important;
  }
  table {
    page-break-inside: auto;
  }
  tr {
    page-break-inside: avoid;
    page-break-after: auto;
  }
  @page {
    margin: 15mm 15mm 20mm 15mm;
  }
}
"""


class DossierExporterThread(QThread):
    """
    Asynchronous worker to compile selected or post-meeting 3GPP Technical Documents
    into a unified Navigable HTML Dossier, Bookmarked Vector PDF, and AI Citation Prompt.
    """
    progress = pyqtSignal(str)
    finished = pyqtSignal(bool, str, dict)  # success, message, file_paths

    def __init__(
        self,
        meeting_dir: Path,
        tdocs_list: List[dict],
        docs_ftp_url: str,
        revisions_url: str,
        mtg_info: dict,
        export_format: str = "both",  # "both", "html", "pdf"
        include_prompt: bool = True,
        embed_base64_images: bool = True,
        download_missing: bool = True,
        force_reconvert: bool = False,
        output_dir: Optional[Path] = None,
        parent=None,
    ):
        super().__init__(parent)
        self.meeting_dir = Path(meeting_dir)
        self.tdocs_list = tdocs_list or []
        self.docs_ftp_url = docs_ftp_url or ""
        self.revisions_url = revisions_url or ""
        self.mtg_info = mtg_info or {}
        self.export_format = export_format.lower()
        self.include_prompt = include_prompt
        self.embed_base64_images = embed_base64_images
        self.download_missing = download_missing
        self.force_reconvert = force_reconvert

        # Resolve output directory
        if output_dir:
            self.export_dir = Path(output_dir)
        else:
            timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S")
            wg = str(self.mtg_info.get("wg_name", "3GPP")).strip().replace(" ", "_")
            mtg = str(self.mtg_info.get("meeting_number", "")).strip()
            folder_name = f"Dossier_{wg}_{mtg}_{timestamp}" if (wg or mtg) else f"Dossier_{timestamp}"
            self.export_dir = self.meeting_dir / "Export" / folder_name

    def run(self):
        try:
            if not self.tdocs_list:
                self.finished.emit(False, "No TDocs selected for dossier generation.", {})
                return

            os.makedirs(to_extended_path(self.export_dir), exist_ok=True)
            media_export_dir = self.export_dir / "media"
            os.makedirs(to_extended_path(media_export_dir), exist_ok=True)

            total_docs = len(self.tdocs_list)
            self.progress.emit(f"Preparing {total_docs} document(s) for dossier...")

            # --- Phase 1: Parallel Pre-download of Missing Documents ---
            if self.download_missing:
                missing_download_tasks: List[Tuple[str, Path]] = []
                seen_tdocs = set()

                for tdoc_data in self.tdocs_list:
                    tdoc_id = str(tdoc_data.get("TDoc", "")).strip()
                    if not tdoc_id or tdoc_id in seen_tdocs:
                        continue
                    seen_tdocs.add(tdoc_id)

                    base_match = re.search(r'^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$', tdoc_id, re.IGNORECASE)
                    base_tdoc = base_match.group(1).upper() if base_match else tdoc_id.upper()
                    doc_type = str(tdoc_data.get("Type", "Other"))
                    expected_mode = resolve_revision_mode_for_type(doc_type, default_mode="markup")

                    tdoc_folder = self.meeting_dir / base_tdoc
                    local_md_file = tdoc_folder / f"{sanitize_filename(tdoc_id)}.md"
                    doc_path = self._find_word_doc(tdoc_folder, tdoc_id)

                    needs_conversion = self.force_reconvert or is_conversion_stale(
                        doc_path, local_md_file, expected_version=CONVERTER_VERSION, expected_mode=expected_mode
                    )
                    if needs_conversion and not doc_path:
                        missing_download_tasks.append((tdoc_id, tdoc_folder))

                if missing_download_tasks:
                    total_missing = len(missing_download_tasks)
                    dl_workers = min(6, total_missing)
                    self.progress.emit(f"Downloading {total_missing} missing TDoc(s) in parallel...")
                    logger.info(f"[DossierExporter] Pre-downloading {total_missing} TDocs using {dl_workers} workers...")
                    dl_done = 0
                    with ThreadPoolExecutor(max_workers=dl_workers) as executor:
                        future_to_tid = {
                            executor.submit(self._download_and_extract_tdoc, tid, folder): tid
                            for (tid, folder) in missing_download_tasks
                        }
                        for future in as_completed(future_to_tid):
                            if self.isInterruptionRequested():
                                try:
                                    executor.shutdown(wait=False, cancel_futures=True)
                                except TypeError:
                                    executor.shutdown(wait=False)
                                self.finished.emit(False, "Export cancelled by user.", {})
                                return
                            tid = future_to_tid[future]
                            dl_done += 1
                            try:
                                res = future.result()
                                self.progress.emit(f"Downloaded TDoc {tid} ({dl_done}/{total_missing})")
                            except Exception as dl_err:
                                logger.error(f"[DossierExporter] Error downloading {tid}: {dl_err}")

            if self.isInterruptionRequested():
                self.finished.emit(False, "Export cancelled by user.", {})
                return

            # --- Phase 2: Convert / Load Markdown for Each Document ---
            converted_docs: List[Dict[str, Any]] = []
            for idx, tdoc_data in enumerate(self.tdocs_list, 1):
                if self.isInterruptionRequested():
                    self.finished.emit(False, "Export cancelled by user.", {})
                    return

                tdoc_id = str(tdoc_data.get("TDoc", "")).strip()
                if not tdoc_id:
                    continue

                base_match = re.search(r'^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$', tdoc_id, re.IGNORECASE)
                base_tdoc = base_match.group(1).upper() if base_match else tdoc_id.upper()
                doc_type = str(tdoc_data.get("Type", "Other"))
                raw_ai = str(tdoc_data.get("Agenda Item", "Unassigned")).strip()
                title = str(tdoc_data.get("Title", "")).strip()
                source = str(tdoc_data.get("Source", "")).strip()
                status = str(tdoc_data.get("TDoc Status", "")).strip()
                expected_mode = resolve_revision_mode_for_type(doc_type, default_mode="markup")

                self.progress.emit(f"Processing Markdown for {tdoc_id} ({idx}/{total_docs})...")

                tdoc_folder = self.meeting_dir / base_tdoc
                local_md_file = tdoc_folder / f"{sanitize_filename(tdoc_id)}.md"
                local_media_dir = tdoc_folder / "media"
                doc_path = self._find_word_doc(tdoc_folder, tdoc_id)

                needs_conversion = self.force_reconvert or is_conversion_stale(
                    doc_path, local_md_file, expected_version=CONVERTER_VERSION, expected_mode=expected_mode
                )

                if needs_conversion and doc_path:
                    try:
                        os.makedirs(to_extended_path(tdoc_folder), exist_ok=True)
                        os.makedirs(to_extended_path(local_media_dir), exist_ok=True)
                        convert_document_to_markdown(
                            doc_path,
                            out_dir=tdoc_folder,
                            tdoc_id=tdoc_id,
                            meta=tdoc_data,
                            revisions=expected_mode,
                            media_dir=local_media_dir,
                            media_rel="media",
                            cancellation_check=self.isInterruptionRequested,
                        )
                    except Exception as conv_err:
                        logger.warning(f"[DossierExporter] Direct conversion failed for {tdoc_id}: {conv_err}")

                md_text = ""
                if os.path.exists(to_extended_path(local_md_file)):
                    try:
                        with open(to_extended_path(local_md_file), "r", encoding="utf-8", errors="ignore") as f:
                            md_text = f.read()
                    except Exception as read_err:
                        logger.error(f"[DossierExporter] Failed reading markdown for {tdoc_id}: {read_err}")

                # Copy media files to export/media/
                doc_images = []
                if local_media_dir.is_dir():
                    for img in local_media_dir.iterdir():
                        if img.is_file() and (tdoc_id.upper() in img.name.upper() or not img.name.startswith("S")):
                            dst = media_export_dir / img.name
                            try:
                                if not dst.exists():
                                    shutil.copy2(to_extended_path(img), to_extended_path(dst))
                                doc_images.append(dst)
                            except Exception:
                                pass

                tdoc_download_url = resolve_tdoc_download_url(
                    tdoc_id, docs_url=self.docs_ftp_url, revisions_url=self.revisions_url, meta=tdoc_data
                )

                converted_docs.append({
                    "tdoc_id": tdoc_id,
                    "title": title,
                    "source": source,
                    "type": doc_type,
                    "status": status,
                    "ai": raw_ai,
                    "row_data": tdoc_data,
                    "md_text": md_text,
                    "media_dir": local_media_dir,
                    "tdoc_url": tdoc_download_url,
                })

            if not converted_docs:
                self.finished.emit(False, "Failed to compile any documents into the dossier.", {})
                return

            # --- Phase 3: Synthesize Interactive HTML Dossier ---
            self.progress.emit("Synthesizing interactive HTML dossier...")
            html_content, tdoc_html_map = self._build_html_dossier(converted_docs, media_export_dir)

            html_path = self.export_dir / "dossier.html"
            with open(to_extended_path(html_path), "w", encoding="utf-8") as f:
                f.write(html_content)
            logger.info(f"[DossierExporter] Saved HTML dossier: {html_path}")

            result_files = {"html": str(html_path)}

            # --- Phase 4: Generate Navigable PDF with Bookmarks ---
            page_spans: Dict[str, Tuple[int, int]] = {}
            pdf_path = self.export_dir / "dossier.pdf"

            if self.export_format in ("both", "pdf"):
                self.progress.emit("Rendering vector PDF with hierarchical bookmarks...")
                page_spans = self._render_bookmarked_pdf(html_content, pdf_path, converted_docs)
                result_files["pdf"] = str(pdf_path)

            # --- Phase 5: Generate Companion AI Citation Prompt ---
            if self.include_prompt:
                self.progress.emit("Synthesizing companion AI citation map...")
                prompt_path = self.export_dir / "dossier_prompt.txt"
                self._generate_ai_prompt(converted_docs, page_spans, prompt_path)
                result_files["prompt"] = str(prompt_path)

            msg = f"Successfully generated Navigable Dossier in:\n{self.export_dir}"
            self.finished.emit(True, msg, result_files)

        except Exception as e:
            logger.error(f"[DossierExporter] Critical error: {e}", exc_info=True)
            self.finished.emit(False, str(e), {})

    # -------------------------------------------------------------------------
    # HTML Dossier Synthesizer
    # -------------------------------------------------------------------------

    def _build_html_dossier(
        self, docs: List[Dict[str, Any]], media_export_dir: Path
    ) -> Tuple[str, Dict[str, str]]:
        """
        Synthesizes the complete HTML dossier with:
        - Floating navigation bar with dropdown and quick actions
        - Executive cover card and statistics
        - Master Table of Contents with jump anchors
        - Fully rendered TDoc sections with highlighted diffs and embedded images
        """
        wg_name = str(self.mtg_info.get("wg_name", "3GPP"))
        mtg_num = str(self.mtg_info.get("meeting_number", ""))
        timestamp_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")

        # Sort docs naturally by Agenda Item then TDoc ID
        def _sort_key(d):
            ai = str(d.get("ai", ""))
            chunks = [int(c) if c.isdigit() else c.lower() for c in re.split(r'(\d+)', ai)]
            return (chunks, str(d.get("tdoc_id", "")))

        sorted_docs = sorted(docs, key=_sort_key)

        # Count categories
        type_counts = {}
        for d in sorted_docs:
            t = d.get("type", "Other")
            type_counts[t] = type_counts.get(t, 0) + 1

        # Build dropdown options
        options_html = ['<option value="#master-index">📑 Master Index & Summary</option>']
        for d in sorted_docs:
            tid = d["tdoc_id"]
            short_t = html.escape(d["title"][:45] + "..." if len(d["title"]) > 45 else d["title"])
            options_html.append(f'<option value="#tdoc-{tid}">{tid} · {short_t}</option>')
        options_markup = "\n".join(options_html)

        # Build stats chips
        stats_chips = [
            f'<div class="stat-chip"><div class="val">{len(sorted_docs)}</div><div class="lbl">Total Documents</div></div>'
        ]
        for t, cnt in sorted(type_counts.items(), key=lambda x: -x[1])[:4]:
            stats_chips.append(f'<div class="stat-chip"><div class="val">{cnt}</div><div class="lbl">{html.escape(t)}</div></div>')
        stats_markup = "\n".join(stats_chips)

        # Build Master Index Table
        table_rows = []
        for d in sorted_docs:
            tid = d["tdoc_id"]
            t_type = d.get("type", "")
            t_status = d.get("status", "")
            badge_type = "badge-cr" if "CR" in t_type.upper() and "P" not in t_type.upper() else (
                "badge-pcr" if "PCR" in t_type.upper() or "P-CR" in t_type.upper() else (
                    "badge-ls" if "LS" in t_type.upper() else "badge-disc"
                )
            )
            badge_status = "badge-agreed" if any(x in t_status.lower() for x in ["agreed", "approved"]) else "badge-noted"

            row_html = f"""
            <tr>
              <td><a href="#tdoc-{tid}"><strong>{tid}</strong></a></td>
              <td><span class="badge {badge_type}">{html.escape(t_type)}</span></td>
              <td>{html.escape(d.get("ai", ""))}</td>
              <td>{html.escape(d.get("title", ""))}</td>
              <td>{html.escape(d.get("source", ""))}</td>
              <td><span class="badge {badge_status}">{html.escape(t_status or "-")}</span></td>
              <td><a href="#tdoc-{tid}">Jump ➔</a></td>
            </tr>
            """
            table_rows.append(row_html)
        index_table_markup = "\n".join(table_rows)

        # Build Document Sections
        sections_html = []
        tdoc_html_map: Dict[str, str] = {}

        for d in sorted_docs:
            tid = d["tdoc_id"]
            raw_md = d.get("md_text", "")
            body_html = self._convert_markdown_to_html(raw_md, d["media_dir"], media_export_dir, tid)

            url_html = f'<a href="{d["tdoc_url"]}" target="_blank">🌐 3GPP Server</a>' if d.get("tdoc_url") else ""
            portal_url = f"https://portal.3gpp.org/ngppapp/CreateTdoc.aspx?mode=view&tdocId={tid}"

            section_markup = f"""
            <section id="tdoc-{tid}" class="tdoc-section" data-tdoc="{tid}">
              <!-- BEGIN_TDOC:{tid} -->
              <div class="tdoc-header">
                <div class="tdoc-header-top">
                  <h1 class="tdoc-id-title"><span style="color:#005A9E;">[{tid}]</span> {html.escape(d.get("title", ""))}</h1>
                  <div class="tdoc-actions no-print">
                    <a href="#master-index">⬆ Back to Index</a>
                    <a href="{portal_url}" target="_blank">🌐 3GU Portal</a>
                    {url_html}
                  </div>
                </div>
                <div class="tdoc-meta-row">
                  <div><strong>Source:</strong> {html.escape(d.get("source", "Unknown"))}</div>
                  <div><strong>Type:</strong> {html.escape(d.get("type", "Unknown"))}</div>
                  <div><strong>Agenda Item:</strong> {html.escape(d.get("ai", "Unassigned"))}</div>
                  <div><strong>Status:</strong> {html.escape(d.get("status", "Unknown"))}</div>
                </div>
              </div>
              <div class="tdoc-body">
                {body_html}
              </div>
              <div style="margin-top: 24px; text-align: right;" class="no-print">
                <a href="#master-index" style="font-size: 11px; color: #64748B; text-decoration: none;">⬆ Back to Top</a>
              </div>
            </section>
            """
            sections_html.append(section_markup)
            tdoc_html_map[tid] = section_markup

        all_sections_markup = "\n".join(sections_html)

        full_html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>3GPP Dossier - {html.escape(wg_name)} #{html.escape(mtg_num)}</title>
  <style>
{DOSSIER_CSS}
  </style>
</head>
<body>

  <!-- Floating Navigation Toolbar -->
  <div class="floating-nav no-print">
    <div class="title-pill">
      📑 <span>{html.escape(wg_name)} #{html.escape(mtg_num)} Dossier</span>
    </div>
    <div class="nav-controls">
      <select onchange="if(this.value) window.location.hash = this.value;">
        {options_markup}
      </select>
      <a href="#master-index" class="btn">📋 Index</a>
      <button class="primary" onclick="window.print();">🖨️ Print / PDF</button>
    </div>
  </div>

  <div class="dossier-wrapper">

    <!-- Cover & Statistics Card -->
    <div id="master-index" class="cover-card">
      <h1>3GPP Technical Dossier</h1>
      <div class="meta-subtitle">
        <strong>Working Group:</strong> {html.escape(wg_name)} | 
        <strong>Meeting:</strong> #{html.escape(mtg_num)} | 
        <strong>Compiled:</strong> {timestamp_str}
      </div>
      <div class="stats-grid">
        {stats_markup}
      </div>
    </div>

    <!-- Master Index Table -->
    <div class="index-card">
      <h2>📑 Master Table of Contents</h2>
      <table class="index-table">
        <thead>
          <tr>
            <th>TDoc</th>
            <th>Type</th>
            <th>AI</th>
            <th>Title</th>
            <th>Source</th>
            <th>Status</th>
            <th>Action</th>
          </tr>
        </thead>
        <tbody>
          {index_table_markup}
        </tbody>
      </table>
    </div>

    <!-- TDoc Content Sections -->
    {all_sections_markup}

  </div>

</body>
</html>
"""
        return full_html, tdoc_html_map

    def _convert_markdown_to_html(
        self, md_text: str, local_media_dir: Path, media_export_dir: Path, tdoc_id: str
    ) -> str:
        """
        Converts document Markdown to semantic HTML, styling inline revision marks,
        3GPP change markers, and inlining images as Base64 Data URIs for a self-contained dossier.
        """
        if not md_text.strip():
            return "<p><em>(No text content available or document could not be converted)</em></p>"

        # Pre-process revision tags before markdown parser
        # Convert [INSERTED: ...] and [DELETED: ...] into styled HTML
        processed = md_text

        # Change markers (* * * * Start of changes * * * *)
        processed = re.sub(
            r'(\*\s*\*\s*\*\s*\*?\s*(?:Start|Next|End)\s+of\s+changes\s*\*?\s*\*\s*\*\s*\*?)',
            r'\n\n<div class="change-marker">\1</div>\n\n',
            processed,
            flags=re.IGNORECASE,
        )

        # Style [INSERTED: ...] and [DELETED: ...]
        processed = re.sub(
            r'\[INSERTED:\s*(.*?)\]',
            r'<ins class="diff-ins">\1</ins>',
            processed,
            flags=re.DOTALL,
        )
        processed = re.sub(
            r'\[DELETED:\s*(.*?)\]',
            r'<del class="diff-del">\1</del>',
            processed,
            flags=re.DOTALL,
        )

        # Render Markdown to HTML
        try:
            body_html = markdown.markdown(
                processed,
                extensions=["tables", "fenced_code", "toc", "attr_list"],
            )
        except Exception as e:
            logger.warning(f"[DossierExporter] markdown library parsing error: {e}. Falling back.")
            doc = QTextDocument()
            doc.setMarkdown(processed)
            body_html = doc.toHtml()

        # In-line images as Base64 Data URIs (or resolve paths)
        if self.embed_base64_images:
            def _replace_img(match):
                full_tag = match.group(0)
                src = match.group(1)
                img_name = Path(src).name

                # Look for image in local_media_dir or media_export_dir
                candidates = [
                    local_media_dir / img_name,
                    media_export_dir / img_name,
                    self.meeting_dir / tdoc_id / "media" / img_name,
                ]
                found_img = next((c for c in candidates if c.is_file()), None)
                if found_img:
                    try:
                        ext = found_img.suffix.lower()
                        mime = "image/png" if ext == ".png" else ("image/jpeg" if ext in [".jpg", ".jpeg"] else "image/png")
                        b64 = base64.b64encode(found_img.read_bytes()).decode("ascii")
                        return f'<img src="data:{mime};base64,{b64}" alt="{html.escape(img_name)}" />'
                    except Exception as err:
                        logger.debug(f"[DossierExporter] Base64 encode failed for {found_img}: {err}")
                return full_tag

            body_html = re.sub(r'<img[^>]+src=["\']([^"\']+)["\'][^>]*>', _replace_img, body_html)

        return body_html

    # -------------------------------------------------------------------------
    # Vector PDF Generation & Outline Bookmark Mapping
    # -------------------------------------------------------------------------

    def _render_bookmarked_pdf(
        self, html_content: str, output_pdf_path: Path, docs: List[Dict[str, Any]]
    ) -> Dict[str, Tuple[int, int]]:
        """
        Renders HTML to a vector PDF via QTextDocument + QPrinter, scans pages with pypdf
        to identify exact 1-indexed page boundaries for each TDoc, and injects hierarchical
        PDF outline bookmarks (Agenda Item > TDoc ID).
        """
        temp_pdf = output_pdf_path.parent / f"temp_{output_pdf_path.name}"
        if temp_pdf.exists():
            try:
                temp_pdf.unlink()
            except Exception:
                pass

        # 1. Print HTML to intermediate PDF
        doc = QTextDocument()
        doc.setHtml(html_content)

        printer = QPrinter(QPrinter.HighResolution)
        printer.setOutputFormat(QPrinter.PdfFormat)
        printer.setOutputFileName(str(temp_pdf))
        printer.setPageMargins(15, 15, 15, 15, QPrinter.Millimeter)
        doc.print_(printer)

        page_spans: Dict[str, Tuple[int, int]] = {}

        if not temp_pdf.exists():
            logger.error("[DossierExporter] Failed to produce intermediate PDF.")
            return page_spans

        # 2. Read with pypdf and scan pages to find TDoc start positions
        reader = pypdf.PdfReader(str(temp_pdf))
        total_pages = len(reader.pages)
        logger.info(f"[DossierExporter] Intermediate PDF rendered with {total_pages} page(s).")

        start_pages: Dict[str, int] = {}
        tdoc_ids = [d["tdoc_id"].upper() for d in docs]

        # Scan each page for the distinctive BEGIN_TDOC or [TDoc] marker
        for page_idx, page in enumerate(reader.pages):
            try:
                text = page.extract_text() or ""
            except Exception:
                text = ""

            for tid in tdoc_ids:
                if tid in start_pages:
                    continue  # Already found
                # Check for marker or prominent header on this page
                if f"BEGIN_TDOC:{tid}" in text or f"[{tid}]" in text or f"TDoc: {tid}" in text:
                    # Make sure it's not the master index table on page 0
                    if page_idx == 0 and "Master Table of Contents" in text:
                        continue
                    start_pages[tid] = page_idx

        # Fallback for any TDoc not explicitly detected: assign sequential estimate
        last_known = 1
        for tid in tdoc_ids:
            if tid not in start_pages:
                start_pages[tid] = min(last_known + 1, max(0, total_pages - 1))
            last_known = start_pages[tid]

        # Calculate exact spans [start_page, end_page] (1-indexed for human readability)
        sorted_by_start = sorted(docs, key=lambda d: start_pages.get(d["tdoc_id"].upper(), 0))
        for i, d in enumerate(sorted_by_start):
            tid = d["tdoc_id"].upper()
            st_0 = start_pages.get(tid, 0)
            if i + 1 < len(sorted_by_start):
                next_tid = sorted_by_start[i + 1]["tdoc_id"].upper()
                next_st_0 = start_pages.get(next_tid, total_pages)
                end_0 = max(st_0, next_st_0 - 1)
            else:
                end_0 = max(0, total_pages - 1)

            page_spans[tid] = (st_0 + 1, end_0 + 1)

        # 3. Assemble Bookmarked PDF
        writer = pypdf.PdfWriter()
        for p in reader.pages:
            writer.add_page(p)

        # Inset master index outline item at page 0
        writer.add_outline_item("📑 Master Index & Summary", 0)

        # Group by Agenda Item for hierarchical outline tree
        ai_groups: Dict[str, List[Dict[str, Any]]] = {}
        for d in docs:
            ai = str(d.get("ai", "Unassigned")).strip()
            ai_groups.setdefault(ai, []).append(d)

        def _ai_sort(ai_name):
            return [int(c) if c.isdigit() else c.lower() for c in re.split(r'(\d+)', ai_name)]

        for ai, group_docs in sorted(ai_groups.items(), key=lambda x: _ai_sort(x[0])):
            first_tid = group_docs[0]["tdoc_id"].upper()
            first_page_0 = start_pages.get(first_tid, 0)

            ai_title = f"📁 Agenda Item {ai} ({len(group_docs)} Docs)"
            ai_bm = writer.add_outline_item(ai_title, first_page_0)

            for d in group_docs:
                tid = d["tdoc_id"]
                tid_upper = tid.upper()
                p_0 = start_pages.get(tid_upper, 0)
                src = d.get("source", "")
                src_tag = f"[{src}] " if src else ""
                t_title = d.get("title", "")
                short_title = t_title[:40] + "..." if len(t_title) > 40 else t_title
                doc_bm_title = f"📄 {tid} · {src_tag}{short_title}"
                writer.add_outline_item(doc_bm_title, p_0, parent=ai_bm)

        # Write final bookmarked PDF
        with open(to_extended_path(output_pdf_path), "wb") as f_out:
            writer.write(f_out)

        # Cleanup temp file
        try:
            if temp_pdf.exists():
                temp_pdf.unlink()
        except Exception:
            pass

        logger.info(f"[DossierExporter] Saved bookmarked PDF dossier: {output_pdf_path}")
        return page_spans

    # -------------------------------------------------------------------------
    # AI Companion Prompt Synthesizer
    # -------------------------------------------------------------------------

    def _generate_ai_prompt(
        self,
        docs: List[Dict[str, Any]],
        page_spans: Dict[str, Tuple[int, int]],
        prompt_path: Path,
    ):
        """
        Synthesizes a companion AI prompt detailing the exact page boundaries
        and structured evaluation directives for NotebookLM, Claude 3.5 Sonnet,
        Gemini 1.5 Pro, and Ollama.
        """
        wg_name = str(self.mtg_info.get("wg_name", "3GPP"))
        mtg_num = str(self.mtg_info.get("meeting_number", ""))
        timestamp_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")

        lines = [
            "# ================================================================================",
            f"# 3GPP TECHNICAL DOSSIER — AI INGESTION & CITATION MAP",
            f"# Working Group : {wg_name} #{mtg_num}",
            f"# Generated     : {timestamp_str}",
            f"# Total TDocs   : {len(docs)}",
            "# ================================================================================",
            "",
            "## 1. DOCUMENT CITATION & EXACT PAGE MAP",
            "Use the exact page ranges below when answering questions or citing proposals:",
            "",
        ]

        for d in docs:
            tid = d["tdoc_id"]
            span = page_spans.get(tid.upper(), (1, 1))
            st, en = span
            span_text = f"pages {st} to {en}" if st != en else f"page {st}"
            title = d.get("title", "Unknown Title")
            source = d.get("source", "Unknown Source")
            dtype = d.get("type", "Unknown Type")
            ai = d.get("ai", "Unknown AI")
            status = d.get("status", "")

            status_str = f" [Status: {status}]" if status else ""
            lines.append(
                f"- **{tid}**: \"{title}\"\n"
                f"  - Sourced by: {source} | Type: {dtype} | AI: {ai}{status_str}\n"
                f"  - PDF Span: {span_text} ({en - st + 1} page(s))\n"
            )

        lines.extend([
            "",
            "# ================================================================================",
            "## 2. 3GPP TECHNICAL CONTRIBUTION EVALUATION DIRECTIVES",
            "# ================================================================================",
            "When analyzing the attached dossier, please follow these standardized 3GPP conventions:",
            "",
            "1. **Change Requests (CRs):**",
            "   - The first page contains the formal CR cover page. The 'Reason for change' and",
            "     'Summary of change' fields summarize the authoritative technical intent.",
            "   - The specification clauses that follow show proposed normative modifications.",
            "   - Focus on revision marks: text enclosed in '[INSERTED: ...]' (or green underline)",
            "     represents proposed additions; '[DELETED: ...]' (or red strikethrough) represents",
            "     proposed deletions.",
            "",
            "2. **Pseudo-CRs (pCRs):**",
            "   - The initial Discussion / Motivation section outlines the technical rationale.",
            "   - The subsequent clauses provide the proposed text for the Technical Report (TR)",
            "     or Technical Specification (TS). Both sections are relevant.",
            "",
            "3. **Discussion Papers & Study Item Contributions:**",
            "   - The entire document outlines technical positions, evaluation results, and architectural options.",
            "",
            "4. **Liaison Statements (LS IN / LS OUT):**",
            "   - Pay particular attention to the 'Actions Requested' and target deadlines.",
            "",
            "5. **Company Positions & Alliances:**",
            "   - Do not consider company names in brackets or with question marks as confirmed co-signers.",
            "   - Explicitly highlight which companies support or oppose specific solutions.",
            "",
            "# ================================================================================",
            "## 3. READY-TO-USE PROMPTS FOR NOTEBOOKLM / CLAUDE / GEMINI / CHATGPT",
            "# ================================================================================",
            "",
            "### [PROMPT A: Executive Agenda Item Summary]",
            "\"Please analyze the attached 3GPP dossier using the page citation map above.",
            "For each Agenda Item included in the document:",
            "1. Provide a concise executive overview of the technical problem being addressed.",
            "2. Summarize each contribution, citing its TDoc ID and exact page numbers.",
            "3. Identify areas of consensus and points of technical disagreement between companies.",
            "4. Highlight any proposed changes to normative 3GPP specifications.\"",
            "",
            "### [PROMPT B: Company Stance & Conflict Analysis]",
            "\"Based on the attached dossier, map out the positions taken by different companies.",
            "Which companies align on specific architectural choices, and which propose competing",
            "solutions? Reference each position with [TDoc ID, page X].\"",
            "",
            "### [PROMPT C: Change Request & Normative Impact Audit]",
            "\"Focus strictly on the CRs and pCRs in this dossier. For each:",
            "- Summarize the 'Reason for change' from the cover sheet.",
            "- Extract the exact normative alterations proposed in the revision-marked clauses.",
            "- Flag any potential backward compatibility risks or architectural side effects.\"",
            "",
        ])

        with open(to_extended_path(prompt_path), "w", encoding="utf-8") as f:
            f.write("\n".join(lines) + "\n")
        logger.info(f"[DossierExporter] Saved companion AI prompt: {prompt_path}")

    # -------------------------------------------------------------------------
    # Helper Methods
    # -------------------------------------------------------------------------

    def _find_word_doc(self, folder: Path, tdoc_id: str) -> Optional[Path]:
        if not folder.exists():
            return None
        for ext in [".docx", ".doc"]:
            files = list(folder.glob(f"*{ext}"))
            if files:
                for f in files:
                    if tdoc_id.lower() in f.name.lower():
                        return f
                return files[0]
        return None

    def _download_and_extract_tdoc(self, tdoc_id: str, tdoc_folder: Path) -> Optional[Path]:
        base_match = re.search(r'^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$', tdoc_id, re.IGNORECASE)
        is_revision = bool(base_match)

        candidate_urls: List[str] = []
        if is_revision:
            if self.revisions_url:
                candidate_urls.append(self.revisions_url)
            if self.docs_ftp_url and self.docs_ftp_url not in candidate_urls:
                candidate_urls.append(self.docs_ftp_url)
        else:
            if self.docs_ftp_url:
                candidate_urls.append(self.docs_ftp_url)
            if self.revisions_url and self.revisions_url not in candidate_urls:
                candidate_urls.append(self.revisions_url)

        if not candidate_urls:
            return None

        last_error = None
        for base_url in candidate_urls:
            if self.isInterruptionRequested():
                return None
            try:
                TDocFileHandler.download_and_extract_tdoc(
                    tdoc_id,
                    base_url,
                    tdoc_folder,
                    timeout=15,
                    cancel_check=self.isInterruptionRequested,
                )
                doc_path = self._find_word_doc(tdoc_folder, tdoc_id)
                if doc_path:
                    return doc_path
            except Exception as e:
                last_error = e

        return None
