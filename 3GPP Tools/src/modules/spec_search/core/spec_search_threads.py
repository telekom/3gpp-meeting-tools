"""
Background QThread workers for asynchronous FTP downloading, DOCX extraction, and fast query execution.
Filters out 3GPP revision mark (-rm) files during extraction and indexing.
"""

import logging
from pathlib import Path
import re
from typing import Any, Dict, List, Optional
import zipfile
import pandas as pd
from PyQt5.QtCore import QThread, pyqtSignal

from core.network.session import NetworkSession
from modules.nas.core.nas_threads import find_cached_spec_file
from modules.nas.core.parsing.protocol_parser_common import ProtocolDocxDispatcher
from modules.spec_search.core.spec_search_db import SpecSearchDatabase
from modules.spec_search.core.spec_text_extractor import SpecDocxExtractor
from modules.word_tools.core.word_converter import convert_doc_to_docx


def is_change_mark_file(file_path_or_name: Any) -> bool:
    """
    Identifies 3GPP Word change/revision-marked documents (-rm).
    Matches patterns like '24501-i40-rm.docx', '38331-h20_rm.doc', or 'spec_s00_s04-rm.docx'.
    """
    stem = Path(str(file_path_or_name)).stem.lower()
    return bool(re.search(r"[-_]rm(?=[-._]|$)", stem, re.IGNORECASE))


class SpecSearchImportThread(QThread):
    """Background worker to fetch, extract, and index multiple specification versions with release dates."""

    progress = pyqtSignal(str, int)
    finished_success = pyqtSignal(int, int)  # specs_count, total_clauses
    error = pyqtSignal(str)

    def __init__(
        self,
        search_db_path: Path,
        tasks: List[Dict[str, Any]],
        cache_dir: Path,
    ):
        super().__init__()
        self.db = SpecSearchDatabase(search_db_path)
        self.tasks = tasks
        self.cache_dir = Path(cache_dir)

    def run(self):
        total_tasks = len(self.tasks)
        if total_tasks == 0:
            self.error.emit("No specifications selected for indexing.")
            return

        total_clauses_indexed = 0
        successful_specs = 0

        for t_idx, task in enumerate(self.tasks):
            spec_number = task.get("spec_number", "")
            version = task.get("version", "")
            filename = task.get("filename", "")
            file_url = task.get("file_url", "")
            release_date = task.get("release_date") or task.get("upload_date")
            local_docx_input = task.get("local_docx_paths") or task.get("local_docx_path")

            base_progress = int((t_idx / total_tasks) * 100)
            task_weight = 1.0 / total_tasks

            def emit_task_progress(msg: str, step_pct: int):
                overall = base_progress + int(step_pct * task_weight)
                self.progress.emit(f"[{t_idx + 1}/{total_tasks}] {msg}", min(overall, 99))

            try:
                target_docs: List[Path] = []

                # 1. Local files (filter out any -rm files)
                if local_docx_input:
                    if isinstance(local_docx_input, list):
                        target_docs = [
                            Path(p) for p in local_docx_input
                            if Path(p).exists() and not is_change_mark_file(p)
                        ]
                    elif Path(local_docx_input).exists() and not is_change_mark_file(local_docx_input):
                        target_docs = [Path(local_docx_input)]

                # 2. Cache Lookup / FTP Download
                if not target_docs:
                    spec_cache_dir = self.cache_dir / spec_number
                    spec_cache_dir.mkdir(parents=True, exist_ok=True)

                    cached_hit = find_cached_spec_file(filename, spec_number)
                    zip_to_extract: Optional[Path] = None

                    if cached_hit and cached_hit.suffix.lower() == ".zip":
                        zip_to_extract = cached_hit
                    elif cached_hit and cached_hit.suffix.lower() in [".docx", ".doc"]:
                        if not is_change_mark_file(cached_hit.name):
                            base_prefix = re.sub(r"_\d+_.*$", "", cached_hit.stem)
                            siblings = list(cached_hit.parent.glob(f"{base_prefix}*.docx")) + list(
                                cached_hit.parent.glob(f"{base_prefix}*.doc")
                            )
                            target_docs = [
                                p for p in set(siblings)
                                if not is_change_mark_file(p.name)
                            ]
                            target_docs.sort(
                                key=lambda p: ProtocolDocxDispatcher._extract_part_index(p.name)
                            )
                    else:
                        zip_path = spec_cache_dir / filename
                        emit_task_progress(f"Downloading {filename} from 3GPP FTP...", 20)
                        NetworkSession.download_file(file_url, zip_path)
                        zip_to_extract = zip_path

                    if zip_to_extract and zip_to_extract.exists():
                        emit_task_progress(f"Extracting {zip_to_extract.name}...", 35)
                        with zipfile.ZipFile(zip_to_extract, "r") as zf:
                            for member in zf.namelist():
                                # Ignore macOS metadata and any revision mark (-rm) documents
                                if (
                                    member.lower().endswith((".docx", ".doc"))
                                    and not member.startswith("._")
                                    and "__MACOSX" not in member
                                    and not is_change_mark_file(member)
                                ):
                                    zf.extract(member, spec_cache_dir)
                                    target_docs.append(spec_cache_dir / member)

                if not target_docs:
                    self.progress.emit(f"⚠️ Could not locate clean Word file for {filename}. Skipping...", base_progress)
                    continue

                # Final guard against revision mark files
                target_docs = [p for p in target_docs if not is_change_mark_file(p.name)]

                # Convert legacy .doc to .docx if necessary
                converted_docs: List[Path] = []
                for doc_file in target_docs:
                    if doc_file.suffix.lower() == ".doc":
                        emit_task_progress(f"Converting legacy .doc: {doc_file.name}...", 45)
                        converted_docs.append(convert_doc_to_docx(doc_file))
                    else:
                        converted_docs.append(doc_file)

                # Extract Clauses directly via XML
                emit_task_progress(f"Parsing clauses in TS {spec_number} v{version}...", 60)
                extractor = SpecDocxExtractor(converted_docs)
                clauses = extractor.extract_clauses(
                    progress_callback=lambda msg, p: emit_task_progress(msg, 60 + int(p * 0.25))
                )

                # Index in SQLite with release date
                emit_task_progress(f"Indexing {len(clauses)} clauses into Trigram DB...", 90)
                success = self.db.insert_parsed_spec(spec_number, version, clauses, release_date=release_date)
                if success:
                    successful_specs += 1
                    total_clauses_indexed += len(clauses)

            except Exception as e:
                self.progress.emit(f"⚠️ Ingestion error on {filename}: {e}", base_progress)

        self.progress.emit("Indexing completed.", 100)
        self.finished_success.emit(successful_specs, total_clauses_indexed)


class SpecSearchQueryWorker(QThread):
    """Background query executor to keep the UI smooth during searches."""

    results_ready = pyqtSignal(object, int)  # DataFrame, request_id

    def __init__(
        self,
        db: SpecSearchDatabase,
        query_str: str,
        version_ids: List[int],
        clause_filter: Optional[str],
        request_id: int,
    ):
        super().__init__()
        self.db = db
        self.query_str = query_str
        self.version_ids = version_ids
        self.clause_filter = clause_filter
        self.request_id = request_id

    def run(self):
        try:
            df = self.db.search_substring(self.query_str, self.version_ids, self.clause_filter)
            self.results_ready.emit(df, self.request_id)
        except Exception:
            self.results_ready.emit(pd.DataFrame(), self.request_id)