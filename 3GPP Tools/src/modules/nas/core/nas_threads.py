import logging
from pathlib import Path
import re
from typing import Any, Dict, List, Optional
import zipfile
from PyQt5.QtCore import QThread, pyqtSignal

from core.network.session import NetworkSession
from core.utils.paths import get_project_root
from modules.meetings.core.settings import MeetingsSettings
from modules.nas.core.nas_db import NASDatabase
from modules.nas.core.parsing.nas_parser import ProtocolDocxDispatcher
from modules.word_tools.core.word_converter import convert_doc_to_docx


def get_candidate_cache_dirs() -> List[Path]:
    """Resolves all candidate paths using MeetingsSettings and project directories."""
    candidate_paths: List[Path] = []

    try:
        settings = MeetingsSettings()
        cfg_download_dir = Path(settings.cache_dir)
        candidate_paths.extend([
            cfg_download_dir,
            cfg_download_dir / "specs",
            cfg_download_dir.parent / "specs",
            cfg_download_dir.parent / "cache" / "specs",
            cfg_download_dir.parent,
        ])
    except Exception as e:
        logging.warning(f"Could not load MeetingsSettings for NAS cache resolution: {e}")

    home_helper = Path.home() / "3GPP_Delegate_Helper"
    candidate_paths.extend([
        home_helper / "specs",
        home_helper / "cache" / "specs",
        home_helper,
    ])

    root = get_project_root()
    candidate_paths.extend([
        root / "cache" / "specs",
        root / "cache",
        root.parent / "cache" / "specs",
        root.parent / "cache",
    ])

    seen = set()
    unique_dirs = []
    for p in candidate_paths:
        resolved = str(p.resolve()) if p.exists() else str(p)
        if resolved not in seen:
            seen.add(resolved)
            unique_dirs.append(p)

    return unique_dirs


def find_cached_spec_file(filename: str, spec_number: str = "24.501") -> Optional[Path]:
    """Searches candidate directories for cached .zip, .docx, or .doc files."""
    doc_base = filename.replace(".zip", "").lower()
    clean_spec = spec_number.replace(".", "")

    for base_dir in get_candidate_cache_dirs():
        if not base_dir.exists():
            continue

        search_dirs = [
            base_dir / spec_number,
            base_dir / clean_spec,
            base_dir,
        ]

        for s_dir in search_dirs:
            if not s_dir.exists():
                continue

            # Prioritize exact zip archive to guarantee full multi-part extraction
            if filename.lower().endswith(".zip"):
                zip_target = s_dir / filename
                if zip_target.exists() and zip_target.is_file():
                    return zip_target

            # Check for generic zip matching the document base name
            for zip_file in s_dir.glob(f"*{doc_base}*.zip"):
                if zip_file.is_file() and not zip_file.name.startswith("._") and "__MACOSX" not in zip_file.parts:
                    return zip_file

            # Fallback to cached docx/doc files
            for ext in [".docx", ".doc"]:
                for file_path in s_dir.glob(f"*{doc_base}*{ext}"):
                    if file_path.is_file() and not file_path.name.startswith("._") and "__MACOSX" not in file_path.parts:
                        return file_path

    return None


class NASFetchAndImportThread(QThread):
    """Background worker that sequentially ingests single or split TS 24.501 / TS 24.301 specifications."""

    progress = pyqtSignal(str, int)
    finished_success = pyqtSignal(int, int)
    error = pyqtSignal(str)

    def __init__(
        self,
        nas_db_path: Path,
        tasks: List[Dict[str, Any]],
        cache_dir: Path,
    ):
        super().__init__()
        self.nas_db = NASDatabase(nas_db_path)
        self.tasks = tasks
        self.cache_dir = Path(cache_dir)

    def run(self):
        total_tasks = len(self.tasks)
        if total_tasks == 0:
            self.error.emit("No specifications selected for ingestion.")
            return

        total_messages_imported = 0
        successful_specs = 0

        for t_idx, task in enumerate(self.tasks):
            spec_number = task.get("spec_number", "24.501")
            version = task.get("version", "")
            filename = task.get("filename", "")
            file_url = task.get("file_url", "")
            local_docx_input = task.get("local_docx_paths") or task.get("local_docx_path")

            base_progress = int((t_idx / total_tasks) * 100)
            task_weight = 1.0 / total_tasks

            def emit_task_progress(msg: str, step_pct: int):
                overall = base_progress + int(step_pct * task_weight)
                self.progress.emit(f"[{t_idx + 1}/{total_tasks}] {msg}", min(overall, 99))

            try:
                target_docs: List[Path] = []

                # 1. Direct Local Files
                if local_docx_input:
                    if isinstance(local_docx_input, list):
                        target_docs = [Path(p) for p in local_docx_input if Path(p).exists()]
                    elif Path(local_docx_input).exists():
                        target_docs = [Path(local_docx_input)]

                    if target_docs:
                        emit_task_progress(f"Loading local file(s): {len(target_docs)} part(s)...", 10)

                # 2. Automated Cache Lookup / FTP Download
                if not target_docs:
                    spec_cache_dir = self.cache_dir / spec_number
                    spec_cache_dir.mkdir(parents=True, exist_ok=True)

                    cached_hit = find_cached_spec_file(filename, spec_number)
                    zip_to_extract: Optional[Path] = None

                    if cached_hit and cached_hit.suffix.lower() == ".zip":
                        zip_to_extract = cached_hit
                        emit_task_progress(f"Found cached archive: {cached_hit.name}...", 20)
                    elif cached_hit and cached_hit.suffix.lower() in [".docx", ".doc"]:
                        # Look for potential split sibling parts cached locally
                        base_prefix = re.sub(r"_\d+_.*$", "", cached_hit.stem)
                        siblings = list(cached_hit.parent.glob(f"{base_prefix}*.docx")) + list(cached_hit.parent.glob(f"{base_prefix}*.doc"))
                        target_docs = sorted(list(set(siblings)), key=lambda p: ProtocolDocxDispatcher._extract_part_index(p.name))
                        emit_task_progress(f"Found cached document(s): {len(target_docs)} part(s)", 20)
                    else:
                        zip_path = spec_cache_dir / filename
                        emit_task_progress(f"Downloading {filename} from 3GPP FTP...", 25)
                        NetworkSession.download_file(file_url, zip_path)
                        zip_to_extract = zip_path

                    if zip_to_extract and zip_to_extract.exists():
                        emit_task_progress(f"Extracting all parts from {zip_to_extract.name}...", 40)
                        with zipfile.ZipFile(zip_to_extract, "r") as zf:
                            for member in zf.namelist():
                                if (
                                    member.lower().endswith((".docx", ".doc"))
                                    and not member.startswith("._")
                                    and "__MACOSX" not in member
                                ):
                                    zf.extract(member, spec_cache_dir)
                                    target_docs.append(spec_cache_dir / member)

                if not target_docs:
                    self.progress.emit(f"⚠️ Could not locate Word doc(s) for {filename}. Skipping...", base_progress)
                    continue

                # Convert legacy .doc to .docx if required
                converted_docs: List[Path] = []
                for doc_file in target_docs:
                    if doc_file.suffix.lower() == ".doc":
                        emit_task_progress(f"Converting legacy .doc: {doc_file.name}...", 48)
                        converted_docs.append(convert_doc_to_docx(doc_file))
                    else:
                        converted_docs.append(doc_file)

                # Initialize Parser with all parts
                parser = ProtocolDocxDispatcher(converted_docs)
                if not version:
                    version = parser.extract_version_from_filename()

                num_parts_str = f" ({len(converted_docs)} parts)" if len(converted_docs) > 1 else ""
                emit_task_progress(f"Parsing TS {spec_number} v{version}{num_parts_str}...", 55)

                messages, ie_defs = parser.parse(
                    progress_callback=lambda msg, p: emit_task_progress(msg, 55 + int(p * 0.35))
                )

                emit_task_progress("Saving records to database...", 92)
                success = self.nas_db.insert_parsed_spec(spec_number, version, messages, ie_defs)

                if success:
                    successful_specs += 1
                    total_messages_imported += len(messages)

            except Exception as e:
                self.progress.emit(f"⚠️ Error ingesting {filename}: {e}", base_progress)

        self.progress.emit("Batch ingestion complete.", 100)
        self.finished_success.emit(successful_specs, total_messages_imported)