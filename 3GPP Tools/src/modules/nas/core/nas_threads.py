import logging
from pathlib import Path
from typing import List, Optional
import zipfile
from PyQt5.QtCore import QThread, pyqtSignal

from core.network.session import NetworkSession
from core.utils.paths import get_project_root
from modules.meetings.core.settings import MeetingsSettings
from modules.nas.core.nas_db import NASDatabase
from modules.nas.core.nas_parser import NASDocxParser


def get_candidate_cache_dirs() -> List[Path]:
    """Resolves all candidate paths using MeetingsSettings and project directories."""
    candidate_paths: List[Path] = []

    # 1. Paths derived from MeetingsSettings (meetings_config.json / 3GPP_Delegate_Helper)
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

    # 2. Standard user directory fallback
    home_helper = Path.home() / "3GPP_Delegate_Helper"
    candidate_paths.extend([
        home_helper / "specs",
        home_helper / "cache" / "specs",
        home_helper,
    ])

    # 3. Project root relative paths
    root = get_project_root()
    candidate_paths.extend([
        root / "cache" / "specs",
        root / "cache",
        root.parent / "cache" / "specs",
        root.parent / "cache",
    ])

    # Deduplicate while preserving order
    seen = set()
    unique_dirs = []
    for p in candidate_paths:
        resolved = str(p.resolve()) if p.exists() else str(p)
        if resolved not in seen:
            seen.add(resolved)
            unique_dirs.append(p)

    return unique_dirs


def find_cached_spec_file(filename: str, spec_number: str = "24.501") -> Optional[Path]:
    """
    Searches candidate directories for cached .docx, .doc, or .zip files matching the spec.
    """
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

            # 1. Look for extracted Word documents
            for ext in [".docx", ".doc"]:
                for file_path in s_dir.glob(f"*{doc_base}*{ext}"):
                    if file_path.is_file() and not file_path.name.startswith(
                            "._") and "__MACOSX" not in file_path.parts:
                        return file_path

            # 2. Look for existing ZIP file
            zip_target = s_dir / filename
            if zip_target.exists() and zip_target.is_file():
                return zip_target

            for zip_file in s_dir.glob(f"*{doc_base}*.zip"):
                if zip_file.is_file() and not zip_file.name.startswith("._") and "__MACOSX" not in zip_file.parts:
                    return zip_file

    return None


class NASFetchAndImportThread(QThread):
    """
    Background worker that resolves local cache, downloads missing .zip archives,
    extracts the .docx file, and parses Clauses 8 and 9 into nas_data.db.
    """

    progress = pyqtSignal(str, int)
    finished_success = pyqtSignal(str, str, int)
    error = pyqtSignal(str)

    def __init__(
            self,
            nas_db_path: Path,
            spec_number: str,
            version: str,
            filename: str,
            file_url: str,
            cache_dir: Path,
            local_docx_path: Optional[Path] = None,
    ):
        super().__init__()
        self.nas_db = NASDatabase(nas_db_path)
        self.spec_number = spec_number
        self.version = version
        self.filename = filename
        self.file_url = file_url
        self.cache_dir = Path(cache_dir)
        self.local_docx_path = Path(local_docx_path) if local_docx_path else None

    def run(self):
        try:
            target_docx: Optional[Path] = None

            # Case A: Manual Local .docx Import
            if self.local_docx_path and self.local_docx_path.exists():
                target_docx = self.local_docx_path
                self.progress.emit(f"Loading local file: {target_docx.name}...", 10)

            # Case B: Automated Lookup from Cache / Remote Fetch
            else:
                spec_cache_dir = self.cache_dir / self.spec_number
                spec_cache_dir.mkdir(parents=True, exist_ok=True)

                cached_hit = find_cached_spec_file(self.filename, self.spec_number)

                # 1. Existing extracted .docx or .doc file
                if cached_hit and cached_hit.suffix.lower() in [".docx", ".doc"]:
                    target_docx = cached_hit
                    self.progress.emit(f"Found cached document: {target_docx.name}", 20)

                # 2. Existing cached .zip file -> extract to specs directory
                elif cached_hit and cached_hit.suffix.lower() == ".zip":
                    self.progress.emit(f"Extracting cached archive: {cached_hit.name}...", 25)
                    with zipfile.ZipFile(cached_hit, "r") as zf:
                        for member in zf.namelist():
                            if (
                                    member.lower().endswith((".docx", ".doc"))
                                    and not member.startswith("._")
                                    and "__MACOSX" not in member
                            ):
                                zf.extract(member, spec_cache_dir)
                                target_docx = spec_cache_dir / member
                                break

                # 3. Not cached -> download .zip from FTP
                else:
                    zip_path = spec_cache_dir / self.filename
                    self.progress.emit(f"Downloading {self.filename} from 3GPP FTP...", 25)
                    NetworkSession.download_file(self.file_url, zip_path)

                    self.progress.emit(f"Extracting {self.filename}...", 45)
                    with zipfile.ZipFile(zip_path, "r") as zf:
                        for member in zf.namelist():
                            if (
                                    member.lower().endswith((".docx", ".doc"))
                                    and not member.startswith("._")
                                    and "__MACOSX" not in member
                            ):
                                zf.extract(member, spec_cache_dir)
                                target_docx = spec_cache_dir / member
                                break

            if not target_docx or not target_docx.exists():
                self.error.emit(f"Could not find or extract Word document for TS {self.spec_number} v{self.version}")
                return

            # Parse Clauses 8 & 9
            self.progress.emit(f"Parsing TS {self.spec_number} v{self.version} ({target_docx.name})...", 55)
            parser = NASDocxParser(target_docx)
            messages, ie_defs = parser.parse(
                progress_callback=lambda msg, p: self.progress.emit(msg, 55 + int(p * 0.35))
            )

            # Insert into SQLite
            self.progress.emit("Saving records to NAS database...", 92)
            success = self.nas_db.insert_parsed_spec(
                self.spec_number, self.version, messages, ie_defs
            )

            if success:
                self.progress.emit("Import complete.", 100)
                self.finished_success.emit(self.spec_number, self.version, len(messages))
            else:
                self.error.emit("Failed to save parsed data into nas_data.db.")

        except Exception as e:
            self.error.emit(f"Ingestion error: {str(e)}")