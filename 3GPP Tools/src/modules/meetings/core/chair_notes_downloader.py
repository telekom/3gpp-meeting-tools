# --- File: src/modules/meetings/core/chair_notes_downloader.py ---
import os
import re
import logging
import zipfile
import urllib.parse
from pathlib import Path
from typing import List, Tuple
from bs4 import BeautifulSoup
from PyQt5.QtCore import QThread, pyqtSignal

from core.utils.paths import to_extended_path
from core.network.session import NetworkSession


# Standard 3GPP pure TDoc ZIP pattern (e.g. R2-2601234.zip, S2-2605678r01.zip)
_PURE_TDOC_ZIP_RE = re.compile(
    r'^[A-Z][A-Z0-9]{1,3}-\d{4,8}(?:[rR]\d{1,2}[a-zA-Z]?)?\.zip$',
    re.IGNORECASE,
)

# Extensions eligible for Chairman's Notes / Agenda notes
_VALID_NOTES_EXTENSIONS = (".docx", ".doc", ".zip", ".htm", ".html")


class ChairNotesListingThread(QThread):
    """
    Lightweight asynchronous probe thread that checks candidate 3GPP meeting URLs
    for an active /Inbox/Chair_Notes or /Inbox/Chairs_Notes directory, parses all available
    candidate files without downloading them, and emits the candidate file list.
    """
    progress = pyqtSignal(str)
    # Emits: (success: bool, files: list of (fname, url), target_dir_url: str, message: str)
    finished = pyqtSignal(bool, list, str, str)

    def __init__(self, candidate_base_urls: list, parent=None):
        super().__init__(parent)
        self.candidate_base_urls = candidate_base_urls

    def run(self):
        try:
            session = NetworkSession.get_instance()
            NetworkSession.apply_humanness(session)

            target_dir_url, html_content = ChairNotesDownloaderThread.locate_chair_notes_dir(
                session,
                self.candidate_base_urls,
                progress_callback=lambda msg: self.progress.emit(msg),
                cancel_check=self.isInterruptionRequested,
            )

            if self.isInterruptionRequested():
                return

            if not target_dir_url or not html_content:
                self.finished.emit(
                    False,
                    [],
                    "",
                    "Could not locate an active /Inbox/Chairs_Notes or /Inbox/Chair_Notes folder on available servers.",
                )
                return

            files = ChairNotesDownloaderThread.parse_file_list_static(html_content, target_dir_url)
            if not files:
                self.finished.emit(
                    True,
                    [],
                    target_dir_url,
                    "Connected to Chairman's Notes folder, but no matching files were found.",
                )
                return

            self.finished.emit(True, files, target_dir_url, "")
        except Exception as ex:
            logging.error(f"[ChairNotesListing] Unexpected failure: {ex}", exc_info=True)
            self.finished.emit(False, [], "", f"Error scanning Chairman's Notes: {ex}")


class ChairNotesDownloaderThread(QThread):
    """
    Asynchronously crawls /Inbox/Chairs_Notes and /Inbox/Chair_Notes directory listings
    across candidate URLs, filters Chairman's Notes / Agenda EOM / Session Report files
     across all 3GPP WGs (SA2, RAN1, RAN2, RAN3, CT, etc.), downloads them to
    Agenda/ChairNotes/, and automatically extracts .zip archives.
    """
    progress = pyqtSignal(str)
    # Emits: (success: bool, downloaded_count: int, downloaded_files: list, message: str)
    finished = pyqtSignal(bool, int, list, str)

    SUBPATHS = [
        "Inbox/Chairs_Notes",
        "INBOX/Chairs_Notes",
        "Inbox/Chair_Notes",
        "INBOX/Chair_Notes",
        "Inbox/Chairs_notes",
        "INBOX/Chairs_notes",
        "Inbox/Chair_notes",
        "INBOX/Chair_notes",
        "Chairs_Notes",
        "Chair_Notes",
    ]

    def __init__(
        self,
        candidate_base_urls: list,
        target_dir: Path,
        parent=None,
        only_final_or_latest: bool = True,
        selected_files: list = None,
        overwrite: bool = True,
    ):
        super().__init__(parent)
        self.candidate_base_urls = candidate_base_urls
        self.target_dir = Path(target_dir)
        self.only_final_or_latest = only_final_or_latest
        self.selected_files = selected_files
        self.overwrite = overwrite

    def run(self):
        try:
            os.makedirs(to_extended_path(self.target_dir), exist_ok=True)
            session = NetworkSession.get_instance()
            NetworkSession.apply_humanness(session)

            if self.selected_files is not None:
                files_to_download = list(self.selected_files)
            else:
                target_dir_url, html_content = self.locate_chair_notes_dir(
                    session,
                    self.candidate_base_urls,
                    progress_callback=lambda msg: self.progress.emit(msg),
                    cancel_check=self.isInterruptionRequested,
                )

                if self.isInterruptionRequested():
                    return

                if not target_dir_url or not html_content:
                    self.finished.emit(
                        False,
                        0,
                        [],
                        "Could not locate an active /Inbox/Chairs_Notes or /Inbox/Chair_Notes folder on available servers.",
                    )
                    return

                # 2. Extract valid Chairman's Notes files
                files_to_download = self._parse_file_list(html_content, target_dir_url)
                if self.only_final_or_latest and len(files_to_download) > 1:
                    files_to_download = self.select_best_chair_notes_files(files_to_download)

            if not files_to_download:
                self.finished.emit(
                    True,
                    0,
                    [],
                    "No matching Chairman's Notes files were selected or found.",
                )
                return

            # 3. Stream downloads to Agenda/ChairNotes/ and extract .zip packages
            downloaded = self.download_files_sync(
                session,
                files_to_download,
                self.target_dir,
                progress_callback=lambda msg: self.progress.emit(msg),
                cancel_check=self.isInterruptionRequested,
                overwrite=self.overwrite,
            )

            if self.isInterruptionRequested():
                return

            success = len(downloaded) > 0
            msg = f"Downloaded {len(downloaded)} Chairman's Notes file(s) to:\n{self.target_dir}"
            self.finished.emit(success, len(downloaded), downloaded, msg)

        except Exception as ex:
            logging.error(f"[ChairNotes] Unexpected failure: {ex}", exc_info=True)
            self.finished.emit(False, 0, [], f"Error downloading Chairman's Notes: {ex}")

    @classmethod
    def locate_chair_notes_dir(
        cls,
        session,
        candidate_base_urls: list,
        progress_callback=None,
        cancel_check=None,
    ) -> Tuple[str, str]:
        """Probes candidate base URLs across Chairs_Notes / Chair_Notes subpaths."""
        for base_url in candidate_base_urls:
            if cancel_check and cancel_check():
                return None, ""
            if not base_url:
                continue
            clean_base = base_url.rstrip("/")
            for sub in cls.SUBPATHS:
                if cancel_check and cancel_check():
                    return None, ""

                probe_url = f"{clean_base}/{sub}"
                try:
                    if progress_callback:
                        progress_callback(f"Probing {probe_url[:35]}...")
                    resp = session.get(probe_url, timeout=15)
                    if resp.status_code == 200 and cls._looks_like_directory_listing(resp.text):
                        # Verify at least one candidate file is inside (or accept if valid listing)
                        parsed = cls.parse_file_list_static(resp.text, probe_url)
                        if parsed or "ChairNotes" in resp.text or "Directory Listing" in resp.text:
                            logging.info(f"[ChairNotes] Found active directory listing at: {probe_url}")
                            return probe_url, resp.text
                except Exception as e:
                    logging.debug(f"[ChairNotes] Failed to probe {probe_url}: {e}")
        return None, ""

    @staticmethod
    def _looks_like_directory_listing(html_text: str) -> bool:
        if not html_text:
            return False
        low = html_text.lower()
        return any(
            token in low
            for token in (
                "chairnotes",
                "chair_notes",
                "chairs_notes",
                "directory listing",
                "parent directory",
                "downloadinput",
                "sort by",
                "<table",
                ".docx",
                ".doc",
                ".zip",
            )
        )

    @classmethod
    def is_valid_chair_notes_filename(cls, filename: str) -> bool:
        """
        Determines whether a filename inside a Chair_Notes / Chairs_Notes folder
        is a valid Chairman's Notes, Agenda EOM, or Session Report file.
        """
        if not filename:
            return False
        name = filename.strip()
        if name.startswith("~$") or name.startswith("."):
            return False
        up = name.upper()
        if up in ("OLDER", "PARENT DIRECTORY", "ARCHIVES", "ARCHIVE", "OLD", "BACKUP"):
            return False

        low = name.lower()
        if not low.endswith(_VALID_NOTES_EXTENSIONS):
            return False

        # Exclude standalone TDoc archives (e.g. R2-2607187.zip) sometimes dropped into Chair_Notes
        if _PURE_TDOC_ZIP_RE.match(name):
            return False

        # Legacy & explicit ChairNotes keyword match
        if "chairnotes" in low or "chair_notes" in low or "chairs_notes" in low or "chair notes" in low or "chairs notes" in low:
            return True

        # Broader RAN1/RAN2/RAN3/SA/CT patterns inside a dedicated Chair_Notes folder:
        keywords = (
            "chair",
            "note",
            "agenda",
            "session",
            "report",
            "eom",
            "final",
            "summary",
            "status",
            "minutes",
            "agreement",
            "conclusion",
        )
        return any(kw in low for kw in keywords)

    def _parse_file_list(self, html: str, base_url: str) -> List[Tuple[str, str]]:
        """Extracts and validates ChairNotes filenames and URLs from HTML directory listings."""
        return self.parse_file_list_static(html, base_url)

    @classmethod
    def parse_file_list_static(cls, html: str, base_url: str) -> List[Tuple[str, str]]:
        """Static helper to extract and validate ChairNotes filenames and URLs from HTML listings."""
        soup = BeautifulSoup(html, "html.parser")
        file_map = {}

        # 1. Check ASP.NET checkbox input controls: value="ChairNotes_Andy_...doc"
        for inp in soup.find_all("input", class_="downloadInput"):
            val = inp.get("value", "").strip()
            if val and cls.is_valid_chair_notes_filename(val):
                file_map[val] = f"{base_url.rstrip('/')}/{urllib.parse.quote(val)}"

        # 2. Check standard anchor tags: href=".../ChairNotes_...doc" or RAN3_133bis_agenda_..._EOM.docx
        for a in soup.find_all("a", href=True):
            href = a["href"].strip()
            if not href or href.startswith(("?", "#", "javascript:", "mailto:")):
                continue
            parsed_name = urllib.parse.unquote(os.path.basename(href.rstrip("/")))
            link_text = a.get_text().strip()
            target_name = parsed_name if parsed_name and "." in parsed_name else link_text

            if target_name and cls.is_valid_chair_notes_filename(target_name):
                full_url = urllib.parse.urljoin(base_url.rstrip("/") + "/", href)
                if target_name not in file_map:
                    file_map[target_name] = full_url

        return list(file_map.items())

    @classmethod
    def select_best_chair_notes_files(
        cls, file_tuples: List[Tuple[str, str]], only_final_or_latest: bool = True
    ) -> List[Tuple[str, str]]:
        """
        Smart selector that prioritizes authoritative Final / EOM Chairman's Notes
        and filters out redundant daily interim snapshots or tracked-change duplicates.

        Handles:
        - RAN3: 20+ `RAN3_133bis_agenda_YYYYMMDD_HHMM.docx` files + `RAN3_133bis_agenda_YYYYMMDD_EOM.docx`
          -> Selects only the `_EOM.docx` file (or latest timestamped snapshot if no EOM exists yet).
        - RAN2: `R2_135b_Chair_Notes_10-16-26_final.zip` (or `_final_clean.docx`) + session `_EOM.docx` files
          + intra-week drafts -> Selects the `_final` combined file(s) and `_EOM` session reports,
          preferring `_clean`/`_cl` over `_tracked_changes`/`_rm`.
        - SA2: Per-chair track files (`ChairNotes_Andy_...`, `ChairNotes_hector_...`) -> Keeps latest file
          per chair/track.
        """
        if not only_final_or_latest or len(file_tuples) <= 1:
            return list(file_tuples)

        # 1. Deduplicate clean vs tracked_changes / revision-marked variants at the listing level
        deduped = cls._prefer_clean_variants(file_tuples)

        # Classify files
        combined_finals = []
        session_finals = []
        interim_files = []

        for fname, url in deduped:
            low = fname.lower()
            stem = Path(fname).stem.lower()
            is_final_or_eom = bool(
                re.search(r'(?:^|[_\-\s\.])(final|eom|end[_\-\s]?of[_\-\s]?meeting|closing)(?:$|[_\-\s\.])', stem)
                or stem.endswith(("final", "eom", "_cl", "_clean"))
                and ("final" in stem or "eom" in stem)
            )
            is_session_or_track = bool(
                "session" in low
                or re.search(r'\([a-z\s]+\)', low)
                or re.match(r'^chair_?notes_[a-z]+', low)
            )

            if is_final_or_eom:
                if is_session_or_track:
                    session_finals.append((fname, url))
                else:
                    combined_finals.append((fname, url))
            else:
                interim_files.append((fname, url))

        # Case A: We have at least one combined full-meeting final/EOM file (e.g. RAN3 EOM or RAN2 _final)
        if combined_finals:
            # If multiple combined finals exist (e.g. v1 and v2), pick the latest/best combined final
            best_combined = [cls._pick_latest_file(combined_finals)]
            # Also include any session EOM reports (e.g. RAN2 breakout session EOMs), deduplicated by track
            best_sessions = cls._group_and_pick_latest_by_track(session_finals)
            return best_combined + best_sessions

        # Case B: No combined final file, but per-session EOM files exist
        if session_finals:
            return cls._group_and_pick_latest_by_track(session_finals)

        # Case C: Ongoing meeting (only interim snapshots or per-chair files exist)
        return cls._group_and_pick_latest_by_track(deduped)

    @classmethod
    def _prefer_clean_variants(cls, file_tuples: List[Tuple[str, str]]) -> List[Tuple[str, str]]:
        """
        When both a `_clean` / `_cl` file and a `_tracked_changes` / `_rm` / `_rev_marks` file
        exist for the same base document, keeps only the clean variant.
        """
        by_norm = {}
        for fname, url in file_tuples:
            stem = Path(fname).stem.lower()
            norm = re.sub(
                r'[_\-\s]*(?:tracked[_\-\s]*changes|track[_\-\s]*changes|rev[_\-\s]*marks|clean|cl|rm)$',
                '',
                stem,
                flags=re.IGNORECASE,
            )
            by_norm.setdefault(norm, []).append((fname, url))

        result = []
        for norm, items in by_norm.items():
            if len(items) == 1:
                result.append(items[0])
                continue
            clean_items = [
                (f, u) for f, u in items
                if re.search(r'(?:^|[_\-\s])(clean|cl)(?:$|[_\-\s\.])', Path(f).stem.lower())
                and "track" not in f.lower()
            ]
            non_tracked = [
                (f, u) for f, u in items
                if not re.search(r'(?:tracked|track[_\-\s]*changes|[_\-]rm$)', Path(f).stem.lower())
            ]
            if clean_items:
                result.extend(clean_items)
            elif non_tracked:
                result.extend(non_tracked)
            else:
                result.append(items[-1])
        return result

    @classmethod
    def _extract_timestamp_sort_key(cls, filename: str) -> Tuple:
        """Extracts a chronological sort key from a filename if a date/time is present."""
        stem = Path(filename).stem
        # Pattern 1: YYYYMMDD_HHMM or YYYYMMDD (e.g. RAN3_133bis_agenda_20261016_0900)
        m_iso = re.search(r'(20\d{2})(\d{2})(\d{2})(?:[_\-](\d{2,4}))?', stem)
        if m_iso:
            y, m, d, hm = m_iso.groups()
            return (int(y), int(m), int(d), int(hm or 0), filename.lower())

        # Pattern 2: MM-DD-YY or DD-MM-YY (e.g. R2_135b_Chair_Notes_10-16-26)
        m_dash = re.search(r'(\d{2})-(\d{2})-(\d{2,4})', stem)
        if m_dash:
            p1, p2, p3 = int(m_dash.group(1)), int(m_dash.group(2)), int(m_dash.group(3))
            year = p3 + 2000 if p3 < 100 else p3
            # Disambiguate MM-DD vs DD-MM
            if p1 > 12:
                day, month = p1, p2
            else:
                month, day = p1, p2
            return (year, month, day, 0, filename.lower())

        return (0, 0, 0, 0, filename.lower())

    @classmethod
    def _pick_latest_file(cls, items: List[Tuple[str, str]]) -> Tuple[str, str]:
        return max(items, key=lambda item: cls._extract_timestamp_sort_key(item[0]))

    @classmethod
    def _group_and_pick_latest_by_track(cls, items: List[Tuple[str, str]]) -> List[Tuple[str, str]]:
        """
        Groups files by logical track/author (stripping dates, timestamps, and EOM/final suffixes)
        and picks the latest file per track.
        """
        if not items:
            return []
        tracks = {}
        for fname, url in items:
            stem = Path(fname).stem.lower()
            # Normalize out dates, timestamps, and final/eom markers to identify the logical track
            track_key = re.sub(r'20\d{6}(?:[_\-]\d{2,4})?', '', stem)
            track_key = re.sub(r'\d{2}-\d{2}-\d{2,4}', '', track_key)
            track_key = re.sub(r'[_\-\s]+(?:final|eom|clean|cl|rm|v\d+|\d{4})(?=$|[_\-\s])', '', track_key)
            track_key = re.sub(r'[_\-\s]+', '_', track_key).strip('_') or "main"
            tracks.setdefault(track_key, []).append((fname, url))

        selected = []
        for track_key, group in tracks.items():
            selected.append(cls._pick_latest_file(group))
        return selected

    @classmethod
    def download_files_sync(
        cls,
        session,
        files_to_download: List[Tuple[str, str]],
        target_dir: Path,
        progress_callback=None,
        cancel_check=None,
        overwrite: bool = True,
    ) -> List[str]:
        """
        Downloads the specified `(filename, url)` tuples into `target_dir`.
        If a downloaded file is a `.zip` archive, extracts valid Word/HTML documents
        into `target_dir` and returns the clean extracted document filenames.
        """
        os.makedirs(to_extended_path(target_dir), exist_ok=True)
        downloaded: List[str] = []
        total_files = len(files_to_download)

        for idx, (filename, file_url) in enumerate(files_to_download, 1):
            if cancel_check and cancel_check():
                break

            if progress_callback:
                progress_callback(f"({idx}/{total_files}) Downloading {filename[:25]}...")
            dest_path = target_dir / filename

            # Skip download if file exists and overwrite is disabled
            if not overwrite and dest_path.exists() and dest_path.stat().st_size > 0:
                logging.info(f"[ChairNotes] Skipping already downloaded file: {dest_path}")
                if dest_path.suffix.lower() == ".zip":
                    extracted_names = cls.extract_chair_notes_zip(dest_path, target_dir)
                    if extracted_names:
                        downloaded.extend(extracted_names)
                    else:
                        downloaded.append(filename)
                else:
                    downloaded.append(filename)
                continue

            try:
                dl_resp = session.get(file_url, stream=True, timeout=60)
                dl_resp.raise_for_status()

                with open(to_extended_path(dest_path), "wb") as f:
                    for chunk in dl_resp.iter_content(chunk_size=65536):
                        if cancel_check and cancel_check():
                            return downloaded
                        if chunk:
                            f.write(chunk)

                logging.info(f"[ChairNotes] Saved: {dest_path}")

                if dest_path.suffix.lower() == ".zip":
                    extracted_names = cls.extract_chair_notes_zip(dest_path, target_dir)
                    if extracted_names:
                        downloaded.extend(extracted_names)
                    else:
                        downloaded.append(filename)
                else:
                    downloaded.append(filename)

            except Exception as dl_err:
                logging.error(f"[ChairNotes] Error downloading {filename}: {dl_err}")

        return downloaded

    @classmethod
    def extract_chair_notes_zip(cls, zip_path: Path, target_dir: Path) -> List[str]:
        """
        Extracts `.docx`, `.doc`, `.htm`, `.html` files from a Chairman's Notes `.zip` archive
        into `target_dir`. If both `_clean` and `_tracked_changes` variants exist inside the ZIP,
        returns only the clean variant(s) for downstream parsing so agreements are not duplicated.
        """
        os.makedirs(to_extended_path(target_dir), exist_ok=True)
        extracted_files: List[Tuple[str, str]] = []
        try:
            with zipfile.ZipFile(to_extended_path(zip_path), "r") as zf:
                for member in zf.infolist():
                    if member.is_dir():
                        continue
                    member_name = os.path.basename(member.filename)
                    if not member_name or member_name.startswith("~$") or member_name.startswith("."):
                        continue
                    ext = Path(member_name).suffix.lower()
                    if ext not in (".docx", ".doc", ".htm", ".html"):
                        continue

                    out_path = target_dir / member_name
                    with zf.open(member) as src, open(to_extended_path(out_path), "wb") as dst:
                        dst.write(src.read())
                    extracted_files.append((member_name, str(out_path)))
                    logging.info(f"[ChairNotes] Extracted from ZIP: {out_path}")
        except Exception as e:
            logging.warning(f"[ChairNotes] Could not extract ZIP {zip_path}: {e}")
            return []

        preferred = cls._prefer_clean_variants(extracted_files)
        return [fname for fname, _ in preferred]
