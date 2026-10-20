# --- File: modules/meetings/core/agenda_manager.py ---
import csv
import logging
import re
from pathlib import Path
from typing import Dict, Optional

from PyQt5.QtCore import QThread, pyqtSignal
from core.network.session import NetworkSession


class AgendaItem:
    def __init__(self, ai_num: str, description: str):
        self.ai_num = ai_num.strip()
        self.description = description.strip()
        self.acronym = self._extract_acronym(self.description)
        self.topic_suffix = self._extract_topic_suffix(self.description, self.acronym)

    @staticmethod
    def _extract_acronym(desc: str) -> Optional[str]:
        """Extracts WI/SI acronyms enclosed in parentheses at or near the end of the title."""
        matches = list(re.finditer(r'\(([A-Za-z0-9_\-\/]+)\)', desc))
        if not matches:
            return None

        for m in reversed(matches):
            candidate = m.group(1).strip()
            # 3GPP acronyms are 2-25 characters long and contain at least one uppercase letter or digit
            if not (2 <= len(candidate) <= 25 and any(c.isupper() or c.isdigit() for c in candidate)):
                continue

            # Skip generic notes or common phrases
            if candidate.lower().startswith(("excluding", "only", "rel-", "stage", "phase")):
                continue

            start_pos, end_pos = m.start(), m.end()
            # Ensure this parenthetical is not embedded inside a word, e.g. "(un)availability" or "(re)configuration"
            if start_pos > 0 and desc[start_pos - 1].isalpha():
                continue
            after = desc[end_pos:].strip()
            if after and after[0].isalpha():
                continue

            # In 3GPP, WI tags appear at end of title OR followed by a delimiter like ' - suffix'
            if after and not re.match(r'^[\-–—:;]', after):
                continue

            return candidate

        return None

    @staticmethod
    def _extract_topic_suffix(desc: str, acronym: Optional[str]) -> Optional[str]:
        """Return a concise sub-topic following the WI/SI acronym, when present."""
        if not acronym:
            return None
        marker = f"({acronym})"
        pos = desc.rfind(marker)
        if pos < 0:
            return None
        suffix = desc[pos + len(marker):].strip()
        suffix = re.sub(r"^[\s\-–—:;]+", "", suffix).strip()
        return suffix or None

    @property
    def display_label(self) -> str:
        """Compact label for dropdown menus."""
        if self.acronym:
            if self.topic_suffix:
                clean_sfx = re.sub(r'\s+', ' ', self.topic_suffix)
                combined = f"{self.ai_num} ({self.acronym} – {clean_sfx})"
                if len(combined) > 42:
                    avail = max(10, 35 - len(self.ai_num) - len(self.acronym))
                    return f"{self.ai_num} ({self.acronym} – {clean_sfx[:avail].rstrip()}...)"
                return combined
            return f"{self.ai_num} ({self.acronym})"

        clean = re.sub(r'\s+', ' ', self.description)
        max_desc_len = max(15, 38 - len(self.ai_num))
        if len(clean) > max_desc_len:
            return f"{self.ai_num} - {clean[:max_desc_len - 3].rstrip()}..."
        return f"{self.ai_num} - {clean}"

    @property
    def full_tooltip(self) -> str:
        """Full unabbreviated text for Qt.ToolTipRole."""
        return f"AI {self.ai_num}: {self.description}"


class AgendaManager:
    """Handles local loading of agenda.csv without triggering any network activity."""

    @classmethod
    def load_local_agenda(cls, agenda_dir: Path) -> Dict[str, AgendaItem]:
        """
        Loads mapping strictly from the local file system.
        Returns an empty dict if the file does not exist locally.
        """
        agenda_dir = Path(agenda_dir)
        candidates = [agenda_dir / "agenda.csv", agenda_dir / "Agenda.csv"]
        csv_path = next((p for p in candidates if p.is_file()), None)

        if not csv_path:
            logging.debug(f"[Agenda] No local agenda.csv in {agenda_dir}")
            return {}

        mapping = {}
        for encoding in ("utf-8-sig", "utf-8", "latin1", "cp1252"):
            try:
                with open(csv_path, mode="r", encoding=encoding, errors="replace") as f:
                    reader = csv.reader(f)
                    for row in reader:
                        if not row or len(row) < 2:
                            continue
                        ai_num = row[0].strip()
                        ai_desc = row[1].strip()
                        if ai_num:
                            mapping[ai_num] = AgendaItem(ai_num, ai_desc)
                logging.info(f"[Agenda] Loaded {len(mapping)} AIs from local cache: {csv_path}")
                break
            except Exception as e:
                logging.debug(f"[Agenda] Failed loading {csv_path} with {encoding}: {e}")

        return mapping


class AgendaDownloaderThread(QThread):
    """
    Manually triggered asynchronous worker to download agenda.csv from 3GPP servers.
    """
    progress = pyqtSignal(str)
    # Emits: (success: bool, message: str)
    finished = pyqtSignal(bool, str)

    def __init__(self, candidate_base_urls: list, target_agenda_dir: Path, parent=None):
        super().__init__(parent)
        self.candidate_base_urls = candidate_base_urls
        self.target_agenda_dir = Path(target_agenda_dir)

    def run(self):
        try:
            self.target_agenda_dir.mkdir(parents=True, exist_ok=True)
            dest_file = self.target_agenda_dir / "agenda.csv"

            session = NetworkSession.get_instance()
            NetworkSession.apply_humanness(session)

            # Build prioritized probes for the remote Agenda folder
            probe_urls = []
            for base in self.candidate_base_urls:
                if not base:
                    continue
                clean = base.rstrip('/')
                probe_urls.extend([
                    f"{clean}/Agenda/agenda.csv",
                    f"{clean}/Agenda/Agenda.csv",
                    f"{clean}/agenda.csv",
                    f"{clean}/Agenda/agenda.CSV"
                ])

            # Deduplicate preserving order
            seen = set()
            probe_urls = [u for u in probe_urls if not (u in seen or seen.add(u))]

            downloaded = False
            for url in probe_urls:
                try:
                    self.progress.emit(f"Probing {url.split('/')[-2]}/{url.split('/')[-1]}...")
                    resp = session.get(url, stream=True, timeout=15)
                    if resp.status_code == 200 and int(resp.headers.get("content-length", 1)) > 0:
                        with open(dest_file, "wb") as f:
                            for chunk in resp.iter_content(chunk_size=32768):
                                if chunk:
                                    f.write(chunk)
                        downloaded = True
                        logging.info(f"[Agenda] Successfully saved agenda.csv from: {url}")
                        break
                except Exception as probe_err:
                    logging.debug(f"[Agenda] Failed {url}: {probe_err}")

            if downloaded:
                self.finished.emit(True, f"agenda.csv successfully saved to:\n{dest_file}")
            else:
                self.finished.emit(False, "No agenda.csv found on remote 3GPP server for this meeting.")

        except Exception as ex:
            logging.error(f"[Agenda] Download error: {ex}", exc_info=True)
            self.finished.emit(False, f"Error downloading agenda.csv: {ex}")