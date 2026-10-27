# --- File: src/modules/specifications/core/spec_tdoc_report.py ---
"""
Specification TDoc Metadata Report & Dossier Engine.

Resolves metadata (Title, Company/Source, Agenda Item, Type, Status, Abstract, Work Item)
for TDocs and Working Group pseudo-CRs referenced in 3GPP Specification Change History
by querying meeting agenda spreadsheets and cached JSON without downloading actual TDocs.
"""

import csv
import io
import json
import logging
import re
import time
from collections import Counter
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Sequence, Set, Tuple

from PyQt5.QtCore import QThread, pyqtSignal

from core.network.session import NetworkSession
from core.utils.paths import get_cache_root, get_project_root
from modules.meetings.core.meetings_db import MeetingsDatabase
from modules.meetings.core.tdocs_parser import TDocsParser

logger = logging.getLogger(__name__)


class SpecTDocReportGenerator:
    """Extracts TDoc metadata from meeting spreadsheets and generates technical dossier reports."""

    def __init__(
        self,
        data_db_path: Optional[Path] = None,
        cache_root: Optional[Path] = None,
        history_db: Optional[Any] = None,
    ):
        self.data_db_path = Path(data_db_path or (get_project_root() / "db" / "3gpp_data.db"))
        self.cache_root = Path(cache_root or get_cache_root())
        self.meetings_db = MeetingsDatabase(self.data_db_path)
        if history_db is not None:
            self.history_db = history_db
        else:
            try:
                from modules.specifications.core.spec_history_db import SpecHistoryDatabase
                self.history_db = SpecHistoryDatabase(get_project_root() / "db" / "3gpp_spec_history.db")
            except Exception:
                self.history_db = None
        self._meeting_cache: Dict[str, Dict[str, Dict[str, Any]]] = {}

    def _resolve_meeting_spreadsheet(self, meeting: Dict[str, Any]) -> List[Dict[str, Any]]:
        """
        Locates or downloads the official meeting TDocs spreadsheet / cached JSON.
        Returns parsed list of TDoc dictionaries.
        """
        folder_name = meeting.get("folder_name") or meeting.get("meeting_number", "unknown")
        mtg_id = meeting.get("mtg_id", "")
        agenda_dir = self.cache_root / folder_name / "Agenda"

        # 1. Check for existing parsed JSON cache
        if agenda_dir.exists():
            for json_file in agenda_dir.glob("*.xlsx.json"):
                try:
                    with open(json_file, "r", encoding="utf-8") as f:
                        data = json.load(f)
                        if data and isinstance(data, list):
                            return data
                except Exception as e:
                    logger.warning(f"Error reading cached JSON {json_file.name}: {e}")

            # 2. Check for raw Excel file to parse
            for xlsx_file in agenda_dir.glob("*.xlsx"):
                if not xlsx_file.name.startswith("~$"):
                    data = TDocsParser.parse_tdocs_excel(str(xlsx_file))
                    if data:
                        return data

        # 3. If neither exists and mtg_id is present, download spreadsheet on demand (<0.5s)
        if mtg_id:
            try:
                logger.info(f"⬇️ [TDoc Report] Fetching meeting document list for MtgId={mtg_id} ({folder_name})...")
                agenda_dir.mkdir(parents=True, exist_ok=True)
                url = f"https://portal.3gpp.org/ngppapp/GenerateDocumentList.Aspx?meetingId={mtg_id}"
                session = NetworkSession.get_instance()
                NetworkSession.apply_humanness(session)
                resp = session.get(url, stream=True, timeout=30)
                if resp.status_code == 200:
                    dest_file = agenda_dir / f"TDoc_List_Meeting_{mtg_id}.xlsx"
                    with open(dest_file, "wb") as f:
                        for chunk in resp.iter_content(chunk_size=16384):
                            if chunk:
                                f.write(chunk)
                    logger.info(f"✅ [TDoc Report] Downloaded meeting spreadsheet: {dest_file.name}")
                    return TDocsParser.parse_tdocs_excel(str(dest_file))
            except Exception as e:
                logger.warning(f"Could not download meeting document list for {mtg_id}: {e}")

        return []

    def get_meeting_tdoc_map(self, meeting: Dict[str, Any]) -> Dict[str, Dict[str, Any]]:
        """Caches and returns an $O(1)$ uppercase lookup dictionary for a meeting's TDocs."""
        mtg_key = meeting.get("folder_name") or meeting.get("meeting_number") or str(meeting.get("id"))
        if mtg_key in self._meeting_cache:
            return self._meeting_cache[mtg_key]

        rows = self._resolve_meeting_spreadsheet(meeting)
        tdoc_map: Dict[str, Dict[str, Any]] = {}
        for r in rows:
            tdoc_token = str(r.get("TDoc", "")).strip().upper()
            if tdoc_token:
                tdoc_map[tdoc_token] = r
                # Also index base token without revision
                base_token = tdoc_token.split("R")[0]
                if base_token not in tdoc_map:
                    tdoc_map[base_token] = r

        self._meeting_cache[mtg_key] = tdoc_map
        return tdoc_map

    def resolve_tdocs_metadata(
        self,
        tdoc_list: Sequence[str],
        default_meeting: str = "",
        progress_cb: Optional[Callable[[int, str], None]] = None
    ) -> List[Dict[str, Any]]:
        """
        Resolves rich metadata for a sequence of TDoc tokens from meeting spreadsheets.
        Executes without downloading actual Word documents.
        """
        cleaned_tdocs = [t.strip() for t in tdoc_list if t.strip() and t.strip() != "-"]
        if not cleaned_tdocs:
            return []

        # Deduplicate while preserving order
        seen: Set[str] = set()
        unique_tdocs: List[str] = []
        for t in cleaned_tdocs:
            if t not in seen:
                seen.add(t)
                unique_tdocs.append(t)

        total = len(unique_tdocs)
        logger.info(f"📑 [TDoc Report] Resolving metadata for {total} TDocs from meeting agendas...")

        # 0. Check SQLite persistent cache first (instant O(1))
        cached_map: Dict[str, Dict[str, Any]] = {}
        if self.history_db and hasattr(self.history_db, "get_tdoc_metadata"):
            try:
                cached_map = self.history_db.get_tdoc_metadata(unique_tdocs)
            except Exception as e:
                logger.warning(f"Error reading tdoc_metadata_cache: {e}")

        # Partition into already-cached vs. needing spreadsheet ingestion
        resolved_records: List[Dict[str, Any]] = []
        pending_tdocs: List[str] = []

        for t in unique_tdocs:
            c_row = cached_map.get(t.upper()) or cached_map.get(t.upper().split("R")[0])
            if c_row and c_row.get("title"):
                resolved_records.append({
                    "tdoc": t,
                    "title": c_row.get("title", ""),
                    "source": c_row.get("source", ""),
                    "type": c_row.get("tdoc_type", "") or ("pCR" if re.search(r"-[0-9]{5,7}", t) else "CR"),
                    "status": c_row.get("status", "").lower(),
                    "agenda_item": c_row.get("agenda_item", ""),
                    "ai_desc": c_row.get("ai_desc", ""),
                    "abstract": c_row.get("abstract", ""),
                    "contact": "",
                    "wi": c_row.get("work_item", ""),
                    "release": "",
                    "spec": "",
                    "meeting": c_row.get("meeting", "") or default_meeting,
                    "remarks": c_row.get("remarks", ""),
                    "resolved": True,
                })
            else:
                pending_tdocs.append(t)

        if not pending_tdocs:
            if progress_cb:
                progress_cb(100, f"Loaded {len(resolved_records)} TDocs from cache.")
            logger.info(f"⚡ [TDoc Report] Instant cache hit: All {len(resolved_records)} TDocs loaded from SQLite.")
            return resolved_records

        if progress_cb:
            progress_cb(5, f"Grouping {len(pending_tdocs)} TDocs by 3GPP meeting...")

        # 1. Group pending TDocs by meeting using MeetingsDatabase
        meeting_to_tdocs: Dict[str, Tuple[Dict[str, Any], List[str]]] = {}
        unmapped: List[str] = []

        for idx, tdoc in enumerate(pending_tdocs):
            mtg = self.meetings_db.find_meeting_by_tdoc(tdoc)
            if mtg:
                m_key = mtg.get("folder_name") or mtg.get("meeting_number") or "unknown"
                if m_key not in meeting_to_tdocs:
                    meeting_to_tdocs[m_key] = (mtg, [])
                meeting_to_tdocs[m_key][1].append(tdoc)
            else:
                unmapped.append(tdoc)

        newly_resolved_for_cache: List[Dict[str, Any]] = []
        step = 0
        total_groups = max(1, len(meeting_to_tdocs))

        # 2. Ingest meeting agendas and populate metadata
        for m_key, (mtg, mtg_tdocs) in meeting_to_tdocs.items():
            step += 1
            pct = 10 + int((step / total_groups) * 80)
            mtg_name = mtg.get("name") or mtg.get("meeting_number") or m_key
            if progress_cb:
                progress_cb(pct, f"Processing agenda for {mtg_name} ({len(mtg_tdocs)} TDocs)...")

            tdoc_map = self.get_meeting_tdoc_map(mtg)

            for tdoc in mtg_tdocs:
                clean_up = tdoc.upper()
                base_up = clean_up.split("R")[0]
                m_row = tdoc_map.get(clean_up) or tdoc_map.get(base_up)

                if m_row:
                    rec = {
                        "tdoc": tdoc,
                        "title": m_row.get("Title", "").strip(),
                        "source": m_row.get("Source", "").strip(),
                        "type": m_row.get("Type", "").strip(),
                        "status": m_row.get("TDoc Status", "").strip().lower(),
                        "agenda_item": m_row.get("Agenda Item", "").strip(),
                        "ai_desc": m_row.get("Agenda item description", "").strip(),
                        "abstract": m_row.get("Abstract", "").strip(),
                        "contact": m_row.get("Contact", "").strip(),
                        "wi": m_row.get("Related WIs", "").strip(),
                        "release": m_row.get("Release", "").strip(),
                        "spec": m_row.get("Spec", "").strip(),
                        "meeting": mtg.get("name") or mtg.get("meeting_number") or default_meeting,
                        "remarks": m_row.get("Secretary Remarks", "").strip(),
                        "resolved": True,
                    }
                    resolved_records.append(rec)
                    newly_resolved_for_cache.append(rec)
                else:
                    resolved_records.append({
                        "tdoc": tdoc,
                        "title": "",
                        "source": "",
                        "type": "pCR" if re.search(r"-[0-9]{5,7}", tdoc) else "CR",
                        "status": "",
                        "agenda_item": "",
                        "ai_desc": "",
                        "abstract": "",
                        "contact": "",
                        "wi": "",
                        "release": "",
                        "spec": "",
                        "meeting": mtg.get("name") or mtg.get("meeting_number") or default_meeting,
                        "remarks": "",
                        "resolved": False,
                    })

        # Append any unmapped TDocs with baseline fallback
        for tdoc in unmapped:
            resolved_records.append({
                "tdoc": tdoc,
                "title": "",
                "source": "",
                "type": "pCR" if re.search(r"-[0-9]{5,7}", tdoc) else "CR",
                "status": "",
                "agenda_item": "",
                "ai_desc": "",
                "abstract": "",
                "contact": "",
                "wi": "",
                "release": "",
                "spec": "",
                "meeting": default_meeting,
                "remarks": "",
                "resolved": False,
            })

        # 3. Store newly resolved metadata into SQLite cache
        if newly_resolved_for_cache and self.history_db and hasattr(self.history_db, "upsert_tdoc_metadata"):
            try:
                self.history_db.upsert_tdoc_metadata(newly_resolved_for_cache)
            except Exception as e:
                logger.warning(f"Could not persist to tdoc_metadata_cache: {e}")

        if progress_cb:
            progress_cb(100, f"Completed resolving {len(resolved_records)} TDocs.")

        resolved_count = sum(1 for r in resolved_records if r.get("resolved"))
        logger.info(f"✅ [TDoc Report] Successfully resolved metadata for {resolved_count}/{total} TDocs.")
        return resolved_records

    @staticmethod
    def generate_markdown_report(
        spec_number: str,
        version: str,
        tdoc_entries: Sequence[Dict[str, Any]],
        spec_title: str = ""
    ) -> str:
        """
        Synthesizes a structured Markdown technical report / dossier summarizing
        all TDocs, corporate contributions, work items, and technical solutions.
        """
        total = len(tdoc_entries)
        resolved = sum(1 for r in tdoc_entries if r.get("resolved"))
        sources = [r.get("source") for r in tdoc_entries if r.get("source")]
        types = [r.get("type") or "Unknown" for r in tdoc_entries if r.get("type")]
        statuses = [r.get("status") or "Unknown" for r in tdoc_entries if r.get("status")]

        source_counts = Counter(sources).most_common(10)
        type_counts = Counter(types).most_common(5)
        status_counts = Counter(statuses).most_common(5)

        lines: List[str] = []
        lines.append(f"# 📑 3GPP Technical Contribution Dossier: TS {spec_number} v{version}")
        if spec_title:
            lines.append(f"> **Title:** {spec_title}\n")

        lines.append("## Executive Summary")
        lines.append(f"- **Specification:** TS {spec_number}")
        lines.append(f"- **Target Version:** v{version}")
        lines.append(f"- **Total Integrated / Referenced TDocs:** {total} ({resolved} resolved with full metadata)")
        lines.append(f"- **Generated At:** {time.strftime('%Y-%m-%d %H:%M:%S')}\n")

        if source_counts:
            lines.append("### Top Contributing Organizations")
            for company, cnt in source_counts:
                lines.append(f"- **{company}:** {cnt} contribution(s)")
            lines.append("")

        if type_counts or status_counts:
            lines.append("### Distribution by Type & Status")
            type_str = ", ".join(f"`{k}`: {v}" for k, v in type_counts)
            status_str = ", ".join(f"`{k}`: {v}" for k, v in status_counts)
            if type_str:
                lines.append(f"- **Contribution Types:** {type_str}")
            if status_str:
                lines.append(f"- **Approval Status:** {status_str}")
            lines.append("")

        # -------------------------------------------------------------
        # 1. Master TDoc Summary Table
        # -------------------------------------------------------------
        lines.append("## Integrated Contributions Ledger\n")
        lines.append("| TDoc | Type | Status | Source (Company) | AI | Title / Subject |")
        lines.append("| :--- | :--- | :--- | :--- | :--- | :--- |")

        for r in tdoc_entries:
            tdoc = r.get("tdoc", "")
            ttype = r.get("type") or "pCR"
            stat = r.get("status") or "—"
            src = r.get("source") or "—"
            ai = r.get("agenda_item") or "—"
            title = r.get("title") or "—"
            # Clean title of internal pipes
            clean_title = title.replace("|", "/")
            lines.append(f"| **{tdoc}** | `{ttype}` | {stat} | {src} | {ai} | {clean_title} |")

        lines.append("")

        # -------------------------------------------------------------
        # 2. Detailed Technical Abstracts & Solutions
        # -------------------------------------------------------------
        abstract_entries = [r for r in tdoc_entries if r.get("abstract") or r.get("remarks")]
        if abstract_entries:
            lines.append("## Technical Details & Contribution Abstracts\n")
            for r in abstract_entries:
                tdoc = r.get("tdoc", "")
                title = r.get("title") or "No title"
                src = r.get("source") or "Unknown"
                ai = r.get("agenda_item") or ""
                ai_desc = r.get("ai_desc") or ""
                abstract = r.get("abstract", "")
                remarks = r.get("remarks", "")
                wi = r.get("wi", "")

                lines.append(f"### {tdoc} — {title}")
                meta_bits = []
                if src:
                    meta_bits.append(f"**Source:** {src}")
                if ai:
                    meta_bits.append(f"**AI:** {ai}{f' ({ai_desc})' if ai_desc else ''}")
                if wi:
                    meta_bits.append(f"**Work Item:** {wi}")
                if meta_bits:
                    lines.append(f"> {' | '.join(meta_bits)}\n")

                if abstract:
                    clean_abstract = "\n".join(f"> {line}" for line in abstract.splitlines())
                    lines.append(f"{clean_abstract}\n")

                if remarks:
                    lines.append(f"- **Secretary Remarks / Outcome:** {remarks}\n")

        return "\n".join(lines)

    @staticmethod
    def generate_csv_report(tdoc_entries: Sequence[Dict[str, Any]]) -> str:
        """Exports TDoc metadata to standard CSV format."""
        output = io.StringIO()
        fieldnames = [
            "TDoc", "Title", "Source", "Type", "Status",
            "Agenda Item", "Agenda item description", "Work Item",
            "Contact", "Meeting", "Abstract", "Secretary Remarks"
        ]
        writer = csv.DictWriter(output, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()

        for r in tdoc_entries:
            writer.writerow({
                "TDoc": r.get("tdoc", ""),
                "Title": r.get("title", ""),
                "Source": r.get("source", ""),
                "Type": r.get("type", ""),
                "Status": r.get("status", ""),
                "Agenda Item": r.get("agenda_item", ""),
                "Agenda item description": r.get("ai_desc", ""),
                "Work Item": r.get("wi", ""),
                "Contact": r.get("contact", ""),
                "Meeting": r.get("meeting", ""),
                "Abstract": r.get("abstract", ""),
                "Secretary Remarks": r.get("remarks", ""),
            })

        return output.getvalue()


class SpecTDocReportThread(QThread):
    """Background worker thread for resolving TDoc metadata from meeting spreadsheets."""

    progress = pyqtSignal(int, str)
    finished = pyqtSignal(list, str)
    error = pyqtSignal(str)

    def __init__(
        self,
        spec_number: str,
        version: str,
        tdoc_list: Sequence[str],
        spec_title: str = "",
        default_meeting: str = "",
        parent=None
    ):
        super().__init__(parent)
        self.spec_number = spec_number
        self.version = version
        self.tdoc_list = list(tdoc_list)
        self.spec_title = spec_title
        self.default_meeting = default_meeting

    def run(self):
        try:
            generator = SpecTDocReportGenerator()

            def _on_prog(pct: int, msg: str):
                if not self.isInterruptionRequested():
                    self.progress.emit(pct, msg)

            results = generator.resolve_tdocs_metadata(
                self.tdoc_list,
                default_meeting=self.default_meeting,
                progress_cb=_on_prog
            )

            md_report = SpecTDocReportGenerator.generate_markdown_report(
                self.spec_number,
                self.version,
                results,
                spec_title=self.spec_title
            )

            self.finished.emit(results, md_report)
        except Exception as e:
            logger.error(f"Error in SpecTDocReportThread: {e}", exc_info=True)
            self.error.emit(str(e))
