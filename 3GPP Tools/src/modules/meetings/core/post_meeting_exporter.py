# --- File: src/modules/meetings/core/post_meeting_exporter.py ---
from concurrent.futures import ThreadPoolExecutor, as_completed
import datetime
import logging
import os
import re
import shutil
from pathlib import Path
from typing import Dict, List, Optional, Set, Any, Tuple

from PyQt5.QtCore import QThread, pyqtSignal

from core.utils.paths import to_extended_path, get_project_root
from modules.meetings.core.agenda_manager import AgendaItem
from modules.meetings.core.tdoc_file_handler import TDocFileHandler
from modules.meetings.core.tdoc_url_resolver import resolve_tdoc_download_url
from modules.meetings.core.tdoc_markdown_converter import (
    convert_document_to_markdown,
    sanitize_filename,
    is_conversion_stale,
    resolve_revision_mode_for_type,
    CONVERTER_VERSION,
)
from modules.meetings.ui.tdocs_models import TDocsTableModel
from modules.work_items.core.wi_database import WorkItemsDatabase
from modules.work_items.core.wi_settings import WorkItemsSettings

logger = logging.getLogger(__name__)


def sanitize_folder_name(name: str) -> str:
    """Sanitizes directory names, removing illegal Windows/filesystem characters."""
    clean = re.sub(r'[\\/*?:"<>|]', "_", str(name or "")).strip()
    return re.sub(r'_+', '_', clean).strip("._ ") or "General"


class PostMeetingExporterThread(QThread):
    """
    Asynchronous worker to generate an organized, executive-grade Post-Meeting Markdown Dossier.
    Structure:
      Export/PostMeeting_<WG>_<Meeting>_<Timestamp>/
        ├── index.md (Master Cover & Summary Tables)
        ├── <Release>/
        │   └── <WorkItem>/
        │       ├── index.md (WI Specific Dossier Cover)
        │       ├── LS_IN/ (S2-xxxxxx.md + media/)
        │       ├── pCR/   (S2-xxxxxx.md + media/)
        │       ├── CR/    (S2-xxxxxx.md + media/)
        │       └── LS_OUT/(S2-xxxxxx.md + media/)
        └── General/
            └── LS_IN/ (Overarching incoming LS without a specific WI)
    """
    progress = pyqtSignal(str)
    finished = pyqtSignal(bool, str)

    def __init__(
        self,
        meeting_dir: Path,
        all_rows: list,
        visible_rows: list,
        agenda_map: dict,
        mtg_info: dict,
        docs_ftp_url: str,
        revisions_url: str,
        scope_mode: str = "followed",  # "followed", "starred", "tracked", "visible", "whole", "custom"
        custom_wis: Optional[Set[str]] = None,
        custom_ais: Optional[Set[str]] = None,
        custom_wi_ais: Optional[Dict[str, Set[str]]] = None,
        followed_tokens: Optional[Set[str]] = None,
        starred_tokens: Optional[Set[str]] = None,
        include_ls_in: bool = True,
        include_pcr: bool = True,
        include_cr: bool = True,
        include_ls_out: bool = True,
        include_disc: bool = False,
        include_general_ls: bool = True,
        generate_local_indices: bool = True,
        download_missing: bool = True,
        force_reconvert: bool = False,
        output_dir: Optional[Path] = None,
        parent=None,
    ):
        super().__init__(parent)
        self.meeting_dir = Path(meeting_dir)
        self.all_rows = all_rows or []
        self.visible_rows = visible_rows or []
        self.agenda_map = agenda_map or {}
        self.mtg_info = mtg_info or {}
        self.docs_ftp_url = docs_ftp_url or ""
        self.revisions_url = revisions_url or ""
        self.scope_mode = scope_mode
        self.custom_wis = {str(w).strip().lower() for w in (custom_wis or set()) if str(w).strip()}
        self.custom_ais = {str(a).strip() for a in (custom_ais or set()) if str(a).strip()}
        self.custom_wi_ais: Dict[str, Set[str]] = {}
        if custom_wi_ais:
            for k, v in custom_wi_ais.items():
                self.custom_wi_ais[str(k).strip().lower()] = {str(a).strip() for a in (v or set()) if str(a).strip()}
        self.followed_tokens = {str(t).strip().lower() for t in (followed_tokens or set()) if str(t).strip()}
        self.starred_tokens = {str(t).strip().lower() for t in (starred_tokens or set()) if str(t).strip()}
        self.include_ls_in = include_ls_in
        self.include_pcr = include_pcr
        self.include_cr = include_cr
        self.include_ls_out = include_ls_out
        self.include_disc = include_disc
        self.include_general_ls = include_general_ls
        self.generate_local_indices = generate_local_indices
        self.download_missing = download_missing
        self.force_reconvert = force_reconvert

        # Resolve output directory
        if output_dir:
            self.export_dir = Path(output_dir)
        else:
            timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S")
            wg = str(self.mtg_info.get("wg_name", "3GPP")).strip().replace(" ", "_")
            mtg = str(self.mtg_info.get("meeting_number", "")).strip()
            folder_name = f"PostMeeting_{wg}_{mtg}_{timestamp}" if (wg or mtg) else f"PostMeeting_{timestamp}"
            self.export_dir = self.meeting_dir / "Export" / folder_name

    def run(self):
        try:
            self.progress.emit("Initializing Work Items knowledge base...")
            wi_db_info = self._load_work_items_info()

            self.progress.emit("Categorizing and filtering post-meeting TDocs...")
            filtered_entries = self._filter_and_group_tdocs(wi_db_info)

            if not filtered_entries:
                self.finished.emit(False, "No TDocs match the selected export criteria.")
                return

            total_entries = sum(len(docs) for rel in filtered_entries.values() for wi in rel.values() for docs in wi["categories"].values())
            self.progress.emit(f"Found {total_entries} matching contribution instance(s). Exporting dossiers...")

            os.makedirs(to_extended_path(self.export_dir), exist_ok=True)

            # Convert and assemble all dossiers
            self._export_all_dossiers(filtered_entries)

            if self.isInterruptionRequested():
                self.finished.emit(False, "Export cancelled by user.")
                return

            # Generate index.md files
            self.progress.emit("Generating Markdown index reports...")
            if self.generate_local_indices:
                self._generate_local_wi_indices(filtered_entries)

            self._generate_master_index(filtered_entries)

            msg = f"Successfully generated Post-Meeting Dossier in:\n{self.export_dir}"
            self.finished.emit(True, msg)

        except Exception as e:
            logger.error(f"[PostMeetingExporter] Critical failure during export: {e}", exc_info=True)
            self.finished.emit(False, str(e))

    def _build_ls_reply_map(self) -> tuple:
        """
        Scans all meeting rows to build bidirectional reply mappings:
        - in_to_replies: incoming_tdoc_id -> set of final approved/agreed reply tdoc_ids
        - out_to_in: outgoing_tdoc_id -> set of incoming tdoc_ids
        - tdoc_map: tdoc_id -> row_dict
        Follows revision chains ('Revised to') to identify final approved reply TDocs.
        """
        tdoc_map = {str(r.get("TDoc", "")).strip().upper(): r for r in self.all_rows if r.get("TDoc")}

        def _resolve_final(tdoc_id: str) -> str:
            curr = tdoc_id.upper()
            visited = set()
            while curr and curr not in visited:
                visited.add(curr)
                row = tdoc_map.get(curr)
                if not row:
                    break
                rev = str(row.get("Revised to", "")).strip()
                if rev:
                    matches = re.findall(r'[A-Za-z0-9]+-\d+', rev)
                    if matches:
                        curr = matches[-1].upper()
                        continue
                break
            return curr

        in_to_replies: Dict[str, Set[str]] = {}
        out_to_in: Dict[str, Set[str]] = {}

        for r in self.all_rows:
            tdoc = str(r.get("TDoc", "")).strip().upper()
            if not tdoc:
                continue
            rep_in = str(r.get("Reply in", "")).strip()
            rep_to = str(r.get("Reply to", "")).strip()

            if rep_in:
                targets = [_resolve_final(x) for x in re.findall(r'[A-Za-z0-9]+-\d+', rep_in)]
                for tgt in targets:
                    in_to_replies.setdefault(tdoc, set()).add(tgt)
                    out_to_in.setdefault(tgt, set()).add(tdoc)

            if rep_to:
                incomings = [x.upper() for x in re.findall(r'[A-Za-z0-9]+-\d+', rep_to)]
                final_out = _resolve_final(tdoc)
                for inc in incomings:
                    in_to_replies.setdefault(inc, set()).add(final_out)
                    out_to_in.setdefault(final_out, set()).add(inc)

        return in_to_replies, out_to_in, tdoc_map

    def _load_work_items_info(self) -> dict:
        """Loads Work Items records and tracking tokens from settings and 3gpp_data.db."""
        info = {
            "wi_by_exact_acronym": {}, # acronym.lower() -> item_dict
            "wi_by_code": {},          # code.lower() -> item_dict
            "wi_by_token": {},         # token -> item_dict
            "followed_tokens": set(),
            "starred_tokens": set(),
            "tracked_tokens": set(),
        }
        if self.followed_tokens:
            info["followed_tokens"] = set(self.followed_tokens)
        if self.starred_tokens:
            info["starred_tokens"] = set(self.starred_tokens)
        info["tracked_tokens"] = info["followed_tokens"] | info["starred_tokens"]

        try:
            db_path = get_project_root() / "db" / "3gpp_data.db"
            if db_path.exists():
                wi_db = WorkItemsDatabase(db_path)
                if not info["followed_tokens"] and not info["starred_tokens"]:
                    settings = WorkItemsSettings()
                    starred_codes, following_codes = settings.get_tracking_sets()
                    info["followed_tokens"] = wi_db.get_tracked_wi_tokens(following_codes)
                    info["starred_tokens"] = wi_db.get_tracked_wi_tokens(starred_codes)
                    info["tracked_tokens"] = info["followed_tokens"] | info["starred_tokens"]

                # Build token lookup map for all known work items
                with wi_db._get_connection() as conn:
                    cursor = conn.cursor()
                    cursor.execute("SELECT code, acronym, name, release FROM work_items")
                    for row in cursor.fetchall():
                        code, acronym, name, release = row
                        item_dict = {
                            "code": str(code or "").strip(),
                            "acronym": str(acronym or "").strip(),
                            "name": str(name or "").strip(),
                            "release": str(release or "").strip(),
                        }
                        acr = item_dict["acronym"]
                        if acr:
                            info["wi_by_exact_acronym"][acr.lower()] = item_dict
                            tokens = TDocsTableModel._extract_base_tokens(acr)
                            tokens.add(acr.lower())
                            if item_dict["code"]:
                                info["wi_by_code"][item_dict["code"].lower()] = item_dict
                                tokens.add(item_dict["code"].lower())
                            for t in tokens:
                                if t not in info["wi_by_token"]:
                                    info["wi_by_token"][t] = item_dict
        except Exception as e:
            logger.warning(f"[PostMeetingExporter] Could not fully load Work Items metadata: {e}")

        return info

    def _filter_and_group_tdocs(self, wi_db_info: dict) -> dict:
        """
        Filters all meeting TDocs according to post-meeting rules:
        - Only incoming LS (LS in) and agreed/approved contributions (pCR, CR, LS out).
        - Groups by Release -> Work Item -> Category.
        - Identifies cross-WI overlaps.
        - Connects incoming LS to approved outgoing replies (and guarantees reply export).
        """
        # Data structure:
        # tree[release][wi_acronym] = {
        #     "info": {...},
        #     "ais": set(),
        #     "ai_titles": {},
        #     "categories": {"LS_IN": [], "pCR": [], "CR": [], "LS_OUT": []}
        # }
        tree: Dict[str, Dict[str, Any]] = {}

        in_to_replies, out_to_in, tdoc_map = self._build_ls_reply_map()
        visible_tdoc_ids = {str(r.get("TDoc", "")).strip().upper() for r in self.visible_rows if r.get("TDoc")}

        for row in self.all_rows:
            tdoc_id = str(row.get("TDoc", "")).strip().upper()
            if not tdoc_id:
                continue

            doc_type = str(row.get("Type", "")).strip().lower()
            status = str(row.get("TDoc Status", "")).strip().lower()

            if "withdrawn" in status:
                continue

            is_not_agreed = any(neg in status for neg in ["not agreed", "not approved", "rejected", "postponed", "withdrawn"])
            is_agreed = (not is_not_agreed) and any(s in status for s in ["agreed", "approved"])

            category = None
            if "ls in" in doc_type or doc_type == "ls":
                if self.include_ls_in:
                    category = "LS_IN"
            elif "ls out" in doc_type or "ls_out" in doc_type:
                if self.include_ls_out and is_agreed:
                    category = "LS_OUT"
            elif "pcr" in doc_type or "p-cr" in doc_type:
                if self.include_pcr and is_agreed:
                    category = "pCR"
            elif "cr" in doc_type:
                if self.include_cr and is_agreed:
                    category = "CR"
            elif any(d in doc_type for d in ["disc", "discussion", "report", "wid", "sid"]):
                is_not_final = any(neg in status for neg in ["not agreed", "not approved", "rejected", "postponed", "withdrawn", "revised", "not discussed"])
                if self.include_disc and (not is_not_final) and (status == "" or any(s in status for s in ["noted", "agreed", "approved"])):
                    category = "DISC"

            if not category:
                continue

            if self.scope_mode == "visible" and tdoc_id not in visible_tdoc_ids:
                continue

            # Check if official Related WIs had an assigned WI
            raw_wis = TDocsTableModel._split_related_wis(row.get("Related WIs", ""))
            clean_wis = [w for w in raw_wis if w and w.lower() not in {"(no wi)", "none", "nan", "-"}]
            has_official_wi = bool(clean_wis)

            # Resolve associated Work Items
            associated_wis = self._resolve_tdoc_wis(row, wi_db_info)

            ai_num = str(row.get("Agenda Item", "")).strip() or "Unassigned"
            ai_desc = str(row.get("Agenda item description", "")).strip()

            if not associated_wis:
                # Non-WI contribution
                if (category in ("LS_IN", "LS_OUT")) and self.include_general_ls:
                    associated_wis = ["General"]
                elif self.scope_mode in ["whole", "custom"]:
                    associated_wis = ["General"]
                else:
                    continue

            # Scope filtering
            matching_wis = []
            for wi in associated_wis:
                if wi == "General":
                    if self.scope_mode in ["whole", "visible"] or ((category in ("LS_IN", "LS_OUT")) and self.include_general_ls):
                        matching_wis.append(wi)
                    elif self.scope_mode == "custom":
                        if "general" in self.custom_wis or (self.custom_ais and ai_num in self.custom_ais):
                            matching_wis.append(wi)
                        elif self.custom_wi_ais and "general" in self.custom_wi_ais and ai_num in self.custom_wi_ais["general"]:
                            matching_wis.append(wi)
                    continue

                wi_tokens = TDocsTableModel._extract_base_tokens(wi)
                if self.scope_mode == "followed":
                    if wi_tokens & wi_db_info["followed_tokens"]:
                        matching_wis.append(wi)
                elif self.scope_mode == "starred":
                    if wi_tokens & wi_db_info["starred_tokens"]:
                        matching_wis.append(wi)
                elif self.scope_mode == "tracked":
                    if wi_tokens & wi_db_info["tracked_tokens"]:
                        matching_wis.append(wi)
                elif self.scope_mode == "custom":
                    wi_match = any(t in self.custom_wis for t in wi_tokens) or (wi.lower() in self.custom_wis)
                    if wi_match:
                        matched_ai_set = None
                        if self.custom_wi_ais:
                            if wi.lower() in self.custom_wi_ais:
                                matched_ai_set = self.custom_wi_ais[wi.lower()]
                            else:
                                for t in wi_tokens:
                                    if t in self.custom_wi_ais:
                                        matched_ai_set = self.custom_wi_ais[t]
                                        break
                        if matched_ai_set is not None:
                            if ai_num in matched_ai_set:
                                matching_wis.append(wi)
                        else:
                            matching_wis.append(wi)
                    elif self.custom_ais and ai_num in self.custom_ais:
                        if wi.lower() == "general" or not self.custom_wis:
                            matching_wis.append(wi)
                else:  # "whole" or "visible"
                    matching_wis.append(wi)

            # If this is an LS IN or LS OUT with no official WI assigned in Related WIs,
            # and it didn't match any followed/selected WI in matching_wis,
            # but self.include_general_ls is enabled:
            # Fall back to "General" so unassigned LS documents are preserved!
            if (category in ("LS_IN", "LS_OUT")) and self.include_general_ls and not has_official_wi and not matching_wis:
                matching_wis.append("General")
                if "General" not in associated_wis:
                    associated_wis.append("General")

            if not matching_wis:
                continue

            # Resolve release
            release = self._resolve_tdoc_release(row, matching_wis, wi_db_info)

            # Insert into tree (handle cross-WI duplication)
            for target_wi in matching_wis:
                other_wis = [w for w in associated_wis if w != target_wi]

                if release not in tree:
                    tree[release] = {}
                if target_wi not in tree[release]:
                    tree[release][target_wi] = {
                        "info": self._get_wi_display_info(target_wi, wi_db_info, release),
                        "ais": set(),
                        "ai_titles": {},
                        "categories": {"LS_IN": [], "pCR": [], "CR": [], "LS_OUT": [], "DISC": []},
                    }

                tree[release][target_wi]["ais"].add(ai_num)
                if ai_num and ai_desc and ai_num not in tree[release][target_wi]["ai_titles"]:
                    tree[release][target_wi]["ai_titles"][ai_num] = ai_desc

                tdoc_url = resolve_tdoc_download_url(
                    tdoc_id, docs_url=self.docs_ftp_url, revisions_url=self.revisions_url, meta=row
                )
                entry = {
                    "tdoc_id": tdoc_id,
                    "title": str(row.get("Title", "")).strip(),
                    "source": str(row.get("Source", "")).strip(),
                    "type": str(row.get("Type", "")).strip(),
                    "status": str(row.get("TDoc Status", "")).strip(),
                    "ai": ai_num,
                    "spec": str(row.get("Spec", "")).strip(),
                    "category": category,
                    "cross_wis": other_wis,
                    "replies": sorted(list(in_to_replies.get(tdoc_id, set()))),
                    "reply_to_ids": sorted(list(out_to_in.get(tdoc_id, set()))),
                    "row_data": row,
                    "tdoc_url": tdoc_url,
                }
                tree[release][target_wi]["categories"][category].append(entry)

        # Post-pass: Guarantee that any approved reply LS OUT answering an included LS IN
        # is present in the export. If its WI was not in the export scope (or unassigned),
        # attach it under General/LS_OUT so the reply dossier is created without cluttering
        # the report with unwanted WI sections.
        all_included_tdocs = {
            doc["tdoc_id"]
            for rel_data in tree.values()
            for wi_data in rel_data.values()
            for doc_list in wi_data["categories"].values()
            for doc in doc_list
        }

        needed_replies = set()
        for rel_data in tree.values():
            for wi_data in rel_data.values():
                for ls_in_doc in wi_data["categories"].get("LS_IN", []):
                    for rep_id in ls_in_doc.get("replies", []):
                        if rep_id not in all_included_tdocs:
                            needed_replies.add(rep_id)

        for rep_id in sorted(needed_replies):
            rep_row = tdoc_map.get(rep_id)
            if not rep_row:
                continue

            # Only export if agreed/approved (or not withdrawn)
            rep_status = str(rep_row.get("TDoc Status", "")).strip().lower()
            if any(neg in rep_status for neg in ["withdrawn", "rejected", "not approved", "not agreed"]):
                continue

            rep_rel = self._resolve_tdoc_release(rep_row, ["General"], wi_db_info)
            target_wi = "General"

            if rep_rel not in tree:
                tree[rep_rel] = {}
            if target_wi not in tree[rep_rel]:
                tree[rep_rel][target_wi] = {
                    "info": self._get_wi_display_info(target_wi, wi_db_info, rep_rel),
                    "ais": set(),
                    "ai_titles": {},
                    "categories": {"LS_IN": [], "pCR": [], "CR": [], "LS_OUT": [], "DISC": []},
                }

            rep_ai = str(rep_row.get("Agenda Item", "")).strip() or "Unassigned"
            rep_ai_desc = str(rep_row.get("Agenda item description", "")).strip()
            tree[rep_rel][target_wi]["ais"].add(rep_ai)
            if rep_ai and rep_ai_desc and rep_ai not in tree[rep_rel][target_wi]["ai_titles"]:
                tree[rep_rel][target_wi]["ai_titles"][rep_ai] = rep_ai_desc

            raw_wis = TDocsTableModel._split_related_wis(rep_row.get("Related WIs", ""))
            clean_wis = [w for w in raw_wis if w and w.lower() not in {"(no wi)", "none", "nan", "-"}]

            rep_url = resolve_tdoc_download_url(
                rep_id, docs_url=self.docs_ftp_url, revisions_url=self.revisions_url, meta=rep_row
            )
            rep_entry = {
                "tdoc_id": rep_id,
                "title": str(rep_row.get("Title", "")).strip(),
                "source": str(rep_row.get("Source", "")).strip(),
                "type": str(rep_row.get("Type", "")).strip(),
                "status": str(rep_row.get("TDoc Status", "")).strip(),
                "ai": rep_ai,
                "spec": str(rep_row.get("Spec", "")).strip(),
                "category": "LS_OUT",
                "cross_wis": clean_wis,  # Mention original WIs in cross-WI column
                "replies": sorted(list(in_to_replies.get(rep_id, set()))),
                "reply_to_ids": sorted(list(out_to_in.get(rep_id, set()))),
                "row_data": rep_row,
                "tdoc_url": rep_url,
            }
            tree[rep_rel][target_wi]["categories"]["LS_OUT"].append(rep_entry)
            all_included_tdocs.add(rep_id)

        # Sort entries
        for rel in tree.values():
            for wi in rel.values():
                for cat in wi["categories"]:
                    wi["categories"][cat].sort(key=lambda x: str(x.get("tdoc_id", "")))

        return tree

    def _resolve_tdoc_wis(self, row: dict, wi_db_info: dict) -> List[str]:
        """Extracts canonical Work Item acronyms associated with a TDoc."""
        wis = []
        raw_wis = TDocsTableModel._split_related_wis(row.get("Related WIs", ""))
        for raw in raw_wis:
            clean = str(raw or "").strip()
            if not clean or clean.lower() in {"(no wi)", "none", "nan", "-"}:
                continue
            canonical = self._canonicalize_wi_name(clean, wi_db_info)
            if canonical and canonical not in wis:
                wis.append(canonical)

        if not wis:
            ai_num = str(row.get("Agenda Item", "")).strip()
            if ai_num and self.agenda_map:
                agenda_item = self.agenda_map.get(ai_num)
                if agenda_item and getattr(agenda_item, "acronym", None):
                    canonical = self._canonicalize_wi_name(agenda_item.acronym, wi_db_info)
                    if canonical and canonical not in wis:
                        wis.append(canonical)

            desc = str(row.get("Agenda item description", "")).strip()
            if desc:
                candidate = AgendaItem._extract_acronym(desc)
                if not candidate:
                    match = re.search(r'\(([A-Za-z0-9_\-\/]+)\)', desc)
                    if match:
                        cand = match.group(1).strip()
                        if 2 <= len(cand) <= 25 and any(c.isupper() or c.isdigit() for c in cand):
                            candidate = cand
                if candidate:
                    canonical = self._canonicalize_wi_name(candidate, wi_db_info)
                    if canonical and canonical not in wis:
                        wis.append(canonical)

        return wis

    def _canonicalize_wi_name(self, name: str, wi_db_info: dict) -> str:
        """Finds the best matching canonical Work Item acronym from DB or cleans the string."""
        clean = str(name or "").strip()
        if not clean:
            return ""

        by_exact = wi_db_info.get("wi_by_exact_acronym", {})
        by_code = wi_db_info.get("wi_by_code", {})

        # 1. Exact match on raw clean name (case-insensitive)
        if clean.lower() in by_exact:
            return by_exact[clean.lower()]["acronym"]

        # 2. Check by code
        if clean.lower() in by_code:
            return by_code[clean.lower()]["acronym"]

        # 3. Strip WG suffixes (e.g., -ARC, -CT, _SEC, -Core)
        base = re.sub(
            r'[-_](?:ARC|CT|Core|Perf|SEC|RAN|SA\d?|CT\d?|UEConTest|ServConTest)$',
            '', clean, flags=re.IGNORECASE
        ).strip()
        if base and base.lower() in by_exact:
            return by_exact[base.lower()]["acronym"]

        # 4. Try toggling FS_ prefix on clean and base
        variations = []
        for cand in [clean, base]:
            if not cand:
                continue
            if cand.upper().startswith("FS_"):
                variations.append(cand[3:])
            else:
                variations.append(f"FS_{cand}")
        for v in variations:
            if v.lower() in by_exact:
                return by_exact[v.lower()]["acronym"]

        # 5. If not in DB, return clean stripped base
        return base or clean

    def _resolve_tdoc_release(self, row: dict, matching_wis: List[str], wi_db_info: dict) -> str:
        """Determines the 3GPP Release for this document or its Work Items."""
        raw_rel = str(row.get("Release", "")).strip()
        if raw_rel:
            m = re.search(r'(Rel[-_]?\d{1,2})', raw_rel, re.IGNORECASE)
            if m:
                return m.group(1).replace("_", "-").capitalize()

        # Check WI metadata in DB
        for wi in matching_wis:
            if wi == "General":
                continue
            if wi.lower() in wi_db_info.get("wi_by_exact_acronym", {}):
                db_rel = wi_db_info["wi_by_exact_acronym"][wi.lower()].get("release")
                if db_rel:
                    return str(db_rel).strip()
            tokens = TDocsTableModel._extract_base_tokens(wi)
            for t in tokens:
                db_item = wi_db_info.get("wi_by_token", {}).get(t)
                if db_item and db_item.get("release"):
                    return str(db_item["release"]).strip()

        # Check Agenda item description
        desc = str(row.get("Agenda item description", ""))
        m = re.search(r'(Rel[-_]?\d{1,2})', desc, re.IGNORECASE)
        if m:
            return m.group(1).replace("_", "-").capitalize()

        if "General" in matching_wis:
            return "General"

        return "Unassigned"

    def _get_wi_display_info(self, wi_name: str, wi_db_info: dict, release: str) -> dict:
        """Returns structured metadata for a Work Item."""
        if wi_name == "General":
            return {
                "acronym": "General",
                "name": "General & Overarching Contributions (Non-WI)",
                "code": "",
                "release": release,
            }

        # Check exact acronym first
        by_exact = wi_db_info.get("wi_by_exact_acronym", {})
        if wi_name.lower() in by_exact:
            return by_exact[wi_name.lower()]

        # Check base without suffix
        base = re.sub(
            r'[-_](?:ARC|CT|Core|Perf|SEC|RAN|SA\d?|CT\d?|UEConTest|ServConTest)$',
            '', wi_name, flags=re.IGNORECASE
        ).strip()
        if base and base.lower() in by_exact:
            return by_exact[base.lower()]

        # Fallback to token lookup
        tokens = TDocsTableModel._extract_base_tokens(wi_name)
        for t in tokens:
            if t in wi_db_info.get("wi_by_token", {}):
                return wi_db_info["wi_by_token"][t]

        return {
            "acronym": wi_name,
            "name": wi_name,
            "code": "",
            "release": release,
        }

    def _export_all_dossiers(self, filtered_entries: dict):
        """Converts and copies documents into their respective type subfolders."""
        # Find all unique tdocs across all categories
        unique_tdocs = {}
        for rel in filtered_entries.values():
            for wi in rel.values():
                for cat_docs in wi["categories"].values():
                    for doc in cat_docs:
                        tdoc_id = doc["tdoc_id"]
                        if tdoc_id not in unique_tdocs:
                            unique_tdocs[tdoc_id] = doc["row_data"]

        total_unique = len(unique_tdocs)
        converted_cache = {}  # tdoc_id -> local_md_file

        # Phase 0: Parallel pre-download of missing documents
        if self.download_missing:
            missing_dl_tasks: List[Tuple[str, Path]] = []
            for tdoc_id, row_data in unique_tdocs.items():
                base_match = re.search(r'^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$', tdoc_id, re.IGNORECASE)
                base_tdoc = base_match.group(1).upper() if base_match else tdoc_id.upper()
                tdoc_folder = self.meeting_dir / base_tdoc
                local_md_file = tdoc_folder / f"{sanitize_filename(tdoc_id)}.md"
                legacy_cache_file = tdoc_folder / f"{tdoc_id}_LLM_v{CONVERTER_VERSION}.md"
                doc_path = self._find_word_doc(tdoc_folder, tdoc_id)
                doc_type = row_data.get("Type", "") if row_data else ""
                expected_mode = resolve_revision_mode_for_type(doc_type, default_mode="markup")

                needs_conversion = self.force_reconvert or is_conversion_stale(
                    doc_path, local_md_file, expected_version=CONVERTER_VERSION, expected_mode=expected_mode
                )
                if needs_conversion and not self.force_reconvert and os.path.exists(to_extended_path(legacy_cache_file)):
                    if not is_conversion_stale(doc_path, legacy_cache_file, expected_version=CONVERTER_VERSION, expected_mode=expected_mode):
                        needs_conversion = False

                if needs_conversion and not doc_path:
                    missing_dl_tasks.append((tdoc_id, tdoc_folder))

            if missing_dl_tasks:
                total_dl = len(missing_dl_tasks)
                dl_workers = min(6, total_dl)
                self.progress.emit(f"Downloading {total_dl} missing TDoc(s) in parallel ({dl_workers} workers)...")
                logger.info(f"[PostMeetingExporter] Concurrently downloading {total_dl} missing TDocs using {dl_workers} workers...")
                dl_done = 0
                with ThreadPoolExecutor(max_workers=dl_workers) as executor:
                    future_to_tid = {
                        executor.submit(self._download_and_extract_tdoc, tid, folder): tid
                        for tid, folder in missing_dl_tasks
                    }
                    for future in as_completed(future_to_tid):
                        if self.isInterruptionRequested():
                            try:
                                executor.shutdown(wait=False, cancel_futures=True)
                            except TypeError:
                                executor.shutdown(wait=False)
                            return
                        tid = future_to_tid[future]
                        dl_done += 1
                        try:
                            res_path = future.result()
                            if res_path:
                                self.progress.emit(f"Downloaded TDoc {tid} ({dl_done}/{total_dl})")
                            else:
                                self.progress.emit(f"Failed to download {tid} ({dl_done}/{total_dl})")
                        except Exception as dl_err:
                            logger.error(f"[PostMeetingExporter] Error downloading {tid}: {dl_err}")

        # Phase 1: Ensure all canonical Markdown files exist in meeting cache
        for idx, (tdoc_id, row_data) in enumerate(unique_tdocs.items(), 1):
            if self.isInterruptionRequested():
                return

            self.progress.emit(f"Converting {tdoc_id} to Markdown ({idx}/{total_unique})...")
            local_md_file, local_media_dir = self._ensure_tdoc_markdown(tdoc_id, row_data)
            converted_cache[tdoc_id] = (local_md_file, local_media_dir)

        # Pre-register target paths for all documents across the dossier structure
        self.exported_tdoc_paths: Dict[str, Path] = {}
        self.tdoc_titles: Dict[str, str] = {}

        for rel_name, wis in filtered_entries.items():
            rel_slug = sanitize_folder_name(rel_name)
            for wi_name, wi_data in wis.items():
                wi_slug = sanitize_folder_name(wi_name)
                for cat_name, cat_docs in wi_data["categories"].items():
                    cat_folder = "Discussion" if cat_name == "DISC" else cat_name
                    type_dir = self.export_dir / rel_slug / wi_slug / cat_folder
                    for doc in cat_docs:
                        tdoc_id = doc["tdoc_id"]
                        target_md_file = type_dir / f"{sanitize_filename(tdoc_id)}.md"
                        doc["target_md_abs_path"] = target_md_file
                        doc["target_md_rel_path"] = f"{rel_slug}/{wi_slug}/{cat_folder}/{sanitize_filename(tdoc_id)}.md"
                        doc["wi_local_rel_path"] = f"{cat_folder}/{sanitize_filename(tdoc_id)}.md"
                        if tdoc_id not in self.exported_tdoc_paths:
                            self.exported_tdoc_paths[tdoc_id] = target_md_file
                            self.tdoc_titles[tdoc_id] = doc.get("title", "")

        # Phase 2: Copy dossiers into target directory structure
        for rel_name, wis in filtered_entries.items():
            rel_slug = sanitize_folder_name(rel_name)
            for wi_name, wi_data in wis.items():
                wi_slug = sanitize_folder_name(wi_name)

                for cat_name, cat_docs in wi_data["categories"].items():
                    if not cat_docs:
                        continue

                    cat_folder = "Discussion" if cat_name == "DISC" else cat_name
                    type_dir = self.export_dir / rel_slug / wi_slug / cat_folder
                    type_media_dir = type_dir / "media"
                    os.makedirs(to_extended_path(type_dir), exist_ok=True)

                    for doc in cat_docs:
                        tdoc_id = doc["tdoc_id"]
                        local_md_file, local_media_dir = converted_cache.get(tdoc_id, (None, None))
                        target_md_file = doc.get("target_md_abs_path") or (type_dir / f"{sanitize_filename(tdoc_id)}.md")

                        # Collect header notices (cross-WI, replies, reply-to)
                        header_notices = []

                        # Inject cross-WI notification header if this TDoc overlaps with other WIs
                        if doc.get("cross_wis"):
                            other_wis_str = ", ".join(f"`{w}`" for w in doc["cross_wis"])
                            header_notices.append(
                                f"> [!NOTE]\n"
                                f"> **Cross-WI Contribution:** This TDoc is also associated with {other_wis_str}."
                            )

                        # Inject reciprocal links for LS documents
                        if doc.get("category") == "LS_IN" and doc.get("replies"):
                            rep_links = []
                            for rep_id in doc["replies"]:
                                rep_abs = self.exported_tdoc_paths.get(rep_id)
                                rep_title = self.tdoc_titles.get(rep_id, "")
                                rep_title_str = f" (*{rep_title}*)" if rep_title else ""
                                if rep_abs:
                                    rel = os.path.relpath(rep_abs, target_md_file.parent).replace("\\", "/")
                                    rep_links.append(f"**[{rep_id}]({rel})**{rep_title_str}")
                                else:
                                    rep_links.append(f"**{rep_id}**{rep_title_str}")
                            if rep_links:
                                header_notices.append(
                                    f"> [!TIP]\n"
                                    f"> **Outgoing Reply:** This incoming Liaison Statement was replied to in {', '.join(rep_links)}."
                                )

                        if doc.get("category") == "LS_OUT" and doc.get("reply_to_ids"):
                            in_links = []
                            for in_id in doc["reply_to_ids"]:
                                in_abs = self.exported_tdoc_paths.get(in_id)
                                in_title = self.tdoc_titles.get(in_id, "")
                                in_title_str = f" (*{in_title}*)" if in_title else ""
                                if in_abs:
                                    rel = os.path.relpath(in_abs, target_md_file.parent).replace("\\", "/")
                                    in_links.append(f"**[{in_id}]({rel})**{in_title_str}")
                                else:
                                    in_links.append(f"**{in_id}**{in_title_str}")
                            if in_links:
                                header_notices.append(
                                    f"> [!NOTE]\n"
                                    f"> **Reply Liaison Statement:** This outgoing LS is a reply to incoming LS {', '.join(in_links)}."
                                )

                        notice_block = "\n\n".join(header_notices) + ("\n\n" if header_notices else "")

                        if local_md_file and os.path.exists(to_extended_path(local_md_file)):
                            try:
                                with open(to_extended_path(local_md_file), "r", encoding="utf-8") as f:
                                    md_content = f.read()

                                tdoc_url = doc.get("tdoc_url") or resolve_tdoc_download_url(
                                    tdoc_id, docs_url=self.docs_ftp_url, revisions_url=self.revisions_url, meta=doc.get("row_data")
                                )
                                if tdoc_url and "- **3GPP Server URL:**" not in md_content:
                                    url_bullet = f"- **3GPP Server URL:** [{tdoc_url}]({tdoc_url})\n"
                                    if "- **Original file:**" in md_content:
                                        md_content = md_content.replace("- **Original file:**", f"{url_bullet}- **Original file:**", 1)
                                    elif "- **Exporter:**" in md_content:
                                        md_content = md_content.replace("- **Exporter:**", f"{url_bullet}- **Exporter:**", 1)
                                    elif "\n---\n" in md_content:
                                        md_content = md_content.replace("\n---\n", f"\n{url_bullet}\n---\n", 1)

                                if notice_block:
                                    if md_content.startswith("<!---"):
                                        first_break = md_content.find("\n\n")
                                        if first_break != -1:
                                            md_content = md_content[:first_break + 2] + notice_block + md_content[first_break + 2:]
                                        else:
                                            md_content = notice_block + md_content
                                    else:
                                        md_content = notice_block + md_content

                                with open(to_extended_path(target_md_file), "w", encoding="utf-8") as f:
                                    f.write(md_content)

                                # Copy media files with guaranteed collision-free prefix
                                if local_media_dir and local_media_dir.is_dir():
                                    media_files = list(local_media_dir.glob(f"{tdoc_id}_*"))
                                    if not media_files:
                                        media_files = list(local_media_dir.glob(f"{sanitize_filename(tdoc_id)}_*"))
                                    if media_files:
                                        os.makedirs(to_extended_path(type_media_dir), exist_ok=True)
                                        for mf in media_files:
                                            dst_img = type_media_dir / mf.name
                                            if not dst_img.exists():
                                                shutil.copy2(to_extended_path(mf), to_extended_path(dst_img))
                            except Exception as err:
                                logger.error(f"[PostMeetingExporter] Error writing dossier for {tdoc_id}: {err}")
                        else:
                            # Create fallback stub
                            try:
                                with open(to_extended_path(target_md_file), "w", encoding="utf-8") as f:
                                    f.write(
                                        f"# {tdoc_id}: {doc['title']}\n\n"
                                        f"{notice_block}"
                                        f"* **Source:** {doc['source']}\n"
                                        f"* **Type:** {doc['type']}\n"
                                        f"* **Agenda Item:** {doc['ai']}\n"
                                        f"* **Status:** {doc['status']}\n\n"
                                        f"> ⚠️ *Note: Document conversion was not available in local cache.*\n"
                                    )
                            except Exception:
                                pass

    def _ensure_tdoc_markdown(self, tdoc_id: str, row_data: dict) -> tuple:
        """Finds or converts the canonical Markdown dossier for a TDoc in the meeting directory."""
        base_match = re.search(r'^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$', tdoc_id, re.IGNORECASE)
        base_tdoc = base_match.group(1).upper() if base_match else tdoc_id.upper()

        tdoc_folder = self.meeting_dir / base_tdoc
        local_md_file = tdoc_folder / f"{sanitize_filename(tdoc_id)}.md"
        local_media_dir = tdoc_folder / "media"
        legacy_cache_file = tdoc_folder / f"{tdoc_id}_LLM_v{CONVERTER_VERSION}.md"

        doc_path = self._find_word_doc(tdoc_folder, tdoc_id)
        doc_type = row_data.get("Type", "") if row_data else ""
        expected_mode = resolve_revision_mode_for_type(doc_type, default_mode="markup")

        needs_conversion = self.force_reconvert or is_conversion_stale(
            doc_path, local_md_file, expected_version=CONVERTER_VERSION, expected_mode=expected_mode
        )

        if needs_conversion and not self.force_reconvert and os.path.exists(to_extended_path(legacy_cache_file)):
            if not is_conversion_stale(doc_path, legacy_cache_file, expected_version=CONVERTER_VERSION, expected_mode=expected_mode):
                try:
                    shutil.copy2(to_extended_path(legacy_cache_file), to_extended_path(local_md_file))
                    needs_conversion = False
                except Exception:
                    pass

        if needs_conversion:
            if not doc_path and self.download_missing:
                self.progress.emit(f"Downloading missing TDoc {tdoc_id}...")
                doc_path = self._download_and_extract_tdoc(tdoc_id, tdoc_folder)

            if doc_path:
                try:
                    os.makedirs(to_extended_path(tdoc_folder), exist_ok=True)
                    os.makedirs(to_extended_path(local_media_dir), exist_ok=True)
                    tdoc_download_url = resolve_tdoc_download_url(
                        tdoc_id,
                        docs_url=self.docs_ftp_url,
                        revisions_url=self.revisions_url,
                        meta=row_data,
                    )
                    convert_document_to_markdown(
                        doc_path,
                        out_dir=tdoc_folder,
                        tdoc_id=tdoc_id,
                        meta=row_data,
                        revisions=expected_mode,
                        media_dir=local_media_dir,
                        media_rel="media",
                        cancellation_check=self.isInterruptionRequested,
                        tdoc_url=tdoc_download_url,
                    )
                except Exception as e:
                    logger.warning(f"[PostMeetingExporter] Conversion error on {tdoc_id}: {e}")

        return local_md_file, local_media_dir

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

        # Build candidate routes: prioritize revisions_url for revisions, docs_ftp_url for base
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
            logger.error(f"[PostMeetingExporter] No base URLs configured for downloading {tdoc_id}")
            return None

        last_error = None
        for base_url in candidate_urls:
            if self.isInterruptionRequested():
                return None
            try:
                logger.info(f"[PostMeetingExporter] Attempting download of {tdoc_id} from {base_url}")
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
                logger.debug(f"[PostMeetingExporter] Route {base_url} failed for {tdoc_id}: {e}")

        if last_error:
            logger.warning(f"[PostMeetingExporter] Failed to download {tdoc_id} across all routes: {last_error}")
        return None

    def _generate_local_wi_indices(self, filtered_entries: dict):
        """Writes an index.md dossier cover inside each Work Item directory."""
        for rel_name, wis in filtered_entries.items():
            rel_slug = sanitize_folder_name(rel_name)
            for wi_name, wi_data in wis.items():
                wi_slug = sanitize_folder_name(wi_name)
                wi_dir = self.export_dir / rel_slug / wi_slug
                wi_index_file = wi_dir / "index.md"

                info = wi_data["info"]
                wi_title = info.get("name") or wi_name
                wi_code = info.get("code")

                lines = []
                lines.append(f"# Work Item Dossier: {wi_name}\n")
                if wi_title and wi_title != wi_name:
                    lines.append(f"**Name:** {wi_title}  \n")
                if wi_code:
                    lines.append(f"**WI Code:** `{wi_code}`  \n")
                lines.append(f"**Release:** {rel_name}  \n")

                # List all Agenda Items covered
                if wi_data["ais"]:
                    sorted_ais = sorted(list(wi_data["ais"]))
                    lines.append("\n### 📋 Covered Agenda Items\n")
                    for ai in sorted_ais:
                        desc = wi_data["ai_titles"].get(ai, "")
                        if desc:
                            lines.append(f"* **AI {ai}**: {desc}")
                        else:
                            lines.append(f"* **AI {ai}**")
                    lines.append("")

                lines.append("---\n")

                # Document Tables
                category_titles = [
                    ("LS_IN", "📥 Incoming Liaison Statements (LS in)"),
                    ("pCR", "📝 Agreed Pseudo-CRs (pCR)"),
                    ("CR", "📄 Agreed Change Requests (CR)"),
                    ("LS_OUT", "📤 Agreed Outgoing Liaison Statements (LS out)"),
                    ("DISC", "💬 Discussion & Status Papers (Disc / WID)"),
                ]

                has_docs = False
                for cat_key, cat_title in category_titles:
                    docs = wi_data["categories"].get(cat_key, [])
                    if docs:
                        has_docs = True
                        lines.append(f"### {cat_title}\n")
                        lines.append(self._render_markdown_table(docs, is_local_wi=True, base_dir=wi_dir))
                        lines.append("")

                if not has_docs:
                    lines.append("*No documents exported under this Work Item.*\n")

                try:
                    with open(to_extended_path(wi_index_file), "w", encoding="utf-8") as f:
                        f.write("\n".join(lines))
                except Exception as e:
                    logger.error(f"[PostMeetingExporter] Error writing WI index for {wi_name}: {e}")

    def _generate_master_index(self, filtered_entries: dict):
        """Generates the master index.md cover page with executive summary tables."""
        master_file = self.export_dir / "index.md"
        wg = str(self.mtg_info.get("wg_name", "3GPP")).strip()
        mtg = str(self.mtg_info.get("meeting_number", "")).strip()
        mtg_title = f"{wg} #{mtg}" if (wg and mtg) else "3GPP Meeting"

        # Compute executive counts
        cat_counts = {"LS_IN": 0, "pCR": 0, "CR": 0, "LS_OUT": 0, "DISC": 0}
        total_unique_tdocs = set()
        total_wis = set()

        for rel in filtered_entries.values():
            for wi_name, wi_data in rel.items():
                if wi_name != "General":
                    total_wis.add(wi_name)
                for cat, docs in wi_data["categories"].items():
                    cat_counts[cat] = cat_counts.get(cat, 0) + len(docs)
                    for d in docs:
                        total_unique_tdocs.add(d["tdoc_id"])

        now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
        scope_labels = {
            "followed": "🎯 Followed Work Items only",
            "starred": "⭐ Starred Work Items only",
            "tracked": "🎯 Followed & ⭐ Starred Work Items",
            "visible": "🔍 Current Filtered Table View",
            "whole": "🌐 Whole Meeting (All Work Items & Contributions)",
            "custom": "📋 Custom Work Item & Agenda Item Selection",
        }
        scope_str = scope_labels.get(self.scope_mode, self.scope_mode)

        lines = []
        lines.append(f"# {mtg_title} — Post-Meeting Dossier Report\n")
        lines.append(f"**Meeting:** {mtg_title}  \n")
        lines.append(f"**Generated:** {now_str}  \n")
        lines.append(f"**Export Scope:** {scope_str}  \n")
        lines.append(f"**Work Items Covered:** {len(total_wis)}  \n")
        lines.append(f"**Total Exported TDocs:** {len(total_unique_tdocs)}  \n")
        lines.append("\n---\n")

        # Executive Summary Table
        lines.append("## 📊 Executive Summary\n")
        lines.append("| Category | Total Count | Decision / Status Criteria |")
        lines.append("|:---|:---:|:---|")
        lines.append(f"| 📥 Incoming LS (`LS in`) | **{cat_counts['LS_IN']}** | All received / recorded incoming LS |")
        lines.append(f"| 📝 Agreed Pseudo-CRs (`pCR`) | **{cat_counts['pCR']}** | Agreed / Approved by meeting |")
        lines.append(f"| 📄 Agreed Change Requests (`CR`) | **{cat_counts['CR']}** | Agreed / Approved by meeting |")
        lines.append(f"| 📤 Outgoing LS (`LS out`) | **{cat_counts['LS_OUT']}** | Agreed / Approved by meeting |")
        if self.include_disc and cat_counts.get("DISC", 0) > 0:
            lines.append(f"| 💬 Discussion & Status (`Disc`) | **{cat_counts['DISC']}** | Noted / Agreed discussion & status papers |")
        lines.append(f"| **Total Document Instances** | **{sum(cat_counts.values())}** | *(Unique TDocs: {len(total_unique_tdocs)})* |")
        lines.append("\n---\n")

        # Table of Contents
        lines.append("## 📑 Table of Contents\n")
        for rel_name in sorted(filtered_entries.keys()):
            rel_slug = sanitize_folder_name(rel_name)
            lines.append(f"* **{rel_name}**")
            for wi_name in sorted(filtered_entries[rel_name].keys()):
                wi_slug = sanitize_folder_name(wi_name)
                wi_anchor = f"wi-{wi_slug.lower()}"
                lines.append(f"  * [{wi_name}](#{wi_anchor}) ([Dossier Folder]({rel_slug}/{wi_slug}/))")
        lines.append("\n---\n")

        # Per-Release / Per-WI Detailed Tables
        for rel_name in sorted(filtered_entries.keys()):
            rel_slug = sanitize_folder_name(rel_name)
            lines.append(f"## Release: {rel_name}\n")

            for wi_name in sorted(filtered_entries[rel_name].keys()):
                wi_data = filtered_entries[rel_name][wi_name]
                wi_slug = sanitize_folder_name(wi_name)
                wi_anchor = f"wi-{wi_slug.lower()}"

                info = wi_data["info"]
                wi_title = info.get("name") or wi_name
                wi_code = info.get("code")

                lines.append(f'<a name="{wi_anchor}"></a>\n')
                lines.append(f"### 🎯 Work Item: `{wi_name}`\n")
                if wi_title and wi_title != wi_name:
                    lines.append(f"**Title:** {wi_title}  \n")
                if wi_code:
                    lines.append(f"**WI Code:** `{wi_code}`  \n")
                lines.append(f"**Dossier Subfolder:** [`{rel_slug}/{wi_slug}/`]({rel_slug}/{wi_slug}/index.md)  \n")

                if wi_data["ais"]:
                    ai_bullets = []
                    for ai in sorted(list(wi_data["ais"])):
                        desc = wi_data["ai_titles"].get(ai, "")
                        ai_bullets.append(f"`AI {ai}`: {desc}" if desc else f"`AI {ai}`")
                    lines.append(f"**Covered Agenda Items:** {'; '.join(ai_bullets)}  \n")

                lines.append("")

                # Tables by category
                category_titles = [
                    ("LS_IN", "📥 Incoming Liaison Statements (LS in)"),
                    ("pCR", "📝 Agreed Pseudo-CRs (pCR)"),
                    ("CR", "📄 Agreed Change Requests (CR)"),
                    ("LS_OUT", "📤 Agreed Outgoing Liaison Statements (LS out)"),
                    ("DISC", "💬 Discussion & Status Papers (Disc / WID)"),
                ]

                has_docs = False
                for cat_key, cat_title in category_titles:
                    docs = wi_data["categories"].get(cat_key, [])
                    if docs:
                        has_docs = True
                        lines.append(f"#### {cat_title}\n")
                        lines.append(self._render_markdown_table(docs, is_local_wi=False, base_dir=self.export_dir))
                        lines.append("")

                if not has_docs:
                    lines.append("*No documents exported in this category.*\n")

                lines.append("")

        try:
            with open(to_extended_path(master_file), "w", encoding="utf-8") as f:
                f.write("\n".join(lines))
        except Exception as e:
            logger.error(f"[PostMeetingExporter] Error writing master index.md: {e}", exc_info=True)

    def _render_markdown_table(self, docs: List[dict], is_local_wi: bool = False, base_dir: Optional[Path] = None) -> str:
        """Renders a standard markdown table for a list of TDoc entries."""
        if not docs:
            return "*None*\n"

        if base_dir is None:
            base_dir = self.export_dir

        headers = ["TDoc", "Type", "Status", "AI", "Title", "Source", "Cross-WI", "Links"]
        lines = [
            "| " + " | ".join(headers) + " |",
            "|" + "|".join([":---"] * len(headers)) + "|",
        ]

        for d in docs:
            tdoc_id = d["tdoc_id"]
            target_abs = d.get("target_md_abs_path")
            if target_abs and base_dir:
                link_path = os.path.relpath(target_abs, base_dir).replace("\\", "/")
            elif is_local_wi:
                link_path = d.get("wi_local_rel_path")
            else:
                link_path = d.get("target_md_rel_path")

            link_str = f"[{tdoc_id}]({link_path})" if link_path else tdoc_id

            title_safe = str(d.get("title", "")).replace("\n", " ").replace("\r", "").replace("|", "\\|")
            source_safe = str(d.get("source", "")).replace("\n", " ").replace("\r", "").replace("|", "\\|")
            cross_wi_str = ", ".join(f"`{w}`" for w in d.get("cross_wis", [])) if d.get("cross_wis") else "—"

            raw_status = str(d.get("status", "")).strip()
            cat = d.get("category", "")
            if cat == "LS_IN" and d.get("replies"):
                rep_items = []
                for rep_id in d["replies"]:
                    rep_abs = getattr(self, "exported_tdoc_paths", {}).get(rep_id)
                    if rep_abs and base_dir:
                        r_rel = os.path.relpath(rep_abs, base_dir).replace("\\", "/")
                        rep_items.append(f"[{rep_id}]({r_rel})")
                    else:
                        rep_items.append(rep_id)
                rep_str = ", ".join(rep_items)
                if "replied" in raw_status.lower():
                    status_cell = f"replied to in {rep_str}"
                elif raw_status:
                    status_cell = f"{raw_status} (replied in {rep_str})"
                else:
                    status_cell = f"replied to in {rep_str}"
            elif cat == "LS_OUT" and d.get("reply_to_ids"):
                in_items = []
                for in_id in d["reply_to_ids"]:
                    in_abs = getattr(self, "exported_tdoc_paths", {}).get(in_id)
                    if in_abs and base_dir:
                        i_rel = os.path.relpath(in_abs, base_dir).replace("\\", "/")
                        in_items.append(f"[{in_id}]({i_rel})")
                    else:
                        in_items.append(in_id)
                in_str = ", ".join(in_items)
                if raw_status:
                    status_cell = f"{raw_status} (reply to {in_str})"
                else:
                    status_cell = f"reply to {in_str}"
            else:
                status_cell = raw_status

            tdoc_url = d.get("tdoc_url") or resolve_tdoc_download_url(
                tdoc_id,
                docs_url=self.docs_ftp_url,
                revisions_url=self.revisions_url,
                meta=d.get("row_data"),
            )
            if link_path and tdoc_url:
                links_cell = f"[Dossier]({link_path}) · [📥 3GPP]({tdoc_url})"
            elif link_path:
                links_cell = f"[Dossier]({link_path})"
            elif tdoc_url:
                links_cell = f"[📥 3GPP]({tdoc_url})"
            else:
                links_cell = "—"

            row_cells = [
                link_str,
                str(d.get("type", "")),
                status_cell,
                str(d.get("ai", "")),
                title_safe,
                source_safe,
                cross_wi_str,
                links_cell,
            ]
            lines.append("| " + " | ".join(row_cells) + " |")

        return "\n".join(lines)
