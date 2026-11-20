# --- File: src/modules/meetings/core/exporters/post_meeting/mixins/grouping.py ---
"""
LS reply tracking, Work Items DB loading, scope filtering, Rel-XX/WI tree grouping,
and WI / Release resolution.

Mixin of PostMeetingExporterThread: relies on attributes set in __init__.
"""
import re
from typing import (
    Dict,
    List,
    Set,
    Any,
)
from core.utils.paths import get_project_root
from modules.meetings.core.agenda_manager import AgendaItem
from modules.meetings.core.tdocs.tdoc_url_resolver import resolve_tdoc_download_url
from modules.meetings.ui.tdocs_models import TDocsTableModel
from modules.work_items.core.wi_database import WorkItemsDatabase
from modules.work_items.core.wi_settings import WorkItemsSettings
from modules.meetings.core.exporters.post_meeting.path_utils import logger


class GroupingMixin:
    """LS reply tracking, Work Items DB loading, scope filtering, Rel-XX/WI tree grouping,"""

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
                        "agreements": [],
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

        # Attach Chairman's Notes Agreements if enabled
        if self.include_chair_agreements:
            self._attach_chair_notes_agreements(tree, tdoc_map, wi_db_info)

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
