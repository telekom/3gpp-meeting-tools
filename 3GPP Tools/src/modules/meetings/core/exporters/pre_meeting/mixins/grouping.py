# --- File: src/modules/meetings/core/exporters/pre_meeting/mixins/grouping.py ---
"""Revision lineage, multi-candidate LS reply mapping, and Pre-Meeting TDoc grouping mixin."""

import re
from typing import Any, Dict, List, Set, Tuple

from modules.meetings.core.tdocs.tdoc_url_resolver import resolve_tdoc_download_url
from modules.meetings.ui.tdocs_models import TDocsTableModel


class PreMeetingGroupingMixin:
    """Builds revision lineages, LS reply associations, and Release -> WI -> Category trees."""

    def _build_revision_lineage_map(self) -> Tuple[Set[str], Dict[str, List[str]], Dict[str, str]]:
        """
        Builds complete revision lineage relationships across all meeting rows:
        - superseded_set: Set of TDoc IDs that have a newer active (non-withdrawn) revision in all_rows.
        - predecessors_map: latest_leaf_tdoc_id -> List of superseded predecessor TDoc IDs in chronological order.
        - leaf_map: tdoc_id -> latest_active_leaf_tdoc_id
        """
        tdoc_map = {str(r.get("TDoc", "")).strip().upper(): r for r in self.all_rows if r.get("TDoc")}

        # Build forward edges: parent -> list of child revisions
        forward_edges: Dict[str, Set[str]] = {}
        for tdoc_id, row in tdoc_map.items():
            # 1. Inspect 'Revised to'
            rev_to = str(row.get("Revised to", "")).strip()
            if rev_to and rev_to.lower() != "none":
                for cand in re.findall(r'[A-Za-z0-9]+-\d+(?:r\d+[a-zA-Z]?)?', rev_to):
                    cand_upper = cand.upper()
                    if cand_upper in tdoc_map and cand_upper != tdoc_id:
                        forward_edges.setdefault(tdoc_id, set()).add(cand_upper)

            # 2. Inspect 'Is revision of'
            rev_of = str(row.get("Is revision of", "")).strip()
            if rev_of and rev_of.lower() != "none":
                for cand in re.findall(r'[A-Za-z0-9]+-\d+(?:r\d+[a-zA-Z]?)?', rev_of):
                    cand_upper = cand.upper()
                    if cand_upper in tdoc_map and cand_upper != tdoc_id:
                        forward_edges.setdefault(cand_upper, set()).add(tdoc_id)

            # 3. Base TDoc regex match (e.g. S2-260100r01 is revision of S2-260100)
            base_match = re.search(r'^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$', tdoc_id, re.IGNORECASE)
            if base_match:
                base_tdoc = base_match.group(1).upper()
                if base_tdoc in tdoc_map and base_tdoc != tdoc_id:
                    forward_edges.setdefault(base_tdoc, set()).add(tdoc_id)

        def _is_active_tdoc(tid: str) -> bool:
            r = tdoc_map.get(tid)
            if not r:
                return False
            st = str(r.get("TDoc Status", "")).strip().lower()
            return "withdrawn" not in st

        # Trace each node to its latest active leaf
        leaf_map: Dict[str, str] = {}
        superseded_set: Set[str] = set()
        predecessors_map: Dict[str, List[str]] = {}

        for tdoc_id in tdoc_map:
            if not _is_active_tdoc(tdoc_id):
                continue

            curr = tdoc_id
            visited = {curr}
            chain = [curr]

            while True:
                children = forward_edges.get(curr, set())
                active_children = [c for c in children if c not in visited and _is_active_tdoc(c)]
                if not active_children:
                    break
                # Sort naturally to pick newest revision (e.g. r02 over r01)
                active_children.sort(key=lambda x: str(x))
                next_node = active_children[-1]
                visited.add(next_node)
                chain.append(next_node)
                curr = next_node

            leaf_tdoc = curr
            leaf_map[tdoc_id] = leaf_tdoc

            if leaf_tdoc != tdoc_id:
                superseded_set.add(tdoc_id)
                preds = predecessors_map.setdefault(leaf_tdoc, [])
                if tdoc_id not in preds:
                    preds.append(tdoc_id)

        # Sort predecessor lists for clean display
        for _leaf, preds in predecessors_map.items():
            preds.sort()

        return superseded_set, predecessors_map, leaf_map

    def _build_ls_reply_map(self) -> Tuple[Dict[str, List[str]], Dict[str, List[str]], Dict[str, dict]]:
        """
        Builds multi-candidate reply mappings for Pre-Meeting preparation:
        - in_to_replies: incoming_tdoc_id -> List of candidate draft LS out tdoc_ids (sorted by TDoc ID)
        - out_to_in: outgoing_tdoc_id -> List of incoming LS tdoc_ids it answers
        - tdoc_map: tdoc_id -> row_dict
        Under deduplicate_revisions, resolves candidate reply LS out documents to their latest active revision.
        """
        tdoc_map = {str(r.get("TDoc", "")).strip().upper(): r for r in self.all_rows if r.get("TDoc")}
        _, _, leaf_map = self._build_revision_lineage_map()

        def _resolve_candidate(tid: str) -> str:
            upper = tid.upper()
            if self.deduplicate_revisions and upper in leaf_map:
                return leaf_map[upper]
            return upper

        in_to_replies_set: Dict[str, Set[str]] = {}
        out_to_in_set: Dict[str, Set[str]] = {}

        for r in self.all_rows:
            tdoc = str(r.get("TDoc", "")).strip().upper()
            if not tdoc:
                continue

            # Check if tdoc itself is withdrawn
            status = str(r.get("TDoc Status", "")).strip().lower()
            if "withdrawn" in status:
                continue

            rep_in = str(r.get("Reply in", "")).strip()
            rep_to = str(r.get("Reply to", "")).strip()
            orig_ls = str(r.get("Original LS", "")).strip()

            resolved_out = _resolve_candidate(tdoc)

            # 1. Incoming LS pointing forward to draft replies via 'Reply in'
            if rep_in:
                targets = [_resolve_candidate(x) for x in re.findall(r'[A-Za-z0-9]+-\d+(?:r\d+[a-zA-Z]?)?', rep_in)]
                for tgt in targets:
                    if tgt in tdoc_map:
                        tgt_status = str(tdoc_map[tgt].get("TDoc Status", "")).strip().lower()
                        if "withdrawn" not in tgt_status:
                            in_to_replies_set.setdefault(tdoc, set()).add(tgt)
                            out_to_in_set.setdefault(tgt, set()).add(tdoc)

            # 2. Outgoing proposed LS pointing back to incoming LS via 'Reply to' or 'Original LS'
            back_sources = []
            if rep_to:
                back_sources.extend(re.findall(r'[A-Za-z0-9]+-\d+(?:r\d+[a-zA-Z]?)?', rep_to))
            if orig_ls:
                back_sources.extend(re.findall(r'[A-Za-z0-9]+-\d+(?:r\d+[a-zA-Z]?)?', orig_ls))

            for inc in back_sources:
                inc_upper = inc.upper()
                in_to_replies_set.setdefault(inc_upper, set()).add(resolved_out)
                out_to_in_set.setdefault(resolved_out, set()).add(inc_upper)

        in_to_replies = {k: sorted(list(v)) for k, v in in_to_replies_set.items()}
        out_to_in = {k: sorted(list(v)) for k, v in out_to_in_set.items()}

        return in_to_replies, out_to_in, tdoc_map

    def _filter_and_group_tdocs(self, wi_db_info: dict) -> dict:
        """
        Filters meeting TDocs according to Pre-Meeting rules:
        - Captures active company proposals (Discussion papers, WID/SIDs, proposed CRs, proposed pCRs)
          and incoming liaison statements (LS in).
        - Connects 1-to-many candidate LS reply proposals and allows candidate LS out to inherit parent WI.
        - Deduplicates earlier pre-meeting revisions when deduplicate_revisions=True.
        - Groups by Release -> Work Item -> Category.
        """
        tree: Dict[str, Dict[str, Any]] = {}

        superseded_set, predecessors_map, _ = self._build_revision_lineage_map()
        in_to_replies, out_to_in, tdoc_map = self._build_ls_reply_map()
        visible_tdoc_ids = {str(r.get("TDoc", "")).strip().upper() for r in self.visible_rows if r.get("TDoc")}

        for row in self.all_rows:
            tdoc_id = str(row.get("TDoc", "")).strip().upper()
            if not tdoc_id:
                continue

            status = str(row.get("TDoc Status", "")).strip().lower()

            # Always skip explicitly withdrawn contributions
            if "withdrawn" in status:
                continue

            # Skip superseded revisions if deduplication is enabled
            if self.deduplicate_revisions and tdoc_id in superseded_set:
                continue

            is_reserved = any(res in status for res in ["reserved", "not uploaded", "not yet uploaded"])
            if is_reserved:
                base_m = re.search(r'^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$', tdoc_id, re.IGNORECASE)
                b_tdoc = base_m.group(1).upper() if base_m else tdoc_id.upper()
                if self._find_word_doc(self.meeting_dir / b_tdoc, tdoc_id):
                    is_reserved = False

            if is_reserved and not self.include_reserved_in_index:
                continue

            # Mid-meeting tolerance check (if user chose to exclude already-treated docs)
            if not self.include_postponed_or_treated:
                if any(tr in status for tr in ["agreed", "approved", "noted", "rejected", "postponed", "not agreed"]):
                    continue

            doc_type = str(row.get("Type", "")).strip().lower()

            category = None
            if "ls in" in doc_type or doc_type == "ls":
                if self.include_ls_in:
                    category = "LS_IN"
            elif "ls out" in doc_type or "ls_out" in doc_type:
                if self.include_ls_out:
                    category = "LS_OUT"
            elif "pcr" in doc_type or "p-cr" in doc_type:
                if self.include_pcr:
                    category = "pCR"
            elif "cr" in doc_type:
                if self.include_cr:
                    category = "CR"
            elif any(d in doc_type for d in ["disc", "discussion", "report", "wid", "sid", "other", "draft", "evaluation"]):
                if self.include_disc:
                    category = "DISC"

            if not category:
                continue

            if self.scope_mode == "visible" and tdoc_id not in visible_tdoc_ids:
                continue

            # Contributor company filter (if active)
            row_companies = set(row.get("_All_Sanitized_Companies", row.get("_Sanitized_Companies", [])))
            row_companies_lower = {str(c).strip().lower() for c in row_companies}

            if self.company_filter:
                company_matched = bool(row_companies_lower & self.company_filter)
                # Keep LS_IN if it has candidate replies from a matching company
                if not company_matched and category == "LS_IN":
                    for cand_rep in in_to_replies.get(tdoc_id, []):
                        rep_row = tdoc_map.get(cand_rep, {})
                        rep_comps = {
                            str(c).strip().lower()
                            for c in rep_row.get("_All_Sanitized_Companies", rep_row.get("_Sanitized_Companies", []))
                        }
                        if rep_comps & self.company_filter:
                            company_matched = True
                            break
                if not company_matched:
                    continue

            # Resolve associated Work Items
            associated_wis = self._resolve_tdoc_wis(row, wi_db_info)

            # Work Item Inheritance for candidate LS OUT:
            # If an LS out has no official WI, inherit the WI from the parent LS in it answers
            if category == "LS_OUT" and (not associated_wis or associated_wis == ["General"]):
                inherited_wis = []
                for inc_id in out_to_in.get(tdoc_id, []):
                    inc_row = tdoc_map.get(inc_id)
                    if inc_row:
                        inc_wis = self._resolve_tdoc_wis(inc_row, wi_db_info)
                        for iw in inc_wis:
                            if iw != "General" and iw not in inherited_wis:
                                inherited_wis.append(iw)
                if inherited_wis:
                    associated_wis = inherited_wis

            raw_wis = TDocsTableModel._split_related_wis(row.get("Related WIs", ""))
            clean_wis = [w for w in raw_wis if w and w.lower() not in {"(no wi)", "none", "nan", "-"}]
            has_official_wi = bool(clean_wis)

            ai_num = str(row.get("Agenda Item", "")).strip() or "Unassigned"
            ai_desc = str(row.get("Agenda item description", "")).strip()

            if not associated_wis:
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

            # Fallback to General for unassigned LS
            if (category in ("LS_IN", "LS_OUT")) and self.include_general_ls and not has_official_wi and not matching_wis:
                matching_wis.append("General")
                if "General" not in associated_wis:
                    associated_wis.append("General")

            if not matching_wis:
                continue

            # Resolve release
            release = self._resolve_tdoc_release(row, matching_wis, wi_db_info)

            # Sibling replies for LS_OUT (competing proposals answering the same incoming LS)
            sibling_replies = set()
            if category == "LS_OUT":
                for parent_in in out_to_in.get(tdoc_id, []):
                    for other_out in in_to_replies.get(parent_in, []):
                        if other_out != tdoc_id:
                            sibling_replies.add(other_out)

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
                        "categories": {"LS_IN": [], "LS_OUT": [], "pCR": [], "CR": [], "DISC": []},
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
                    "replies": in_to_replies.get(tdoc_id, []),
                    "reply_to_ids": out_to_in.get(tdoc_id, []),
                    "sibling_replies": sorted(list(sibling_replies)),
                    "supersedes": predecessors_map.get(tdoc_id, []),
                    "is_reserved": is_reserved,
                    "my_status": str(row.get("My Status", "⚪ Neutral")).replace("🔄 ", "").strip(),
                    "my_notes": str(row.get("My Notes", "")).replace("🔄 [From Base]: ", "").strip(),
                    "for_action": str(row.get("For", "")).strip(),
                    "sanitized_companies": row.get("_Sanitized_Companies", []),
                    "row_data": row,
                    "tdoc_url": tdoc_url,
                }
                tree[release][target_wi]["categories"][category].append(entry)

        # Post-pass: Guarantee that any candidate LS OUT replying to an included LS IN is present
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

            rep_status = str(rep_row.get("TDoc Status", "")).strip().lower()
            if "withdrawn" in rep_status:
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
                    "categories": {"LS_IN": [], "LS_OUT": [], "pCR": [], "CR": [], "DISC": []},
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
            rep_b_m = re.search(r'^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$', rep_id, re.IGNORECASE)
            rep_b_tdoc = rep_b_m.group(1).upper() if rep_b_m else rep_id.upper()
            rep_is_res = any(res in rep_status for res in ["reserved", "not uploaded"])
            if rep_is_res and self._find_word_doc(self.meeting_dir / rep_b_tdoc, rep_id):
                rep_is_res = False

            rep_entry = {
                "tdoc_id": rep_id,
                "title": str(rep_row.get("Title", "")).strip(),
                "source": str(rep_row.get("Source", "")).strip(),
                "type": str(rep_row.get("Type", "")).strip(),
                "status": str(rep_row.get("TDoc Status", "")).strip(),
                "ai": rep_ai,
                "spec": str(rep_row.get("Spec", "")).strip(),
                "category": "LS_OUT",
                "cross_wis": clean_wis,
                "replies": in_to_replies.get(rep_id, []),
                "reply_to_ids": out_to_in.get(rep_id, []),
                "sibling_replies": [],
                "supersedes": predecessors_map.get(rep_id, []),
                "is_reserved": rep_is_res,
                "my_status": str(rep_row.get("My Status", "⚪ Neutral")).replace("🔄 ", "").strip(),
                "my_notes": str(rep_row.get("My Notes", "")).replace("🔄 [From Base]: ", "").strip(),
                "for_action": str(rep_row.get("For", "")).strip(),
                "sanitized_companies": rep_row.get("_Sanitized_Companies", []),
                "row_data": rep_row,
                "tdoc_url": rep_url,
            }
            tree[rep_rel][target_wi]["categories"]["LS_OUT"].append(rep_entry)
            all_included_tdocs.add(rep_id)

        # Sort entries within each category
        for rel in tree.values():
            for wi in rel.values():
                for cat in wi["categories"]:
                    wi["categories"][cat].sort(key=lambda x: str(x.get("tdoc_id", "")))

        return tree

