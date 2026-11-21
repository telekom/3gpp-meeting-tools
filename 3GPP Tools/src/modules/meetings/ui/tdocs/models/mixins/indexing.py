"""O(1) WI token, agenda sort tuple, sanitization, lineage family, and email indexing mixin."""

import re

from core.utils.company_sanitizer import CompanySanitizer
from modules.meetings.core.agenda_manager import AgendaItem
from modules.meetings.ui.tdocs.models.sorting import natural_sort_key


class TDocsIndexingMixin:
    """Precomputes WI tokens, agenda sort keys, sanitized companies, and revision families."""

    @staticmethod
    def _extract_base_tokens(text: str) -> set:
        """Generates lowercase base tokens (with/without FS_ prefix and with stripped WG suffixes)."""
        clean = str(text or "").strip()
        if not clean or clean.lower() in {"(no wi)", "none", "nan", "-"}:
            return set()
        tokens = {clean.lower()}
        base = re.sub(
            r'[-_](?:ARC|CT|Core|Perf|SEC|RAN|SA\d?|CT\d?|UEConTest|ServConTest)$',
            '', clean, flags=re.IGNORECASE
        ).strip()
        if base:
            tokens.add(base.lower())
        if clean.upper().startswith("FS_"):
            tokens.add(clean[3:].lower())
            if base and base.upper().startswith("FS_"):
                tokens.add(base[3:].lower())
        else:
            tokens.add(f"fs_{clean.lower()}")
            if base:
                tokens.add(f"fs_{base.lower()}")
        return tokens

    def _extract_row_wi_tokens(self, row: dict) -> set:
        """Precomputes all WI tokens associated with a TDoc row."""
        tokens = set()
        raw_wis = self._split_related_wis(row.get("Related WIs", ""))
        for wi in raw_wis:
            tokens.update(self._extract_base_tokens(wi))

        ai_num = str(row.get("Agenda Item", "")).strip()
        if ai_num and self.agenda_map:
            agenda_item = self.agenda_map.get(ai_num)
            if agenda_item and getattr(agenda_item, "acronym", None):
                tokens.update(self._extract_base_tokens(agenda_item.acronym))

        desc = str(row.get("Agenda item description", "")).strip()
        if desc:
            cand = AgendaItem._extract_acronym(desc)
            if cand:
                tokens.update(self._extract_base_tokens(cand))
            else:
                match = re.search(r'\(([A-Za-z0-9_\-\/]+)\)', desc)
                if match:
                    candidate = match.group(1).strip()
                    if 2 <= len(candidate) <= 25 and any(c.isupper() or c.isdigit() for c in candidate):
                        tokens.update(self._extract_base_tokens(candidate))

        return tokens

    def _precompute_row_wi_tokens(self, rows: list):
        """Precomputes and caches _wi_tokens on each row dictionary for fast O(1) filtering."""
        for row in rows:
            row["_wi_tokens"] = self._extract_row_wi_tokens(row)

    def _precompute_agenda_sort_tuples(self, rows: list):
        """
        Precomputes _agenda_sort_tuple on each row dictionary for high-performance O(1)
        composite sorting in TDocsFilterProxyModel.

        Ordering Hierarchy:
          1. Agenda Item (natural numeric: 5.1 < 5.2 < 5.10 < 6.1; unassigned/empty at the end)
          2. Precedence Tier:
             - Tier 0: TDocsByAgenda / Chair notes order (live meeting order; supersedes Excel)
             - Tier 1: Excel baseline 'TDoc sort order within agenda item'
             - Tier 2: Unindexed / fallback
          3. Sub-Item sequence (order in which Sub-Item dividers appeared in TDocsByAgenda)
          4. Document sequence within Agenda Item / Sub-Item
          5. Revision offset (anchors unindexed r01, r02 directly beneath their parent TDoc)
          6. TDoc alphanumeric identifier
        """
        tdoc_map = {str(r.get("TDoc", "")).strip().upper(): r for r in self._data if r.get("TDoc")}

        for row in rows:
            raw_ai = str(row.get("Agenda Item", "") or "").strip()
            is_unassigned = 1 if (not raw_ai or raw_ai.lower() in ("n/a", "none", "unknown", "-")) else 0
            ai_key = natural_sort_key(raw_ai) if not is_unassigned else ()

            tier = 2
            sub_item_order = int(row.get('_sub_item_order') or 0)
            order_val = float('inf')
            rev_offset = 0

            tdoc_id = str(row.get("TDoc", "") or "").strip().upper()
            pred = str(row.get("Is revision of", "") or "").strip().upper()
            base_tdoc = ""
            rev_num = 0

            rev_match = re.search(r'^(.*?)-?(?:r|rev)(\d{1,2}[a-zA-Z]?)$', tdoc_id, re.IGNORECASE)
            if rev_match:
                base_tdoc = rev_match.group(1).upper()
                try:
                    rev_num = int(re.match(r'\d+', rev_match.group(2)).group())
                except Exception:
                    rev_num = 1
            elif pred:
                base_match = re.search(r'^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$', pred, re.IGNORECASE)
                base_tdoc = base_match.group(1).upper() if base_match else pred
                rev_num = 1

            # 1. Tier 0: TDocsByAgenda explicit order
            if row.get('_agenda_order') is not None:
                try:
                    order_val = float(row['_agenda_order'])
                    tier = 0
                except (ValueError, TypeError):
                    pass
            elif base_tdoc and base_tdoc in tdoc_map and tdoc_map[base_tdoc].get('_agenda_order') is not None:
                try:
                    order_val = float(tdoc_map[base_tdoc]['_agenda_order'])
                    sub_item_order = int(tdoc_map[base_tdoc].get('_sub_item_order') or 0)
                    tier = 0
                    rev_offset = rev_num
                except (ValueError, TypeError):
                    pass

            # 2. Tier 1: Excel 'TDoc sort order within agenda item'
            if tier > 0:
                raw_excel_order = row.get("TDoc sort order within agenda item")
                if raw_excel_order not in (None, "", "NaN", "nan", "-"):
                    try:
                        order_val = float(raw_excel_order)
                        tier = 1
                    except (ValueError, TypeError):
                        pass
                elif base_tdoc and base_tdoc in tdoc_map:
                    base_excel = tdoc_map[base_tdoc].get("TDoc sort order within agenda item")
                    if base_excel not in (None, "", "NaN", "nan", "-"):
                        try:
                            order_val = float(base_excel)
                            tier = 1
                            rev_offset = rev_num
                        except (ValueError, TypeError):
                            pass

            tdoc_key = natural_sort_key(tdoc_id)
            row['_agenda_sort_tuple'] = (
                is_unassigned,
                ai_key,
                tier,
                sub_item_order,
                order_val,
                rev_offset,
                tdoc_key
            )

    def _apply_company_sanitization(self, rows_to_process: list):
        """Passes the raw Source string through the Sanitizer and caches confirmed and tentative contributors."""
        for row in rows_to_process:
            source_str = str(row.get('Source', ''))
            confirmed, tentative = CompanySanitizer.parse_contributors_detailed(source_str)
            confirmed_list = list(confirmed)
            tentative_list = list(tentative)
            all_list = confirmed_list + [c for c in tentative_list if c not in confirmed_list]
            # Confirmed contributors only (used by default for filtering and statistics)
            row['_Sanitized_Companies'] = confirmed_list if confirmed_list else ["Other"]
            # Tentative / unconfirmed contributors ('?' or '[...]')
            row['_Tentative_Companies'] = tentative_list
            # Union of confirmed + tentative (used when 'Include unconfirmed (?, []) sources' is enabled)
            row['_All_Sanitized_Companies'] = all_list if all_list else ["Other"]

    def _rebuild_family_cache(self):
        """Precomputes complete bidirectional revision lineage in a single O(N) pass."""
        self._tdoc_map = {str(r.get("TDoc", "")).strip().upper(): r for r in self._data if r.get("TDoc")}
        self._family_cache = {}

        for tdoc_upper, row in self._tdoc_map.items():
            family = {tdoc_upper}

            # Parent revision
            parent = str(row.get("Is revision of", "")).strip().upper()
            if parent and parent in self._tdoc_map:
                family.add(parent)
                # Include siblings listed in parent's "Revised to"
                parent_row = self._tdoc_map[parent]
                p_children = str(parent_row.get("Revised to", "")).strip()
                if p_children and p_children.lower() != "none":
                    for child in p_children.split(","):
                        c_clean = child.strip().upper()
                        if c_clean:
                            family.add(c_clean)

            # Children revisions
            children_str = str(row.get("Revised to", "")).strip()
            if children_str and children_str.lower() != "none":
                for child in children_str.split(","):
                    c_clean = child.strip().upper()
                    if c_clean:
                        family.add(c_clean)

            self._family_cache[tdoc_upper] = sorted(list(family))

    def _recompute_cached_email_stats(self):
        """Precomputes integer totals and sort tuples on each row using unique ID set unions."""
        for row in self._data:
            if row.get("_is_tracked_thread"):
                tot = int(row.get("_email_tot", 0))
                unread = int(row.get("_email_unread", 0))
                flagged = int(row.get("_email_flagged", 0))
                row["_email_sort_key"] = (tot, unread, flagged)
                continue
            tdoc = str(row.get("TDoc", "")).strip().upper()
            family = self.get_family_tdocs(tdoc)

            family_all_ids = set()
            family_unread_ids = set()
            family_flagged_ids = set()

            for t in family:
                stats = self.email_counts.get(t)
                if stats:
                    if "all_ids" in stats:
                        family_all_ids.update(stats["all_ids"])
                        family_unread_ids.update(stats["unread_ids"])
                        family_flagged_ids.update(stats.get("flagged_ids", ()))
                    else:
                        # Fallback for mock or legacy scalar dictionaries
                        tot_num = stats.get("total", 0)
                        un_num = stats.get("unread", 0)
                        fl_num = stats.get("flagged", 0)
                        if tot_num:
                            family_all_ids.add(f"{t}_tot_{tot_num}")
                        if un_num:
                            family_unread_ids.add(f"{t}_un_{un_num}")
                        if fl_num:
                            family_flagged_ids.add(f"{t}_fl_{fl_num}")

            tot = len(family_all_ids)
            unread = len(family_unread_ids)
            flagged = len(family_flagged_ids)
            row['_email_tot'] = tot
            row['_email_unread'] = unread
            row['_email_flagged'] = flagged
            row['_email_sort_key'] = (tot, unread, flagged)

    def get_unmatched_sources(self) -> list:
        """Returns a sorted list of unique raw 'Source' strings that evaluated to 'Other' (neither confirmed nor tentative matched)."""
        unmatched = set()
        for row in self._data:
            if row.get('_All_Sanitized_Companies', row.get('_Sanitized_Companies')) == ["Other"]:
                raw_source = str(row.get('Source', '')).strip()
                if raw_source:
                    unmatched.add(raw_source)
        return sorted(list(unmatched))

    def get_family_tdocs(self, base_tdoc: str) -> list:
        """Instant O(1) lineage retrieval for a TDoc and its revisions."""
        if not base_tdoc:
            return []
        tdoc_clean = str(base_tdoc).strip().upper()
        if tdoc_clean in self._family_cache:
            return self._family_cache[tdoc_clean]

        # Dynamic fallback for external or unindexed documents
        family = {tdoc_clean}
        row = self._tdoc_map.get(tdoc_clean)
        if row:
            parent = str(row.get("Is revision of", "")).strip().upper()
            if parent and parent in self._tdoc_map:
                family.add(parent)
            children_str = str(row.get("Revised to", "")).strip()
            if children_str and children_str.lower() != "none":
                for child in children_str.split(","):
                    c_clean = child.strip().upper()
                    if c_clean:
                        family.add(c_clean)
        res = sorted(list(family))
        self._family_cache[tdoc_clean] = res
        return res

    @staticmethod
    def _split_related_wis(value) -> list:
        """Normalizes the Related WIs cell into individual WI/SI identifiers."""
        if value is None:
            return ["(No WI)"]
        text = str(value).strip()
        if not text or text.lower() in {"none", "nan", "-"}:
            return ["(No WI)"]

        # 3GPP lists may contain more than one WI. Accept the common separators
        # without splitting characters that legitimately occur inside WI acronyms.
        parts = [part.strip() for part in re.split(r"[,;\n]+", text) if part.strip()]
        return parts or ["(No WI)"]

    def get_related_wis(self, row: int) -> list:
        if row < 0 or row >= len(self._data):
            return []
        return self._split_related_wis(self._data[row].get("Related WIs", ""))
