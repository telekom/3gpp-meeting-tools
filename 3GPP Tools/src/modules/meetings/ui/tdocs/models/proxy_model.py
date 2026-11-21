"""Multi-facet filter and custom sorting proxy model for TDocsTableModel."""

import re
from typing import Optional

from PyQt5.QtCore import QSortFilterProxyModel, Qt

from modules.meetings.ui.tdocs.models.sorting import natural_sort_key


class TDocsFilterProxyModel(QSortFilterProxyModel):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.include_terms = []
        self.exclude_terms = []
        self.type_filters = set()
        self.status_filters = set()
        self.ai_filters = set()
        self.sub_item_filters = set()
        self.wi_filters = set()
        self.company_filters = set()
        self.my_status_filters = set()
        self.filter_no_comments = False
        self.filter_followed_wis = False
        self.filter_starred_wis = False
        self.filter_agreements_only = False
        self.include_unassigned_ls_in = True
        self.include_tentative_sources = False
        self.followed_wi_tokens = set()
        self.starred_wi_tokens = set()
        self.sort_mode = "AGENDA"
        self.row_mode_filter = "ALL"
        self.email_state_filter = "ALL"
        self.tag_filters = set()
        self.tag_filters_active = False

    def setTagFilters(self, tags, active: Optional[bool] = None):
        self.tag_filters = set(tags or [])
        if active is not None:
            self.tag_filters_active = bool(active)
        self.invalidateFilter()

    def setRowModeFilter(self, mode: str):
        norm = str(mode or "ALL").upper()
        if self.row_mode_filter != norm:
            self.row_mode_filter = norm
            self.invalidateFilter()

    def setEmailStateFilter(self, mode: str):
        norm = str(mode or "ALL").upper()
        if self.email_state_filter != norm:
            self.email_state_filter = norm
            self.invalidateFilter()

    def setIncludeTentativeSources(self, include_tentative: bool):
        new_val = bool(include_tentative)
        if self.include_tentative_sources != new_val:
            self.include_tentative_sources = new_val
            self.invalidateFilter()

    def setSortMode(self, mode: str):
        self.sort_mode = mode
        self.invalidate()

    def setSubItemFilters(self, sub_items):
        self.sub_item_filters = set(sub_items)
        self.invalidateFilter()

    def _parse_query(self, query: str):
        includes = []
        excludes = []
        pattern = r'(-|!)?(?:["\']([^"\']*)["\']|(\S+))'
        for match in re.finditer(pattern, query):
            is_neg, quoted, unquoted = match.groups()
            term = (quoted if quoted is not None else unquoted).lower().strip()
            if not term:
                continue
            if is_neg:
                excludes.append(term)
            else:
                includes.append(term)
        return includes, excludes

    def setNoCommentsFilter(self, enabled: bool):
        self.filter_no_comments = enabled
        self.invalidateFilter()

    def setGlobalFilter(self, text):
        query = str(text).strip()
        self.include_terms, self.exclude_terms = self._parse_query(query)
        self.invalidateFilter()

    def setTypeFilters(self, types):
        self.type_filters = set(types)
        self.invalidateFilter()

    def setStatusFilters(self, statuses):
        self.status_filters = set(statuses)
        self.invalidateFilter()

    def setAIFilters(self, ais):
        self.ai_filters = set(ais)
        self.invalidateFilter()

    def setWIFilters(self, wis):
        self.wi_filters = set(wis)
        self.invalidateFilter()

    def setCompanyFilters(self, companies):
        self.company_filters = set(companies)
        self.invalidateFilter()

    def setMyStatusFilters(self, statuses):
        self.my_status_filters = set(statuses)
        self.invalidateFilter()

    def setTrackedWITokens(self, followed_tokens: set, starred_tokens: set):
        self.followed_wi_tokens = {str(t).lower().strip() for t in followed_tokens if str(t).strip()}
        self.starred_wi_tokens = {str(t).lower().strip() for t in starred_tokens if str(t).strip()}
        if self.filter_followed_wis or self.filter_starred_wis:
            self.invalidateFilter()

    def setAgreementsFilter(self, enabled: bool):
        new_val = bool(enabled)
        if self.filter_agreements_only != new_val:
            self.filter_agreements_only = new_val
            self.invalidateFilter()

    def setTrackedWIFilters(self, filter_followed: bool, filter_starred: bool):
        self.filter_followed_wis = bool(filter_followed)
        self.filter_starred_wis = bool(filter_starred)
        self.invalidateFilter()

    def setIncludeUnassignedLsIn(self, include_unassigned: bool):
        new_val = bool(include_unassigned)
        if self.include_unassigned_ls_in != new_val:
            self.include_unassigned_ls_in = new_val
            if self.filter_followed_wis or self.filter_starred_wis:
                self.invalidateFilter()

    def rowMatchesFilters(self, source_row, exclude_filter=None):
        """Return whether a source row matches active filters, optionally ignoring one facet."""
        model = self.sourceModel()
        if not model or source_row < 0 or source_row >= model.rowCount():
            return False

        row = model._data[source_row]

        if exclude_filter != "row_mode":
            is_thread = bool(row.get("_is_tracked_thread"))
            if self.row_mode_filter == "TDOCS" and is_thread:
                return False
            if self.row_mode_filter == "THREADS" and not is_thread:
                return False

        if exclude_filter != "email_state" and self.email_state_filter != "ALL":
            unread_cnt = int(row.get("_email_unread", 0) or 0)
            flagged_cnt = int(row.get("_email_flagged", 0) or 0)
            tot_cnt = int(row.get("_email_tot", 0) or 0)
            if self.email_state_filter == "UNREAD" and unread_cnt <= 0:
                return False
            if self.email_state_filter == "FLAGGED" and flagged_cnt <= 0:
                return False
            if self.email_state_filter == "ANY" and tot_cnt <= 0:
                return False

        if exclude_filter != "agreements" and self.filter_agreements_only:
            if not row.get("_has_agreement"):
                return False

        if exclude_filter != "no_comments" and self.filter_no_comments:
            if str(row.get("Secretary Remarks", "") or "").strip():
                return False

        if exclude_filter != "tracked_wi" and (self.filter_followed_wis or self.filter_starred_wis):
            target_tokens = set()
            if self.filter_followed_wis:
                target_tokens |= self.followed_wi_tokens
            if self.filter_starred_wis:
                target_tokens |= self.starred_wi_tokens
            row_tokens = row.get("_wi_tokens")
            if row_tokens is None:
                row_tokens = model._extract_row_wi_tokens(row)
                row["_wi_tokens"] = row_tokens

            doc_type = str(row.get("Type", "") or "").strip().lower()
            raw_wis = model._split_related_wis(row.get("Related WIs", ""))
            has_assigned_wi = any(w and w.lower() not in {"(no wi)", "none", "nan", "-"} for w in raw_wis)
            is_unassigned_ls = ("ls in" in doc_type or doc_type == "ls") and not has_assigned_wi

            wi_match = bool(row_tokens & target_tokens)
            ls_match = bool(self.include_unassigned_ls_in and is_unassigned_ls)

            if not (wi_match or ls_match):
                return False

        if exclude_filter != "type" and str(row.get("Type", "") or "").strip() not in self.type_filters:
            return False
        if exclude_filter != "ai":
            linked_ais = row.get("_linked_ais")
            if linked_ais:
                if not (linked_ais & self.ai_filters):
                    return False
            elif str(row.get("Agenda Item", "") or "").strip() not in self.ai_filters:
                return False
        if exclude_filter != "sub_item" and self.sub_item_filters:
            linked_subs = row.get("_linked_sub_items")
            raw_sub = str(row.get("Sub-Item", "") or "").strip()
            if linked_subs:
                if not (linked_subs & self.sub_item_filters):
                    return False
            elif row.get("_is_tracked_thread") and not raw_sub:
                pass
            else:
                row_sub = raw_sub or "(No Sub-Item)"
                if row_sub not in self.sub_item_filters:
                    return False
        if exclude_filter != "wi":
            row_wis = model.get_related_wis(source_row)
            if not row_wis or not self.wi_filters.intersection(row_wis):
                return False
        if exclude_filter != "status" and str(row.get("TDoc Status", "") or "").strip() not in self.status_filters:
            return False
        if exclude_filter != "my_status" and str(row.get("My Status", "") or "").strip() not in self.my_status_filters:
            return False
        if exclude_filter != "tags" and self.tag_filters_active:
            row_tags = row.get("_tag_list") or []
            if row_tags:
                if not (set(row_tags) & self.tag_filters):
                    return False
            elif "(No Tags)" not in self.tag_filters:
                return False
        if exclude_filter != "company":
            comp_key = "_All_Sanitized_Companies" if getattr(self, "include_tentative_sources", False) else "_Sanitized_Companies"
            row_companies = row.get(comp_key) or row.get("_Sanitized_Companies", ["Other"])
            if not row_companies or not self.company_filters.intersection(row_companies):
                return False

        if self.include_terms or self.exclude_terms:
            searchable_parts = []
            for key in ["TDoc", "Title", "Source", "Abstract", "Secretary Remarks", "My Notes", "My Tags", "Sub-Item"]:
                data = row.get(key, "")
                if data:
                    searchable_parts.append(str(data).lower())
            searchable_parts.append(model._format_related_tdocs(row, html=False).lower())
            row_text = " ".join(searchable_parts)
            if any(exc in row_text for exc in self.exclude_terms):
                return False
            if any(inc not in row_text for inc in self.include_terms):
                return False

        return True

    def lessThan(self, left, right):
        col_idx = left.column()
        model = self.sourceModel()
        if not model or col_idx >= len(model._headers):
            return super().lessThan(left, right)

        col_name = model._headers[col_idx]

        # 1. Agenda Order (applied when clicking Agenda Item header or when in AGENDA sort mode)
        if col_name == "Agenda Item" or getattr(self, "sort_mode", "AGENDA") == "AGENDA":
            left_row = model._data[left.row()]
            right_row = model._data[right.row()]
            left_tuple = left_row.get('_agenda_sort_tuple')
            right_tuple = right_row.get('_agenda_sort_tuple')
            if left_tuple is not None and right_tuple is not None:
                return left_tuple < right_tuple

        # 2. Numerical tuple sorting for Emails (tot, unread)
        if col_name == "Emails":
            left_key = model.data(left, Qt.UserRole + 2) or (0, 0)
            right_key = model.data(right, Qt.UserRole + 2) or (0, 0)
            return left_key < right_key

        # 3. Natural sorting with LRU caching for TDoc
        if col_name == "TDoc":
            left_data = model.data(left, Qt.UserRole + 2) or ""
            right_data = model.data(right, Qt.UserRole + 2) or ""
            return natural_sort_key(left_data) < natural_sort_key(right_data)

        # 4. Standard string sorting for remaining columns without redundant lookups
        return super().lessThan(left, right)

    def filterAcceptsRow(self, source_row, source_parent):
        return self.rowMatchesFilters(source_row)
