# --- File: src/modules/meetings/ui/tdocs/window/mixins/filter_facets.py ---
"""
Filter facet counts, combobox population, AI tooltips and clear-all-filters.

Mixin of TDocsWindow: relies on attributes set in TDocsWindow.__init__ / _setup_*.
"""
import re
from PyQt5.QtCore import Qt, QTimer
from modules.meetings.ui.tdocs_models import natural_sort_key


class FilterFacetsMixin:
    """Filter facet counts, combobox population, AI tooltips and clear-all-filters."""

    def _update_filter_facets(self):
        """Update zero-intersection styling for every metadata filter."""
        if not hasattr(self, "proxy") or not hasattr(self, "model"):
            return

        facets = {
            "company": set(),
            "type": set(),
            "ai": set(),
            "sub_item": set(),
            "wi": set(),
            "status": set(),
            "my_status": set(),
            "tags": set(),
        }

        followed_count = 0
        starred_count = 0
        has_followed_tokens = bool(self.proxy.followed_wi_tokens)
        has_starred_tokens = bool(self.proxy.starred_wi_tokens)

        # Each facet is evaluated with all active constraints except itself.
        # This answers: "Would selecting this option yield rows given my other filters?"
        for row_idx, row in enumerate(self.model._data):
            if self.proxy.rowMatchesFilters(row_idx, exclude_filter="company"):
                comp_key = "_All_Sanitized_Companies" if getattr(self.proxy, "include_tentative_sources", False) else "_Sanitized_Companies"
                facets["company"].update(row.get(comp_key) or row.get("_Sanitized_Companies", ["Other"]))
            if self.proxy.rowMatchesFilters(row_idx, exclude_filter="type"):
                facets["type"].add(str(row.get("Type", "") or "").strip())
            if self.proxy.rowMatchesFilters(row_idx, exclude_filter="ai"):
                if row.get("_linked_ais"):
                    facets["ai"].update(row["_linked_ais"])
                else:
                    facets["ai"].add(str(row.get("Agenda Item", "") or "").strip())
            if self.proxy.rowMatchesFilters(row_idx, exclude_filter="sub_item"):
                if row.get("_linked_sub_items"):
                    facets["sub_item"].update(row["_linked_sub_items"])
                else:
                    sub_val = str(row.get("Sub-Item", "") or "").strip()
                    if sub_val:
                        facets["sub_item"].add(sub_val)
            if self.proxy.rowMatchesFilters(row_idx, exclude_filter="wi"):
                facets["wi"].update(self.model.get_related_wis(row_idx))
            if self.proxy.rowMatchesFilters(row_idx, exclude_filter="status"):
                facets["status"].add(str(row.get("TDoc Status", "") or "").strip())
            if self.proxy.rowMatchesFilters(row_idx, exclude_filter="my_status"):
                facets["my_status"].add(str(row.get("My Status", "") or "").strip())
            if self.proxy.rowMatchesFilters(row_idx, exclude_filter="tags"):
                row_tags = row.get("_tag_list") or []
                if row_tags:
                    facets["tags"].update(row_tags)
                else:
                    facets["tags"].add("(No Tags)")

            if has_followed_tokens or has_starred_tokens:
                if self.proxy.rowMatchesFilters(row_idx, exclude_filter="tracked_wi"):
                    row_tokens = row.get("_wi_tokens")
                    if row_tokens is None:
                        row_tokens = self.model._extract_row_wi_tokens(row)
                        row["_wi_tokens"] = row_tokens

                    doc_type = str(row.get("Type", "") or "").strip().lower()
                    raw_wis = self.model._split_related_wis(row.get("Related WIs", ""))
                    has_assigned_wi = any(w and w.lower() not in {"(no wi)", "none", "nan", "-"} for w in raw_wis)
                    is_unassigned_ls = ("ls in" in doc_type or doc_type == "ls") and not has_assigned_wi

                    wi_followed = bool(has_followed_tokens and (row_tokens & self.proxy.followed_wi_tokens))
                    wi_starred = bool(has_starred_tokens and (row_tokens & self.proxy.starred_wi_tokens))
                    ls_included = bool(self.proxy.include_unassigned_ls_in and is_unassigned_ls)

                    if wi_followed or (has_followed_tokens and ls_included):
                        followed_count += 1
                    if wi_starred or (has_starred_tokens and ls_included):
                        starred_count += 1

        self.company_combo.setAvailability(facets["company"])
        self.type_combo.setAvailability(facets["type"])
        self.ai_combo.setAvailability(facets["ai"])
        if hasattr(self, "sub_item_combo") and self.sub_item_combo.isVisible():
            self.sub_item_combo.setAvailability(facets["sub_item"])
        self.wi_combo.setAvailability(facets["wi"])
        self.status_combo.setAvailability(facets["status"])
        self.my_status_combo.setAvailability(facets["my_status"])
        if hasattr(self, "tags_combo") and self.tags_combo.isVisible():
            self.tags_combo.setAvailability(facets["tags"])

        ls_suffix = " (incl. unassigned LS in)" if getattr(self.proxy, "include_unassigned_ls_in", True) else ""

        if hasattr(self, "btn_followed_wis"):
            self.btn_followed_wis.setText("🎯")
            self.btn_followed_wis.setToolTip(
                f"🎯 Followed Work Items ({followed_count} matching{ls_suffix})\n"
                "Filter to TDocs associated with Work Items you are following (assigned responsibility).\n"
                "Right-click to toggle including unassigned Incoming LS."
            )
        if hasattr(self, "btn_starred_wis"):
            self.btn_starred_wis.setText("⭐")
            self.btn_starred_wis.setToolTip(
                f"⭐ Starred Work Items ({starred_count} matching{ls_suffix})\n"
                "Filter to TDocs associated with Work Items you have starred (personal watchlist).\n"
                "Right-click to toggle including unassigned Incoming LS."
            )

    def _refresh_comboboxes(self):
        def sanitize(val): return str(val).strip() if val is not None else ""

        unique_types = sorted(list(set(sanitize(r.get("Type", "")) for r in self.model._data)))
        all_ais = set()
        for r in self.model._data:
            if r.get("_linked_ais"):
                all_ais.update(sanitize(a) for a in r["_linked_ais"])
            all_ais.add(sanitize(r.get("Agenda Item", "")))
        unique_ais = sorted(list(all_ais), key=natural_sort_key)
        unique_wis = sorted({
            wi
            for row_idx in range(self.model.rowCount())
            for wi in self.model.get_related_wis(row_idx)
        }, key=lambda x: x.lower())
        unique_statuses = sorted(list(set(sanitize(r.get("TDoc Status", "")) for r in self.model._data)))
        unique_my_statuses = sorted(list(set(sanitize(r.get("My Status", "")) for r in self.model._data)))

        unique_companies = set()
        comp_key = "_All_Sanitized_Companies" if getattr(self.proxy, "include_tentative_sources", False) else "_Sanitized_Companies"
        for r in self.model._data:
            unique_companies.update(r.get(comp_key) or r.get("_Sanitized_Companies", ["Other"]))
        sorted_companies = sorted(list(unique_companies), key=lambda x: x.lower())

        ai_display_labels = {}
        for ai_num in unique_ais:
            agenda_item = self.agenda_map.get(ai_num)
            if agenda_item:
                ai_display_labels[ai_num] = agenda_item.display_label
            else:
                excel_desc = next((str(r.get("Agenda item description", "")).strip()
                                   for r in self.model._data
                                   if str(r.get("Agenda Item", "")).strip() == ai_num and r.get("Agenda item description")), "")
                if excel_desc:
                    clean = re.sub(r'\s+', ' ', excel_desc)
                    max_desc_len = max(15, 38 - len(ai_num))
                    label_desc = clean if len(clean) <= max_desc_len else clean[:max_desc_len - 3].rstrip() + "..."
                    ai_display_labels[ai_num] = f"{ai_num} - {label_desc}"
                else:
                    ai_display_labels[ai_num] = ai_num

        self.type_combo.updateItems(unique_types)
        self.ai_combo.updateItems(unique_ais, display_labels=ai_display_labels)
        self._update_ai_tooltips()
        self.wi_combo.updateItems(unique_wis)
        self.status_combo.updateItems(unique_statuses)
        self.company_combo.updateItems(sorted_companies)
        self.my_status_combo.updateItems(unique_my_statuses)
        self._load_tracked_wi_tokens()
        self._update_tracked_wi_button_styles()

        unique_sub_items = []
        seen_subs = set()
        for r in self.model._data:
            for sub in (list(r.get("_linked_sub_items") or []) or [str(r.get("Sub-Item", "") or "").strip()]):
                sub = str(sub or "").strip()
                if sub and sub not in seen_subs:
                    seen_subs.add(sub)
                    unique_sub_items.append(sub)

        if hasattr(self, "sub_item_combo"):
            if unique_sub_items:
                self.sub_item_combo.setVisible(True)
                self.sub_item_combo.updateItems(unique_sub_items)
                self.proxy.setSubItemFilters(self.sub_item_combo.getCheckedItems())
            else:
                self.sub_item_combo.setVisible(False)
                self.proxy.setSubItemFilters([])

        unique_tags = self.db.get_all_unique_tags() if hasattr(self, "db") else []
        if hasattr(self, "tags_combo"):
            if unique_tags:
                tag_items = list(unique_tags) + ["(No Tags)"]
                self.tags_combo.setVisible(True)
                self.tags_combo.updateItems(tag_items)
                all_checked = (
                    self.tags_combo.model().rowCount() > 0
                    and self.tags_combo.model().item(0).checkState() == Qt.Checked
                )
                self.proxy.setTagFilters(self.tags_combo.getCheckedItems(), active=not all_checked)
            else:
                self.tags_combo.setVisible(False)
                self.proxy.setTagFilters([], active=False)

        self.proxy.setTypeFilters(self.type_combo.getCheckedItems())
        self.proxy.setAIFilters(self.ai_combo.getCheckedItems())
        self.proxy.setWIFilters(self.wi_combo.getCheckedItems())
        self.proxy.setStatusFilters(self.status_combo.getCheckedItems())
        self.proxy.setCompanyFilters(self.company_combo.getCheckedItems())
        self.proxy.setMyStatusFilters(self.my_status_combo.getCheckedItems())
        self._update_filter_facets()

        QTimer.singleShot(0, self._update_count_label)

    def _update_ai_tooltips(self):
        if not hasattr(self, 'ai_combo') or not self.ai_combo.model():
            return
        model = self.ai_combo.model()
        for i in range(1, model.rowCount()):
            item = model.item(i)
            if not item:
                continue
            ai_num = str(item.data(Qt.UserRole) or "").strip()
            agenda_item = self.agenda_map.get(ai_num)
            if agenda_item:
                if hasattr(agenda_item, 'full_tooltip'):
                    full_desc = agenda_item.full_tooltip
                elif hasattr(agenda_item, 'description'):
                    full_desc = f"AI {ai_num}: {agenda_item.description}"
                else:
                    full_desc = f"AI {ai_num}: {agenda_item}"
                item.setData(full_desc, Qt.ToolTipRole)
            else:
                excel_desc = next((str(r.get("Agenda item description", "")).strip()
                                   for r in self.model._data
                                   if str(r.get("Agenda Item", "")).strip() == ai_num and r.get("Agenda item description")), "")
                if excel_desc:
                    item.setData(f"AI {ai_num}: {excel_desc}", Qt.ToolTipRole)
                else:
                    item.setData(f"Agenda Item {ai_num}", Qt.ToolTipRole)

    def _clear_all_filters(self):
        self.search_input.blockSignals(True)
        self.search_input.clear()
        self.search_input.blockSignals(False)
        self.proxy.setGlobalFilter("")
        if self.is_sa2:
            self.chk_no_comments.blockSignals(True)
            self.chk_no_comments.setChecked(False)
            self.chk_no_comments.blockSignals(False)
            self.proxy.setNoCommentsFilter(False)

        if hasattr(self, "btn_followed_wis"):
            self.btn_followed_wis.blockSignals(True)
            self.btn_followed_wis.setChecked(False)
            self.btn_followed_wis.blockSignals(False)

        if hasattr(self, "btn_starred_wis"):
            self.btn_starred_wis.blockSignals(True)
            self.btn_starred_wis.setChecked(False)
            self.btn_starred_wis.blockSignals(False)

        if hasattr(self, "btn_agreements_filter"):
            self.btn_agreements_filter.blockSignals(True)
            self.btn_agreements_filter.setChecked(False)
            self.btn_agreements_filter.blockSignals(False)

        for btn_attr in ("btn_unread_emails", "btn_flagged_emails"):
            if hasattr(self, btn_attr):
                b = getattr(self, btn_attr)
                b.blockSignals(True)
                b.setChecked(False)
                b.blockSignals(False)

        self.proxy.setTrackedWIFilters(False, False)
        self.proxy.setAgreementsFilter(False)
        self.proxy.setEmailStateFilter("ALL")
        if hasattr(self, "row_mode_combo"):
            self.row_mode_combo.blockSignals(True)
            self.row_mode_combo.setCurrentIndex(0)
            self.row_mode_combo.blockSignals(False)
        self.proxy.setRowModeFilter("ALL")
        self._update_tracked_wi_button_styles()

        for combo in [self.company_combo, self.type_combo, self.ai_combo, self.wi_combo, self.status_combo, self.my_status_combo]:
            combo.blockSignals(True)
            combo.model().item(0).setCheckState(Qt.Checked)
            for i in range(1, combo.model().rowCount()):
                combo.model().item(i).setCheckState(Qt.Checked)
            combo.updateText()
            combo.blockSignals(False)

        if hasattr(self, 'sub_item_combo') and self.sub_item_combo.isVisible():
            self.sub_item_combo.blockSignals(True)
            if self.sub_item_combo.model() and self.sub_item_combo.model().rowCount() > 0:
                self.sub_item_combo.model().item(0).setCheckState(Qt.Checked)
                for i in range(1, self.sub_item_combo.model().rowCount()):
                    self.sub_item_combo.model().item(i).setCheckState(Qt.Checked)
                self.sub_item_combo.updateText()
            self.sub_item_combo.blockSignals(False)
            self.proxy.setSubItemFilters(self.sub_item_combo.getCheckedItems())

        if hasattr(self, 'tags_combo') and self.tags_combo.isVisible():
            self.tags_combo.blockSignals(True)
            if self.tags_combo.model() and self.tags_combo.model().rowCount() > 0:
                self.tags_combo.model().item(0).setCheckState(Qt.Checked)
                for i in range(1, self.tags_combo.model().rowCount()):
                    self.tags_combo.model().item(i).setCheckState(Qt.Checked)
                self.tags_combo.updateText()
            self.tags_combo.blockSignals(False)
            self.proxy.setTagFilters(self.tags_combo.getCheckedItems(), active=False)

        self.proxy.setTypeFilters(self.type_combo.getCheckedItems())
        self.proxy.setAIFilters(self.ai_combo.getCheckedItems())
        self.proxy.setWIFilters(self.wi_combo.getCheckedItems())
        self.proxy.setStatusFilters(self.status_combo.getCheckedItems())
        self.proxy.setCompanyFilters(self.company_combo.getCheckedItems())
        self.proxy.setMyStatusFilters(self.my_status_combo.getCheckedItems())
        self._update_filter_facets()

        QTimer.singleShot(0, self._update_count_label)
