"""Contribution indexing, revision lineage analysis, and live scope estimation mixin."""

import re
from typing import Dict, List, Set

from PyQt5.QtWidgets import QTreeWidget

from core.ui.theme import C
from core.utils.company_sanitizer import CompanySanitizer
from modules.meetings.core.agenda_manager import AgendaItem
from modules.meetings.ui.dialogs.pre_meeting_export.styles import natural_ai_sort_key
from modules.meetings.ui.tdocs_models import TDocsTableModel


class PreMeetingAnalysisMixin:
    """Computes revision supersession, contributor analytics, and live TDoc scope estimates."""

    def _extract_row_wis(self, row: dict) -> List[str]:
        """Resolves canonical Work Items associated with a row."""
        raw_wis = TDocsTableModel._split_related_wis(row.get("Related WIs", ""))
        wis = []
        for rw in raw_wis:
            clean = rw.strip()
            if clean and clean.lower() not in {"(no wi)", "none", "nan", "-"}:
                if clean not in wis:
                    wis.append(clean)
        if not wis:
            ai_num = str(row.get("Agenda Item", "")).strip()
            if ai_num and self.agenda_map:
                item = self.agenda_map.get(ai_num)
                if item and getattr(item, "acronym", None):
                    acr = item.acronym.strip()
                    if acr and acr not in wis:
                        wis.append(acr)
            if not wis:
                desc = str(row.get("Agenda item description", "")).strip()
                if desc:
                    cand = AgendaItem._extract_acronym(desc)
                    if cand and cand not in wis:
                        wis.append(cand.strip())
        if not wis:
            wis = ["General"]
        return wis

    def _analyze_meeting_contributions(self):
        """Builds indices of Work Items, Agenda Items, TDoc distributions, and contributor companies."""
        wis_set = set()
        ais_set = set()
        companies_set = set()
        self.ai_descriptions: Dict[str, str] = {}
        self.wi_to_ais: Dict[str, Dict[str, int]] = {}
        self.ai_to_wis: Dict[str, Dict[str, int]] = {}
        self.wi_counts: Dict[str, int] = {}
        self.ai_counts: Dict[str, int] = {}

        # Precompute revision lineage to support live estimation of deduplication
        tdoc_map = {str(r.get("TDoc", "")).strip().upper(): r for r in self.all_rows if r.get("TDoc")}
        forward_edges: Dict[str, Set[str]] = {}
        for tdoc_id, r in tdoc_map.items():
            rev_to = str(r.get("Revised to", "")).strip()
            if rev_to and rev_to.lower() != "none":
                for cand in re.findall(r'[A-Za-z0-9]+-\d+(?:r\d+[a-zA-Z]?)?', rev_to):
                    cand_upper = cand.upper()
                    if cand_upper in tdoc_map and cand_upper != tdoc_id:
                        forward_edges.setdefault(tdoc_id, set()).add(cand_upper)

            rev_of = str(r.get("Is revision of", "")).strip()
            if rev_of and rev_of.lower() != "none":
                for cand in re.findall(r'[A-Za-z0-9]+-\d+(?:r\d+[a-zA-Z]?)?', rev_of):
                    cand_upper = cand.upper()
                    if cand_upper in tdoc_map and cand_upper != tdoc_id:
                        forward_edges.setdefault(cand_upper, set()).add(tdoc_id)

            base_match = re.search(r'^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$', tdoc_id, re.IGNORECASE)
            if base_match:
                base_tdoc = base_match.group(1).upper()
                if base_tdoc in tdoc_map and base_tdoc != tdoc_id:
                    forward_edges.setdefault(base_tdoc, set()).add(tdoc_id)

        self.superseded_tdocs: Set[str] = set()
        for tdoc_id, r in tdoc_map.items():
            st = str(r.get("TDoc Status", "")).strip().lower()
            if "withdrawn" in st:
                continue
            curr = tdoc_id
            visited = {curr}
            while True:
                children = forward_edges.get(curr, set())
                active_children = [c for c in children if c not in visited and "withdrawn" not in str(tdoc_map[c].get("TDoc Status", "")).strip().lower()]
                if not active_children:
                    break
                active_children.sort()
                next_node = active_children[-1]
                visited.add(next_node)
                curr = next_node
            if curr != tdoc_id:
                self.superseded_tdocs.add(tdoc_id)

        for r in self.all_rows:
            ai_num = str(r.get("Agenda Item", "")).strip() or "Unassigned"
            ais_set.add(ai_num)

            # Store best description for AI
            if ai_num not in self.ai_descriptions:
                desc = ""
                if ai_num in self.agenda_map:
                    item = self.agenda_map[ai_num]
                    desc = getattr(item, "description", "") or ""
                if not desc:
                    desc = str(r.get("Agenda item description", "")).strip()
                self.ai_descriptions[ai_num] = desc

            # Collect companies
            companies = r.get("_All_Sanitized_Companies", r.get("_Sanitized_Companies", []))
            if not companies:
                confirmed, tentative = CompanySanitizer.parse_contributors_detailed(r.get("Source", ""))
                companies = list(confirmed) or ["Other"]
            for comp in companies:
                if comp and comp != "Other":
                    companies_set.add(comp)

            row_wis = self._extract_row_wis(r)
            for w in row_wis:
                if w != "General":
                    wis_set.add(w)

                if w not in self.wi_to_ais:
                    self.wi_to_ais[w] = {}
                self.wi_to_ais[w][ai_num] = self.wi_to_ais[w].get(ai_num, 0) + 1
                self.wi_counts[w] = self.wi_counts.get(w, 0) + 1

                if ai_num not in self.ai_to_wis:
                    self.ai_to_wis[ai_num] = {}
                self.ai_to_wis[ai_num][w] = self.ai_to_wis[ai_num].get(w, 0) + 1
                self.ai_counts[ai_num] = self.ai_counts.get(ai_num, 0) + 1

        if "General" in self.wi_to_ais:
            wis_set.add("General")

        self.meeting_wis = sorted(list(wis_set), key=lambda x: (x == "General", x.lower()))
        self.meeting_ais = sorted(list(ais_set), key=natural_ai_sort_key)
        self.meeting_companies = sorted(list(companies_set), key=lambda x: x.lower())

        self.followed_wis = [w for w in self.meeting_wis if (TDocsTableModel._extract_base_tokens(w) & self.followed_tokens)]
        self.starred_wis = [w for w in self.meeting_wis if (TDocsTableModel._extract_base_tokens(w) & self.starred_tokens)]
        self.tracked_wis = [w for w in self.meeting_wis if (TDocsTableModel._extract_base_tokens(w) & self.tracked_tokens)]

    def _initialize_default_custom_selection(self):
        """Pre-checks Followed & Starred WIs (and all their child AIs) automatically."""
        self.selected_wi_ais.clear()
        has_tracked = bool(self.tracked_wis)

        for wi in self.meeting_wis:
            wi_tokens = TDocsTableModel._extract_base_tokens(wi)
            is_tracked = bool(wi_tokens & self.tracked_tokens)
            if is_tracked or (not has_tracked):
                for ai in self.wi_to_ais.get(wi, {}):
                    self.selected_wi_ais.add((wi, ai))

    def _sync_scope_ui_state(self):
        is_custom = self.rb_custom.isChecked()
        if is_custom:
            self.lbl_custom_hint.setText("⚡ <b>Custom Selection:</b> Choose specific Work Items and Agenda Items to export:")
            self.lbl_custom_hint.setStyleSheet(f"font-size: 11px; color: {C.TEAL_DARK};")
        else:
            self.lbl_custom_hint.setText("💡 <i>Check any item below or select 'Custom Selection' to customize scope:</i>")
            self.lbl_custom_hint.setStyleSheet(f"font-size: 11px; color: {C.TEXT_MUTED};")

    def _on_scope_radio_toggled(self, checked: bool):
        self._sync_scope_ui_state()
        if self.rb_custom.isChecked():
            active_tree = self.tree_stack.currentWidget()
            if isinstance(active_tree, QTreeWidget):
                self._sync_tree_from_selection(active_tree)
        self._update_estimates()

    def _update_estimates(self):
        scope_mode = self.get_scope_mode()
        include_ls_in = self.chk_ls_in.isChecked()
        include_ls_out = self.chk_ls_out.isChecked()
        include_pcr = self.chk_pcr.isChecked()
        include_cr = self.chk_cr.isChecked()
        include_disc = self.chk_disc.isChecked()
        include_general_ls = self.chk_general_ls.isChecked()
        dedup_revs = self.chk_dedup_revs.isChecked()
        include_reserved = self.chk_include_reserved.isChecked()

        selected_comp = self.combo_company.currentData()
        company_filter = {selected_comp.lower()} if selected_comp else set()

        matching_tdocs = set()
        superseded_hidden_count = 0
        scope_counts = {"LS_IN": 0, "LS_OUT": 0, "pCR": 0, "CR": 0, "DISC": 0}
        general_ls_count = 0
        matching_wis = set()
        matching_ais = set()
        matching_companies = set()

        visible_ids = {str(r.get("TDoc", "")).strip().upper() for r in self.visible_rows if r.get("TDoc")}

        for row in self.all_rows:
            tdoc_id = str(row.get("TDoc", "")).strip().upper()
            if not tdoc_id:
                continue

            status = str(row.get("TDoc Status", "")).strip().lower()
            if "withdrawn" in status:
                continue

            if dedup_revs and tdoc_id in self.superseded_tdocs:
                superseded_hidden_count += 1
                continue

            is_reserved = any(res in status for res in ["reserved", "not uploaded", "not yet uploaded"])
            if is_reserved and not include_reserved:
                continue

            doc_type = str(row.get("Type", "")).strip().lower()

            category = None
            if "ls in" in doc_type or doc_type == "ls":
                category = "LS_IN"
            elif "ls out" in doc_type or "ls_out" in doc_type:
                category = "LS_OUT"
            elif "pcr" in doc_type or "p-cr" in doc_type:
                category = "pCR"
            elif "cr" in doc_type:
                category = "CR"
            elif any(d in doc_type for d in ["disc", "discussion", "report", "wid", "sid", "other", "draft", "evaluation"]):
                category = "DISC"

            if not category:
                continue

            cat_enabled = (
                (category == "LS_IN" and include_ls_in) or
                (category == "LS_OUT" and include_ls_out) or
                (category == "pCR" and include_pcr) or
                (category == "CR" and include_cr) or
                (category == "DISC" and include_disc)
            )

            if scope_mode == "visible" and tdoc_id not in visible_ids:
                continue

            # Check company filter
            row_comps = row.get("_All_Sanitized_Companies", row.get("_Sanitized_Companies", []))
            row_comps_lower = {str(c).strip().lower() for c in row_comps}
            if company_filter and not (row_comps_lower & company_filter):
                continue

            row_wis = self._extract_row_wis(row)
            ai_num = str(row.get("Agenda Item", "")).strip() or "Unassigned"

            matches_scope = False
            matched_wis_for_row = []

            for wi in row_wis:
                if wi == "General":
                    if scope_mode in ["whole", "visible"] or ((category in ("LS_IN", "LS_OUT")) and include_general_ls):
                        matches_scope = True
                        matched_wis_for_row.append(wi)
                    elif scope_mode == "custom" and (
                        ("General", ai_num) in self.selected_wi_ais
                        or ("General", "Unassigned") in self.selected_wi_ais
                        or (include_general_ls and category in ("LS_IN", "LS_OUT"))
                    ):
                        matches_scope = True
                        matched_wis_for_row.append(wi)
                    continue

                wi_tokens = TDocsTableModel._extract_base_tokens(wi)
                if scope_mode == "followed":
                    if wi_tokens & self.followed_tokens:
                        matches_scope = True
                        matched_wis_for_row.append(wi)
                elif scope_mode == "starred":
                    if wi_tokens & self.starred_tokens:
                        matches_scope = True
                        matched_wis_for_row.append(wi)
                elif scope_mode == "tracked":
                    if wi_tokens & self.tracked_tokens:
                        matches_scope = True
                        matched_wis_for_row.append(wi)
                elif scope_mode == "custom":
                    if (wi, ai_num) in self.selected_wi_ais:
                        matches_scope = True
                        matched_wis_for_row.append(wi)
                    else:
                        for sel_wi, sel_ai in self.selected_wi_ais:
                            if sel_ai == ai_num and (TDocsTableModel._extract_base_tokens(sel_wi) & wi_tokens):
                                matches_scope = True
                                matched_wis_for_row.append(wi)
                                break
                else:  # whole
                    matches_scope = True
                    matched_wis_for_row.append(wi)

            if not matches_scope and (category in ("LS_IN", "LS_OUT")) and include_general_ls:
                matches_scope = True
                matched_wis_for_row.append("General")

            if matches_scope:
                scope_counts[category] += 1
                if "General" in matched_wis_for_row and category in ("LS_IN", "LS_OUT"):
                    general_ls_count += 1
                if cat_enabled:
                    matching_tdocs.add(tdoc_id)
                    for w in matched_wis_for_row:
                        matching_wis.add(w)
                    matching_ais.add(ai_num)
                    for c in row_comps:
                        if c and c != "Other":
                            matching_companies.add(c)

        # Update Section 2 checkbox labels with dynamic counts
        self.chk_ls_in.setText(f"📥 Incoming Liaison Statements (LS in) — {scope_counts['LS_IN']} docs")
        self.chk_ls_out.setText(f"📤 Proposed Outgoing LS Replies (LS out) — {scope_counts['LS_OUT']} docs")
        self.chk_pcr.setText(f"📝 Proposed Pseudo-CRs (pCR) — {scope_counts['pCR']} docs")
        self.chk_cr.setText(f"📄 Proposed Change Requests (CR) — {scope_counts['CR']} docs")
        self.chk_disc.setText(f"💬 Discussion, Evaluation && WID Papers (Disc) — {scope_counts['DISC']} docs")
        self.chk_general_ls.setText(f"📦 Include unassigned && general LS — {general_ls_count} docs")

        total = len(matching_tdocs)
        if total > 0:
            parts = []
            if self.chk_pcr.isChecked() and scope_counts['pCR']:
                parts.append(f"{scope_counts['pCR']} pCRs")
            if self.chk_cr.isChecked() and scope_counts['CR']:
                parts.append(f"{scope_counts['CR']} CRs")
            if self.chk_disc.isChecked() and scope_counts['DISC']:
                parts.append(f"{scope_counts['DISC']} Disc")
            if self.chk_ls_in.isChecked() and scope_counts['LS_IN']:
                parts.append(f"{scope_counts['LS_IN']} LS in")
            if self.chk_ls_out.isChecked() and scope_counts['LS_OUT']:
                parts.append(f"{scope_counts['LS_OUT']} LS out")
            summary_str = f" ({', '.join(parts)})" if parts else ""

            dedup_note = f" [{superseded_hidden_count} older revisions hidden]" if (dedup_revs and superseded_hidden_count > 0) else ""
            n_wis = len(matching_wis)
            n_ais = len(matching_ais)
            n_comps = len(matching_companies)
            self.lbl_estimate.setText(
                f"📊 <b>Estimated Pre-Meeting Scope:</b> {total} Active Proposals across {n_wis} Work Items, "
                f"{n_ais} Agenda Items, and {n_comps} Companies{summary_str}{dedup_note}"
            )
            self.lbl_estimate.setStyleSheet(f"""
                QLabel {{
                    background-color: {C.SUCCESS_SOFT};
                    border: 1px solid {C.SUCCESS_BORDER};
                    border-radius: 6px;
                    padding: 6px 10px;
                    font-size: 11px;
                    color: {C.SUCCESS_DARK};
                    font-weight: 500;
                }}
            """)
        else:
            self.lbl_estimate.setText(
                "⚠️ <b>Estimated Scope:</b> 0 TDocs match current criteria. Adjust scope, filters, or document types."
            )
            self.lbl_estimate.setStyleSheet(f"""
                QLabel {{
                    background-color: {C.WARNING_SOFT};
                    border: 1px solid {C.WARNING_BORDER};
                    border-radius: 6px;
                    padding: 6px 10px;
                    font-size: 11px;
                    color: {C.WARNING};
                    font-weight: 500;
                }}
            """)
