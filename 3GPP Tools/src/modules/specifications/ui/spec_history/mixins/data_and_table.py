"""Numeric sort item, data loading, filtering, and table rendering for SpecVersionHistoryDialog."""

from typing import Dict

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QBrush, QColor
from PyQt5.QtWidgets import QTableWidgetItem

from core.ui.theme import C
from modules.specifications.core.spec_history_extractor import SpecHistoryExtractor


class NumericSortItem(QTableWidgetItem):
    """Sorts table rows numerically when possible, falling back to natural string sorting."""

    def __lt__(self, other):
        try:
            return int(self.text().strip()) < int(other.text().strip())
        except (ValueError, TypeError):
            return super().__lt__(other)


class _HistoryDataMixin:
    """Handles history DB loading, version/clause filtering, origin CR ranking, and table rows."""

    def _load_data(self, auto_extract_if_empty: bool = True):
        """Loads change history from 3gpp_spec_history.db, auto-extracting from local docx if empty."""
        records = self.history_db.get_spec_history(self.spec_number)

        if not records and auto_extract_if_empty:
            # Check if local docx exists and trigger fast extraction
            doc_file = SpecHistoryExtractor.resolve_spec_document(self.spec_number)
            if doc_file and doc_file.exists():
                extracted = SpecHistoryExtractor.extract_from_docx(doc_file, self.spec_number)
                if extracted:
                    self.history_db.upsert_history_records(extracted)
                    records = self.history_db.get_spec_history(self.spec_number)

        self.all_records = records

        # Populate version combo box
        self.version_combo.blockSignals(True)
        self.version_combo.clear()
        self.version_combo.addItem("All Versions", userData="")

        version_counts: Dict[str, int] = {}
        for r in self.all_records:
            v = r.get("version", "")
            if v:
                version_counts[v] = version_counts.get(v, 0) + 1

        # Natural descending sort of version strings
        def ver_key(v):
            return [(0, int(x)) if x.isdigit() else (1, x) for x in v.split(".")]

        sorted_vers = sorted(version_counts.keys(), key=ver_key, reverse=True)
        target_idx = 0
        for idx, ver in enumerate(sorted_vers):
            cnt = version_counts[ver]
            self.version_combo.addItem(f"v{ver} ({cnt} entry)", userData=ver)
            if self.initial_version and ver == self.initial_version:
                target_idx = idx + 1

        if self.sort_by_origin_relevance and not self.initial_version:
            target_idx = 0
        self.version_combo.setCurrentIndex(target_idx)
        self.version_combo.blockSignals(False)

        self._apply_filters()

    def _on_version_changed(self):
        self._apply_filters()

    def _apply_filters(self):
        """Applies version dropdown filter, clause filter, and origin ranking."""
        if not hasattr(self, "table") or not hasattr(self, "version_combo") or not hasattr(self, "search_input"):
            return
        target_ver = self.version_combo.currentData() or ""
        search_txt = self.search_input.text().strip()
        search_lower = search_txt.lower()
        clause_val = self.clause_input.text().strip() if hasattr(self, "clause_input") and self.chk_clause_filter.isChecked() else ""
        is_strict_clause = bool(clause_val) and hasattr(self, "clause_mode_combo") and self.clause_mode_combo.currentText() == "Strict Clause Filter"
        is_origin_ranking = hasattr(self, "chk_origin_ranking") and self.chk_origin_ranking.isChecked()

        if is_origin_ranking and (search_txt or clause_val):
            from modules.specifications.core.spec_clause_attribution import SpecClauseAttributionEngine
            if not self._attribution_engine:
                self._attribution_engine = SpecClauseAttributionEngine(history_db=self.history_db)

            hint_clauses = [clause_val] if clause_val else None
            ranked = self._attribution_engine.rank_origin_crs_for_query(
                spec_number=self.spec_number,
                query=search_txt or clause_val,
                target_version=target_ver or None,
                affected_clauses_hint=hint_clauses,
            )

            # Apply version filter if explicitly selected
            if target_ver:
                ranked = [r for r in ranked if r.get("version") == target_ver]

            # Apply strict clause filter if enabled
            if is_strict_clause:
                import re as _re
                ranked = [
                    r for r in ranked
                    if clause_val in r.get("affected_clauses", "") or _re.search(rf"\b{_re.escape(clause_val)}\b", r.get("subject", ""))
                ]

            filtered = ranked
        else:
            filtered = []
            import re as _re
            for r in self.all_records:
                # 1. Version filter
                if target_ver and r.get("version") != target_ver:
                    continue

                # 2. Strict Clause filter
                if is_strict_clause:
                    aff = r.get("affected_clauses", "")
                    subj = r.get("subject", "")
                    if not (clause_val in aff or _re.search(rf"\b{_re.escape(clause_val)}\b", subj)):
                        continue

                # 3. Substring filter
                if search_lower:
                    row_str = " ".join([
                        r.get("meeting", ""),
                        r.get("source_tdoc", ""),
                        r.get("cr_number", ""),
                        r.get("subject", ""),
                        r.get("category", ""),
                        r.get("affected_clauses", ""),
                        " ".join(t["raw_tdoc"] for t in r.get("tdocs", []))
                    ]).lower()
                    if search_lower not in row_str:
                        continue

                # Additive clause boost if additive mode is active
                if clause_val and not is_strict_clause:
                    aff = r.get("affected_clauses", "")
                    subj = r.get("subject", "")
                    if clause_val in aff or _re.search(rf"\b{_re.escape(clause_val)}\b", subj):
                        r = dict(r)
                        r["score"] = r.get("score", 0) + 30
                        r.setdefault("evidence_badges", []).append(f"🎯 Clause Match: {clause_val}")

                filtered.append(r)

            if clause_val and not is_strict_clause:
                filtered.sort(key=lambda x: x.get("score", 0), reverse=True)

        self.filtered_records = filtered
        self._populate_table()

    def _populate_table(self):
        """Renders the filtered records into QTableWidget."""
        self.table.setRowCount(0)
        self.table.setSortingEnabled(False)

        is_ranking_active = hasattr(self, "chk_origin_ranking") and self.chk_origin_ranking.isChecked()

        if is_ranking_active:
            self.table.setColumnCount(9)
            self.table.setHorizontalHeaderLabels([
                "Relevance", "Date", "Meeting", "Version", "Type", "CR #", "Cat", "Primary TDoc", "Subject / Rapporteur Notes"
            ])
        else:
            self.table.setColumnCount(8)
            self.table.setHorizontalHeaderLabels([
                "Date", "Meeting", "Version", "Type", "CR #", "Cat", "Primary TDoc", "Subject / Rapporteur Notes"
            ])

        total_rows = len(self.filtered_records)
        self.stats_lbl.setText(f"{total_rows} record(s)")

        for row_idx, r in enumerate(self.filtered_records):
            self.table.insertRow(row_idx)
            col_offset = 1 if is_ranking_active else 0

            if is_ranking_active:
                score = r.get("score", 0)
                role = r.get("role", "")
                role_icon = role.split()[0] if role else "⭐"
                rel_item = NumericSortItem(str(score))
                rel_item.setText(f"{role_icon} {score}")
                rel_item.setTextAlignment(Qt.AlignCenter)
                f_rel = rel_item.font()
                f_rel.setBold(True)
                rel_item.setFont(f_rel)
                if score >= 70:
                    rel_item.setForeground(QBrush(QColor(C.ACCENT_SKY)))
                elif score >= 50:
                    rel_item.setForeground(QBrush(QColor(C.TEAL_DARK)))
                else:
                    rel_item.setForeground(QBrush(QColor(C.TEXT_MUTED)))
                badges = r.get("evidence_badges", [])
                tip = f"<b>{role}</b> (Score: {score}/100)<br/>" + "<br/>".join(f"• {b}" for b in badges)
                rel_item.setToolTip(tip)
                self.table.setItem(row_idx, 0, rel_item)

            # Date
            date_item = QTableWidgetItem(r.get("date", ""))
            date_item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row_idx, col_offset + 0, date_item)

            # Meeting
            meeting_item = QTableWidgetItem(r.get("meeting", ""))
            meeting_item.setTextAlignment(Qt.AlignCenter)
            meeting_item.setForeground(QBrush(QColor(C.ACCENT_SKY_DARK)))
            self.table.setItem(row_idx, col_offset + 1, meeting_item)

            # Version
            ver_text = f"v{r.get('version', '')}" if r.get('version') else "—"
            ver_item = QTableWidgetItem(ver_text)
            ver_item.setTextAlignment(Qt.AlignCenter)
            f = ver_item.font()
            f.setBold(True)
            ver_item.setFont(f)
            self.table.setItem(row_idx, col_offset + 2, ver_item)

            # Type
            cr_num = r.get("cr_number", "")
            pcr_cnt = len(r.get("tdocs", []))
            if cr_num and cr_num != "-":
                type_str = "CR"
            elif pcr_cnt > 0:
                type_str = f"p-CR ({pcr_cnt})"
            else:
                type_str = "Editorial"
            type_item = QTableWidgetItem(type_str)
            type_item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row_idx, col_offset + 3, type_item)

            # CR #
            cr_item = NumericSortItem(cr_num if cr_num != "-" else "—")
            cr_item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row_idx, col_offset + 4, cr_item)

            # Cat
            cat_val = r.get("category", "")
            cat_item = QTableWidgetItem(cat_val if cat_val != "-" else "—")
            cat_item.setTextAlignment(Qt.AlignCenter)
            if cat_val == "B":
                cat_item.setForeground(QBrush(QColor(C.ACCENT_SKY)))
                f_cat = cat_item.font()
                f_cat.setBold(True)
                cat_item.setFont(f_cat)
            self.table.setItem(row_idx, col_offset + 5, cat_item)

            # Primary TDoc
            source_tdoc = r.get("source_tdoc", "")
            if not source_tdoc and r.get("tdocs"):
                source_tdoc = r["tdocs"][0]["raw_tdoc"]
            tdoc_item = QTableWidgetItem(source_tdoc if source_tdoc else "—")
            tdoc_item.setTextAlignment(Qt.AlignCenter)
            if source_tdoc and source_tdoc != "—":
                tdoc_item.setForeground(QBrush(QColor(C.ACCENT_SKY)))
                f_tdoc = tdoc_item.font()
                f_tdoc.setBold(True)
                tdoc_item.setFont(f_tdoc)
            self.table.setItem(row_idx, col_offset + 6, tdoc_item)

            # Subject / Comment
            subj_item = QTableWidgetItem(r.get("subject", "").replace("\n", " "))
            self.table.setItem(row_idx, col_offset + 7, subj_item)

        # Set sensible initial column widths
        if is_ranking_active:
            self.table.setColumnWidth(0, 85)   # Relevance
            self.table.setColumnWidth(1, 85)   # Date
            self.table.setColumnWidth(2, 115)  # Meeting
            self.table.setColumnWidth(3, 80)   # Version
            self.table.setColumnWidth(4, 75)   # Type
            self.table.setColumnWidth(5, 65)   # CR #
            self.table.setColumnWidth(6, 50)   # Cat
            self.table.setColumnWidth(7, 115)  # Primary TDoc
        else:
            self.table.setColumnWidth(0, 85)   # Date
            self.table.setColumnWidth(1, 115)  # Meeting
            self.table.setColumnWidth(2, 80)   # Version
            self.table.setColumnWidth(3, 85)   # Type
            self.table.setColumnWidth(4, 65)   # CR #
            self.table.setColumnWidth(5, 50)   # Cat
            self.table.setColumnWidth(6, 115)  # Primary TDoc

        if total_rows > 0:
            self.table.selectRow(0)
