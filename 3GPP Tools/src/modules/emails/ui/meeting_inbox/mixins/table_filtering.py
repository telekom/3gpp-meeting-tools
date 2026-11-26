"""Filter combo population, mode pill counters, row cell formatting, and inbox table filtering."""

from __future__ import annotations

from datetime import datetime, timedelta
from typing import Dict, List, Optional

from PyQt5.QtCore import QItemSelectionModel, Qt
from PyQt5.QtGui import QColor
from PyQt5.QtWidgets import QTableWidgetItem

from core.ui.theme import C
from modules.emails.core.emeeting_signals import format_signal_badges
from modules.emails.core.thread_markdown_exporter import format_sender_display
from modules.meetings.ui.tdocs_dialogs import MY_STATUS_OPTIONS
from ..tdoc_grouping import is_item_effectively_parked, is_priority_item
from ..widgets import _status_color_hex


class InboxTableFilteringMixin:
    """Updates filter combos/pills, formats row state badges, and populates the inbox table."""

    def _populate_filter_combos(self):
        cutoff = self._recent_cutoff_str()
        dataset = [it for it in self._active_dataset() if self._item_matches_mode(it, self._active_mode, cutoff)]

        companies = sorted(
            {
                c
                for it in dataset
                for c in (it.get("_companies") if it.get("_is_group") else [it.get("company")])
                if c
            },
            key=str.lower,
        )
        ais = sorted({ai for it in dataset for ai in (it.get("linked_ais") or []) if ai})
        unassigned_cnt = sum(1 for it in dataset if not (it.get("linked_ais") or []))

        self.company_combo.blockSignals(True)
        self.company_combo.updateItems(companies)
        if self._pending_company_filter is not None and companies:
            wanted_lower = {c.lower() for c in self._pending_company_filter}
            matched_comps = [c for c in companies if c.lower() in wanted_lower]
            if matched_comps:
                self.company_combo.setCheckedItems(matched_comps)
            self._pending_company_filter = None
        self.company_combo.blockSignals(False)

        ai_items = ["__UNASSIGNED__"] + ais
        ai_labels: Dict[str, str] = {
            "__UNASSIGNED__": f"❓ Unassigned (No AI) ({unassigned_cnt})",
        }
        ai_tooltips: Dict[str, str] = {
            "__UNASSIGNED__": f"Emails or threads without any linked AI ({unassigned_cnt})",
        }
        for ai in ais:
            full_label = self._ai_label_map.get(ai, ai)
            ai_tooltips[ai] = full_label
            ai_labels[ai] = (full_label[:45] + "...") if len(full_label) > 48 else full_label

        self.ai_combo.blockSignals(True)
        self.ai_combo.updateItems(ai_items, display_labels=ai_labels, item_tooltips=ai_tooltips)
        if self._pending_ai_filter is not None and dataset:
            wanted_ais = set(self._pending_ai_filter)
            matched_ais = [a for a in ai_items if a in wanted_ais]
            if matched_ais:
                self.ai_combo.setCheckedItems(matched_ais)
            self._pending_ai_filter = None
        self.ai_combo.blockSignals(False)

        if hasattr(self, "status_combo"):
            status_counts: Dict[str, int] = {s: 0 for s in MY_STATUS_OPTIONS}
            for it in dataset:
                st = str(it.get("user_status") or "⚪ Neutral").strip() or "⚪ Neutral"
                status_counts[st] = status_counts.get(st, 0) + 1
            st_items = list(MY_STATUS_OPTIONS)
            st_labels = {s: f"{s} ({status_counts.get(s, 0)})" for s in st_items}
            self.status_combo.blockSignals(True)
            self.status_combo.updateItems(st_items, display_labels=st_labels)
            if self._pending_status_filter is not None and dataset:
                wanted_st = set(self._pending_status_filter)
                matched_st = [s for s in st_items if s in wanted_st]
                if matched_st:
                    self.status_combo.setCheckedItems(matched_st)
                self._pending_status_filter = None
            self.status_combo.blockSignals(False)

    @staticmethod
    def _recent_cutoff_str() -> str:
        return (datetime.now() - timedelta(hours=24)).strftime("%Y-%m-%d %H:%M")

    def _update_pill_counts(self):
        cutoff = self._recent_cutoff_str()
        dataset = self._active_dataset()
        followed_ais = getattr(self, "_followed_ais", set())
        priority_cnt = sum(1 for it in dataset if is_priority_item(it, followed_ais))
        unread_cnt = sum(1 for it in dataset if not it["is_read"])
        flagged_cnt = sum(1 for it in dataset if it["is_flagged"])
        recent_cnt = sum(1 for it in dataset if str(it.get("date_str") or "") >= cutoff)
        total_cnt = len(dataset)

        if hasattr(self, "pill_priority"):
            self.pill_priority.setText(f"🔥 Priority ({priority_cnt})")
        self.pill_unread.setText(f"🔵 Unread Only ({unread_cnt})")
        self.pill_flagged.setText(f"🚩 Flagged ({flagged_cnt})")
        self.pill_recent.setText(f"🕒 Last 24h ({recent_cnt})")
        self.pill_all.setText(f"📋 All ({total_cnt})")

    def _format_row_state_cell_text_and_tip(self, it: dict) -> tuple[str, str]:
        u_status = str(it.get("user_status") or "⚪ Neutral").strip() or "⚪ Neutral"
        u_notes = str(it.get("user_notes") or "").strip()
        note_badge = "📝" if u_notes else ""
        sig = it.get("_signals") or {}
        sig_badges = format_signal_badges(sig)

        if it.get("_is_group"):
            u_cnt = int(it.get("_unread_count", 0))
            t_cnt = int(it.get("_total_count", 1))
            f_part = "🚩" if it["is_flagged"] else ""
            core_s = f"🔵 {u_cnt}/{t_cnt} {f_part}".strip() if u_cnt > 0 else f"✓ {t_cnt} {f_part}".strip()
            state_txt = " ".join(p for p in (core_s, sig_badges, note_badge) if p)
        else:
            state_parts = []
            if not it["is_read"]:
                state_parts.append("🔵")
            if it["is_flagged"]:
                state_parts.append("🚩")
            if not state_parts:
                state_parts.append("✓")
            if sig_badges:
                state_parts.append(sig_badges)
            if note_badge:
                state_parts.append(note_badge)
            state_txt = " ".join(state_parts)

        st_tip: List[str] = []
        if u_status != "⚪ Neutral":
            st_tip.append(f"My Status: {u_status}")
        if sig.get("is_objected"):
            rev_info = f" at {sig.get('objected_at_rev')}" if sig.get("objected_at_rev") else ""
            if sig.get("unparked_by_new_rev"):
                st_tip.append(f"🛑→🆕 Unparked! Marked Objected{rev_info}, but newer revision {sig.get('max_rev_tag')} arrived!")
            else:
                st_tip.append(f"🛑 Marked Objected / Parked{rev_info}")
        elif sig.get("max_rev_tag"):
            st_tip.append(f"🆕 Latest Revision in Emails: {sig['max_rev_tag']}")
        if sig.get("has_template"):
            st_tip.append("📋 Contains formal e-Meeting <<START>>...<<END>> comment block")
        if sig.get("has_objection"):
            snips = sig.get("objection_snippets") or ([sig["objection_snippet"]] if sig.get("objection_snippet") else [])
            st_tip.append("⚠️ Objection keyword(s) detected in email body" + (":\n  " + "\n  ".join(snips[:3]) if snips else ""))
        if it.get("_is_group") and hasattr(self, "_resolve_item_summary_scope") and getattr(self, "_summary_store", None) is not None:
            scope, scope_msgs, _ = self._resolve_item_summary_scope(it)
            cache = getattr(self, "_summary_freshness_cache", {})
            ckey = f"{scope.scope_type}:{scope.scope_key}"
            fresh = cache.get(ckey)
            if fresh is not None and fresh.has_summary:
                sum_badge = f"✨+{fresh.new_count}" if fresh.new_count > 0 else "✨"
                state_txt = f"{state_txt} {sum_badge}".strip()
                st_tip.append(
                    f"✨ AI Summary: +{fresh.new_count} new message(s)"
                    if fresh.new_count > 0
                    else f"✨ AI Summary: Up to date ({fresh.covered_count} msgs)"
                )
        if u_notes:
            preview_n = (u_notes[:180] + "…") if len(u_notes) > 180 else u_notes
            st_tip.append(f"Comment: {preview_n}")
        return state_txt, "\n".join(st_tip)

    def _apply_filters(
        self,
        *_args,
        preserve_id: Optional[str] = None,
        preserve_ids: Optional[List[str]] = None,
        select_row: Optional[int] = None,
    ):
        self._save_filter_settings()
        target_id = preserve_id if preserve_id is not None else (self._current_item["id"] if self._current_item else None)
        mode = self._active_mode
        cutoff = self._recent_cutoff_str()
        q = self.search_edit.text().strip().lower()
        hide_objected = bool(getattr(self, "_hide_objected", False))
        filter_comp = self.company_combo.isFiltered()
        selected_comps = (
            {str(c).strip().lower() for c in self.company_combo.getCheckedItems() if str(c).strip()}
            if filter_comp
            else set()
        )
        filter_ai = self.ai_combo.isFiltered()
        selected_ais = (
            {str(a).strip() for a in self.ai_combo.getCheckedItems() if str(a).strip()}
            if filter_ai
            else set()
        )
        filter_status = hasattr(self, "status_combo") and self.status_combo.isFiltered()
        selected_statuses = (
            {str(s).strip() for s in self.status_combo.getCheckedItems() if str(s).strip()}
            if filter_status
            else set()
        )
        source_filter = (self.source_combo.currentData() or "").strip()

        dataset = self._active_dataset()
        if hasattr(self, "_refresh_summary_freshness_cache"):
            self._refresh_summary_freshness_cache(dataset)
        visible: List[dict] = []
        for it in dataset:
            if not self._item_matches_mode(it, mode, cutoff):
                continue
            if hide_objected and is_item_effectively_parked(it):
                continue
            if source_filter == "THREAD" and not it["in_tracked_db"]:
                continue
            if source_filter == "UNPINNED" and it["in_tracked_db"]:
                continue
            if source_filter == "NEEDS_TRIAGE" and it["in_tracked_db"] and bool(it.get("linked_ais")):
                continue
            if filter_comp:
                item_comps = (
                    {str(c).strip().lower() for c in (it.get("_companies") or []) if str(c).strip()}
                    if it.get("_is_group")
                    else {str(it.get("company") or "").strip().lower()}
                )
                if not (item_comps & selected_comps):
                    continue
            if filter_ai:
                item_ais = {str(a).strip() for a in (it.get("linked_ais") or []) if str(a).strip()}
                matches_unassigned = ("__UNASSIGNED__" in selected_ais) and (not item_ais)
                matches_ai = bool(item_ais & (selected_ais - {"__UNASSIGNED__"}))
                if not (matches_unassigned or matches_ai):
                    continue
            if filter_status:
                item_status = str(it.get("user_status") or "⚪ Neutral").strip() or "⚪ Neutral"
                if item_status not in selected_statuses:
                    continue
            if q:
                hay = " ".join([
                    str(it.get("subject") or ""),
                    str(it.get("thread_title") or ""),
                    str(it.get("_tdoc_source") or ""),
                    str(it.get("sender_name") or ""),
                    str(it.get("company") or ""),
                    " ".join(it.get("linked_tdocs") or []),
                    " ".join(it.get("linked_ais") or []),
                    str(it.get("body_text") or "")[:1200],
                ]).lower()
                if q not in hay:
                    continue
            visible.append(it)

        self._visible_items = visible
        self.table.blockSignals(True)
        self.table.setRowCount(len(visible))
        row_to_select = -1
        is_tdoc_view = self._view_mode == "TDOCS"
        followed_ais = getattr(self, "_followed_ais", set())

        for r, it in enumerate(visible):
            u_status = str(it.get("user_status") or "⚪ Neutral").strip() or "⚪ Neutral"
            state_txt, state_tip = self._format_row_state_cell_text_and_tip(it)
            state_item = QTableWidgetItem(state_txt)
            state_item.setData(Qt.UserRole, it["id"])
            state_item.setData(Qt.UserRole + 1, _status_color_hex(u_status))
            state_item.setTextAlignment(Qt.AlignCenter)
            if state_tip:
                state_item.setToolTip(state_tip)

            if is_tdoc_view and it.get("_primary_tdoc"):
                btd = it["_primary_tdoc"]
                td_type = it.get("_tdoc_type") or ""
                td_src = it.get("_tdoc_source") or ""
                type_tag = f" [{td_type}]" if td_type else ""
                src_short = f" • {td_src[:22]}" + ("…" if len(td_src) > 22 else "") if td_src else ""
                ctx = f"📄 {btd}{type_tag}{src_short}"
                ctx_full = f"TDoc: {btd}{type_tag}\nAuthor/Source: {td_src or 'Unknown'}\nOfficial Status: {it.get('_tdoc_status') or 'N/A'}"
            elif it["in_tracked_db"] and it.get("thread_title"):
                ctx_full = f"🧵 {it['thread_title']}"
                ctx = f"🧵 {it['thread_title'][:34]}"
                if len(it["thread_title"]) > 34:
                    ctx += "…"
            elif it.get("linked_tdocs"):
                ctx_full = "📄 " + ", ".join(it["linked_tdocs"])
                ctx = "📄 " + ", ".join(it["linked_tdocs"][:2])
            else:
                ctx_full = "📧 Email"
                ctx = "📧 Email"
            ctx_item = QTableWidgetItem(ctx)
            ctx_item.setToolTip(ctx_full)

            ais = it.get("linked_ais") or []
            wis = it.get("linked_wis") or []
            is_followed_ai = bool(followed_ais and (set(ais) & followed_ais))
            ai_prefix = "★ " if is_followed_ai else ""
            if is_tdoc_view and ais and wis:
                ai_disp = f"{ai_prefix}{', '.join(ais)} • {wis[0]}"
            else:
                ai_disp = f"{ai_prefix}{', '.join(ais)}" if ais else "❓ —"
            ai_item = QTableWidgetItem(ai_disp)
            if not ais:
                ai_item.setForeground(QColor(C.WARNING))
                ai_item.setToolTip("Unassigned (No AI)")
            else:
                tip_lines = [self._ai_label_map.get(a, a) for a in ais]
                if wis:
                    tip_lines.append("WIs: " + ", ".join(wis))
                if is_followed_ai:
                    tip_lines.append("★ Followed AI (on your Priority Radar)")
                ai_item.setToolTip("\n".join(tip_lines))

            subj_full = str(it.get("subject") or "")
            subj_item = QTableWidgetItem(subj_full)
            subj_item.setToolTip(subj_full)
            if not it["is_read"]:
                f = subj_item.font()
                f.setBold(True)
                subj_item.setFont(f)

            if is_tdoc_view and it.get("_is_group"):
                comps = it.get("_companies") or []
                comp_counts = it.get("_company_counts") or {}
                latest_c = it.get("_latest_company") or "Unknown"
                if len(comps) > 1:
                    sender_disp = f"{len(comps)} cos (latest: {latest_c})"
                elif comps:
                    sender_disp = comps[0]
                else:
                    sender_disp = latest_c
                sender_tip = ", ".join(f"{c} ({comp_counts.get(c, 1)})" for c in comps) or sender_disp
            else:
                sender_disp = format_sender_display(
                    it["sender_name"] or it["sender_email"] or "Unknown",
                    it.get("company") or "",
                )
                sender_tip = f"{sender_disp} <{it['sender_email']}>" if it.get("sender_email") else sender_disp
            sender_item = QTableWidgetItem(sender_disp)
            sender_item.setToolTip(sender_tip)

            date_full = str(it.get("date_str") or "")
            date_item = QTableWidgetItem(date_full[:16])
            date_item.setToolTip(date_full)

            for cell in (state_item, ctx_item, ai_item, subj_item, sender_item, date_item):
                if not it["is_read"]:
                    cell.setBackground(QColor(C.INFO_SOFT))

            self.table.setItem(r, 0, state_item)
            self.table.setItem(r, 1, ctx_item)
            self.table.setItem(r, 2, ai_item)
            self.table.setItem(r, 3, subj_item)
            self.table.setItem(r, 4, sender_item)
            self.table.setItem(r, 5, date_item)

            if target_id and it["id"] == target_id:
                row_to_select = r

        if visible:
            rows_to_select: List[int] = []
            if preserve_ids:
                pid_set = set(preserve_ids)
                rows_to_select = [r for r, it in enumerate(visible) if it["id"] in pid_set]
            if not rows_to_select:
                if select_row is not None and 0 <= select_row < len(visible):
                    rows_to_select = [select_row]
                elif select_row is not None and select_row >= len(visible):
                    rows_to_select = [len(visible) - 1]
                elif row_to_select >= 0:
                    rows_to_select = [row_to_select]
                else:
                    rows_to_select = [0]

            sel_model = self.table.selectionModel()
            if sel_model is None:
                self.table.selectRow(rows_to_select[0])
            else:
                sel_model.clearSelection()
                for r in rows_to_select:
                    sel_model.select(
                        self.table.model().index(r, 0),
                        QItemSelectionModel.Select | QItemSelectionModel.Rows,
                    )
                sel_model.setCurrentIndex(
                    self.table.model().index(rows_to_select[0], 0),
                    QItemSelectionModel.NoUpdate,
                )
        self.table.blockSignals(False)

        if self._view_mode == "TDOCS":
            noun = "TDoc group(s)"
        elif self._view_mode == "THREADS":
            noun = "thread(s)"
        else:
            noun = "indexed email(s)"
        self.status_lbl.setText(f"Showing {len(visible)} of {len(dataset)} {noun}")
        self._update_sync_button_state(
            total_count=len(dataset),
            visible_count=len(visible),
            has_narrow_filter=bool(
                filter_comp
                or filter_ai
                or filter_status
                or q
                or source_filter
                or hide_objected
                or mode in ("FLAGGED", "PRIORITY")
            ),
        )
        if hasattr(self, "_update_ai_briefing_strip"):
            self._update_ai_briefing_strip(selected_ais if filter_ai else set())
        self._on_selection_changed()
