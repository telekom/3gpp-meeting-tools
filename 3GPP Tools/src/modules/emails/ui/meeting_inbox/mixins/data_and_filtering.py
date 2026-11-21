"""Dataset loading, thread grouping, filter population, and inbox table filtering mixin."""

from __future__ import annotations

from datetime import datetime, timedelta
from typing import Any, Dict, List, Optional

from PyQt5.QtCore import QItemSelectionModel, Qt
from PyQt5.QtGui import QColor
from PyQt5.QtWidgets import QTableWidgetItem

from core.ui.theme import C
from modules.emails.core.thread_markdown_exporter import clean_tdoc_id, format_sender_display
from modules.emails.core.tracked_threads_service import normalize_subject
from modules.meetings.ui.tdocs_dialogs import MY_STATUS_OPTIONS
from ..widgets import _status_color_hex


class InboxDataAndFilteringMixin:
    """Loads DB items, groups threads, updates filter combos/pills, and populates the inbox table."""

    def _set_mode(self, mode: str):
        norm = str(mode or "UNREAD").strip().upper()
        if norm not in ("UNREAD", "FLAGGED", "RECENT", "ALL"):
            norm = "UNREAD"
        self._active_mode = norm
        for pill, key in (
            (getattr(self, "pill_unread", None), "UNREAD"),
            (getattr(self, "pill_flagged", None), "FLAGGED"),
            (getattr(self, "pill_recent", None), "RECENT"),
            (getattr(self, "pill_all", None), "ALL"),
        ):
            if pill is not None:
                pill.blockSignals(True)
                pill.setChecked(norm == key)
                pill.blockSignals(False)
        self._populate_filter_combos()
        self._apply_filters()

    def _set_view_mode(self, view_mode: str):
        norm = "THREADS" if str(view_mode or "").upper() == "THREADS" else "MESSAGES"
        self._view_mode = norm
        if hasattr(self, "btn_view_messages"):
            self.btn_view_messages.blockSignals(True)
            self.btn_view_messages.setChecked(norm == "MESSAGES")
            self.btn_view_messages.blockSignals(False)
        if hasattr(self, "btn_view_threads"):
            self.btn_view_threads.blockSignals(True)
            self.btn_view_threads.setChecked(norm == "THREADS")
            self.btn_view_threads.blockSignals(False)
        self._populate_filter_combos()
        self._update_pill_counts()
        self._apply_filters()

    def _active_dataset(self) -> List[dict]:
        return self._grouped_items if self._view_mode == "THREADS" else self._all_items

    def _build_grouped_thread_items(self) -> List[dict]:
        buckets: Dict[str, List[dict]] = {}
        for it in self._all_items:
            if it.get("in_tracked_db") and it.get("thread_id"):
                key = f"THR::{it['thread_id']}"
            else:
                norm = normalize_subject(it.get("subject") or "").lower()
                key = f"UNPIN::{norm or it['id']}"
            buckets.setdefault(key, []).append(it)

        grouped: List[dict] = []
        for gkey, msgs in buckets.items():
            msgs_asc = sorted(msgs, key=lambda m: str(m.get("date_str") or ""))
            latest = msgs_asc[-1]
            in_tracked = any(m.get("in_tracked_db") for m in msgs_asc)
            tid = next((m.get("thread_id") for m in msgs_asc if m.get("thread_id")), "")
            th_title = next((m.get("thread_title") for m in msgs_asc if m.get("thread_title")), "")
            linked_ais = list(dict.fromkeys(ai for m in msgs_asc for ai in (m.get("linked_ais") or []) if ai))
            linked_subs = list(dict.fromkeys(s for m in msgs_asc for s in (m.get("linked_sub_items") or []) if s))
            linked_wis = list(dict.fromkeys(w for m in msgs_asc for w in (m.get("linked_wis") or []) if w))
            linked_tdocs = list(dict.fromkeys(td for m in msgs_asc for td in (m.get("linked_tdocs") or []) if td))
            companies = list(dict.fromkeys(m.get("company") for m in msgs_asc if m.get("company")))
            unread_cnt = sum(1 for m in msgs_asc if not m["is_read"])
            flagged_cnt = sum(1 for m in msgs_asc if m["is_flagged"])
            user_status = next(
                (str(m.get("user_status") or "").strip() for m in msgs_asc if str(m.get("user_status") or "").strip() and m.get("user_status") != "⚪ Neutral"),
                str(msgs_asc[0].get("user_status") or "⚪ Neutral").strip() if msgs_asc else "⚪ Neutral",
            )
            user_notes = next((str(m.get("user_notes") or "").strip() for m in msgs_asc if str(m.get("user_notes") or "").strip()), "")

            grouped.append({
                "id": gkey,
                "_is_group": True,
                "_group_messages": msgs_asc,
                "_unread_count": unread_cnt,
                "_total_count": len(msgs_asc),
                "_flagged_count": flagged_cnt,
                "_companies": companies,
                "source_kind": "thread" if in_tracked else "tdoc",
                "thread_id": tid,
                "thread_title": th_title or normalize_subject(latest.get("subject") or ""),
                "subject": th_title or latest.get("subject") or "(No Subject)",
                "sender_name": latest.get("sender_name") or "",
                "sender_email": latest.get("sender_email") or "",
                "company": ", ".join(companies) if companies else (latest.get("company") or ""),
                "date_str": latest.get("date_str") or "",
                "body_text": latest.get("body_text") or "",
                "is_read": unread_cnt == 0,
                "is_flagged": flagged_cnt > 0,
                "attachments": latest.get("attachments") or [],
                "linked_ais": linked_ais,
                "linked_sub_items": linked_subs,
                "linked_wis": linked_wis,
                "linked_tdocs": linked_tdocs,
                "in_tracked_db": in_tracked,
                "in_general_db": any(m.get("in_general_db") for m in msgs_asc),
                "conversation_id": next((m.get("conversation_id") for m in msgs_asc if m.get("conversation_id")), ""),
                "user_status": user_status or "⚪ Neutral",
                "user_notes": user_notes,
            })
        return sorted(grouped, key=lambda x: x.get("date_str", ""), reverse=True)

    def reload_from_dbs(
        self,
        preserve_id: Optional[str] = None,
        preserve_ids: Optional[List[str]] = None,
    ):
        """Loads and deduplicates emails from TrackedThreadsDatabase and GeneralEmailDatabase."""
        if not getattr(self, "_restoring_filters", False) and hasattr(self, "note_edit"):
            self._flush_pending_note_save(reload_view=False)
        by_id: Dict[str, dict] = {}

        if self.tracked_db is not None:
            threads = self.tracked_db.get_threads_for_meeting(self.meeting_id)
            for th in threads:
                tid = str(th.get("thread_id") or "")
                th_title = str(th.get("title") or "").strip()
                th_ais = [str(a).strip() for a in (th.get("linked_ais") or []) if str(a).strip()]
                th_subs = [str(s).strip() for s in (th.get("linked_sub_items") or []) if str(s).strip()]
                th_wis = [str(w).strip() for w in (th.get("linked_wis") or []) if str(w).strip()]
                th_tdocs = list(dict.fromkeys(
                    self._clean_tdoc_id(t) for t in (th.get("linked_tdocs") or []) if self._clean_tdoc_id(t)
                ))
                th_status = str(th.get("user_status") or "⚪ Neutral").strip() or "⚪ Neutral"
                th_notes = str(th.get("user_notes") or "").strip()
                for m in self.tracked_db.get_thread_messages(tid):
                    eid = str(m.get("id") or "").strip()
                    if not eid:
                        continue
                    by_id[eid] = {
                        "id": eid,
                        "source_kind": "thread",
                        "thread_id": tid,
                        "thread_title": th_title,
                        "subject": str(m.get("subject") or th_title or "(No Subject)").strip(),
                        "sender_name": str(m.get("sender_name") or "").strip(),
                        "sender_email": str(m.get("sender_email") or "").strip(),
                        "company": str(m.get("company") or "").strip(),
                        "date_str": str(m.get("date_sent") or "").strip(),
                        "body_text": str(m.get("clean_body") or m.get("body_text") or ""),
                        "clean_body": str(m.get("clean_body") or ""),
                        "raw_body_text": str(m.get("body_text") or ""),
                        "html_body": "",
                        "is_read": int(m.get("is_read", 0) or 0) == 1,
                        "is_flagged": int(m.get("flag_status", 0) or 0) == 2,
                        "attachments": list(m.get("attachments") or []),
                        "linked_ais": list(th_ais),
                        "linked_sub_items": list(th_subs),
                        "linked_wis": list(th_wis),
                        "linked_tdocs": list(th_tdocs),
                        "conversation_id": str(m.get("conversation_id") or ""),
                        "in_tracked_db": True,
                        "in_general_db": False,
                        "user_status": th_status,
                        "user_notes": th_notes,
                    }

        if self.general_email_db is not None and hasattr(self.general_email_db, "get_all_emails_with_tdocs"):
            for row in self.general_email_db.get_all_emails_with_tdocs(show_ignored=False):
                eid = str(row.get("id") or "").strip()
                if not eid:
                    continue
                matched_tdocs = list(dict.fromkeys(
                    self._clean_tdoc_id(t)
                    for t in str(row.get("matched_tdocs") or "").split(",")
                    if self._clean_tdoc_id(t)
                ))
                inferred_ais = []
                inferred_subs = []
                inferred_wis = []
                for td in matched_tdocs:
                    ai = self.tdoc_ai_lookup.get(td, "")
                    meta = self._tdoc_meta_map.get(td) or {}
                    if not ai:
                        ai = str(meta.get("ai") or "").strip()
                    if ai and ai not in inferred_ais:
                        inferred_ais.append(ai)
                    sub = str(meta.get("sub_item") or "").strip()
                    if sub and sub not in inferred_subs:
                        inferred_subs.append(sub)
                    for w in (meta.get("wis") or []):
                        if w and w not in inferred_wis:
                            inferred_wis.append(w)

                if eid in by_id:
                    item = by_id[eid]
                    item["in_general_db"] = True
                    for td in matched_tdocs:
                        if td not in item["linked_tdocs"]:
                            item["linked_tdocs"].append(td)
                    for ai in inferred_ais:
                        if ai not in item["linked_ais"]:
                            item["linked_ais"].append(ai)
                    if row.get("html_body") and not item.get("html_body"):
                        item["html_body"] = str(row["html_body"])
                    comp_val = str(row.get("company") or row.get("sender_company") or "").strip()
                    if not item.get("company") and comp_val:
                        item["company"] = comp_val
                else:
                    by_id[eid] = {
                        "id": eid,
                        "source_kind": "tdoc",
                        "thread_id": "",
                        "thread_title": "",
                        "subject": str(row.get("subject") or "(No Subject)").strip(),
                        "sender_name": str(row.get("sender_name") or "").strip(),
                        "sender_email": str(row.get("sender_email") or "").strip(),
                        "company": str(row.get("company") or row.get("sender_company") or "").strip(),
                        "date_str": str(row.get("date_received") or "").strip(),
                        "body_text": str(row.get("clean_body") or row.get("body_text") or ""),
                        "clean_body": str(row.get("clean_body") or ""),
                        "raw_body_text": str(row.get("body_text") or ""),
                        "html_body": str(row.get("html_body") or ""),
                        "is_read": int(row.get("is_read", 0) or 0) == 1,
                        "is_flagged": int(row.get("flag_status", 0) or 0) == 2,
                        "attachments": list(row.get("attachments") or []),
                        "linked_ais": inferred_ais,
                        "linked_sub_items": inferred_subs,
                        "linked_wis": inferred_wis,
                        "linked_tdocs": matched_tdocs,
                        "conversation_id": "",
                        "in_tracked_db": False,
                        "in_general_db": True,
                        "user_status": "⚪ Neutral",
                        "user_notes": "",
                    }

        self._all_items = sorted(by_id.values(), key=lambda x: x.get("date_str", ""), reverse=True)
        self._grouped_items = self._build_grouped_thread_items()
        self._items_by_id = {it["id"]: it for it in self._all_items}

        self._populate_filter_combos()
        self._update_pill_counts()
        self._apply_filters(preserve_id=preserve_id, preserve_ids=preserve_ids)

    @staticmethod
    def _clean_tdoc_id(raw: Any) -> str:
        return clean_tdoc_id(raw)

    def _item_matches_mode(self, it: dict, mode: str, cutoff: str) -> bool:
        if mode == "UNREAD" and it["is_read"]:
            return False
        if mode == "FLAGGED" and not it["is_flagged"]:
            return False
        if mode == "RECENT" and str(it.get("date_str") or "") < cutoff:
            return False
        return True

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
        unread_cnt = sum(1 for it in dataset if not it["is_read"])
        flagged_cnt = sum(1 for it in dataset if it["is_flagged"])
        recent_cnt = sum(1 for it in dataset if str(it.get("date_str") or "") >= cutoff)
        total_cnt = len(dataset)

        self.pill_unread.setText(f"🔵 Unread Only ({unread_cnt})")
        self.pill_flagged.setText(f"🚩 Flagged ({flagged_cnt})")
        self.pill_recent.setText(f"🕒 Last 24h ({recent_cnt})")
        self.pill_all.setText(f"📋 All ({total_cnt})")

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
        visible: List[dict] = []
        for it in dataset:
            if mode == "UNREAD" and it["is_read"]:
                continue
            if mode == "FLAGGED" and not it["is_flagged"]:
                continue
            if mode == "RECENT" and str(it.get("date_str") or "") < cutoff:
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
        for r, it in enumerate(visible):
            u_status = str(it.get("user_status") or "⚪ Neutral").strip() or "⚪ Neutral"
            u_notes = str(it.get("user_notes") or "").strip()
            note_badge = "📝" if u_notes else ""
            if it.get("_is_group"):
                u_cnt = int(it.get("_unread_count", 0))
                t_cnt = int(it.get("_total_count", 1))
                f_part = "🚩" if it["is_flagged"] else ""
                core_s = f"🔵 {u_cnt}/{t_cnt} {f_part}".strip() if u_cnt > 0 else f"✓ {t_cnt} {f_part}".strip()
                state_txt = f"{core_s} {note_badge}".strip()
            else:
                state_parts = []
                if not it["is_read"]:
                    state_parts.append("🔵")
                if it["is_flagged"]:
                    state_parts.append("🚩")
                if not state_parts:
                    state_parts.append("✓")
                if note_badge:
                    state_parts.append(note_badge)
                state_txt = " ".join(state_parts)
            state_item = QTableWidgetItem(state_txt)
            state_item.setData(Qt.UserRole, it["id"])
            state_item.setData(Qt.UserRole + 1, _status_color_hex(u_status))
            state_item.setTextAlignment(Qt.AlignCenter)
            st_tip = [f"Status: {u_status}"] if u_status != "⚪ Neutral" else []
            if u_notes:
                preview_n = (u_notes[:180] + "…") if len(u_notes) > 180 else u_notes
                st_tip.append(f"Comment: {preview_n}")
            if st_tip:
                state_item.setToolTip("\n".join(st_tip))

            if it["in_tracked_db"] and it.get("thread_title"):
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
            ai_item = QTableWidgetItem(", ".join(ais) if ais else "❓ —")
            if not ais:
                ai_item.setForeground(QColor(C.WARNING))
                ai_item.setToolTip("Unassigned (No AI)")
            else:
                ai_item.setToolTip("\n".join(self._ai_label_map.get(a, a) for a in ais))

            subj_full = str(it.get("subject") or "")
            subj_item = QTableWidgetItem(subj_full)
            subj_item.setToolTip(subj_full)
            if not it["is_read"]:
                f = subj_item.font()
                f.setBold(True)
                subj_item.setFont(f)

            sender_disp = format_sender_display(
                it["sender_name"] or it["sender_email"] or "Unknown",
                it.get("company") or "",
            )
            sender_item = QTableWidgetItem(sender_disp)
            sender_item.setToolTip(
                f"{sender_disp} <{it['sender_email']}>" if it.get("sender_email") else sender_disp
            )
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

        noun = "thread(s)" if self._view_mode == "THREADS" else "indexed email(s)"
        self.status_lbl.setText(f"Showing {len(visible)} of {len(dataset)} {noun}")
        self._update_sync_button_state(
            total_count=len(dataset),
            visible_count=len(visible),
            has_narrow_filter=bool(filter_comp or filter_ai or filter_status or q or source_filter or mode == "FLAGGED"),
        )
        self._on_selection_changed()
