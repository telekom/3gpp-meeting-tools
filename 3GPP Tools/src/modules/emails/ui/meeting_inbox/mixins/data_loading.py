"""Dataset loading, thread/TDoc grouping, mode switching, and Followed-AIs scope mixin."""

from __future__ import annotations

from typing import Any, Dict, List, Optional

from PyQt5.QtGui import QCursor
from PyQt5.QtWidgets import QMenu

from modules.emails.core.thread_markdown_exporter import clean_tdoc_id
from ..followed_ais_helper import load_effective_followed_ais, save_followed_ais_override
from ..tdoc_grouping import (
    build_grouped_tdoc_items,
    build_grouped_thread_items,
    enrich_message_item_signals,
    is_priority_item,
)


class InboxDataLoadingMixin:
    """Loads DB items, builds grouped datasets, manages mode/view toggles, and syncs Followed AIs."""

    def _set_mode(self, mode: str):
        norm = str(mode or "UNREAD").strip().upper()
        if norm not in ("PRIORITY", "UNREAD", "FLAGGED", "RECENT", "ALL"):
            norm = "UNREAD"
        self._active_mode = norm
        for pill, key in (
            (getattr(self, "pill_priority", None), "PRIORITY"),
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
        raw = str(view_mode or "").strip().upper()
        norm = raw if raw in ("MESSAGES", "THREADS", "TDOCS") else "MESSAGES"
        self._view_mode = norm
        for btn, key in (
            (getattr(self, "btn_view_messages", None), "MESSAGES"),
            (getattr(self, "btn_view_threads", None), "THREADS"),
            (getattr(self, "btn_view_tdocs", None), "TDOCS"),
        ):
            if btn is not None:
                btn.blockSignals(True)
                btn.setChecked(norm == key)
                btn.blockSignals(False)
        if hasattr(self, "_update_table_headers_for_view_mode"):
            self._update_table_headers_for_view_mode()
        self._populate_filter_combos()
        self._update_pill_counts()
        self._apply_filters()

    def _active_dataset(self) -> List[dict]:
        if self._view_mode == "TDOCS":
            return getattr(self, "_tdoc_grouped_items", [])
        if self._view_mode == "THREADS":
            return self._grouped_items
        return self._all_items

    def _build_grouped_thread_items(self) -> List[dict]:
        return build_grouped_thread_items(self._all_items, tdoc_meta_map=self._tdoc_meta_map)

    def _build_grouped_tdoc_items(self) -> List[dict]:
        return build_grouped_tdoc_items(
            self._all_items,
            tdoc_meta_map=self._tdoc_meta_map,
            tdoc_ai_lookup=self.tdoc_ai_lookup,
        )

    def _rebuild_all_grouped_datasets(self):
        for m in self._all_items:
            enrich_message_item_signals(m, tdoc_meta_map=self._tdoc_meta_map)
        self._grouped_items = self._build_grouped_thread_items()
        self._tdoc_grouped_items = self._build_grouped_tdoc_items()

    def _refresh_tdoc_meta_from_db(self):
        """Refreshes user_tdocs.db status, notes, and objection flags into `_tdoc_meta_map`."""
        if getattr(self, "tdocs_db", None) is None or not hasattr(self.tdocs_db, "get_all_user_data"):
            return
        try:
            user_rows = self.tdocs_db.get_all_user_data() or {}
            for tid, u in user_rows.items():
                meta = self._tdoc_meta_map.setdefault(str(tid).upper(), {})
                meta["my_status"] = str(u.get("my_status") or u.get("status") or "⚪ Neutral").strip() or "⚪ Neutral"
                meta["my_notes"] = str(u.get("my_notes") or u.get("notes") or "").strip()
                meta["is_objected"] = bool(u.get("is_objected"))
                meta["objected_at_rev"] = str(u.get("objected_at_rev") or "").strip()
        except Exception:
            pass

    def reload_from_dbs(
        self,
        preserve_id: Optional[str] = None,
        preserve_ids: Optional[List[str]] = None,
    ):
        """Loads and deduplicates emails from TrackedThreadsDatabase and GeneralEmailDatabase."""
        if not getattr(self, "_restoring_filters", False) and hasattr(self, "note_edit"):
            self._flush_pending_note_save(reload_view=False)
        self._refresh_tdoc_meta_from_db()
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
                th_obj = bool(th.get("is_objected"))
                th_obj_rev = str(th.get("objected_at_rev") or "").strip()
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
                        "is_objected": th_obj,
                        "objected_at_rev": th_obj_rev,
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
                raw_subj = str(row.get("subject") or "")
                from modules.emails.core.tracked_threads.parsing import _TDOC_RE
                subj_tdocs = {self._clean_tdoc_id(m) for m in _TDOC_RE.findall(raw_subj) if self._clean_tdoc_id(m)}
                scope_tdocs = [t for t in matched_tdocs if t in subj_tdocs] if subj_tdocs else matched_tdocs
                for td in scope_tdocs:
                    ai = self.tdoc_ai_lookup.get(td, "")
                    meta = self._tdoc_meta_map.get(td) or {}
                    if not ai:
                        ai = str(meta.get("ai") or "").strip()
                    if ai and ai not in inferred_ais:
                        inferred_ais.append(ai)
                    sub = str(meta.get("sub_item") or "").strip()
                    if sub and sub not in inferred_subs:
                        inferred_subs.append(sub)
                for td in matched_tdocs:
                    meta = self._tdoc_meta_map.get(td) or {}
                    for w in (meta.get("wis") or []):
                        if w and w not in inferred_wis:
                            inferred_wis.append(w)

                if not subj_tdocs:
                    if len(inferred_ais) > 1:
                        inferred_ais = []
                        inferred_subs = []
                    elif len(inferred_subs) > 1:
                        inferred_subs = []

                if eid in by_id:
                    item = by_id[eid]
                    item["in_general_db"] = True
                    for td in matched_tdocs:
                        if td not in item["linked_tdocs"]:
                            item["linked_tdocs"].append(td)
                    if not item["linked_ais"] and len(inferred_ais) == 1:
                        item["linked_ais"].append(inferred_ais[0])
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
                        "is_objected": False,
                        "objected_at_rev": "",
                    }

        self._all_items = sorted(by_id.values(), key=lambda x: x.get("date_str", ""), reverse=True)
        self._rebuild_all_grouped_datasets()
        self._items_by_id = {it["id"]: it for it in self._all_items}

        self._refresh_followed_ais_scope()
        if hasattr(self, "_update_table_headers_for_view_mode"):
            self._update_table_headers_for_view_mode()
        self._populate_filter_combos()
        self._update_pill_counts()
        self._apply_filters(preserve_id=preserve_id, preserve_ids=preserve_ids)

    def _refresh_followed_ais_scope(self):
        gdb_path = getattr(self.general_email_db, "db_path", None) if self.general_email_db is not None else None
        all_ais = sorted({
            ai
            for it in self._all_items
            for ai in (it.get("linked_ais") or [])
            if ai
        } | set(self._ai_label_map.keys()))
        eff, inh, customized = load_effective_followed_ais(
            self._settings,
            self._meeting_settings_prefix(),
            tdoc_meta_map=self._tdoc_meta_map,
            general_email_db_path=gdb_path,
            available_ais=all_ais,
        )
        self._followed_ais = eff
        self._inherited_followed_ais = inh
        self._followed_ais_customized = customized
        if hasattr(self, "followed_ais_combo") and all_ais:
            labels = {}
            tips = {}
            for ai in all_ais:
                full = self._ai_label_map.get(ai, ai)
                inh_tag = " (inherited)" if ai in inh else ""
                tips[ai] = f"{full}{inh_tag}"
                short = (full[:42] + "...") if len(full) > 45 else full
                labels[ai] = f"★ {short}" if ai in inh else short
            self.followed_ais_combo.blockSignals(True)
            self.followed_ais_combo.updateItems(all_ais, display_labels=labels, item_tooltips=tips)
            if eff and len(eff) < len(all_ais):
                self.followed_ais_combo.setCheckedItems(sorted(eff))
            self.followed_ais_combo.blockSignals(False)

    def _on_followed_ais_combo_changed(self, *_args):
        if getattr(self, "_restoring_filters", False) or not hasattr(self, "followed_ais_combo"):
            return
        if self.followed_ais_combo.isFiltered():
            checked = {str(a).strip() for a in self.followed_ais_combo.getCheckedItems() if str(a).strip()}
            self._followed_ais = checked
            self._followed_ais_customized = True
            save_followed_ais_override(self._settings, self._meeting_settings_prefix(), checked)
        else:
            self._followed_ais = set(self._inherited_followed_ais)
            self._followed_ais_customized = False
            save_followed_ais_override(
                self._settings,
                self._meeting_settings_prefix(),
                self._followed_ais,
                reset_to_inherited=True,
            )
        self._update_pill_counts()
        self._apply_filters()

    def _on_followed_ais_context_menu(self, _pos):
        menu = QMenu(self)
        act_reset = menu.addAction("🔄 Reset to Inherited Followed AIs")
        chosen = menu.exec_(QCursor.pos())
        if chosen == act_reset:
            save_followed_ais_override(
                self._settings,
                self._meeting_settings_prefix(),
                [],
                reset_to_inherited=True,
            )
            self._refresh_followed_ais_scope()
            self._update_pill_counts()
            self._apply_filters()

    def _on_hide_objected_toggled(self, checked: bool):
        self._hide_objected = bool(checked)
        self._save_filter_settings()
        self._apply_filters()

    @staticmethod
    def _clean_tdoc_id(raw: Any) -> str:
        return clean_tdoc_id(raw)

    def _item_matches_mode(self, it: dict, mode: str, cutoff: str) -> bool:
        if mode == "PRIORITY":
            return is_priority_item(it, getattr(self, "_followed_ais", set()))
        if mode == "UNREAD" and it["is_read"]:
            return False
        if mode == "FLAGGED" and not it["is_flagged"]:
            return False
        if mode == "RECENT" and str(it.get("date_str") or "") < cutoff:
            return False
        return True
