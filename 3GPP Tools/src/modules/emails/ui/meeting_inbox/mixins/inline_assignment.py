"""In-place thread pinning, auto-suggestion, and AI/Sub-Item/WI assignment mixin."""

from __future__ import annotations

from typing import List, Optional

from PyQt5.QtGui import QColor

from core.ui.theme import C
from core.ui.toast import notify_success
from modules.emails.core.tracked_threads_service import (
    auto_suggest_links,
    infer_wis_for_ais,
    normalize_subject,
    pin_thread_by_subject_or_cid,
)


class InboxInlineAssignmentMixin:
    """Handles inline pinning, auto-linking AIs/WIs, and batch inline assignment updates."""

    def _seed_messages_for_item(self, it: dict) -> List[dict]:
        raw_msgs = it.get("_group_messages") if it.get("_is_group") else [it]
        out: List[dict] = []
        for m in (raw_msgs or []):
            out.append({
                "id": m["id"],
                "subject": m.get("subject", ""),
                "sender_name": m.get("sender_name", ""),
                "sender_email": m.get("sender_email", ""),
                "company": m.get("company", ""),
                "date_sent": m.get("date_str", ""),
                "body_text": m.get("raw_body_text") or m.get("body_text", ""),
                "clean_body": m.get("clean_body") or m.get("body_text", ""),
                "conversation_id": m.get("conversation_id", ""),
                "is_read": 1 if m.get("is_read") else 0,
                "flag_status": 2 if m.get("is_flagged") else 0,
                "attachments": list(m.get("attachments") or []),
            })
        return out

    def _pin_item_in_place(
        self,
        it: dict,
        *,
        linked_ais: Optional[List[str]] = None,
        linked_wis: Optional[List[str]] = None,
        linked_sub_items: Optional[List[str]] = None,
        reload: bool = True,
    ) -> Optional[str]:
        if self.tracked_db is None or not it:
            return None
        subj = str(it.get("thread_title") or it.get("subject") or "Email Discussion").strip()
        norm = normalize_subject(subj) or subj
        tid = pin_thread_by_subject_or_cid(
            self.tracked_db,
            title=norm,
            normalized_subject=norm,
            conversation_id=str(it.get("conversation_id") or ""),
            meeting_id=self.meeting_id,
            wg=self.wg,
            agenda_map=self._agenda_map,
            valid_wis=self._valid_wis if self._wi_suggestions else None,
            tdoc_meta_map=self._tdoc_meta_map,
            linked_ais=linked_ais,
            linked_wis=linked_wis,
            linked_tdocs=list(it.get("linked_tdocs") or []),
            linked_sub_items=linked_sub_items,
            seed_messages=self._seed_messages_for_item(it),
            # Already seeded from local DBs; never run synchronous Outlook COM scans on the UI thread
            refresh_from_outlook=False,
        )
        new_preserve = f"THR::{tid}" if self._view_mode == "THREADS" else it["id"]
        if reload:
            self.reload_from_dbs(preserve_id=new_preserve)
            self.data_changed.emit()
        return tid

    def _on_pin_inline_clicked(self):
        targets = self._selected_items or ([self._current_item] if self._current_item else [])
        unpinned = [it for it in targets if not it.get("in_tracked_db")]
        if not unpinned or self.tracked_db is None:
            return
        if len(targets) == 1:
            tid = self._pin_item_in_place(unpinned[0])
            if tid:
                notify_success(self, "Thread Pinned", "Email conversation pinned as a Tracked Thread.")
            return

        preserved_ids: List[str] = []
        pinned_count = 0
        for it in targets:
            if it.get("in_tracked_db"):
                preserved_ids.append(it["id"])
                continue
            tid = self._pin_item_in_place(it, reload=False)
            if tid:
                pinned_count += 1
                preserved_ids.append(f"THR::{tid}" if self._view_mode == "THREADS" else it["id"])
        self.reload_from_dbs(preserve_ids=list(dict.fromkeys(preserved_ids)))
        self.data_changed.emit()
        if pinned_count:
            notify_success(self, "Threads Pinned", f"Pinned {pinned_count} conversation(s) as Tracked Threads.")

    def _on_auto_assign_clicked(self):
        targets = self._selected_items or ([self._current_item] if self._current_item else [])
        if not targets or self.tracked_db is None:
            return

        preserved_ids: List[str] = []
        seen_tids = set()
        for it in targets:
            sugg = auto_suggest_links(
                it.get("subject", ""),
                it.get("body_text", ""),
                agenda_map=self._agenda_map,
                valid_wis=self._valid_wis if self._wi_suggestions else None,
                tdoc_meta_map=self._tdoc_meta_map,
                extra_tdocs=it.get("linked_tdocs") or [],
            )
            ais = list(dict.fromkeys((it.get("linked_ais") or []) + (sugg.get("ais") or [])))
            subs = list(dict.fromkeys((it.get("linked_sub_items") or []) + (sugg.get("sub_items") or [])))
            wis = list(dict.fromkeys((it.get("linked_wis") or []) + (sugg.get("wis") or [])))
            if not it.get("in_tracked_db"):
                tid = self._pin_item_in_place(it, linked_ais=ais, linked_wis=wis, linked_sub_items=subs, reload=False)
                if tid:
                    preserved_ids.append(f"THR::{tid}" if self._view_mode == "THREADS" else it["id"])
            else:
                tid = it["thread_id"]
                if tid not in seen_tids:
                    seen_tids.add(tid)
                    self.tracked_db.update_thread_links(tid, linked_ais=ais, linked_sub_items=subs, linked_wis=wis)
                preserved_ids.append(it["id"])

        dedup_ids = list(dict.fromkeys(preserved_ids))
        if len(targets) == 1 and len(dedup_ids) == 1:
            self.reload_from_dbs(preserve_id=dedup_ids[0])
        else:
            self.reload_from_dbs(preserve_ids=dedup_ids)
        self.data_changed.emit()

    def _on_inline_links_changed(self):
        targets = self._selected_items or ([self._current_item] if self._current_item else [])
        if self._updating_assign_bar or not targets or self.tracked_db is None:
            return

        if len(targets) >= 2:
            cur_ais = self.assign_ais_pills.tags()
            cur_subs = self.assign_subs_pills.tags()
            cur_wis = self.assign_wis_pills.tags()

            base_ais = set(self._multi_base_common.get("ais", []))
            base_subs = set(self._multi_base_common.get("subs", []))
            base_wis = set(self._multi_base_common.get("wis", []))

            added_ais = [a for a in cur_ais if a not in base_ais]
            removed_ais = base_ais - set(cur_ais)
            added_subs = [s for s in cur_subs if s not in base_subs]
            removed_subs = base_subs - set(cur_subs)
            added_wis = [w for w in cur_wis if w not in base_wis]
            removed_wis = base_wis - set(cur_wis)

            preserved_ids: List[str] = []
            seen_tids = set()
            all_resulting_ais = list(cur_ais)
            for it in targets:
                it_ais = [a for a in (it.get("linked_ais") or []) if a not in removed_ais]
                for a in added_ais:
                    if a not in it_ais:
                        it_ais.append(a)
                for a in it_ais:
                    if a not in all_resulting_ais:
                        all_resulting_ais.append(a)

                it_wis = [w for w in (it.get("linked_wis") or []) if w not in removed_wis]
                for w in added_wis:
                    if w not in it_wis:
                        it_wis.append(w)

                it_subs = [s for s in (it.get("linked_sub_items") or []) if s not in removed_subs]
                for s in added_subs:
                    if s not in it_subs:
                        it_subs.append(s)
                if self._ai_to_sub_items:
                    it_allowed = {sub for ai in it_ais for sub in self._ai_to_sub_items.get(ai, [])}
                    it_subs = [s for s in it_subs if s in it_allowed]

                if not it.get("in_tracked_db"):
                    tid = self._pin_item_in_place(
                        it,
                        linked_ais=it_ais,
                        linked_wis=it_wis,
                        linked_sub_items=it_subs,
                        reload=False,
                    )
                    if tid:
                        preserved_ids.append(f"THR::{tid}" if self._view_mode == "THREADS" else it["id"])
                else:
                    tid = it["thread_id"]
                    if tid not in seen_tids:
                        seen_tids.add(tid)
                        self.tracked_db.update_thread_links(
                            tid,
                            linked_ais=it_ais,
                            linked_sub_items=it_subs,
                            linked_wis=it_wis,
                        )
                    preserved_ids.append(it["id"])

            self._multi_base_common = {
                "ais": list(cur_ais),
                "subs": list(cur_subs),
                "wis": list(cur_wis),
            }
            self.reload_from_dbs(preserve_ids=list(dict.fromkeys(preserved_ids)))
            self.data_changed.emit()
            return

        it = targets[0]
        selected_ais = self.assign_ais_pills.tags()
        self._refresh_assign_sub_suggestions(selected_ais)
        allowed_subs = {
            sub
            for ai in selected_ais
            for sub in self._ai_to_sub_items.get(ai, [])
        }
        selected_subs = self.assign_subs_pills.tags()
        if self._ai_to_sub_items:
            pruned_subs = [s for s in selected_subs if s in allowed_subs]
            if pruned_subs != selected_subs:
                selected_subs = pruned_subs
                self._updating_assign_bar = True
                try:
                    self.assign_subs_pills.setText(", ".join(pruned_subs))
                finally:
                    self._updating_assign_bar = False
        selected_wis = self.assign_wis_pills.tags()
        if selected_ais and not selected_wis:
            suggested_wis = infer_wis_for_ais(
                selected_ais,
                agenda_map=self._agenda_map,
                tdoc_meta_map=self._tdoc_meta_map,
                valid_wis=self._valid_wis if self._wi_suggestions else None,
            )
            if suggested_wis:
                selected_wis = suggested_wis
                self._updating_assign_bar = True
                try:
                    self.assign_wis_pills.setText(", ".join(suggested_wis))
                finally:
                    self._updating_assign_bar = False

        if not it.get("in_tracked_db"):
            self._pin_item_in_place(
                it,
                linked_ais=selected_ais,
                linked_wis=selected_wis,
                linked_sub_items=selected_subs,
            )
            return

        tid = it["thread_id"]
        self.tracked_db.update_thread_links(
            tid,
            linked_ais=selected_ais,
            linked_sub_items=selected_subs,
            linked_wis=selected_wis,
        )
        for m in self._all_items:
            if m.get("thread_id") == tid:
                m["linked_ais"] = list(selected_ais)
                m["linked_sub_items"] = list(selected_subs)
                m["linked_wis"] = list(selected_wis)
        self._grouped_items = self._build_grouped_thread_items()
        cur_row = self.table.currentRow()
        if 0 <= cur_row < len(self._visible_items):
            self._visible_items[cur_row]["linked_ais"] = list(selected_ais)
            self._visible_items[cur_row]["linked_sub_items"] = list(selected_subs)
            self._visible_items[cur_row]["linked_wis"] = list(selected_wis)
            ai_cell = self.table.item(cur_row, 2)
            if ai_cell:
                ai_cell.setText(", ".join(selected_ais) if selected_ais else "❓ —")
                ai_cell.setForeground(QColor(C.TEXT_PRIMARY if selected_ais else C.WARNING))
        self._populate_filter_combos()
        self.data_changed.emit()
