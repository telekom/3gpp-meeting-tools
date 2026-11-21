"""AI, Sub-Topic, WI, TDoc, Tag, and delegate notes management mixin."""

from typing import Dict, List, Optional


class TrackedThreadLinksNotesMixin:
    """Cascading AI->Sub-Topic pills, TDoc family expansion, and notes/status dialog."""

    @staticmethod
    def _normalize_ai_sub_map(
        explicit_map: Optional[Dict[str, List[str]]],
        sub_item_suggestions: List[str],
    ) -> Dict[str, List[str]]:
        out: Dict[str, List[str]] = {}
        for ai_key, subs in (explicit_map or {}).items():
            ak = str(ai_key or "").strip()
            if not ak:
                continue
            bucket = out.setdefault(ak, [])
            for s in (subs or []):
                sc = str(s or "").strip()
                if sc and sc not in bucket:
                    bucket.append(sc)
        if not explicit_map:
            for item in (sub_item_suggestions or []):
                txt = str(item or "").strip()
                if " — AI " in txt:
                    sub_part, ai_part = txt.rsplit(" — AI ", 1)
                    ak = ai_part.strip()
                    sc = sub_part.strip()
                    if ak and sc:
                        bucket = out.setdefault(ak, [])
                        if sc not in bucket:
                            bucket.append(sc)
        return out

    def _get_valid_sub_items_for_ais(self, ais: List[str]) -> List[str]:
        by_lower = {k.lower(): v for k, v in self._ai_to_sub_items.items()}
        valid: List[str] = []
        seen_lower: set = set()
        for ai in (ais or []):
            for sub in by_lower.get(str(ai).strip().lower(), []):
                if sub.lower() not in seen_lower:
                    seen_lower.add(sub.lower())
                    valid.append(sub)
        return valid

    def _refresh_sub_items_dropdown(self, ais: List[str]) -> List[str]:
        valid_subs = self._get_valid_sub_items_for_ais(ais)
        if not hasattr(self, "sub_items_pills"):
            return valid_subs
        if not ais:
            self.sub_items_pills.set_empty_dropdown_label("Assign AI first")
        elif not valid_subs:
            self.sub_items_pills.set_empty_dropdown_label("No sub-topics for AI")
        self.sub_items_pills.set_suggestions(valid_subs, value_extractor=lambda s: s.split(" — ", 1)[0].strip())
        return valid_subs

    def _on_links_changed(self):
        if self._loading:
            return
        ais = self.ais_pills.tags()
        valid_subs = self._refresh_sub_items_dropdown(ais)
        valid_subs_lower = {s.lower() for s in valid_subs}
        sub_items = self.sub_items_pills.tags()
        # Automatically prune sub-topics whose parent AI is no longer assigned
        pruned_subs = [s for s in sub_items if s.lower() in valid_subs_lower]
        if pruned_subs != sub_items:
            self._loading = True
            try:
                self.sub_items_pills.setText(", ".join(pruned_subs))
            finally:
                self._loading = False
            sub_items = pruned_subs

        wis = self.wis_pills.tags()
        tdocs = self.tdocs_pills.tags()
        user_tags = self.tags_pills.tags()
        if self._tdoc_family_expander:
            expanded = self._tdoc_family_expander(tdocs)
            if [t.upper() for t in expanded] != [t.upper() for t in tdocs]:
                self._loading = True
                try:
                    self.tdocs_pills.setText(", ".join(expanded))
                finally:
                    self._loading = False
                tdocs = expanded
        self.db.update_thread_links(
            self.thread_id,
            linked_ais=ais,
            linked_wis=wis,
            linked_tdocs=tdocs,
            linked_sub_items=sub_items,
        )
        self.db.update_thread_annotations(self.thread_id, user_tags=user_tags)
        t = self.db.get_thread(self.thread_id) or {}
        self._update_links_toggle_label(t)
        if self._annotations_save_callback:
            self._annotations_save_callback(
                self.thread_id,
                str(t.get("user_status") or "⚪ Neutral"),
                str(t.get("user_notes") or ""),
                ", ".join(user_tags),
            )
        self.thread_changed.emit()
        self._render_messages()

    def _on_edit_notes_and_status(self):
        from modules.meetings.ui.tdocs_dialogs import InteractiveNotesDialog

        thread = self.db.get_thread(self.thread_id) or {}
        row_data = {
            "TDoc": self.thread_id,
            "Title": thread.get("title", ""),
            "Secretary Remarks": f"Tracked Email Thread ({thread.get('message_count', 0)} messages)",
            "My Status": thread.get("user_status") or "⚪ Neutral",
            "My Notes": thread.get("user_notes") or "",
            "My Tags": ", ".join(thread.get("user_tags") or []),
            "_is_tracked_thread": True,
        }

        def _save(tid: str, status: str, notes: str, tags_str: str = ""):
            tag_list = [x.strip() for x in str(tags_str or "").replace(";", ",").split(",") if x.strip()]
            self.db.update_thread_annotations(tid, user_status=status, user_notes=notes, user_tags=tag_list)
            if self._annotations_save_callback:
                self._annotations_save_callback(tid, status, notes, tags_str)
            self._load_data()
            self.thread_changed.emit()

        all_tags = sorted(set(self._all_existing_tags) | set(self.tags_pills.tags()), key=str.lower)
        InteractiveNotesDialog(self, self.thread_id, row_data, _save, all_existing_tags=all_tags).exec_()

