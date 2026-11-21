"""User annotation propagation, tracked thread rows, and dynamic state setters mixin."""

from PyQt5.QtCore import Qt


class TDocsStateThreadsMixin:
    """Applies user status/notes, injects tracked thread rows, and notifies view updates."""

    def _apply_user_data_logic(self):
        tdoc_dict = {row.get('TDoc', ''): row for row in self._data}

        for row in self._data:
            tdoc = row.get('TDoc', '')
            meta = self.user_data.get(tdoc, {})
            if row.get("_is_tracked_thread") and not meta:
                status = row.get("My Status") or "⚪ Neutral"
                notes = row.get("My Notes") or ""
                tags_str = row.get("My Tags") or ""
            else:
                status = meta.get('status', row.get("My Status", '⚪ Neutral') if row.get("_is_tracked_thread") else '⚪ Neutral')
                notes = meta.get('notes', row.get("My Notes", '') if row.get("_is_tracked_thread") else '')
                tags_str = meta.get('tags', row.get("My Tags", '') if row.get("_is_tracked_thread") else '')
            row['My Status'] = status or '⚪ Neutral'
            row['My Notes'] = notes or ''
            row['My Tags'] = tags_str or ''
            row['_tag_list'] = [t.strip() for t in str(row['My Tags']).replace(';', ',').split(',') if t.strip()]

        for row in self._data:
            if not row.get('My Notes') and row.get('My Status') == '⚪ Neutral' and not row.get('My Tags'):
                parent = row.get('Is revision of', '')
                if parent and parent in tdoc_dict:
                    parent_row = tdoc_dict[parent]
                    p_status = parent_row.get('My Status', '⚪ Neutral')
                    p_notes = parent_row.get('My Notes', '')
                    p_tags = parent_row.get('My Tags', '')

                    if p_status != '⚪ Neutral' or p_notes or p_tags:
                        row['My Status'] = p_status
                        row['My Notes'] = f"🔄 [From Base]: {p_notes}" if p_notes else ("🔄 [From Base]" if p_status != '⚪ Neutral' else "")
                        row['My Tags'] = p_tags
                        row['_tag_list'] = list(parent_row.get('_tag_list') or [])

    def sync_tracked_thread_rows(self, threads: list):
        """Injects or updates pinned/tracked email threads as synthetic rows in the meeting table."""
        self.beginResetModel()
        self._data = [r for r in self._data if not r.get("_is_tracked_thread")]
        thread_rows = []
        for t in (threads or []):
            tid = str(t.get("thread_id", "")).strip()
            if not tid:
                continue
            ais = [str(a).strip() for a in (t.get("linked_ais") or []) if str(a).strip()]
            primary_ai = next((a for a in ais if a in (self.agenda_map or {})), ais[0] if ais else "")
            explicit_subs = [str(s).strip() for s in (t.get("linked_sub_items") or []) if str(s).strip()]
            wis = t.get("linked_wis") or []
            tdocs = t.get("linked_tdocs") or []
            tot = int(t.get("message_count", 0))
            unread = int(t.get("unread_count", 0))
            flagged = int(t.get("flagged_count", 0))
            last_dt = t.get("last_message_date") or ""
            status_str = "📧 Active Thread"
            stat_bits = [f"{tot} msgs"]
            if unread > 0:
                stat_bits.append(f"{unread} new")
            if flagged > 0:
                stat_bits.append(f"{flagged} flagged")
            remarks_parts = [f"Tracked Email Thread ({', '.join(stat_bits)}) • Last: {last_dt or 'N/A'}"]
            if ais:
                remarks_parts.append(f"AIs: {', '.join(ais)}")
            if explicit_subs:
                remarks_parts.append(f"Sub-Topics: {', '.join(explicit_subs)}")
            if wis:
                remarks_parts.append(f"WIs: {', '.join(wis)}")
            if tdocs:
                remarks_parts.append(f"TDocs: {', '.join(tdocs)}")

            # Use explicit Sub-Topic if assigned, else inherit from an existing TDoc in the same Agenda Item
            primary_sub = explicit_subs[0] if explicit_subs else ""
            sub_item_order = 0
            for existing_row in self._data:
                if existing_row.get("_is_tracked_thread"):
                    continue
                ex_ai = str(existing_row.get("Agenda Item", "")).strip()
                ex_sub = str(existing_row.get("Sub-Item", "")).strip()
                if primary_sub and ex_sub.lower() == primary_sub.lower():
                    if not primary_ai or ex_ai == primary_ai:
                        sub_item_order = int(existing_row.get("_sub_item_order") or 0)
                        break
                elif not primary_sub and primary_ai and ex_ai == primary_ai and ex_sub:
                    primary_sub = ex_sub
                    sub_item_order = int(existing_row.get("_sub_item_order") or 0)
                    break

            all_subs = set(explicit_subs)
            if primary_sub:
                all_subs.add(primary_sub)

            row = {
                "TDoc": tid,
                "Title": f"📧 {t.get('title', 'Email Thread')}",
                "Source": t.get("participants_companies") or t.get("wg") or "Email Thread",
                "Type": "📧 Email Thread",
                "For": "Discussion",
                "Abstract": "",
                "Secretary Remarks": " | ".join(remarks_parts),
                "My Status": t.get("user_status") or "⚪ Neutral",
                "My Notes": t.get("user_notes") or "",
                "My Tags": ", ".join(t.get("user_tags") or []),
                "Agenda Item": primary_ai,
                "Sub-Item": primary_sub,
                "TDoc Status": status_str,
                "Related WIs": ", ".join(wis),
                "Related TDocs": ", ".join(tdocs),
                "Reply in": ", ".join(tdocs),
                "_is_tracked_thread": True,
                "_thread_id": tid,
                "_linked_ais": set(ais),
                "_linked_sub_items": all_subs,
                "_sub_item_order": sub_item_order,
                "_email_tot": tot,
                "_email_unread": unread,
                "_email_flagged": flagged,
                "_email_sort_key": (tot, unread, flagged),
            }
            thread_rows.append(row)

        if thread_rows:
            self._apply_company_sanitization(thread_rows)
            self._precompute_row_wi_tokens(thread_rows)
            self._data.extend(thread_rows)

        self.valid_tdocs = {str(r.get("TDoc", "")) for r in self._data if r.get("TDoc")}
        self._apply_user_data_logic()
        self._rebuild_family_cache()
        self._recompute_cached_email_stats()
        self._precompute_agenda_sort_tuples(self._data)
        self.endResetModel()

    def apply_user_data_refresh(self):
        self.beginResetModel()
        self._apply_user_data_logic()
        self.endResetModel()

    def update_data(self, new_data):
        self.beginResetModel()
        existing_threads = [r for r in self._data if r.get("_is_tracked_thread")]
        self._data = list(new_data or []) + existing_threads
        self.valid_tdocs = {str(r.get("TDoc", "")) for r in self._data if r.get("TDoc")}
        self.loading_tdocs.clear()

        # Pre-compute for a completely fresh dataset update
        self._apply_company_sanitization(self._data)
        self._precompute_row_wi_tokens(self._data)
        self._precompute_agenda_sort_tuples(self._data)
        self._apply_user_data_logic()
        self._rebuild_family_cache()
        self._recompute_cached_email_stats()
        self.endResetModel()

    def set_loading(self, tdoc: str, is_loading: bool):
        if is_loading:
            self.loading_tdocs.add(tdoc)
        else:
            self.loading_tdocs.discard(tdoc)
        for r in range(self.rowCount()):
            if self._data[r].get("TDoc") == tdoc:
                idx = self.index(r, 0)
                self.dataChanged.emit(idx, idx)
                break

    def set_highlighted_tdoc(self, tdoc: str = None):
        """Highlights a single TDoc row temporarily (or clears when None)."""
        normalized = str(tdoc).strip().upper() if tdoc else None
        prev = self._highlighted_tdoc
        if prev == normalized:
            return
        self._highlighted_tdoc = normalized
        targets = {t for t in (prev, normalized) if t}
        if not targets or self.columnCount() <= 0:
            return
        last_col = self.columnCount() - 1
        for r in range(self.rowCount()):
            row_tdoc = str(self._data[r].get("TDoc", "")).strip().upper()
            if row_tdoc in targets:
                self.dataChanged.emit(self.index(r, 0), self.index(r, last_col), [Qt.BackgroundRole])

    def set_email_counts(self, counts: dict):
        """Updates the cached email counts dictionary, recomputes keys, and repaints the Emails column."""
        self.email_counts = counts or {}
        self._email_counts = self.email_counts
        self._recompute_cached_email_stats()
        if "Emails" in self._headers:
            col_idx = self._headers.index("Emails")
            self.dataChanged.emit(self.index(0, col_idx), self.index(self.rowCount() - 1, col_idx))

    def set_agenda_map(self, agenda_map: dict):
        """Updates the cached agenda mapping and repaints the Agenda Item column."""
        self.agenda_map = agenda_map or {}
        self._precompute_row_wi_tokens(self._data)
        if "Agenda Item" in self._headers:
            col_idx = self._headers.index("Agenda Item")
            self.dataChanged.emit(self.index(0, col_idx), self.index(self.rowCount() - 1, col_idx))
