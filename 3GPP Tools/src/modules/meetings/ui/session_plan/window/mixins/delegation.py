# --- File: src/modules/meetings/ui/session_plan/window/mixins/delegation.py ---
"""
Delegation logic: effective delegates, collisions, availability, delegate focus buttons,
roster member CRUD, topic ownership, dispatch, and delegation import/export.

Mixin of SessionPlanWindow: relies on attributes set in SessionPlanWindow.__init__.
"""
from typing import (
    Dict,
    List,
    Optional,
    Set,
    Tuple,
)
from PyQt5.QtCore import (
    QPoint,
    QSettings,
    Qt,
)
from PyQt5.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QDialog,
    QMenu,
    QMessageBox,
    QPushButton,
)
from modules.meetings.core.session_plan.session_plan_exporter import SessionPlanExporter
from modules.meetings.core.session_plan.session_plan_models import (
    DelegationMember,
    SessionSlot,
    compute_unique_initials,
    resolve_slot_effective_delegates,
)
from modules.meetings.ui.session_plan.session_plan_dialogs import (
    DelegationImportDialog,
    DelegationMemberDialog,
    MultiDelegateSelectDialog,
    SessionDispatchDialog,
)
from core.ui.theme import C


class DelegationMixin:
    """Delegation logic: effective delegates, collisions, availability, delegate focus buttons,"""

    # --- Delegation Engine & Team View ---

    def _get_slot_effective_delegates(self, slot: SessionSlot) -> Tuple[List[str], bool, bool]:
        """
        Returns (delegates_list, is_appointment_override, is_all).
        Priority:
        1. Explicit session appointment (slot override).
        2. Topic ownership from extracted slot items.
        3. Boundary-aware raw text match against assigned AI tokens
           (``20.6.2`` must never match ``20.6.24``).
        Delegates to the shared ``resolve_slot_effective_delegates`` so the grid,
        clash board, and exporters always agree.
        """
        return resolve_slot_effective_delegates(slot, self.delegation_plan, agenda_map=self.agenda_map)

    def _get_slot_effective_delegate(self, slot: SessionSlot) -> Tuple[Optional[str], bool]:
        """Backward-compatible single-string delegate helper."""
        dels, is_cov, is_all = self._get_slot_effective_delegates(slot)
        if not dels:
            return None, False
        if is_all:
            return "ALL", is_cov
        return ", ".join(dels), is_cov

    def _slots_time_overlap(self, s1: SessionSlot, s2: SessionSlot) -> bool:
        """
        Determines whether two slots physically overlap in meeting time.
        - Must be the same day and quarter.
        - If either slot spans the full quarter (sub_slot == 0), they overlap.
        - If both are half-quarters (sub_slot 1 or 2), they overlap iff sub_slot == sub_slot.
        """
        if s1.day_index != s2.day_index or s1.quarter != s2.quarter:
            return False
        if s1.sub_slot == 0 or s2.sub_slot == 0:
            return True
        return s1.sub_slot == s2.sub_slot

    def _detect_delegation_collisions(self) -> Dict[Tuple[str, int, str], List[SessionSlot]]:
        """
        Detects parallel stream collisions for delegates taking half-quarter TU1/TU2 into account.
        Key: (delegate_name, day_index, quarter) -> list of colliding SessionSlots.
        Only slots that physically overlap in time across different active streams (>0) are flagged.
        """
        if not self.active_snapshot or not self.delegation_plan:
            return {}

        slot_map: Dict[Tuple[str, int, str], List[SessionSlot]] = {}
        for s in self.active_snapshot.slots:
            del_names, _, _ = self._get_slot_effective_delegates(s)
            for del_name in del_names:
                key = (del_name, s.day_index, s.quarter)
                slot_map.setdefault(key, []).append(s)

        collisions: Dict[Tuple[str, int, str], List[SessionSlot]] = {}
        for (del_name, d_idx, q_name), slots in slot_map.items():
            colliding_slots: List[SessionSlot] = []
            for i, s1 in enumerate(slots):
                for j, s2 in enumerate(slots):
                    if i != j and s1.stream_number != s2.stream_number and s1.stream_number > 0 and s2.stream_number > 0:
                        if self._slots_time_overlap(s1, s2):
                            if s1 not in colliding_slots:
                                colliding_slots.append(s1)
                            if s2 not in colliding_slots:
                                colliding_slots.append(s2)
            if colliding_slots:
                collisions[(del_name, d_idx, q_name)] = colliding_slots

        return collisions

    def _compute_delegate_availability(self, target_slot: SessionSlot) -> Dict[str, dict]:
        """
        For each member in the roster, determines availability for target_slot:
        returns: member_name -> {'status': 'free'|'busy', 'reason': str}
        """
        res = {}
        if not self.active_snapshot or not self.delegation_plan:
            return res

        for name, m in self.delegation_plan.members.items():
            busy_slots = []
            for s in self.active_snapshot.slots:
                if (
                    s.stream_number != target_slot.stream_number
                    and s.stream_number > 0
                    and self._slots_time_overlap(s, target_slot)
                ):
                    del_names, _, _ = self._get_slot_effective_delegates(s)
                    if any(dn.lower() == name.lower() for dn in del_names):
                        busy_slots.append(s)
                        busy_slots.append(s)

            if busy_slots:
                streams_str = ", ".join(f"Str {bs.stream_number}" for bs in busy_slots)
                res[name] = {"status": "busy", "reason": f"Busy in {streams_str} ({busy_slots[0].raw_text[:28]}...)"}
            else:
                res[name] = {"status": "free", "reason": "Available (no parallel session)"}

        return res

    def _delegate_initials_map(self) -> Dict[str, str]:
        """Returns roster-unique badge initials (cached until the roster changes)."""
        names = tuple(self.delegation_plan.members.keys()) if self.delegation_plan else ()
        if self._initials_cache[0] != names:
            self._initials_cache = (names, compute_unique_initials(names))
        return self._initials_cache[1]

    def _self_member(self) -> Optional[DelegationMember]:
        if not self.delegation_plan:
            return None
        return next((m for m in self.delegation_plan.members.values() if m.is_self), None)

    def _update_overview_delegate_strip(self):
        """Rebuilds the multi-select Delegate Focus checkboxes above the Overview Grid."""
        if not hasattr(self, "delegate_chips_layout"):
            return

        while self.delegate_chips_layout.count():
            item = self.delegate_chips_layout.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()
        self.dynamic_delegate_chips.clear()
        self.btn_chip_me = None

        roster = self.delegation_plan.members if self.delegation_plan else {}
        # Drop focus entries for delegates that were removed/renamed
        self.grid_focus_delegates = {n for n in self.grid_focus_delegates if n in roster}

        initials = self._delegate_initials_map()
        for name, m in roster.items():
            chk = QCheckBox(f"{initials.get(name, m.initials)} · {name}")
            chk.setProperty("delegate_name", name)
            chk.setCursor(Qt.PointingHandCursor)
            chk.setToolTip(f"Highlight sessions covered by {name}. Tick several delegates to combine.")
            chk.setChecked(name in self.grid_focus_delegates)
            chk.toggled.connect(lambda checked, n=name: self._on_focus_delegate_toggled(n, checked))
            self.delegate_chips_layout.addWidget(chk)
            self.dynamic_delegate_chips.append(chk)

        self._style_focus_buttons()

    def _style_focus_buttons(self):
        """Applies active/inactive styling to the focus strip according to the focus set."""
        base = "padding: 2px 8px; font-size: 11px; font-weight: 600; border-radius: 12px;"
        active = f"border: 1px solid {C.INFO_ACCENT}; background: {C.INFO_ACCENT}; color: {C.BG_CARD};"
        idle = f"border: 1px solid {C.BORDER_STRONG}; background: {C.BG_CARD}; color: {C.TEXT_BODY};"
        self.btn_chip_all.setStyleSheet(f"QPushButton {{ {base} {active if not self.grid_focus_delegates else idle} }}")

        self_member = self._self_member()
        if self.btn_chip_me is not None:
            only_me = bool(self_member) and self.grid_focus_delegates == {self_member.name}
            col = self_member.color if self_member else C.TEXT_PRIMARY
            style = f"border: 1px solid {col}; background: {col}; color: {C.BG_CARD};" if only_me else idle
            self.btn_chip_me.setStyleSheet(f"QPushButton {{ {base} {style} }}")

        for chk in self.dynamic_delegate_chips:
            name = chk.property("delegate_name")
            m = self.delegation_plan.members.get(name) if self.delegation_plan else None
            col = m.color if m else C.TEXT_BODY
            bg = m.bg_tint if m else C.BG_CARD
            bdr = m.border_tint if m else C.BORDER_STRONG
            if chk.isChecked():
                bg, bdr = (m.compute_tint(col, 0.22) if m else C.BORDER_SUBTLE), col
            chk.setStyleSheet(
                f"QCheckBox {{ {base} padding: 2px 8px 2px 6px; color: {col}; background: {bg}; border: 1px solid {bdr}; }}"
                f"QCheckBox::indicator {{ width: 12px; height: 12px; }}"
            )

    def _sync_focus_mirror(self):
        """Keeps the legacy single-value ``grid_highlight_delegate`` attribute coherent."""
        self_member = self._self_member()
        if not self.grid_focus_delegates:
            self.grid_highlight_delegate = "ALL"
        elif self_member and self.grid_focus_delegates == {self_member.name}:
            self.grid_highlight_delegate = "SELF"
        elif len(self.grid_focus_delegates) == 1:
            self.grid_highlight_delegate = next(iter(self.grid_focus_delegates))
        else:
            self.grid_highlight_delegate = "MULTI"

    def _set_focus_delegates(self, names: Set[str]):
        """Sets the focus set, syncs checkboxes without signal storms, and re-renders the grid once."""
        self.grid_focus_delegates = set(names)
        for chk in self.dynamic_delegate_chips:
            chk.blockSignals(True)
            chk.setChecked(chk.property("delegate_name") in self.grid_focus_delegates)
            chk.blockSignals(False)
        self._sync_focus_mirror()
        self._style_focus_buttons()
        if getattr(self, "active_snapshot", None) is not None:
            self._render_overview_grid()

    def _on_focus_delegate_toggled(self, name: str, checked: bool):
        names = set(self.grid_focus_delegates)
        if checked:
            names.add(name)
        else:
            names.discard(name)
        self._set_focus_delegates(names)

    def _on_delegate_highlight_clicked(self, key: str):
        """
        Backward-compatible focus entry point:
        "ALL" clears the focus, "SELF" focuses only the local user, a roster name toggles that delegate.
        """
        if key == "ALL":
            self._set_focus_delegates(set())
        elif key == "SELF":
            me = self._self_member()
            self._set_focus_delegates({me.name} if me else set())
        else:
            names = set(self.grid_focus_delegates)
            names.symmetric_difference_update({key})
            self._set_focus_delegates(names)

    def _on_delegate_agenda_changed(self):
        self._render_tracked_topics()

    def _get_scheduled_ais_map(self) -> Dict[str, List[SessionSlot]]:
        """Maps AI numbers to the list of SessionSlots in which they are scheduled."""
        scheduled_ais: Dict[str, List[SessionSlot]] = {}
        if self.active_snapshot:
            for s in self.active_snapshot.slots:
                for it in s.items:
                    if it.ai_number:
                        scheduled_ais.setdefault(it.ai_number, []).append(s)
        return scheduled_ais

    def _show_member_context_menu(self, pos: QPoint, name: str, member: DelegationMember, btn: QPushButton):
        menu = QMenu(self)
        menu.addAction("✏️ Edit Member...", lambda: self._on_edit_delegation_member(member))
        menu.addSeparator()
        menu.addAction("🗑️ Remove Member", lambda: self._on_delete_delegation_member(name))
        menu.exec_(btn.mapToGlobal(pos))

    def _on_add_delegation_member(self):
        existing = list(self.delegation_plan.members.keys()) if self.delegation_plan else []
        dlg = DelegationMemberDialog(parent=self, member=None, existing_names=existing)
        if dlg.exec_() == QDialog.Accepted:
            new_mem = dlg.get_member()
            self.db.save_delegation_member(self.meeting_number, new_mem)
            self.delegation_plan = self.db.get_delegation_plan(self.meeting_number)
            self._render_current_views()

    def _on_edit_delegation_member(self, member: DelegationMember):
        existing = list(self.delegation_plan.members.keys()) if self.delegation_plan else []
        dlg = DelegationMemberDialog(parent=self, member=member, existing_names=existing)
        if dlg.exec_() == QDialog.Accepted:
            updated_mem = dlg.get_member()
            self.db.save_delegation_member(self.meeting_number, updated_mem)
            self.delegation_plan = self.db.get_delegation_plan(self.meeting_number)
            self._render_current_views()

    def _on_delete_delegation_member(self, name: str):
        ans = QMessageBox.question(
            self,
            "Confirm Delete",
            f"Are you sure you want to remove '{name}' from the delegation roster?\nTheir assigned topics and appointments will be cleared.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if ans == QMessageBox.Yes:
            self.db.delete_delegation_member(self.meeting_number, name)
            self.delegation_plan = self.db.get_delegation_plan(self.meeting_number)
            self._render_current_views()

    def _on_set_self_delegate(self, name: str):
        self.db.set_self_delegate(self.meeting_number, name)
        self.delegation_plan = self.db.get_delegation_plan(self.meeting_number)
        self._render_current_views()

    def _on_topic_owner_changed(self, ai_number: str, owner_name: str):
        self.db.set_topic_owner(self.meeting_number, ai_number, owner_name)
        self.delegation_plan = self.db.get_delegation_plan(self.meeting_number)
        self._render_overview_grid()
        self._render_tracked_topics()
        self._render_team_clashes()
        self._update_team_topic_status_cells()

    def _on_topic_owner_combo_activated(self, ai_num: str, desc: str, combo: QComboBox):
        choice = combo.currentData()
        if choice == "__CUSTOM__":
            members = list(self.delegation_plan.members.values()) if self.delegation_plan else []
            to = self.delegation_plan.topic_owners.get(ai_num) if self.delegation_plan else None
            cur = to.owner_name if to else ""
            dlg = MultiDelegateSelectDialog(f"AI {ai_num}: {desc}", members, cur, self)
            if dlg.exec_() == QDialog.Accepted:
                new_owner = dlg.get_owner_string()
                self._on_topic_owner_changed(ai_num, new_owner)
                self._render_team_topics()
            else:
                to = self.delegation_plan.topic_owners.get(ai_num) if self.delegation_plan else None
                cur = to.owner_name if to else ""
                idx = combo.findData(cur)
                if idx >= 0:
                    combo.setCurrentIndex(idx)
        else:
            self._on_topic_owner_changed(ai_num, choice)

    def _on_dispatch_slot(self, slot: SessionSlot, original_owner: str):
        members = list(self.delegation_plan.members.values()) if self.delegation_plan else []
        if not members:
            QMessageBox.information(self, "No Roster", "Please add delegation members before appointing coverage.")
            return

        avail = self._compute_delegate_availability(slot)
        dlg = SessionDispatchDialog(slot, original_owner, members, avail, self)
        if dlg.exec_() == QDialog.Accepted:
            appointed, reason = dlg.get_result()
            if appointed:
                self.db.set_session_appointment(
                    meeting_number=self.meeting_number,
                    slot_key=slot.slot_key,
                    appointed_delegate=appointed,
                    original_owner=original_owner,
                    reason=reason,
                )
            else:
                self.db.delete_session_appointment(self.meeting_number, slot.slot_key)

            self.delegation_plan = self.db.get_delegation_plan(self.meeting_number)
            self._render_current_views()

    def _on_revert_dispatch(self, slot: SessionSlot):
        """Removes ad-hoc appointment override, restoring ownership to original topic owner."""
        self.db.delete_session_appointment(self.meeting_number, slot.slot_key)
        self.delegation_plan = self.db.get_delegation_plan(self.meeting_number)
        self._render_current_views()

    def _style_scheduled_toggle(self, checked: bool):
        base = "padding: 5px 12px; font-size: 11.5px; font-weight: 600; border-radius: 6px;"
        if checked:
            self.btn_toggle_scheduled_only.setStyleSheet(
                f"QPushButton {{ {base} background: {C.INFO_SOFT}; border: 1.5px solid {C.ACCENT_ALT}; color: {C.INFO}; }}"
                f"QPushButton:hover {{ background: {C.INFO_SOFT_STRONG}; }}"
            )
        else:
            self.btn_toggle_scheduled_only.setStyleSheet(
                f"QPushButton {{ {base} background: {C.BG_SURFACE}; border: 1px solid {C.BORDER_STRONG}; color: {C.TEXT_SECONDARY}; }}"
                f"QPushButton:hover {{ background: {C.BG_CANVAS}; color: {C.TEXT_STRONG}; border-color: {C.TEXT_FAINT}; }}"
            )

    def _on_scheduled_only_toggled(self, checked: bool):
        settings = QSettings("3GPP_Tools", "SessionPlanWindow")
        settings.setValue("team_scheduled_only", checked)
        if hasattr(self, "btn_toggle_scheduled_only"):
            total_ais = self.team_topics_table.rowCount()
            num_sched = sum(
                1 for r in range(total_ais)
                if (self.team_topics_table.item(r, 2) and (self.team_topics_table.item(r, 2).data(Qt.UserRole) or 0) > 0)
            )
            lbl = f"✓ 📅 Scheduled Only ({num_sched}/{total_ais})" if checked else f"📅 Scheduled Only ({num_sched}/{total_ais})"
            self.btn_toggle_scheduled_only.setText(lbl)
            self._style_scheduled_toggle(checked)
        self._filter_topic_ownership_table()

    def _filter_topic_ownership_table(self, query: Optional[str] = None):
        q = (self.team_ai_search.text() if hasattr(self, "team_ai_search") else (query or "")).lower().strip()
        scheduled_only = self.btn_toggle_scheduled_only.isChecked() if hasattr(self, "btn_toggle_scheduled_only") else False

        for r in range(self.team_topics_table.rowCount()):
            ai_item = self.team_topics_table.item(r, 0)
            desc_item = self.team_topics_table.item(r, 1)
            slots_item = self.team_topics_table.item(r, 2)
            picker = self.team_topics_table.cellWidget(r, 3)

            slot_count = slots_item.data(Qt.UserRole) if slots_item else 0
            if slot_count is None:
                slot_count = 0 if (not slots_item or "(not scheduled)" in slots_item.text().lower() or "—" in slots_item.text()) else 1

            if scheduled_only and slot_count == 0:
                self.team_topics_table.setRowHidden(r, True)
                continue

            owner_txt = ""
            if picker:
                if hasattr(picker, "text"):
                    owner_txt = picker.text()
                elif hasattr(picker, "currentText"):
                    owner_txt = picker.currentText()

            txt = f"{ai_item.text()} {desc_item.text()} {owner_txt}".lower() if ai_item and desc_item else ""
            self.team_topics_table.setRowHidden(r, bool(q and q not in txt))

    def _auto_assign_followed_wis_to_self(self):
        if not self.delegation_plan or not self.delegation_plan.members:
            QMessageBox.warning(self, "No Roster", "Please add delegation members before assigning topics.")
            return

        names = list(self.delegation_plan.members.keys())
        if len(names) == 1:
            target_name = names[0]
        else:
            from PyQt5.QtWidgets import QInputDialog
            target_name, ok = QInputDialog.getItem(
                self,
                "Auto-Assign Followed WIs",
                "Select delegate to assign followed Work Item topics to:",
                names,
                0,
                False,
            )
            if not ok or not target_name:
                return

        count = 0
        if self.agenda_map:
            for ai_num, item in self.agenda_map.items():
                acronym = getattr(item, "acronym", "") or ""
                desc = getattr(item, "description", "") or ""
                matches = (
                    ai_num.lower() in self.tracked_tokens
                    or (acronym and acronym.lower() in self.tracked_tokens)
                    or any(t in desc.lower() for t in self.tracked_tokens if t)
                )
                if matches:
                    self.db.set_topic_owner(self.meeting_number, ai_num, target_name)
                    count += 1

        self.delegation_plan = self.db.get_delegation_plan(self.meeting_number)
        self._render_current_views()
        QMessageBox.information(self, "Auto-Assignment Complete", f"Assigned {count} topic(s) to {target_name}.")

    def _on_export_delegation_email(self):
        if not self.active_snapshot or not self.delegation_plan:
            return
        subject, html_body, plain_text = SessionPlanExporter.generate_delegation_email_summary(
            snapshot=self.active_snapshot,
            plan=self.delegation_plan,
            meeting_tz_offset_min=self.tz_offset_min,
            tz_name=self.tz_name,
            agenda_map=self.agenda_map,
        )
        QApplication.clipboard().setText(html_body)
        QMessageBox.information(
            self,
            "Delegation Summary Copied",
            f"Subject: {subject}\n\nPre-formatted executive rich HTML email summary copied to clipboard!\nYou can paste it directly into Microsoft Outlook or Teams.",
        )

    def _on_copy_delegation_json(self):
        if not self.delegation_plan:
            return
        payload = SessionPlanExporter.serialize_delegation_plan_json(self.delegation_plan)
        QApplication.clipboard().setText(payload)
        QMessageBox.information(
            self,
            "Plan JSON Copied",
            "Delegation Plan JSON copied to clipboard.\nColleagues can import this using 'Import Team Plan'.",
        )

    def _on_import_delegation_plan(self):
        dlg = DelegationImportDialog(
            current_roster=list(self.delegation_plan.members.keys()) if self.delegation_plan else [],
            parent=self,
        )
        if dlg.exec_() == QDialog.Accepted:
            data, is_merge, claimed_self = dlg.get_data()
            if data:
                count = self.db.import_delegation_plan_dict(self.meeting_number, data, merge=is_merge)
                if claimed_self:
                    self.db.set_self_delegate(self.meeting_number, claimed_self)
                self.delegation_plan = self.db.get_delegation_plan(self.meeting_number)
                self._render_current_views()
                QMessageBox.information(
                    self,
                    "Import Complete",
                    f"Successfully imported delegation plan ({count} items updated).\nYour identity is set to: {claimed_self or 'Default'}.",
                )
