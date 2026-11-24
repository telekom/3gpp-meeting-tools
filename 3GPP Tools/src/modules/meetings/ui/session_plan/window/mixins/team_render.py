# --- File: src/modules/meetings/ui/session_plan/window/mixins/team_render.py ---
"""
Rendering of the Team view: roster chips, topic ownership table, status cells, clash board.

Mixin of SessionPlanWindow: relies on attributes set in SessionPlanWindow.__init__.
"""
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor
from PyQt5.QtWidgets import (
    QMenu,
    QMessageBox,
    QPushButton,
    QTableWidgetItem,
)
from modules.meetings.core.session_plan.session_plan_models import SessionSlot, format_venue_time_range
from modules.meetings.ui.delegate_owner_picker import DelegateOwnerPicker
from core.ui.theme import C


class TeamRenderMixin:
    """Rendering of the Team view: roster chips, topic ownership table, status cells, clash board."""

    def _render_team_roster_chips(self):
        """Refreshes the horizontal roster chips strip at the top of the Team view."""
        if not self.delegation_plan:
            return
        while self.team_member_chips_layout.count():
            item = self.team_member_chips_layout.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()

        for name, m in self.delegation_plan.members.items():
            chip = QPushButton(f"● {name}")
            chip.setStyleSheet(
                f"padding: 3px 10px; font-size: 11px; font-weight: 600; border-radius: 12px; "
                f"border: 1px solid {m.border_tint}; background: {m.bg_tint}; color: {m.color};"
            )
            chip.setCursor(Qt.PointingHandCursor)
            chip.clicked.connect(lambda _, mem=m: self._on_edit_delegation_member(mem))
            chip.setContextMenuPolicy(Qt.CustomContextMenu)
            chip.customContextMenuRequested.connect(
                lambda pos, n=name, mem=m, btn=chip: self._show_member_context_menu(pos, n, mem, btn)
            )
            self.team_member_chips_layout.addWidget(chip)

    def _render_team_topics(self):
        """Populates the Topic Ownership Matrix table with comboboxes and status tags."""
        if not self.delegation_plan:
            return

        scheduled_ais = self._get_scheduled_ais_map()
        all_ais = self._combined_ai_keys(scheduled_ais)

        self.team_topics_table.setRowCount(len(all_ais))
        collisions = self._detect_delegation_collisions()
        members_list = list(self.delegation_plan.members.values())
        initials_map = self._delegate_initials_map()

        for row, ai_num in enumerate(all_ais):
            desc, _ = self._resolve_ai_description(ai_num, scheduled_ais)

            slots_for_ai = self._slots_for_ai(ai_num, scheduled_ais)
            slots_summary = ", ".join(f"{s.day_name[:3]} {s.quarter} (Str {s.stream_number})" for s in slots_for_ai[:3])
            if len(slots_for_ai) > 3:
                slots_summary += f" +{len(slots_for_ai)-3} more"
            if not slots_summary:
                slots_summary = "— (Not scheduled)"

            to = self.delegation_plan.topic_owners.get(ai_num)
            owner_name = to.owner_name if to else ""
            owners_list = to.get_owners([m.name for m in members_list]) if to else []

            has_clash = False
            if owners_list and slots_for_ai:
                for s in slots_for_ai:
                    for o_name in owners_list:
                        quarter_clashes = collisions.get((o_name, s.day_index, s.quarter), [])
                        if any(s.id == cs.id or s.slot_key == cs.slot_key for cs in quarter_clashes):
                            has_clash = True
                            break
                    if has_clash:
                        break

            item_ai = QTableWidgetItem(ai_num)
            item_ai.setTextAlignment(Qt.AlignCenter)
            item_ai.setToolTip(f"Agenda Item {ai_num}")

            item_desc = QTableWidgetItem(desc)
            item_desc.setTextAlignment(Qt.AlignVCenter | Qt.AlignLeft)
            item_desc.setToolTip(desc)

            item_slots = QTableWidgetItem(slots_summary)
            item_slots.setTextAlignment(Qt.AlignVCenter | Qt.AlignLeft)
            item_slots.setData(Qt.UserRole, len(slots_for_ai))
            item_slots.setToolTip(f"Scheduled slots for AI {ai_num}:\n{slots_summary}")
            if not slots_for_ai:
                item_slots.setForeground(QColor(C.TEXT_FAINT))  # Muted slate for unscheduled
            else:
                item_slots.setForeground(QColor(C.TEXT_PRIMARY))
                font_slots = item_slots.font()
                font_slots.setBold(True)
                item_slots.setFont(font_slots)

            # Checkbox-based multi-select owner picker (replaces combo + custom dialog round-trip)
            picker = DelegateOwnerPicker(members_list, owner_name, initials=initials_map)
            picker.ownerChanged.connect(lambda new_owner, a=ai_num: self._on_topic_owner_changed(a, new_owner))

            override_slots = [
                s for s in slots_for_ai
                if self.delegation_plan and s.slot_key in self.delegation_plan.appointments
            ]
            override_suffix = f" · ⚡ {len(override_slots)} session override(s)" if override_slots else ""

            if has_clash:
                status_txt = f"🔴 Collision in parallel stream{override_suffix}"
                status_col = C.DANGER_DARK
                status_bg = C.DANGER_SOFT
            elif to and to.is_not_followed():
                status_txt = f"⚪ Not Followed{override_suffix}"
                status_col = C.TEXT_MUTED
                status_bg = C.BG_CANVAS
            elif owner_name:
                cov_lbl = "All Team" if to and to.is_all() else (f"{len(owners_list)} delegates" if len(owners_list) > 1 else owner_name)
                status_txt = (
                    f"🟢 Covered ({cov_lbl} · {len(slots_for_ai)} slot(s){override_suffix})"
                    if slots_for_ai
                    else f"⚪ Assigned (Default: {cov_lbl})"
                )
                status_col = C.WG_RAN_TEXT
                status_bg = C.WG_RAN_BG
            elif override_slots:
                status_txt = f"⚡ Covered via {len(override_slots)} session override(s) (No default owner)"
                status_col = C.INFO
                status_bg = C.INFO_SOFT
            else:
                status_txt = "🟡 Unassigned"
                status_col = C.WARNING_DARK
                status_bg = C.WARNING_SOFT

            item_status = QTableWidgetItem(status_txt)
            item_status.setTextAlignment(Qt.AlignVCenter | Qt.AlignLeft)
            item_status.setForeground(QColor(status_col))
            item_status.setBackground(QColor(status_bg))
            item_status.setToolTip(status_txt)
            font_status = item_status.font()
            font_status.setBold(True)
            item_status.setFont(font_status)

            self.team_topics_table.setItem(row, 0, item_ai)
            self.team_topics_table.setItem(row, 1, item_desc)
            self.team_topics_table.setItem(row, 2, item_slots)
            self.team_topics_table.setCellWidget(row, 3, picker)
            self.team_topics_table.setItem(row, 4, item_status)

        # Update button text with dynamic counts
        num_sched = sum(1 for ai in all_ais if scheduled_ais.get(ai))
        total_ais = len(all_ais)
        if hasattr(self, "btn_toggle_scheduled_only"):
            is_chk = self.btn_toggle_scheduled_only.isChecked()
            lbl = f"✓ 📅 Scheduled Only ({num_sched}/{total_ais})" if is_chk else f"📅 Scheduled Only ({num_sched}/{total_ais})"
            self.btn_toggle_scheduled_only.setText(lbl)
            self._style_scheduled_toggle(is_chk)

        if hasattr(self, "btn_reset_session_overrides"):
            app_count = len(self.delegation_plan.appointments) if self.delegation_plan else 0
            self.btn_reset_session_overrides.setText(f"↺ Reset Session Overrides ({app_count})" if app_count else "↺ Reset Session Overrides")
            self.btn_reset_session_overrides.setEnabled(app_count > 0)

        # Apply current filter
        self._filter_topic_ownership_table()

    def _update_team_topic_status_cells(self):
        """Fast status column refresh for the Topic Ownership Matrix without recreating controls."""
        if not self.delegation_plan:
            return
        collisions = self._detect_delegation_collisions()
        scheduled_ais = self._get_scheduled_ais_map()
        members_list = list(self.delegation_plan.members.values())
        for row in range(self.team_topics_table.rowCount()):
            ai_item = self.team_topics_table.item(row, 0)
            if not ai_item:
                continue
            ai_num = ai_item.text().strip()
            to = self.delegation_plan.topic_owners.get(ai_num)
            owner_name = to.owner_name if to else ""
            owners_list = to.get_owners([m.name for m in members_list]) if to else []
            slots_for_ai = self._slots_for_ai(ai_num, scheduled_ais)

            has_clash = False
            if owners_list and slots_for_ai:
                for s in slots_for_ai:
                    for o_name in owners_list:
                        quarter_clashes = collisions.get((o_name, s.day_index, s.quarter), [])
                        if any(s.id == cs.id or s.slot_key == cs.slot_key for cs in quarter_clashes):
                            has_clash = True
                            break
                    if has_clash:
                        break

            override_slots = [
                s for s in slots_for_ai
                if self.delegation_plan and s.slot_key in self.delegation_plan.appointments
            ]
            override_suffix = f" · ⚡ {len(override_slots)} session override(s)" if override_slots else ""

            if has_clash:
                status_txt = f"🔴 Collision in parallel stream{override_suffix}"
                status_col = C.DANGER_DARK
                status_bg = C.DANGER_SOFT
            elif to and to.is_not_followed():
                status_txt = f"⚪ Not Followed{override_suffix}"
                status_col = C.TEXT_MUTED
                status_bg = C.BG_CANVAS
            elif owner_name:
                cov_lbl = "All Team" if to and to.is_all() else (f"{len(owners_list)} delegates" if len(owners_list) > 1 else owner_name)
                status_txt = (
                    f"🟢 Covered ({cov_lbl} · {len(slots_for_ai)} slot(s){override_suffix})"
                    if slots_for_ai
                    else f"⚪ Assigned (Default: {cov_lbl})"
                )
                status_col = C.WG_RAN_TEXT
                status_bg = C.WG_RAN_BG
            elif override_slots:
                status_txt = f"⚡ Covered via {len(override_slots)} session override(s) (No default owner)"
                status_col = C.INFO
                status_bg = C.INFO_SOFT
            else:
                status_txt = "🟡 Unassigned"
                status_col = C.WARNING_DARK
                status_bg = C.WARNING_SOFT

            item_status = self.team_topics_table.item(row, 4)
            if not item_status:
                item_status = QTableWidgetItem(status_txt)
                self.team_topics_table.setItem(row, 4, item_status)
            else:
                item_status.setText(status_txt)
            item_status.setForeground(QColor(status_col))
            item_status.setBackground(QColor(status_bg))

        if hasattr(self, "btn_reset_session_overrides"):
            app_count = len(self.delegation_plan.appointments) if self.delegation_plan else 0
            self.btn_reset_session_overrides.setText(f"↺ Reset Session Overrides ({app_count})" if app_count else "↺ Reset Session Overrides")
            self.btn_reset_session_overrides.setEnabled(app_count > 0)

    def _jump_to_conflicts_board(self):
        """Switches to the Team Allocations view and selects the Clash & Dispatch Board tab."""
        self.btn_team.setChecked(True)
        self.view_stack.setCurrentWidget(self.team_page)
        if hasattr(self, "team_tabs"):
            self.team_tabs.setCurrentIndex(1)

    def _on_clash_cell_double_clicked(self, row: int, col: int):
        """Jumps from a conflict row in the Clash & Dispatch Board directly to that slot in the Overview Grid."""
        slot_item = self.team_clashes_table.item(row, 0)
        if not slot_item:
            return
        slot_obj = slot_item.data(Qt.UserRole)
        if not slot_obj:
            return
        day_idx = slot_obj.day_index
        q_name = slot_obj.quarter
        stream_num = slot_obj.stream_number or 1
        target_row = None
        for r_i, r_d in enumerate(self.grid_rows):
            if r_d.period == q_name and (r_d.stream_number == stream_num or r_d.row_type == "break"):
                target_row = r_i
                break
        if target_row is not None:
            self._ensure_day_visible(day_idx)
            self.btn_grid.setChecked(True)
            self.view_stack.setCurrentWidget(self.overview_page)
            col_idx = day_idx * 2 + (1 if slot_obj.sub_slot == 2 else 0)
            self.grid_table.setCurrentCell(target_row, col_idx)

    def _quick_appoint_colleague(self, slot: SessionSlot, original_owner: str, appointed_delegate: str):
        """1-click direct appointment of an available colleague to resolve a parallel stream collision."""
        self.db.set_session_appointment(
            meeting_number=self.meeting_number,
            slot_key=slot.slot_key,
            appointed_delegate=appointed_delegate,
            original_owner=original_owner,
            reason=f"Covering parallel stream collision for {original_owner}",
        )
        self.delegation_plan = self.db.get_delegation_plan(self.meeting_number)
        self._render_current_views()

    def _on_reset_all_session_overrides(self):
        """Clears all session-level coverage overrides for this meeting while keeping persistent Default AI Assignments."""
        if not self.delegation_plan or not self.delegation_plan.appointments:
            return
        ans = QMessageBox.question(
            self,
            "Reset Session Overrides",
            f"Clear all {len(self.delegation_plan.appointments)} session-level override(s) for SA2 #{self.meeting_number}?\n\nPersistent Default AI Assignments across meetings will remain untouched.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if ans == QMessageBox.Yes:
            self.db.clear_session_appointments(self.meeting_number)
            self.delegation_plan = self.db.get_delegation_plan(self.meeting_number)
            self._render_current_views()

    def _render_team_clashes(self):
        """Renders the Clash & Dispatch Board with multi-stream choices, free colleague detection, and 1-click actions."""
        if not self.delegation_plan:
            return

        collisions = self._detect_delegation_collisions()
        clash_count = len(collisions)
        self.btn_team.setText(f"👥 Team Allocations ({clash_count} ⚠️)" if clash_count else "👥 Team Allocations (Clean)")
        self.team_tabs.setTabText(1, f"⚡ Clash && Dispatch Board ({clash_count})")
        if hasattr(self, "btn_conflicts_quick"):
            self.btn_conflicts_quick.setText(f"⚠️ Conflicts ({clash_count})")
            if clash_count > 0:
                self.btn_conflicts_quick.setStyleSheet(
                    f"QPushButton {{ padding: 5px 10px; font-size: 12px; font-weight: 700; border: 1px solid {C.DANGER_BORDER_ALT}; background: {C.DANGER_SOFT}; color: {C.DANGER}; border-radius: 6px; }}"
                    f"QPushButton:hover {{ background: {C.DANGER_SOFT_HOVER}; }}"
                )
            else:
                self.btn_conflicts_quick.setStyleSheet(
                    f"QPushButton {{ padding: 5px 10px; font-size: 12px; font-weight: 600; border: 1px solid {C.WG_RAN_BORDER}; background: {C.WG_RAN_BG}; color: {C.WG_RAN_TEXT}; border-radius: 6px; }}"
                    f"QPushButton:hover {{ background: {C.EMERALD_SOFT}; }}"
                )

        if clash_count == 0:
            self.clash_kpi_lbl.setText("✅ No collisions detected! All parallel streams have distinct delegates or single coverage.")
            self.clash_kpi_lbl.setStyleSheet(
                f"background: {C.WG_RAN_BG}; color: {C.WG_RAN_TEXT}; border: 1px solid {C.WG_RAN_BORDER}; font-size: 12px; font-weight: 600; padding: 6px 10px; border-radius: 6px;"
            )
            self.team_clashes_table.setRowCount(0)
        else:
            self.clash_kpi_lbl.setText(
                f"⚠️ {clash_count} Double-Booking Collision(s) Detected! Appoint an available colleague to cover one of the parallel rooms."
            )
            self.clash_kpi_lbl.setStyleSheet(
                f"background: {C.WARNING_SOFT}; color: {C.WARNING_DARK}; border: 1px solid {C.WARNING_BORDER}; font-size: 12px; font-weight: 600; padding: 6px 10px; border-radius: 6px;"
            )
            self.team_clashes_table.setRowCount(clash_count)

            for row, ((del_name, d_idx, q_name), slots) in enumerate(collisions.items()):
                s_a = slots[0]
                time_str = format_venue_time_range(s_a.start_time_local, s_a.end_time_local)

                # Determine overlapping TUs
                has_tu1 = all(s.sub_slot in (0, 1) for s in slots) or any(
                    s1.sub_slot in (0, 1) and s2.sub_slot in (0, 1) and s1.stream_number != s2.stream_number
                    for s1 in slots for s2 in slots
                )
                has_tu2 = all(s.sub_slot in (0, 2) for s in slots) or any(
                    s1.sub_slot in (0, 2) and s2.sub_slot in (0, 2) and s1.stream_number != s2.stream_number
                    for s1 in slots for s2 in slots
                )
                tu_tag = " (TU 1&2)" if (has_tu1 and has_tu2) else (" (TU 1)" if has_tu1 else (" (TU 2)" if has_tu2 else ""))

                item_slot = QTableWidgetItem(f"{s_a.day_name} {s_a.quarter}{tu_tag}")
                item_slot.setData(Qt.UserRole, s_a)
                item_slot.setToolTip("Double-click to jump to this slot in the Overview Grid.")
                item_time = QTableWidgetItem(time_str)
                f_ct = item_time.font()
                f_ct.setBold(True)
                item_time.setFont(f_ct)
                item_time.setForeground(QColor(C.TEXT_PRIMARY))

                item_del = QTableWidgetItem(f"● {del_name}")
                m = self.delegation_plan.members.get(del_name)
                if m:
                    item_del.setForeground(QColor(m.color))
                f = item_del.font()
                f.setBold(True)
                item_del.setFont(f)

                def _format_clash_slot_desc(slot_obj: SessionSlot) -> str:
                    s_inf = self.active_snapshot.streams.get(slot_obj.stream_number) if self.active_snapshot else None
                    rm = getattr(slot_obj, "room_override", "") or (s_inf.room_name if s_inf else f"Stream {slot_obj.stream_number}")
                    _, is_cov_flag, is_all_flag = self._get_slot_effective_delegates(slot_obj)
                    role_lbl = "[All Team]" if is_all_flag else ("[Override]" if is_cov_flag else "[Owner]")
                    tu_sub = f" [TU {slot_obj.sub_slot}]" if slot_obj.sub_slot in (1, 2) else ""
                    return f"Str {slot_obj.stream_number}{tu_sub} ({rm}) {role_lbl}: {slot_obj.raw_text}"

                item_stra = QTableWidgetItem(_format_clash_slot_desc(s_a))
                item_stra.setToolTip(_format_clash_slot_desc(s_a))

                other_desc = [_format_clash_slot_desc(other_s) for other_s in slots[1:]]
                strb_txt = "\n".join(other_desc) if other_desc else _format_clash_slot_desc(slots[-1])
                item_strb = QTableWidgetItem(strb_txt)
                item_strb.setToolTip(strb_txt)

                # Compute free colleagues for this collision
                avail_map = self._compute_delegate_availability(s_a)
                free_colleagues = [
                    c_name for c_name, c_info in avail_map.items()
                    if c_info.get("status") == "free" and c_name.lower() != del_name.lower()
                ]
                if free_colleagues:
                    free_txt = "🟢 " + ", ".join(free_colleagues)
                    free_col = C.WG_RAN_TEXT
                else:
                    free_txt = "⚠️ All busy"
                    free_col = C.DANGER_DARK
                item_free = QTableWidgetItem(free_txt)
                item_free.setForeground(QColor(free_col))
                item_free.setToolTip(
                    f"Available colleagues during {s_a.day_name} {s_a.quarter}: {', '.join(free_colleagues)}"
                    if free_colleagues
                    else "All delegation members are currently assigned to parallel streams in this slot."
                )

                btn_appoint = QPushButton("⚡ Appoint Coverage ▾")
                btn_appoint.setStyleSheet(
                    f"padding: 3px 8px; font-size: 11px; font-weight: 600; border-radius: 6px; background: {C.EMERALD_PRIMARY}; color: {C.BG_CARD};"
                )
                menu = QMenu(btn_appoint)
                if free_colleagues:
                    for fc in free_colleagues:
                        for s_opt in slots:
                            act_q = menu.addAction(f"⚡ Quick-Appoint {fc} (🟢 Free) → Cover Str {s_opt.stream_number}")
                            act_q.triggered.connect(
                                lambda _, slot=s_opt, orig=del_name, target=fc: self._quick_appoint_colleague(slot, orig, target)
                            )
                    menu.addSeparator()

                for s_opt in slots:
                    s_inf = self.active_snapshot.streams.get(s_opt.stream_number) if self.active_snapshot else None
                    rm = getattr(s_opt, "room_override", "") or (s_inf.room_name if s_inf else f"Stream {s_opt.stream_number}")
                    act = menu.addAction(f"🔧 Dispatch Dialog: Cover Str {s_opt.stream_number} ({rm}): {s_opt.raw_text[:28]}...")
                    act.triggered.connect(lambda _, slot=s_opt, orig=del_name: self._on_dispatch_slot(slot, orig))

                menu.addSeparator()
                act_jump = menu.addAction(f"🔍 Show {s_a.day_name} {s_a.quarter} in Overview Grid")
                act_jump.triggered.connect(lambda _, r=row: self._on_clash_cell_double_clicked(r, 0))
                btn_appoint.setMenu(menu)

                self.team_clashes_table.setItem(row, 0, item_slot)
                self.team_clashes_table.setItem(row, 1, item_time)
                self.team_clashes_table.setItem(row, 2, item_del)
                self.team_clashes_table.setItem(row, 3, item_stra)
                self.team_clashes_table.setItem(row, 4, item_strb)
                self.team_clashes_table.setItem(row, 5, item_free)
                self.team_clashes_table.setCellWidget(row, 6, btn_appoint)

    def _render_team_view(self):
        """Refreshes the Team Allocations view: Roster chips, Topic Ownership Matrix, and Clash Board."""
        if not self.delegation_plan:
            return
        self._render_team_roster_chips()
        self._render_team_topics()
        self._render_team_clashes()
