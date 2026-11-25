# --- File: src/modules/meetings/ui/session_plan/window/mixins/interactions.py ---
"""
User interactions on grid/drafting/tracked tables: double-clicks and context menus.

Mixin of SessionPlanWindow: relies on attributes set in SessionPlanWindow.__init__.
"""
import re
from PyQt5.QtCore import QPoint, Qt
from PyQt5.QtWidgets import (
    QApplication,
    QDialog,
    QMenu,
)
from modules.meetings.core.session_plan.session_plan_models import expand_ai_number
from modules.meetings.ui.session_plan.session_plan_dialogs import MultiDelegateSelectDialog, SessionPlanNoteDialog
from modules.meetings.ui.session_plan.window.constants import DAY_NAMES, QUARTER_ORDER


class InteractionsMixin:
    """User interactions on grid/drafting/tracked tables: double-clicks and context menus."""

    def _on_grid_cell_double_clicked(self, row: int, col: int):
        if row < 0 or row >= len(self.grid_rows):
            return
        r_desc = self.grid_rows[row]
        day_idx = col // 2
        if day_idx < 0 or day_idx >= len(DAY_NAMES):
            return
        day = DAY_NAMES[day_idx]
        span = self.grid_table.columnSpan(row, col)
        tu = 1 if col % 2 == 0 else 2
        sub_slot = 0 if span == 2 else tu

        slot_key = f"{day}|{r_desc.period}|{sub_slot}|{r_desc.stream_number}"
        ann = self.annotations.get(slot_key)
        cur_slot_loc = self.slot_locations.get(slot_key, "")
        label = f"{day} {r_desc.period} (Str {r_desc.stream_number})" if r_desc.stream_number > 0 else f"{day} {r_desc.period}"
        dlg = SessionPlanNoteDialog(
            label,
            ann,
            current_slot_location=cur_slot_loc,
            parent=self,
        )
        if dlg.exec_() == QDialog.Accepted:
            badge, colleague, comment = dlg.get_data()
            self.db.set_annotation(
                meeting_number=self.meeting_number,
                target_type="slot",
                target_key=slot_key,
                status_badge=badge,
                assigned_colleague=colleague,
                comment=comment,
            )
            self.db.set_slot_location(
                meeting_number=self.meeting_number,
                slot_key=slot_key,
                location_note=dlg.get_slot_location(),
                meeting_dir=self.meeting_dir,
            )
            self.annotations = self.db.get_annotations(self.meeting_number)
            self.slot_locations = self.db.get_slot_locations(self.meeting_number, meeting_dir=self.meeting_dir)
            self._render_current_views()

    def _on_drafting_cell_double_clicked(self, row: int, col: int):
        d_item = self.drafting_table.item(row, 0)
        q_item = self.drafting_table.item(row, 1)
        s_item = self.drafting_table.item(row, 3)
        if not d_item or not q_item:
            return
        d_name = d_item.text().strip()
        q_name = q_item.text().strip()
        m_s = re.search(r"Str\s*(\d+)", s_item.text() if s_item else "")
        stream_num = int(m_s.group(1)) if m_s else 1

        if d_name in DAY_NAMES:
            day_idx = DAY_NAMES.index(d_name)
            self._jump_to_slot_in_grid(day_idx, q_name, stream_num)

    def _on_tracked_cell_double_clicked(self, row: int, col: int):
        d_item = self.tracked_table.item(row, 0)
        q_item = self.tracked_table.item(row, 1)
        s_item = self.tracked_table.item(row, 3)
        if not d_item or not q_item:
            return
        d_name = d_item.text().strip()
        q_name = q_item.text().strip()
        m_s = re.search(r"Str\s*(\d+)", s_item.text() if s_item else "")
        stream_num = int(m_s.group(1)) if m_s else 1

        if d_name in DAY_NAMES:
            day_idx = DAY_NAMES.index(d_name)
            self._jump_to_slot_in_grid(day_idx, q_name, stream_num)

    def _on_tracked_context_menu(self, pos: QPoint):
        r = self.tracked_table.rowAt(pos.y())
        if r < 0:
            return
        d_item = self.tracked_table.item(r, 0)
        q_item = self.tracked_table.item(r, 1)
        topic_item = self.tracked_table.item(r, 5)
        if not d_item or not q_item or not topic_item:
            return

        d_name = d_item.text().strip()
        q_name = q_item.text().strip()

        menu = QMenu(self)
        menu.setToolTipsVisible(True)

        act_jump = menu.addAction(f"🏢 Show in Overview Grid ({d_name} {q_name})")
        act_jump.triggered.connect(lambda: self._on_tracked_cell_double_clicked(r, 0))

        if d_name in DAY_NAMES and q_name in QUARTER_ORDER:
            c = DAY_NAMES.index(d_name)
            grid_r = QUARTER_ORDER.index(q_name)
            act_note = menu.addAction(f"📝 Edit Note for {d_name} {q_name}...")
            act_note.triggered.connect(lambda: self._on_grid_cell_double_clicked(grid_r, c))

        s = d_item.data(Qt.UserRole)
        if self.delegation_plan and self.delegation_plan.members and s:
            menu.addSeparator()
            del_name, is_cov = self._get_slot_effective_delegate(s)
            disp_name = "All Team" if del_name == "ALL" else (del_name or "")
            app = self.delegation_plan.appointments.get(s.slot_key)
            orig_owner = app.original_owner if (app and app.original_owner) else (del_name or "")
            if is_cov:
                act_disp = menu.addAction(f"⚡ Change Coverage for this Session (Covered by {disp_name})...")
                act_disp.triggered.connect(lambda _, slot=s, orig=orig_owner: self._on_dispatch_slot(slot, orig))
                act_rev = menu.addAction(f"↺ Revert Coverage to {orig_owner or 'Topic Owner'}")
                act_rev.triggered.connect(lambda _, slot=s: self._on_revert_dispatch(slot))
            else:
                curr_str = f"({disp_name})" if disp_name else "(Unassigned)"
                act_disp = menu.addAction(f"⚡ Appoint Coverage for this Session {curr_str}...")
                act_disp.triggered.connect(lambda _, slot=s, orig=orig_owner: self._on_dispatch_slot(slot, orig))

        menu.addSeparator()
        act_copy = menu.addAction("📋 Copy Slot Details")
        act_copy.triggered.connect(lambda: self._copy_tracked_row(r))

        menu.exec_(self.tracked_table.viewport().mapToGlobal(pos))

    def _copy_tracked_row(self, row: int):
        parts = [
            self.tracked_table.item(row, i).text()
            for i in range(self.tracked_table.columnCount())
            if self.tracked_table.item(row, i)
        ]
        QApplication.clipboard().setText(" | ".join(parts))

    def _on_grid_context_menu(self, pos: QPoint):
        r = self.grid_table.rowAt(pos.y())
        c = self.grid_table.columnAt(pos.x())
        if r < 0 or c < 0 or r >= len(self.grid_rows):
            return

        r_desc = self.grid_rows[r]
        day_idx = c // 2
        if day_idx < 0 or day_idx >= len(DAY_NAMES):
            return
        day = DAY_NAMES[day_idx]
        span = self.grid_table.columnSpan(r, c)
        tu = 1 if c % 2 == 0 else 2
        sub_slot = 0 if span == 2 else tu

        # Find slots in this cell
        if r_desc.row_type == "stream":
            slots_here = [
                s for s in (self.active_snapshot.slots if self.active_snapshot else [])
                if s.day_index == day_idx and s.quarter == r_desc.period and s.stream_number == r_desc.stream_number and (s.sub_slot == sub_slot or s.sub_slot == 0)
            ]
        else:
            slots_here = [
                s for s in (self.active_snapshot.slots if self.active_snapshot else [])
                if s.day_index == day_idx and s.quarter == r_desc.period
            ]

        menu = QMenu(self)
        menu.setToolTipsVisible(True)

        lbl_target = f"{day} {r_desc.period} (Str {r_desc.stream_number})" if r_desc.stream_number > 0 else f"{day} {r_desc.period}"
        act_note = menu.addAction(f"📝 Edit Note / Priority for {lbl_target}...")
        act_note.triggered.connect(lambda: self._on_grid_cell_double_clicked(r, c))

        if slots_here:
            menu.addSeparator()
            for s in slots_here:
                for it in s.items:
                    if it.ai_number:
                        expanded_ais = expand_ai_number(it.ai_number, self.agenda_map)
                        for target_ai in expanded_ais[:12]:
                            ai_obj = self.agenda_map.get(target_ai) if self.agenda_map else None
                            raw_lbl = getattr(ai_obj, "description", "") if ai_obj else (it.topic_label or f"AI {target_ai}")
                            label = f"AI {target_ai}: {raw_lbl}" if (len(expanded_ais) > 1 or target_ai != it.ai_number) else (it.topic_label or f"AI {target_ai}")
                            is_tr = target_ai.lower() in self.tracked_tokens
                            act_tr = menu.addAction(f"{'⭐ Untrack' if is_tr else '⭐ Track'}: {label[:44]}")
                            act_tr.triggered.connect(lambda _, tok=target_ai, lbl=raw_lbl: self._toggle_track_ai(tok, lbl))
                    else:
                        token = it.topic_label
                        if not token:
                            continue
                        is_tr = token.lower() in self.tracked_tokens
                        act_tr = menu.addAction(f"{'⭐ Untrack' if is_tr else '⭐ Track'}: {token[:40]}")
                        act_tr.triggered.connect(lambda _, tok=token, lbl=token: self._toggle_track_ai(tok, lbl))

            if self.delegation_plan and self.delegation_plan.members:
                menu.addSeparator()

                # Direct Topic Owner Assignment Submenu for each AI covered (expanding .x wildcards)
                for s in slots_here:
                    for it in s.items:
                        if not it.ai_number:
                            continue
                        expanded_ais = expand_ai_number(it.ai_number, self.agenda_map)
                        for target_ai in expanded_ais[:12]:
                            current_to = self.delegation_plan.topic_owners.get(target_ai)
                            curr_owner_str = current_to.owner_name if current_to else ""
                            curr_disp = " (Not Followed)" if (current_to and current_to.is_not_followed()) else (f" ({curr_owner_str})" if curr_owner_str else " (Unassigned)")
                            ai_menu = menu.addMenu(f"👤 Assign Topic Owner (AI {target_ai}){curr_disp}...")

                            act_all = ai_menu.addAction("👥 All Team")
                            act_all.triggered.connect(lambda _, ai=target_ai: self._set_topic_owner_direct(ai, "ALL"))

                            ai_menu.addSeparator()
                            for m_name in self.delegation_plan.members.keys():
                                is_curr = (curr_owner_str == m_name)
                                prefix = "✓ " if is_curr else "● "
                                act_m = ai_menu.addAction(f"{prefix}{m_name}")
                                act_m.triggered.connect(lambda _, ai=target_ai, n=m_name: self._set_topic_owner_direct(ai, n))

                            ai_menu.addSeparator()
                            act_multi = ai_menu.addAction("⚡ Custom Multi-Delegate...")
                            act_multi.triggered.connect(lambda _, ai=target_ai, cur=curr_owner_str: self._open_multi_delegate_dialog(ai, cur))

                            ai_menu.addSeparator()
                            act_not_followed = ai_menu.addAction("⛔ Mark as Not Followed")
                            act_not_followed.triggered.connect(lambda _, ai=target_ai: self._set_topic_owner_direct(ai, "NOT_FOLLOWED"))

                            if curr_owner_str:
                                act_clear = ai_menu.addAction("✕ Clear (Unassign)")
                                act_clear.triggered.connect(lambda _, ai=target_ai: self._set_topic_owner_direct(ai, ""))

                menu.addSeparator()
                for s in slots_here:
                    del_name, is_cov = self._get_slot_effective_delegate(s)
                    disp_name = "All Team" if del_name == "ALL" else (del_name or "")
                    app = self.delegation_plan.appointments.get(s.slot_key)
                    orig_owner = app.original_owner if (app and app.original_owner) else (del_name or "")
                    str_lbl = f"Str {s.stream_number}: " if len(slots_here) > 1 else ""

                    if is_cov:
                        act_disp = menu.addAction(f"⚡ Change Coverage for {str_lbl}this Session (Covered by {disp_name})...")
                        act_disp.triggered.connect(lambda _, slot=s, orig=orig_owner: self._on_dispatch_slot(slot, orig))
                        act_rev = menu.addAction(f"↺ Revert Coverage for {str_lbl}this Session to {orig_owner or 'Topic Owner'}")
                        act_rev.triggered.connect(lambda _, slot=s: self._on_revert_dispatch(slot))
                    else:
                        curr_str = f"({disp_name})" if disp_name else "(Unassigned)"
                        act_disp = menu.addAction(f"⚡ Appoint Coverage for {str_lbl}this Session {curr_str}...")
                        act_disp.triggered.connect(lambda _, slot=s, orig=orig_owner: self._on_dispatch_slot(slot, orig))

            menu.addSeparator()
            act_copy = menu.addAction("📋 Copy Slot Details")
            act_copy.triggered.connect(lambda: self._copy_slot_text(slots_here[0].raw_text))

        menu.exec_(self.grid_table.viewport().mapToGlobal(pos))

    def _copy_slot_text(self, text: str):
        QApplication.clipboard().setText(text)

    def _set_topic_owner_direct(self, ai_number: str, owner_name: str):
        """Directly updates topic owner from Overview Grid context menu and refreshes views."""
        self.db.set_topic_owner(self.meeting_number, ai_number, owner_name)
        self.delegation_plan = self.db.get_delegation_plan(self.meeting_number)
        self._render_current_views()

    def _open_multi_delegate_dialog(self, ai_number: str, current_owner: str):
        """Opens MultiDelegateSelectDialog for compound multi-delegate ownership from grid."""
        ai_desc = ""
        if self.agenda_map and ai_number in self.agenda_map:
            ai_desc = getattr(self.agenda_map[ai_number], "description", "")
        title = f"AI {ai_number}: {ai_desc}" if ai_desc else f"AI {ai_number}"
        members_list = list(self.delegation_plan.members.values()) if self.delegation_plan else []
        dlg = MultiDelegateSelectDialog(title, members_list, current_owner, parent=self)
        if dlg.exec_() == QDialog.Accepted:
            new_owner = dlg.get_owner_string()
            self._set_topic_owner_direct(ai_number, new_owner)
