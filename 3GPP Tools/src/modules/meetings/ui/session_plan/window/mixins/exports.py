# --- File: src/modules/meetings/ui/session_plan/window/mixins/exports.py ---
"""
Exports: mobile HTML, ICS (all / per delegate), e-mail summary, plans folder.

Mixin of SessionPlanWindow: relies on attributes set in SessionPlanWindow.__init__.
"""
import re
import webbrowser
from pathlib import Path
from typing import Optional
from PyQt5.QtWidgets import QApplication, QMessageBox
from core.utils.folder_opener import open_folder
from modules.meetings.core.session_plan.session_plan_exporter import SessionPlanExporter


class ExportsMixin:
    """Exports: mobile HTML, ICS (all / per delegate), e-mail summary, plans folder."""

    def _export_mobile_html(self) -> Optional[Path]:
        if not self.active_snapshot:
            return None
        out_dir = self.meeting_dir / "Export"
        out_path = out_dir / f"SA2_{self.meeting_number}_Session_Plan_Mobile.html"
        SessionPlanExporter.export_mobile_html(
            snapshot=self.active_snapshot,
            output_path=out_path,
            meeting_tz_offset_min=self.tz_offset_min,
            tz_name=self.tz_name,
            meeting_dates_str=self.start_date_str,
            annotations=self.annotations,
            tracked_tokens=self.tracked_tokens,
            delegation_plan=self.delegation_plan,
            agenda_map=self.agenda_map,
            room_locations=self.room_locations,
            slot_locations=self.slot_locations,
        )
        QMessageBox.information(
            self,
            "Interactive HTML Exported",
            f"Saved interactive Session Plan (Overview Grid + 🔴 Now + Mobile Day Cards) to:\n{out_path}\n\nFeatures live stream/room lookup, delegate toggling, and custom room location notes.",
        )
        return out_path

    def _open_mobile_html_in_browser(self):
        if not self.active_snapshot:
            return
        out_path = self._export_mobile_html()
        if out_path and out_path.exists():
            webbrowser.open(str(out_path))

    def _export_ics(self):
        if not self.active_snapshot:
            return
        out_dir = self.meeting_dir / "Export"
        out_path = out_dir / f"SA2_{self.meeting_number}_Schedule.ics"
        SessionPlanExporter.export_ics(
            snapshot=self.active_snapshot,
            output_path=out_path,
            meeting_start_date_str=self.start_date_str,
            meeting_tz_offset_min=self.tz_offset_min,
            only_tracked=False,
            tracked_tokens=self.tracked_tokens,
            annotations=self.annotations,
            delegation_plan=self.delegation_plan,
            agenda_map=self.agenda_map,
            room_locations=self.room_locations,
            slot_locations=self.slot_locations,
        )
        QMessageBox.information(
            self,
            "Calendar Exported",
            f"Generated .ics calendar file:\n{out_path}\n\nStrictly anchored in UTC so your Outlook and smartphone trigger reminders at the exact venue time.",
        )

    def _export_delegate_ics(self):
        """Exports an .ics calendar file scoped to a specific delegation member."""
        if not self.active_snapshot:
            return
        if not self.delegation_plan or not self.delegation_plan.members:
            QMessageBox.information(self, "No Roster", "Please add delegation members before exporting a delegate calendar.")
            return

        default_name = ""
        if hasattr(self, "tracked_delegate_filter") and self.tracked_delegate_filter.currentIndex() > 0:
            default_name = self.tracked_delegate_filter.currentData()
        if not default_name:
            for name, m in self.delegation_plan.members.items():
                if m.is_self:
                    default_name = name
                    break
        if not default_name:
            default_name = list(self.delegation_plan.members.keys())[0]

        from PyQt5.QtWidgets import QInputDialog
        names = list(self.delegation_plan.members.keys())
        default_idx = names.index(default_name) if default_name in names else 0
        target, ok = QInputDialog.getItem(
            self,
            "Export Delegate Calendar",
            "Select delegate to export .ics calendar for:",
            names,
            default_idx,
            False,
        )
        if not ok or not target:
            return

        out_dir = self.meeting_dir / "Export"
        safe_name = re.sub(r"[^\w\-]", "_", target)
        out_path = out_dir / f"SA2_{self.meeting_number}_{safe_name}.ics"

        SessionPlanExporter.export_ics(
            snapshot=self.active_snapshot,
            output_path=out_path,
            meeting_start_date_str=self.start_date_str,
            meeting_tz_offset_min=self.tz_offset_min,
            only_tracked=False,
            tracked_tokens=self.tracked_tokens,
            annotations=self.annotations,
            delegation_plan=self.delegation_plan,
            target_delegate=target,
            agenda_map=self.agenda_map,
            room_locations=self.room_locations,
            slot_locations=self.slot_locations,
        )
        QMessageBox.information(
            self,
            "Delegate Calendar Exported",
            f"Generated personalized .ics calendar file for {target}:\n{out_path}\n\nStrictly anchored in UTC with assigned session topics.",
        )

    def _copy_email_summary(self):
        if not self.active_snapshot:
            return
        subj, body = SessionPlanExporter.generate_email_summary(
            snapshot=self.active_snapshot,
            meeting_tz_offset_min=self.tz_offset_min,
            tz_name=self.tz_name,
            tracked_tokens=self.tracked_tokens,
            annotations=self.annotations,
        )
        clipboard = QApplication.clipboard()
        clipboard.setText(body)
        QMessageBox.information(self, "Email Summary Copied", f"Subject: {subj}\n\nRich HTML email summary copied to your clipboard!\nYou can paste it directly into Microsoft Outlook or your mail client.")

    def _open_plans_folder(self):
        plans_dir = self.meeting_dir / "Agenda" / "SessionPlans"
        plans_dir.mkdir(parents=True, exist_ok=True)
        open_folder(plans_dir)
