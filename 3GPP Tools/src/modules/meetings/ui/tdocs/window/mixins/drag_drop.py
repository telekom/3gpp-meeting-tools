# --- File: src/modules/meetings/ui/tdocs/window/mixins/drag_drop.py ---
"""
Drag & drop of files / Outlook items onto the TDocs window.

Mixin of TDocsWindow: relies on attributes set in TDocsWindow.__init__ / _setup_*.
"""
from pathlib import Path


class DragDropMixin:
    """Drag & drop of files / Outlook items onto the TDocs window."""

    @staticmethod
    def _has_outlook_drag_payload(mime_data) -> bool:
        if not mime_data:
            return False
        for fmt in mime_data.formats():
            if any(k in fmt for k in ("FileGroupDescriptor", "RenPrivateMessages", "RenPrivateItem")):
                return True
        return False

    def dragEnterEvent(self, event):
        mime = event.mimeData()
        accepted = False
        if mime and mime.hasUrls():
            for url in mime.urls():
                ext = Path(url.toLocalFile()).suffix.lower()
                if ext in ['.docx', '.doc', '.zip', '.htm', '.html', '.msg']:
                    accepted = True
                    break
        if not accepted and self._has_outlook_drag_payload(mime):
            accepted = True
        if accepted:
            self.drop_overlay.setGeometry(self.rect())
            self.drop_overlay.show()
            self.drop_overlay.raise_()
            event.acceptProposedAction()
            return
        event.ignore()

    def dragMoveEvent(self, event):
        mime = event.mimeData()
        if mime and mime.hasUrls():
            for url in mime.urls():
                ext = Path(url.toLocalFile()).suffix.lower()
                if ext in ['.docx', '.doc', '.zip', '.htm', '.html', '.msg']:
                    event.acceptProposedAction()
                    return
        if self._has_outlook_drag_payload(mime):
            event.acceptProposedAction()
            return
        event.ignore()

    def dragLeaveEvent(self, event):
        if hasattr(self, 'drop_overlay'):
            self.drop_overlay.hide()
        event.accept()

    def dropEvent(self, event):
        if hasattr(self, 'drop_overlay'):
            self.drop_overlay.hide()

        mime = event.mimeData()
        valid_files = []
        msg_files = []
        if mime and mime.hasUrls():
            for url in mime.urls():
                file_path = Path(url.toLocalFile())
                ext = file_path.suffix.lower()
                if ext == '.msg':
                    msg_files.append(file_path)
                elif ext in ['.docx', '.doc', '.zip', '.htm', '.html']:
                    valid_files.append(file_path)

        outlook_dropped = False
        if not msg_files and not valid_files and self._has_outlook_drag_payload(mime):
            event.acceptProposedAction()
            outlook_dropped = True
            self._process_dropped_outlook_items(mime)

        if msg_files:
            event.acceptProposedAction()
            self._process_dropped_msg_files(msg_files)

        if valid_files:
            event.acceptProposedAction()
            # Smart routing: check if dropped file is a session plan
            session_plan_files = []
            agenda_files = []
            for vf in valid_files:
                name_lower = vf.name.lower()
                if "session" in name_lower or "time_allocation" in name_lower or "time allocation" in name_lower:
                    session_plan_files.append(vf)
                else:
                    agenda_files.append(vf)

            if session_plan_files:
                for spf in session_plan_files:
                    from modules.meetings.core.session_plan.session_plan_parser import SessionPlanParser
                    from modules.meetings.core.session_plan.session_plan_db import SessionPlanDatabase
                    staged = SessionPlanParser.stage_session_plan_file(spf, self.meeting_dir)
                    db = SessionPlanDatabase(self.meeting_dir / "Agenda" / "session_plans.db")
                    if staged.suffix.lower() == ".docx":
                        snap = SessionPlanParser.parse_docx(staged, meeting_number=str(self.mtg_info.get("meeting_number", "")))
                    else:
                        snap = SessionPlanParser.parse_text_or_markdown(
                            staged.read_text(encoding="utf-8", errors="replace"),
                            meeting_number=str(self.mtg_info.get("meeting_number", "")),
                            cached_file_path=str(staged),
                        )
                    db.save_snapshot(snap)
                self._open_session_plan_window()
                if hasattr(self, "_session_plan_window") and self._session_plan_window:
                    self._session_plan_window._load_data()

            if agenda_files:
                self._process_imported_agenda_files(agenda_files)
        elif not msg_files and not outlook_dropped:
            event.ignore()
