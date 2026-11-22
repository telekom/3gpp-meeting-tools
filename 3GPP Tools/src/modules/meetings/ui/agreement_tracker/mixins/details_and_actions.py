"""Detail card formatting, TDoc download/open dispatch, and table context menu mixin."""

import html
import logging
import os
import re
from pathlib import Path
from typing import Optional

from PyQt5.QtCore import QModelIndex, QPoint, Qt
from PyQt5.QtWidgets import QApplication, QMenu, QMessageBox

from core.ui.theme import C
from core.utils.folder_opener import open_local_file
from core.utils.paths import to_extended_path
from core.utils.portal_utils import open_portal_tdoc
from modules.meetings.core.tdocs.tdocs_threads import TDocActionThread
from modules.meetings.core.url_router import URLRouter

logger = logging.getLogger(__name__)


class _TrackerActionsMixin:
    """Handles row selection details, TDoc archive resolution/opening, and context menus."""

    def _get_selected_row(self) -> Optional[dict]:
        idx = self.table_view.currentIndex()
        if not idx.isValid():
            return None
        return self.table_model.data(idx, Qt.UserRole)

    def _on_current_row_changed(self, current: QModelIndex, _previous: QModelIndex):
        if not current.isValid():
            self.btn_open_tdoc.setEnabled(False)
            self.btn_open_notes.setEnabled(False)
            self.btn_open_mtg.setEnabled(False)
            self.btn_copy_ag.setEnabled(False)
            return

        row = self.table_model.data(current, Qt.UserRole)
        if not row:
            return

        tdoc_id = str(row.get("TDoc", "")).strip()
        chair_file = str(row.get("Chair Notes File", "")).strip()

        self.btn_open_tdoc.setEnabled(bool(tdoc_id))
        self.btn_open_tdoc.setText(f"📄 Open {tdoc_id}" if tdoc_id else "📄 Open Related TDoc")
        self.btn_open_notes.setEnabled(bool(chair_file))
        self.btn_open_mtg.setEnabled(True)
        self.btn_copy_ag.setEnabled(True)

        kind = str(row.get("Kind", "Agreement"))
        low_k = kind.lower()
        badge_bg = C.SUCCESS_SOFT_ALT
        badge_fg = C.SUCCESS_DARK
        if "ffs" in low_k or "open issue" in low_k:
            badge_bg = C.INFO_SOFT_STRONG
            badge_fg = C.INFO_DARK
        elif "working assumption" in low_k:
            badge_bg = C.WARNING_SOFT_ALT
            badge_fg = C.WARNING_DARK
        elif "agreed tdoc" in low_k:
            badge_bg = C.INDIGO_SOFT
            badge_fg = C.WG_SA_TEXT

        mtg = html.escape(str(row.get("Meeting", "")))
        dt = html.escape(str(row.get("start_date", "")))
        ai = html.escape(str(row.get("AI", "") or "N/A"))
        ai_title = html.escape(str(row.get("AI Title", "") or ""))
        wi = html.escape(str(row.get("Work Item", "") or "Unspecified"))
        rel = html.escape(str(row.get("Release", "") or ""))
        tdoc_title = html.escape(str(row.get("TDoc Title", "") or ""))
        src = html.escape(str(row.get("Source", "") or ""))
        cfile = html.escape(chair_file)

        raw_text = str(row.get("Agreement Text", ""))
        formatted_lines = []
        for line in raw_text.split("\n"):
            esc_line = html.escape(line)
            # Highlight TDoc IDs inside agreement text
            esc_line = re.sub(
                r'\b([A-Z][A-Z0-9]{1,3}-\d{4,8}(?:[rR]\d{1,2}[a-zA-Z]?)?)\b',
                fr"<b><span style='color:{C.ACCENT_PRIMARY};'>\1</span></b>",
                esc_line,
            )
            formatted_lines.append(esc_line)
        body_html = "<br>".join(formatted_lines)

        tdoc_meta_html = ""
        if tdoc_id:
            tdoc_meta_html = (
                f"<div style='margin-top:6px; padding:5px 8px; background:{C.INFO_SOFT}; border-left:3px solid {C.INFO_ACCENT};'>"
                f"<b>Related TDoc:</b> <span style='color:{C.ACCENT_PRIMARY}; font-weight:bold;'>{html.escape(tdoc_id)}</span>"
                f"{f' — {tdoc_title}' if tdoc_title else ''}"
                f"{f' <i>({src})</i>' if src else ''}"
                f"</div>"
            )

        html_doc = f"""
        <div style='font-family: Segoe UI, sans-serif;'>
            <div style='margin-bottom:6px;'>
                <span style='background:{badge_bg}; color:{badge_fg}; padding:2px 8px; border-radius:4px; font-weight:bold;'>
                    {html.escape(kind)}
                </span>
                &nbsp;&nbsp;<b>{mtg}</b> ({dt})
                &nbsp;&nbsp;|&nbsp;&nbsp;<b>Work Item:</b> <span style='color:{C.ACCENT_PRIMARY};'>{wi}</span> {f'({rel})' if rel else ''}
                &nbsp;&nbsp;|&nbsp;&nbsp;<b>AI {ai}:</b> {ai_title}
                &nbsp;&nbsp;|&nbsp;&nbsp;<span style='color:{C.TEXT_MUTED};'>Source: {cfile}</span>
            </div>
            <div style='padding:8px 10px; background:{C.BG_CARD}; border:1px solid {C.BORDER_SUBTLE}; border-radius:4px; line-height:1.45;'>
                {body_html}
            </div>
            {tdoc_meta_html}
        </div>
        """
        self.detail_browser.setHtml(html_doc)

    def _open_selected_tdoc(self):
        row = self._get_selected_row()
        if row and row.get("TDoc"):
            self._open_tdoc(row)

    def _open_selected_chair_notes_file(self):
        row = self._get_selected_row()
        if not row:
            return
        folder_name = row.get("folder_name") or row.get("meeting_number", "")
        cfile = row.get("Chair Notes File", "")
        if not folder_name or not cfile:
            return
        mtg_dir = Path(self.settings.cache_dir) / folder_name
        candidates = [
            mtg_dir / "Agenda" / "ChairNotes" / cfile,
            mtg_dir / "Agenda" / cfile,
        ]
        for cand in candidates:
            if os.path.exists(to_extended_path(cand)):
                open_local_file(str(cand))
                return
        # If it was a .xlsx fallback, look for any .xlsx in Agenda/
        agenda_dir = mtg_dir / "Agenda"
        if cfile.endswith(".xlsx") and os.path.exists(to_extended_path(agenda_dir)):
            for f in agenda_dir.iterdir():
                if f.suffix.lower() == ".xlsx":
                    open_local_file(str(f))
                    return
        QMessageBox.information(
            self, "File Not Found Locally", f"Could not find '{cfile}' in {mtg_dir / 'Agenda'}."
        )

    def _open_selected_meeting(self):
        row = self._get_selected_row()
        if row:
            self._request_open_meeting(row)

    def _copy_selected_agreement(self):
        row = self._get_selected_row()
        if not row:
            return
        header = (
            f"[{row.get('Meeting', '')} | AI {row.get('AI', '')} ({row.get('AI Title', '')}) | "
            f"WI: {row.get('Work Item', '')} | {row.get('Kind', '')}]"
        )
        tdoc = row.get("TDoc", "")
        if tdoc:
            header += f" (TDoc: {tdoc})"
        QApplication.clipboard().setText(f"{header}\n{row.get('Agreement Text', '')}")

    def _request_open_meeting(self, row_data: dict):
        wg = str(row_data.get("WG", "")).strip()
        mtg_num = str(row_data.get("meeting_number", "")).strip()
        if not mtg_num:
            mtg_raw = str(row_data.get("Meeting", "")).strip()
            mtg_num = mtg_raw.split("#")[-1].strip() if "#" in mtg_raw else mtg_raw
        if wg and mtg_num:
            self.open_meeting_requested.emit(wg, mtg_num)

    def _open_tdoc(self, row_data: dict):
        tdoc_id = str(row_data.get("TDoc", "")).strip()
        if not tdoc_id:
            return

        match = re.search(r'^(.*?)-?(?:r|rev)(\d{1,2}[a-zA-Z]?)$', tdoc_id, re.IGNORECASE)
        base_tdoc = match.group(1).upper() if match else tdoc_id.upper()
        target_filename = tdoc_id.upper()

        if target_filename in self.active_tdoc_threads:
            th = self.active_tdoc_threads[target_filename]
            if th and th.isRunning():
                return

        wg_name = row_data.get("WG") or row_data.get("wg_name", "")
        folder_name = row_data.get("folder_name") or row_data.get("meeting_number", "")
        raw_url = row_data.get("url_key", "")
        docs_url = row_data.get("docs_folder_url", "")

        main_ftp = (
            raw_url
            if raw_url.startswith("http")
            else (f"https://www.3gpp.org/ftp/{raw_url.lstrip('/')}" if raw_url else "")
        )
        is_active_sync = self.meetings_db.is_active_sync_meeting(
            wg_name,
            row_data.get("start_date", ""),
            row_data.get("end_date", ""),
            int(row_data.get("is_electronic", 0) or 0),
        )
        candidate_urls = URLRouter.build_priority_url_list(
            wg_name, folder_name, main_ftp, is_active_sync, target_filename=target_filename
        )
        if docs_url:
            clean_docs = (
                docs_url
                if docs_url.startswith("http")
                else f"https://www.3gpp.org/ftp/{docs_url.lstrip('/')}"
            ).rstrip("/")
            if clean_docs not in candidate_urls:
                candidate_urls.append(clean_docs)

        if not candidate_urls:
            QMessageBox.warning(
                self, "Missing URL", f"No candidate download routes found for {tdoc_id}."
            )
            return

        meeting_dir = Path(self.settings.cache_dir) / folder_name
        meeting_dir.mkdir(parents=True, exist_ok=True)

        thread = TDocActionThread(
            base_tdoc, target_filename, candidate_urls, meeting_dir, open_file=True
        )
        self.active_tdoc_threads[target_filename] = thread
        thread.finished_action.connect(
            lambda t, s, m, th=thread: self._on_tdoc_action_finished(t, s, m, th)
        )
        thread.start()

    def _on_tdoc_action_finished(
        self, tdoc: str, success: bool, msg: str, thread: TDocActionThread
    ):
        if thread.target_filename in self.active_tdoc_threads:
            del self.active_tdoc_threads[thread.target_filename]
        if not success:
            QMessageBox.warning(self, f"Action Failed: {tdoc}", msg)

    def _on_table_double_clicked(self, index: QModelIndex):
        row = self.table_model.data(index, Qt.UserRole)
        if not row:
            return
        col_name = self.table_model.COLUMNS[index.column()]
        if col_name in ("WG", "Meeting"):
            self._request_open_meeting(row)
        elif col_name == "Chair Notes File":
            self._open_selected_chair_notes_file()
        elif row.get("TDoc"):
            self._open_tdoc(row)

    def _on_table_context_menu(self, pos: QPoint):
        index = self.table_view.indexAt(pos)
        if not index.isValid():
            return
        row = self.table_model.data(index, Qt.UserRole)
        if not row:
            return

        menu = QMenu(self)
        mtg = row.get("Meeting", "")
        tdoc = row.get("TDoc", "")

        act_copy_ag = menu.addAction("📋 Copy Agreement with Metadata")
        act_copy_ag.triggered.connect(self._copy_selected_agreement)

        act_open_notes = menu.addAction(f"📝 Open Chair Notes File ({row.get('Chair Notes File', '')})")
        act_open_notes.triggered.connect(self._open_selected_chair_notes_file)

        act_open_mtg = menu.addAction(f"🗓️ Open Meeting Window ({mtg})")
        act_open_mtg.triggered.connect(lambda: self._request_open_meeting(row))

        if tdoc:
            menu.addSeparator()
            act_open_td = menu.addAction(f"📄 Open Related TDoc ({tdoc})")
            act_open_td.triggered.connect(lambda: self._open_tdoc(row))

            act_portal = menu.addAction(f"🌐 View {tdoc} on 3GU Portal")
            act_portal.triggered.connect(lambda: open_portal_tdoc(tdoc))

        menu.exec_(self.table_view.viewport().mapToGlobal(pos))
