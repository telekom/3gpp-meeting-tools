"""TDocEmailsDialog shell combining UI setup, rendering, and action mixins."""

import logging
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional

from PyQt5.QtCore import QSettings, Qt, QTimer, pyqtSignal
from PyQt5.QtWidgets import QDialog

from modules.emails.core.ai_summary import EmailSummaryStore
from modules.emails.core.general_email_db import GeneralEmailDatabase
from modules.emails.ui.tdoc_emails.mixins import (
    TDocEmailsActionsMixin,
    TDocEmailsRenderingMixin,
    TDocEmailsUiSetupMixin,
)

logger = logging.getLogger("modules.emails.ui.tdoc_emails_dialog")


class TDocEmailsDialog(
    TDocEmailsUiSetupMixin,
    TDocEmailsRenderingMixin,
    TDocEmailsActionsMixin,
    QDialog,
):
    """Unified continuous Markdown timeline viewer for emails linked to a 3GPP TDoc family."""

    data_changed = pyqtSignal()
    tdoc_selected = pyqtSignal(str)

    def __init__(
        self,
        target_tdoc: str,
        family_tdocs: list,
        db_path: Path,
        wg: str = "SA2",
        save_callback=None,
        parent=None,
        pin_thread_callback: Optional[Callable[[dict], None]] = None,
        *,
        tdoc_info: Optional[Dict[str, Any]] = None,
        open_tdoc_callback: Optional[Callable[[str], None]] = None,
        user_data_callback: Optional[Callable[[str, str, str], None]] = None,
    ):
        super().__init__(None)  # Pass None so Win32 does not set parent as GW_OWNER
        self._parent_ref = parent
        if parent:
            try:
                parent.destroyed.connect(self.close)
            except Exception:
                pass
        self.target_tdoc = str(target_tdoc or "").strip().upper()
        self.family_tdocs = [str(t).strip().upper() for t in (family_tdocs or []) if str(t).strip()]
        if self.target_tdoc and self.target_tdoc not in self.family_tdocs:
            self.family_tdocs.insert(0, self.target_tdoc)

        self.wg = str(wg or "SA2").strip().upper()
        self.save_callback = save_callback
        self._pin_thread_callback = pin_thread_callback
        self._tdoc_info: Dict[str, Any] = dict(tdoc_info or {})
        self._open_tdoc_callback = open_tdoc_callback
        self._user_data_callback = user_data_callback
        self._updating_status_bar = False
        self._updating_note_bar = False
        self._note_dirty = False
        self._note_save_timer = QTimer(self)
        self._note_save_timer.setSingleShot(True)
        self._note_save_timer.setInterval(450)
        self._note_save_timer.timeout.connect(self._flush_pending_note_save)
        self.db = GeneralEmailDatabase(db_path)
        db_p = getattr(self.db, "db_path", None)
        self._summary_store: Optional[EmailSummaryStore] = (
            EmailSummaryStore(Path(db_p).parent / "email_summaries.db") if db_p else None
        )
        self._summary_dialog = None
        self.emails: List[dict] = []
        self._filtered_emails: List[dict] = []
        self._unread_blocks: List[Any] = []
        self._unread_cursor: int = -1
        self._initial_scroll_done: bool = False
        self._loading = False
        self._settings = QSettings("3GPP_Tools", "TDocEmailsDialog")

        self.setWindowFlags(Qt.Window)
        self.setWindowModality(Qt.NonModal)
        self.setWindowTitle(f"📧 Related Emails: {self.target_tdoc}")
        self.setMinimumSize(620, 460)
        self.resize(1100, 760)
        saved_geom = self._settings.value("geometry")
        if saved_geom:
            try:
                self.restoreGeometry(saved_geom)
            except Exception:
                pass

        self._init_ui()
        self._load_emails()

    def showEvent(self, event):
        super().showEvent(event)
        if not self._initial_scroll_done:
            QTimer.singleShot(50, self._scroll_to_initial_unread)

    def closeEvent(self, event):
        try:
            self._flush_pending_note_save()
            self._settings.setValue("geometry", self.saveGeometry())
            if hasattr(self, "browser"):
                self._settings.setValue("zoom_steps", int(self.browser.property("_reader_zoom_steps") or 0))
        except Exception:
            pass
        if hasattr(self, "summary_layout_ctrl"):
            try:
                self.summary_layout_ctrl.shutdown()
            except Exception:
                pass
        if hasattr(self, "summary_card"):
            try:
                self.summary_card.shutdown_worker()
            except Exception:
                pass
        if getattr(self, "_summary_dialog", None):
            try:
                self._summary_dialog.close()
            except Exception:
                pass
        super().closeEvent(event)
